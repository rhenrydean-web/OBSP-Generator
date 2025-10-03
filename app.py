# app.py
import io, json
import streamlit as st
from pptx import Presentation

# ---------- PAGE CONFIG (must be first st.* call) ----------
st.set_page_config(page_title="Initiatives Slide Filler", page_icon="🧩", layout="centered")

# ---------- OPENAI WRAPPER ----------
def call_llm(system, user, model):
    if "OPENAI_API_KEY" not in st.secrets or not st.secrets["OPENAI_API_KEY"]:
        raise RuntimeError("Missing OPENAI_API_KEY secret. Add it in Streamlit → Manage app → Settings → Secrets.")
    from openai import OpenAI
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role":"system","content":system},{"role":"user","content":user}],
        temperature=0.2,
    )
    return resp.choices[0].message.content

def extract_json(raw: str):
    start = raw.find("{"); end = raw.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("No JSON object found in LLM response.")
    return json.loads(raw[start:end+1])

# ---------- CONTEXT INGEST (PDF/DOCX/TXT) ----------
from pypdf import PdfReader
from docx import Document as DocxDocument
from rapidfuzz import fuzz

def extract_text_from_file(f):
    name = (f.name or "").lower()
    if name.endswith(".pdf"):
        reader = PdfReader(f)
        pages = min(len(reader.pages), 25)
        return "\n\n".join((reader.pages[i].extract_text() or "") for i in range(pages))
    elif name.endswith(".docx"):
        d = DocxDocument(f)
        return "\n".join(p.text for p in d.paragraphs)
    else:  # .txt
        return f.read().decode(errors="ignore")

def chunk_text(txt, chunk_size=800, overlap=120):
    txt = " ".join(txt.split())
    chunks, i = [], 0
    while i < len(txt):
        chunks.append(txt[i:i+chunk_size])
        i += chunk_size - overlap
    return chunks

def pick_top_snippets(texts, query, k=8):
    chunks = []
    for t in texts:
        chunks.extend(chunk_text(t))
    if not chunks:
        return []
    scored = [(c, fuzz.partial_ratio(query, c)) for c in chunks]
    scored.sort(key=lambda x: x[1], reverse=True)
    return [c for c, s in scored[:k] if s > 40]

# ---------- SMALL UTIL ----------
def _ellipsize(text, max_chars=120):
    t = (text or "").strip()
    return (t[:max_chars-1] + "…") if len(t) > max_chars else t

# ---------- TABLE FIND & FILL HELPERS ----------
def _find_table(prs, table_name="INIT_TABLE"):
    """Find table by name; if not found, try header match."""
    # 1) By exact shape name (including inside groups)
    for slide in prs.slides:
        for sh in slide.shapes:
            if getattr(sh, "name", "") == table_name and hasattr(sh, "table"):
                return sh
            if hasattr(sh, "shapes"):
                for s in sh.shapes:
                    if getattr(s, "name", "") == table_name and hasattr(s, "table"):
                        return s
    # 2) Fallback by header row
    expected_headers = ["Initiative","Step","Success criteria","Owner","Target","Status"]
    for slide in prs.slides:
        for sh in slide.shapes:
            if hasattr(sh, "table"):
                tbl = sh.table
                if len(tbl.columns) >= 6 and len(tbl.rows) >= 1:
                    headers = [tbl.cell(0, c).text.strip() for c in range(6)]
                    if headers == expected_headers:
                        return sh
    return None

def _ensure_min_rows(tbl, min_rows=10):
    # Add rows until there are min_rows total (header + data)
    while len(tbl.rows) < min_rows:
        tbl.add_row()

def fill_initiatives_table(prs, initiatives, table_name="INIT_TABLE"):
    """
    Fill a 6-col table with 9 rows (+ header) for 3 initiatives × 3 steps.
    Columns: [Initiative, Step, Success criteria, Owner, Target, Status]
    Tolerates pre-merged cells and will add rows if the grid is short.
    """
    shp = _find_table(prs, table_name)
    if shp is None or not hasattr(shp, "table"):
        return False
    tbl = shp.table

    # Ensure 6 columns and at least 10 rows in the GRID
    if len(tbl.columns) < 6:
        return False
    _ensure_min_rows(tbl, 10)

    if len(initiatives) != 3:
        return False

    # Clear data rows (row 1..end)
    for r in range(1, len(tbl.rows)):
        for c in range(6):
            try:
                tbl.cell(r, c).text = ""
            except:
                pass

    # Fill rows: 3 initiatives × 3 steps
    row_idx = 1
    for init in initiatives:
        name = _ellipsize(init.get("name",""), 70)
        crit = _ellipsize(init.get("success_criteria",""), 220)
        steps = (init.get("steps") or [])[:3]
        if len(steps) < 3:
            steps += [{"name":"TBD","owner":"TBD","target":"TBD","status":"not started"}] * (3 - len(steps))

        start = row_idx
        for step in steps:
            # Step, Owner, Target, Status
            try: tbl.cell(row_idx, 1).text = _ellipsize(step.get("name",""), 120)
            except: pass
            try: tbl.cell(row_idx, 3).text = _ellipsize(step.get("owner",""), 60)
            except: pass
            try: tbl.cell(row_idx, 4).text = _ellipsize(step.get("target",""), 30)
            except: pass
            try: tbl.cell(row_idx, 5).text = _ellipsize(step.get("status","not started"), 30)
            except: pass
            row_idx += 1
        end = row_idx - 1

        # Initiative (merge rows start..end in col 0 if possible)
        try:
            tbl.cell(start, 0).merge(tbl.cell(end, 0))
        except:
            pass
        try:
            tbl.cell(start, 0).text = name
        except:
            pass

        # Success criteria (merge rows start..end in col 2 if possible)
        try:
            tbl.cell(start, 2).merge(tbl.cell(end, 2))
        except:
            pass
        try:
            tbl.cell(start, 2).text = crit
        except:
            pass

    return True

# ---------- TABLE POST-FORMATTING (autoshrink, widths, padding, resize-to-slide) ----------
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import MSO_AUTO_SIZE, MSO_ANCHOR

def _set_col_widths(tbl, widths_in_inches):
    for i, w in enumerate(widths_in_inches):
        if i < len(tbl.columns):
            tbl.columns[i].width = Inches(w)

def _apply_text_autofit(tf, start_pt=13):
    # Enable wrap + try native autosize
    tf.word_wrap = True
    try:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass
    # Tight spacing + reasonable starting size
    for p in tf.paragraphs:
        p.line_spacing = 1.0
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        for r in p.runs:
            r.font.size = Pt(start_pt)

def _fallback_shrink(tf, hard_min_pt=9):
    # Deterministic shrink if theme blocks autosize
    text_len = len(tf.text or "")
    if text_len <= 80:
        target = 13
    elif text_len <= 140:
        target = 12
    elif text_len <= 200:
        target = 11
    elif text_len <= 260:
        target = 10
    else:
        target = hard_min_pt
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(target)

def _resize_table_to_slide(prs, table_shape, horiz_margin_in=0.5):
    """
    Snap the table to slide width with left/right margins.
    If table is inside a group, attempt to resize the group.
    """
    target_width = prs.slide_width - Inches(horiz_margin_in * 2)
    try:
        table_shape.left = Inches(horiz_margin_in)
        table_shape.width = target_width
    except Exception:
        # Try parent group if direct resize is blocked
        try:
            parent = table_shape._parent
            if hasattr(parent, "left") and hasattr(parent, "width"):
                parent.left = Inches(horiz_margin_in)
                parent.width = target_width
        except Exception:
            pass

def pretty_format_init_table(prs, table_shape, max_pt=13, widths=None):
    """
    Normalize widths, padding, vertical alignment, autoshrink text,
    and resize the table to fit the slide.
    """
    tbl = table_shape.table

    # 0) Resize the table (or its group) to slide width
    _resize_table_to_slide(prs, table_shape, horiz_margin_in=0.5)

    # 1) Column widths — give more room to Initiative & Success criteria
    widths = widths or [2.2, 1.8, 3.4, 1.4, 1.2, 1.2]  # inches
    _set_col_widths(tbl, widths)

    # 2) Cell formatting: padding, vertical centering, text autofit + fallback
    for r in range(len(tbl.rows)):
        for c in range(len(tbl.columns)):
            cell = tbl.cell(r, c)
            try:
                # modest padding
                cell.margin_left = Pt(2)
                cell.margin_right = Pt(2)
                cell.margin_top = Pt(1)
                cell.margin_bottom = Pt(1)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            except Exception:
                pass
            if cell.text_frame:
                _apply_text_autofit(cell.text_frame, start_pt=max_pt)
                _fallback_shrink(cell.text_frame, hard_min_pt=9)

# ---------- UI ----------
st.markdown("## 🧩 Initiatives Slide Filler")
st.caption("Upload your PPT template (with a 6-col table named INIT_TABLE), upload context (QBR/notes), and I’ll auto-fill 3 initiatives × 3 steps — with autoshrink and slide-fit formatting.")

with st.form("form"):
    col1, col2 = st.columns(2)
    customer = col1.text_input("Customer", "Acme Corp")
    industry = col2.text_input("Industry", "Manufacturing")
    arr = col1.text_input("ARR Segment", "$1–5M ARR")
    horizon = col2.text_input("Time Horizon", "12 months")

    objectives = st.text_area("Top objectives (comma-separated)", "Reduce onboarding time by 20%, Improve compliance accuracy, Grow analytics adoption")
    baselines  = st.text_area("Baselines / current state", "NPS=45; 315 active users; WAU=68%; Analytics adoption=32%")
    constraints = st.text_area("Constraints", "Limited training resources; reliance on vendor workshops")
    stakeholders = st.text_area("Stakeholders", "CFO (sponsor); HR Systems Lead (admin); Analytics Lead (champion)")
    risks = st.text_area("Known risks", "Adoption resistance; renewal in 6 months requires ROI proof")

    uploads = st.file_uploader(
        "Upload context files (PDF / DOCX / TXT) — optional",
        type=["pdf", "docx", "txt"], accept_multiple_files=True
    )
    ppt_template = st.file_uploader(
        "Upload your PPT template (.pptx) with table named INIT_TABLE",
        type=["pptx"], accept_multiple_files=False
    )

    model = st.selectbox("Model", ["gpt-4o-mini","gpt-4o","gpt-4-turbo","gpt-3.5-turbo"], index=0)
    submitted = st.form_submit_button("Generate & Fill Slide")

if submitted:
    # Require a template
    if ppt_template is None:
        st.error("Please upload a PPTX template with a table named INIT_TABLE.")
        st.stop()

    # Gather context
    query = f"{customer} {industry} {arr} {horizon} Objectives: {objectives} Baselines: {baselines} Constraints: {constraints}"
    all_texts = []
    for f in (uploads or []):
        try:
            all_texts.append(extract_text_from_file(f))
        except Exception:
            pass
    snippets = pick_top_snippets(all_texts, query, k=8)
    context_blob = "\n\n---\n\n".join(snippets)[:8000]

    # Ask model for exactly 3 initiatives × 3 steps (STRICT JSON)
    system = "You output STRICT JSON only. No markdown or extra text."
    user = f"""
Create exactly 3 initiatives for an HR ERP/HCM account, each with exactly 3 steps.
Ensure each step has owner, target (e.g., Q1/Q2/Month), and status (one of: not started, in progress, completed).

Return ONLY this JSON schema:
{{
  "initiatives": [
    {{
      "name": "short title",
      "success_criteria": "1-2 sentences, measurable",
      "steps": [
        {{"name":"specific step","owner":"role/person","target":"Q1/Q2/Month","status":"not started|in progress|completed"}},
        {{"name":"specific step","owner":"role/person","target":"Q1/Q2/Month","status":"not started|in progress|completed"}},
        {{"name":"specific step","owner":"role/person","target":"Q1/Q2/Month","status":"not started|in progress|completed"}}
      ]
    }},
    {{
      "name":"...",
      "success_criteria":"...",
      "steps":[ ...3 items... ]
    }},
    {{
      "name":"...",
      "success_criteria":"...",
      "steps":[ ...3 items... ]
    }}
  ]
}}

Account context:
Customer: {customer}
Industry: {industry}
ARR: {arr}
Horizon: {horizon}
Objectives: {objectives}
Baselines: {baselines}
Constraints: {constraints}
Stakeholders: {stakeholders}
Risks: {risks}

Relevant uploaded excerpts (paraphrase succinctly):
{context_blob}
"""
    try:
        raw = call_llm(system, user, model)
        data = extract_json(raw)
    except Exception as e:
        st.error(f"Could not generate initiatives. {e}")
        st.stop()

    # Load template
    try:
        prs = Presentation(ppt_template)
        if not prs.slide_layouts:
            raise ValueError("Template has no layouts")
    except Exception as e:
        st.error(f"Could not load template: {e}")
        st.stop()

    # Fill the table
    ok = fill_initiatives_table(prs, data.get("initiatives", []), table_name="INIT_TABLE")
    if not ok:
        st.error("Could not find/fill the table. Check it's named 'INIT_TABLE', has 6 columns, and enough rows (the app will add rows if needed).")
        st.stop()

    # Beautify the table (autoshrink / widths / padding / resize-to-slide)
    shp = _find_table(prs, "INIT_TABLE")
    if shp is not None and hasattr(shp, "table"):
        pretty_format_init_table(prs, shp, max_pt=13, widths=[2.2, 1.8, 3.4, 1.4, 1.2, 1.2])

    # Save+download
    buf = io.BytesIO()
    prs.save(buf); buf.seek(0)
    st.success("Slide filled and formatted!")
    st.download_button(
        "⬇️ Download Updated PowerPoint",
        buf,
        file_name=f"Initiatives_{customer.replace(' ','_')}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
