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

# ---------- TABLE FILL HELPERS ----------
def _find_table(prs, table_name="INIT_TABLE"):
    """Find table by name; if not found, try header match."""
    # 1) By exact shape name
    for slide in prs.slides:
        for sh in slide.shapes:
            if getattr(sh, "name", "") == table_name and hasattr(sh, "table"):
                return sh
            if hasattr(sh, "shapes"):
                for s in sh.shapes:
                    if getattr(s, "name", "") == table_name and hasattr(s, "table"):
                        return s
    # 2) Fallback by header match
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

def fill_initiatives_table(prs, initiatives, table_name="INIT_TABLE"):
    """
    Fills a 6-col table with 9 rows (+ header) for 3 initiatives × 3 steps.
    Columns: [Initiative, Step, Success criteria, Owner, Target, Status]
    Merges Initiative & Success criteria cells across each 3-row block if possible.
    """
    shp = _find_table(prs, table_name)
    if shp is None or not hasattr(shp, "table"):
        return False
    tbl = shp.table

    # Basic checks
    if len(tbl.columns) < 6 or len(tbl.rows) < 10:
        return False
    if len(initiatives) != 3:
        return False

    # Clear data rows
    for r in range(1, len(tbl.rows)):
        for c in range(6):
            tbl.cell(r, c).text = ""

    row_idx = 1
    for init in initiatives:
        name = init.get("name","")
        crit = init.get("success_criteria","")
        steps = init.get("steps", [])[:3]
        if len(steps) < 3:
            steps += [{"name":"TBD","owner":"TBD","target":"TBD","status":"not started"}] * (3 - len(steps))

        start = row_idx
        for step in steps:
            tbl.cell(row_idx, 1).text = step.get("name","")
            tbl.cell(row_idx, 3).text = step.get("owner","")
            tbl.cell(row_idx, 4).text = step.get("target","")
            tbl.cell(row_idx, 5).text = step.get("status","not started")
            row_idx += 1
        end = row_idx - 1

        # Merge & fill Initiative
        try:
            tbl.cell(start, 0).merge(tbl.cell(end, 0))
        except:  # already merged or not supported
            pass
        tbl.cell(start, 0).text = name

        # Merge & fill Success criteria
        try:
            tbl.cell(start, 2).merge(tbl.cell(end, 2))
        except:
            pass
        tbl.cell(start, 2).text = crit

    return True

# ---------- UI ----------
st.markdown("## 🧩 Initiatives Slide Filler")
st.caption("Upload your PPT template (with a 6-col table named INIT_TABLE), upload context (QBR/notes), and I’ll auto-fill 3 initiatives × 3 steps.")

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
    submitted = st.form_submit_button("Generate Initiatives & Fill Slide")

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
        st.error("Could not find/fill the table. Check it's named 'INIT_TABLE', has 6 columns, and 10 rows (header + 9).")
        st.stop()

    # Save+download
    buf = io.BytesIO()
    prs.save(buf); buf.seek(0)
    st.success("Slide filled!")
    st.download_button(
        "⬇️ Download Updated PowerPoint",
        buf,
        file_name=f"Initiatives_{customer.replace(' ','_')}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
