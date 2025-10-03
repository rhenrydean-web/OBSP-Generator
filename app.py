# app.py
import io, os, json
import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="OBSP Generator (Robust)", page_icon="📘", layout="centered")

PRIMARY = RGBColor(30, 64, 175)   # deep blue
BODY_SIZE = 18
TITLE_SIZE = 36

# ---------- UI HELPERS ----------
def add_title(slide, text):
    title = slide.shapes.title
    title.text = text
    p = title.text_frame.paragraphs[0]
    p.font.size = Pt(TITLE_SIZE); p.font.bold = True; p.font.color.rgb = PRIMARY

def add_body(slide, text):
    ph = slide.placeholders[1]
    ph.text = text
    for p in ph.text_frame.paragraphs:
        p.font.size = Pt(BODY_SIZE)

def add_bullets(slide, lines):
    tf = slide.placeholders[1].text_frame
    tf.clear()
    for i, line in enumerate([l for l in lines if l.strip()]):
        p = tf.add_paragraph() if i else tf.paragraphs[0]
        p.text = line.strip(); p.level = 0; p.font.size = Pt(BODY_SIZE)

def add_two_col_table(slide, left_title, left_items, right_title, right_items):
    rows = max(len(left_items), len(right_items)) + 1
    table = slide.shapes.add_table(rows, 2, Pt(20), Pt(120), Pt(900), Pt(360)).table
    table.cell(0,0).text, table.cell(0,1).text = left_title, right_title
    for j in (0,1):
        p = table.cell(0,j).text_frame.paragraphs[0]
        p.font.bold = True; p.font.size = Pt(BODY_SIZE)
    for i in range(1, rows):
        table.cell(i,0).text = left_items[i-1] if i-1 < len(left_items) else ""
        table.cell(i,1).text = right_items[i-1] if i-1 < len(right_items) else ""

def add_kpi_table(slide, kpis):
    rows, cols = len(kpis)+1, 5
    table = slide.shapes.add_table(rows, cols, Pt(20), Pt(120), Pt(900), Pt(360)).table
    headers = ["Metric","Baseline","Target","Cadence","Owner"]
    for j,h in enumerate(headers):
        c = table.cell(0,j); c.text = h
        p = c.text_frame.paragraphs[0]; p.font.bold = True; p.font.size = Pt(BODY_SIZE)
    for i,k in enumerate(kpis, start=1):
        table.cell(i,0).text = k.get("metric","")
        table.cell(i,1).text = k.get("baseline","")
        table.cell(i,2).text = k.get("target","")
        table.cell(i,3).text = k.get("cadence","")
        table.cell(i,4).text = k.get("owner","")

def extract_json(raw: str):
    start = raw.find("{"); end = raw.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("No JSON object found in LLM response.")
    return json.loads(raw[start:end+1])

def call_llm(system, user, model):
    if "OPENAI_API_KEY" not in st.secrets or not st.secrets["OPENAI_API_KEY"]:
        raise RuntimeError("Missing OPENAI_API_KEY secret. Add it in Streamlit Cloud → App → Settings → Secrets.")
    from openai import OpenAI
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role":"system","content":system},{"role":"user","content":user}],
        temperature=0.2,
    )
    return resp.choices[0].message.content

# ---------- Context extraction (lightweight) ----------
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

# ---------- Safe slide creation helper ----------
def safe_add_slide(prs, layout_idx=0):
    try:
        layout = prs.slide_layouts[layout_idx]
    except:
        # fallback to default presentation layout if template is broken
        tmp = Presentation()
        layout = tmp.slide_layouts[0]
    return prs.slides.add_slide(layout)

# ---------- APP UI ----------
st.markdown("## 📘 Outcome-Based Customer Success Plan — Robust Generator")
st.caption("Fill the form, (optionally) upload context + a PPT template, and download a polished deck.")

with st.form("obsp"):
    col1, col2 = st.columns(2)
    customer = col1.text_input("Customer", "Acme Corp")
    industry = col2.text_input("Industry", "SaaS - HR Tech")
    arr = col1.text_input("ARR Segment", "$1–5M ARR")
    horizon = col2.text_input("Time Horizon", "12 months")

    objectives = st.text_area("Customer Outcomes (comma-separated)", "Reduce onboarding time by 20%, Improve compliance accuracy")
    baselines  = st.text_area("Baselines (e.g., NPS=45, Adoption=60%)", "NPS=45, 315 active users, WAU=68%")
    constraints = st.text_area("Constraints", "Limited training resources; integration backlog")
    stakeholders = st.text_area("Stakeholders", "CFO (Exec Sponsor), HR Systems Lead (Admin), Payroll Manager (Champion)")
    risks = st.text_area("Known Risks", "Analytics adoption lagging; renewal in 6 months")

    uploads = st.file_uploader(
        "Upload context files (QBRs, notes, SOWs) — PDF, DOCX, or TXT (optional)",
        type=["pdf","docx","txt"], accept_multiple_files=True
    )
    ppt_template = st.file_uploader(
        "Optional: Upload a PowerPoint template (.pptx) to apply your brand/theme",
        type=["pptx"], accept_multiple_files=False
    )

    model = st.selectbox("Model", ["gpt-4o-mini","gpt-4o","gpt-4-turbo","gpt-3.5-turbo"], index=0)
    submitted = st.form_submit_button("Generate Plan")

if submitted:
    # Gather context snippets
    query = f"{customer} {industry} {arr} {horizon} Objectives: {objectives} Baselines: {baselines}"
    all_texts = []
    for f in (uploads or []):
        try:
            all_texts.append(extract_text_from_file(f))
        except Exception:
            pass
    snippets = pick_top_snippets(all_texts, query, k=8)
    context_blob = "\n\n---\n\n".join(snippets)[:8000]

    # Prompt
    system = "You output STRICT JSON for a robust customer success plan. No markdown, no extra prose."
    user = f"""
JSON schema:
{{
  "cover": {{"title":"Outcome-Based Customer Success Plan — {customer}","account_meta":"Industry: {industry} | ARR: {arr} | Horizon: {horizon}"}},
  "purpose": "Why CS plan matters...",
  "objectives": [{", ".join([f'"{o.strip()}"' for o in objectives.split(",")])}],
  "kpis": [{{"metric":"NPS","baseline":"{baselines}","target":"Improve vs baseline","cadence":"Monthly","owner":"CSM"}}],
  "milestones": [
    {{"phase":"Onboarding (0–30d)","deliverables":["Training complete","Integrations configured"]}},
    {{"phase":"Adoption (30–90d)","deliverables":[">60% WAU","Analytics rollout"]}},
    {{"phase":"Optimization (90–180d)","deliverables":["Advanced features live"]}},
    {{"phase":"Renewal Prep (180–365d)","deliverables":["ROI case study","Exec QBR"]}}
  ],
  "roles_vendor":["CSM — outcomes","SE — integrations","Support — resolution"],
  "roles_customer":["CFO — sponsor","HR Lead — admin","Payroll Mgr — champion"],
  "risks":[{{"risk":"Low adoption","mitigation":"Exec sponsor + champion"}}],
  "governance":"Weekly status, monthly steering, quarterly exec QBR."
}}

Context:
{context_blob}
"""
    try:
        raw = call_llm(system, user, model)
        data = extract_json(raw)
    except Exception as e:
        st.error(f"Could not create plan. {e}")
        st.stop()

    # Load PPT (safe fallback if template is broken)
    try:
        prs = Presentation(ppt_template) if ppt_template is not None else Presentation()
        if not prs.slide_layouts:
            raise ValueError("Template has no layouts")
    except Exception as e:
        st.warning(f"Template invalid ({e}), using default.")
        prs = Presentation()

    # Build slides safely
    s = safe_add_slide(prs, 0)
    add_title(s, data["cover"]["title"])
    s.placeholders[1].text = data["cover"]["account_meta"]

    s = safe_add_slide(prs, 1)
    add_title(s, "Purpose")
    add_body(s, data.get("purpose",""))

    s = safe_add_slide(prs, 1)
    add_title(s, "Objectives")
    add_bullets(s, data.get("objectives", []))

    s = safe_add_slide(prs, 1)
    add_title(s, "KPIs")
    add_kpi_table(s, data.get("kpis", []))

    s = safe_add_slide(prs, 1)
    add_title(s, "Milestones")
    lines = [f"{m.get('phase','')}: " + "; ".join(m.get('deliverables',[])) for m in data.get("milestones",[])]
    add_bullets(s, lines)

    s = safe_add_slide(prs, 1)
    add_title(s, "Roles & Responsibilities")
    add_two_col_table(s, "Vendor Team", data.get("roles_vendor", []), "Customer Team", data.get("roles_customer", []))

    s = safe_add_slide(prs, 1)
    add_title(s, "Risks & Mitigation")
    risk_lines = [f"{r.get('risk','')} — Mitigation: {r.get('mitigation','')}" for r in data.get("risks",[])]
    add_bullets(s, risk_lines)

    s = safe_add_slide(prs, 1)
    add_title(s, "Engagement & Governance")
    add_body(s, data.get("governance",""))

    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    st.success("Plan generated!")
    st.download_button("⬇️ Download PowerPoint", buf, file_name=f"OBSP_{customer.replace(' ','_')}.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
