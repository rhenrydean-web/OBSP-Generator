import io, os, json, re
import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="OBSP Generator", page_icon="📘", layout="centered")
PRIMARY = RGBColor(30, 64, 175)
BODY_SIZE = 18
TITLE_SIZE = 36

def add_title(slide, text):
    title = slide.shapes.title
    title.text = text
    p = title.text_frame.paragraphs[0]
    p.font.size = Pt(TITLE_SIZE); p.font.bold = True; p.font.color.rgb = PRIMARY

def add_body(slide, text):
    ph = slide.placeholders[1]
    ph.text = text
    for p in ph.text_frame.paragraphs: p.font.size = Pt(BODY_SIZE)

def add_bullets(slide, lines):
    tf = slide.placeholders[1].text_frame
    tf.clear()
    for i, line in enumerate([l for l in lines if l.strip()]):
        p = tf.add_paragraph() if i else tf.paragraphs[0]
        p.text = line.strip(); p.level = 0; p.font.size = Pt(BODY_SIZE)

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

def extract_json(raw):
    # robustly grab the first {...} block
    start = raw.find("{"); end = raw.rfind("}")
    if start == -1 or end == -1: raise ValueError("No JSON in response")
    return json.loads(raw[start:end+1])

def call_llm(system, user, model):
    from openai import OpenAI
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
    resp = client.chat.completions.create(
        model=model, messages=[{"role":"system","content":system},{"role":"user","content":user}], temperature=0.2
    )
    return resp.choices[0].message.content

st.markdown("## 📘 Outcome-Based Success Plan (OBSP) Generator")
st.caption("Fill the form and generate a polished PowerPoint plan.")

with st.form("obsp"):
    col1, col2 = st.columns(2)
    customer = col1.text_input("Customer", "Acme Corp")
    industry = col2.text_input("Industry", "SaaS - HR Tech")
    arr = col1.text_input("ARR Segment", "$1–5M ARR")
    horizon = col2.text_input("Time Horizon", "12 months")

    objectives = st.text_area("Objectives (comma-separated)", "Reduce churn by 10%, Improve onboarding time by 20%")
    baselines  = st.text_area("Baselines (e.g., NPS=45, Adoption=60%)", "NPS=45, Adoption=60%")
    constraints= st.text_area("Constraints", "Limited customer resources; integration backlog")
    stakeholders = st.text_area("Stakeholders", "Customer: Exec Sponsor, Admin, Champions; Vendor: CSM, SE, Support")
    risks = st.text_area("Known Risks", "Low admin bandwidth; End-user training gaps")

    model = st.selectbox("Model", ["gpt-4o-mini","gpt-4o","gpt-4-turbo","gpt-3.5-turbo"], index=0)
    submitted = st.form_submit_button("Generate Plan")

if submitted:
    system = "You output STRICT JSON for customer success plans. No markdown."
    user = f"""
JSON schema:
{{
  "exec_summary": "2-3 short paragraphs.",
  "objectives": ["...","...","..."],
  "kpis": [{{"metric":"","baseline":"","target":"","cadence":"","owner":""}}],
  "milestones": [
    {{"phase":"Onboarding (0–30 days)","deliverables":["...","..."]}},
    {{"phase":"Adoption (30–90 days)","deliverables":["...","..."]}},
    {{"phase":"Optimization (90–180 days)","deliverables":["...","..."]}},
    {{"phase":"Renewal Prep (180–365 days)","deliverables":["...","..."]}}
  ],
  "roles_vendor": ["role — responsibility","role — responsibility"],
  "roles_customer": ["role — responsibility","role — responsibility"],
  "risks": [{{"risk":"","mitigation":""}}],
  "governance": "Narrative of cadences and escalation."
}}

Context:
Customer: {customer}
Industry: {industry}
ARR Segment: {arr}
Objectives: {objectives}
Baselines: {baselines}
Time Horizon: {horizon}
Constraints: {constraints}
Stakeholders: {stakeholders}
Known Risks: {risks}
"""
    try:
        raw = call_llm(system, user, model)
        data = extract_json(raw)
    except Exception as e:
        st.error(f"Could not parse model response as JSON. Try a different model or re-run. Error: {e}")
        st.stop()

    # Build PPT
    prs = Presentation()
    # Title
    s = prs.slides.add_slide(prs.slide_layouts[0])
    add_title(s, f"Outcome-Based Success Plan — {customer}")
    s.placeholders[1].text = f"Industry: {industry}  |  ARR: {arr}\nHorizon: {horizon}"
    for p in s.placeholders[1].text_frame.paragraphs: p.font.size = Pt(18)

    # Exec
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "1. Executive Summary")
    add_body(s, data.get("exec_summary",""))

    # Objectives
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "2. Customer Objectives")
    add_bullets(s, data.get("objectives", []))

    # KPIs
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "3. Success Metrics & KPIs")
    add_kpi_table(s, data.get("kpis", []))

    # Milestones
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "4. Milestones & Timeline")
    lines = [f"{m.get('phase','')}: " + "; ".join(m.get('deliverables',[])) for m in data.get("milestones",[])]
    add_bullets(s, lines)

    # Roles
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "5. Roles & Responsibilities")
    add_two_col_table(s, "Vendor Team", data.get("roles_vendor", []), "Customer Team", data.get("roles_customer", []))

    # Risks
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "6. Risks & Mitigation")
    risk_lines = [f"{r.get('risk','')} — Mitigation: {r.get('mitigation','')}" for r in data.get("risks",[])]
    add_bullets(s, risk_lines)

    # Governance
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "7. Engagement & Governance")
    add_body(s, data.get("governance",""))

    buf = io.BytesIO()
    prs.save(buf); buf.seek(0)
    st.success("Plan generated!")
    st.download_button("⬇️ Download PowerPoint", buf, file_name=f"OBSP_{customer.replace(' ','_')}.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
import io, os, json, re
import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="OBSP Generator", page_icon="📘", layout="centered")
PRIMARY = RGBColor(30, 64, 175)
BODY_SIZE = 18
TITLE_SIZE = 36

def add_title(slide, text):
    title = slide.shapes.title
    title.text = text
    p = title.text_frame.paragraphs[0]
    p.font.size = Pt(TITLE_SIZE); p.font.bold = True; p.font.color.rgb = PRIMARY

def add_body(slide, text):
    ph = slide.placeholders[1]
    ph.text = text
    for p in ph.text_frame.paragraphs: p.font.size = Pt(BODY_SIZE)

def add_bullets(slide, lines):
    tf = slide.placeholders[1].text_frame
    tf.clear()
    for i, line in enumerate([l for l in lines if l.strip()]):
        p = tf.add_paragraph() if i else tf.paragraphs[0]
        p.text = line.strip(); p.level = 0; p.font.size = Pt(BODY_SIZE)

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

def extract_json(raw):
    # robustly grab the first {...} block
    start = raw.find("{"); end = raw.rfind("}")
    if start == -1 or end == -1: raise ValueError("No JSON in response")
    return json.loads(raw[start:end+1])

def call_llm(system, user, model):
    from openai import OpenAI
    client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])
    resp = client.chat.completions.create(
        model=model, messages=[{"role":"system","content":system},{"role":"user","content":user}], temperature=0.2
    )
    return resp.choices[0].message.content

st.markdown("## 📘 Outcome-Based Success Plan (OBSP) Generator")
st.caption("Fill the form and generate a polished PowerPoint plan.")

with st.form("obsp"):
    col1, col2 = st.columns(2)
    customer = col1.text_input("Customer", "Acme Corp")
    industry = col2.text_input("Industry", "SaaS - HR Tech")
    arr = col1.text_input("ARR Segment", "$1–5M ARR")
    horizon = col2.text_input("Time Horizon", "12 months")

    objectives = st.text_area("Objectives (comma-separated)", "Reduce churn by 10%, Improve onboarding time by 20%")
    baselines  = st.text_area("Baselines (e.g., NPS=45, Adoption=60%)", "NPS=45, Adoption=60%")
    constraints= st.text_area("Constraints", "Limited customer resources; integration backlog")
    stakeholders = st.text_area("Stakeholders", "Customer: Exec Sponsor, Admin, Champions; Vendor: CSM, SE, Support")
    risks = st.text_area("Known Risks", "Low admin bandwidth; End-user training gaps")

    model = st.selectbox("Model", ["gpt-4o-mini","gpt-4o","gpt-4-turbo","gpt-3.5-turbo"], index=0)
    submitted = st.form_submit_button("Generate Plan")

if submitted:
    system = "You output STRICT JSON for customer success plans. No markdown."
    user = f"""
JSON schema:
{{
  "exec_summary": "2-3 short paragraphs.",
  "objectives": ["...","...","..."],
  "kpis": [{{"metric":"","baseline":"","target":"","cadence":"","owner":""}}],
  "milestones": [
    {{"phase":"Onboarding (0–30 days)","deliverables":["...","..."]}},
    {{"phase":"Adoption (30–90 days)","deliverables":["...","..."]}},
    {{"phase":"Optimization (90–180 days)","deliverables":["...","..."]}},
    {{"phase":"Renewal Prep (180–365 days)","deliverables":["...","..."]}}
  ],
  "roles_vendor": ["role — responsibility","role — responsibility"],
  "roles_customer": ["role — responsibility","role — responsibility"],
  "risks": [{{"risk":"","mitigation":""}}],
  "governance": "Narrative of cadences and escalation."
}}

Context:
Customer: {customer}
Industry: {industry}
ARR Segment: {arr}
Objectives: {objectives}
Baselines: {baselines}
Time Horizon: {horizon}
Constraints: {constraints}
Stakeholders: {stakeholders}
Known Risks: {risks}
"""
    try:
        raw = call_llm(system, user, model)
        data = extract_json(raw)
    except Exception as e:
        st.error(f"Could not parse model response as JSON. Try a different model or re-run. Error: {e}")
        st.stop()

    # Build PPT
    prs = Presentation()
    # Title
    s = prs.slides.add_slide(prs.slide_layouts[0])
    add_title(s, f"Outcome-Based Success Plan — {customer}")
    s.placeholders[1].text = f"Industry: {industry}  |  ARR: {arr}\nHorizon: {horizon}"
    for p in s.placeholders[1].text_frame.paragraphs: p.font.size = Pt(18)

    # Exec
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "1. Executive Summary")
    add_body(s, data.get("exec_summary",""))

    # Objectives
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "2. Customer Objectives")
    add_bullets(s, data.get("objectives", []))

    # KPIs
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "3. Success Metrics & KPIs")
    add_kpi_table(s, data.get("kpis", []))

    # Milestones
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "4. Milestones & Timeline")
    lines = [f"{m.get('phase','')}: " + "; ".join(m.get('deliverables',[])) for m in data.get("milestones",[])]
    add_bullets(s, lines)

    # Roles
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "5. Roles & Responsibilities")
    add_two_col_table(s, "Vendor Team", data.get("roles_vendor", []), "Customer Team", data.get("roles_customer", []))

    # Risks
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "6. Risks & Mitigation")
    risk_lines = [f"{r.get('risk','')} — Mitigation: {r.get('mitigation','')}" for r in data.get("risks",[])]
    add_bullets(s, risk_lines)

    # Governance
    s = prs.slides.add_slide(prs.slide_layouts[1])
    add_title(s, "7. Engagement & Governance")
    add_body(s, data.get("governance",""))

    buf = io.BytesIO()
    prs.save(buf); buf.seek(0)
    st.success("Plan generated!")
    st.download_button("⬇️ Download PowerPoint", buf, file_name=f"OBSP_{customer.replace(' ','_')}.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
