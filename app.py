import io
import json
import os
import re
from datetime import date
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from openai import OpenAI

MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-20b")

st.set_page_config(page_title="Infrastructure Method Statement Builder", page_icon="🏗️", layout="wide")


SCHEMA = {
    "type": "object", "additionalProperties": False,
    "properties": {
        "title": {"type": "string"}, "purpose": {"type": "string"}, "scope": {"type": "string"},
        "resources": {"type": "array", "items": {"type": "string"}},
        "roles": {"type": "array", "items": {"type": "object", "properties": {"role": {"type": "string"}, "duty": {"type": "string"}}, "required": ["role", "duty"], "additionalProperties": False}},
        "preconditions": {"type": "array", "items": {"type": "string"}},
        "sequence": {"type": "array", "items": {"type": "object", "properties": {"step": {"type": "string"}, "method": {"type": "string"}, "release": {"type": "string"}}, "required": ["step", "method", "release"], "additionalProperties": False}},
        "itp": {"type": "array", "items": {"type": "object", "properties": {"stage": {"type": "string"}, "check": {"type": "string"}, "acceptance": {"type": "string"}, "frequency": {"type": "string"}, "record": {"type": "string"}, "point": {"type": "string"}, "owner": {"type": "string"}}, "required": ["stage", "check", "acceptance", "frequency", "record", "point", "owner"], "additionalProperties": False}},
        "checklist": {"type": "array", "items": {"type": "string"}},
        "risks": {"type": "array", "items": {"type": "object", "properties": {"hazard": {"type": "string"}, "effect": {"type": "string"}, "initial": {"type": "string"}, "controls": {"type": "string"}, "residual": {"type": "string"}, "owner": {"type": "string"}}, "required": ["hazard", "effect", "initial", "controls", "residual", "owner"], "additionalProperties": False}},
        "environment": {"type": "array", "items": {"type": "string"}}, "emergency": {"type": "array", "items": {"type": "string"}},
        "sketch": {"type": "array", "items": {"type": "string"}}, "gaps": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["title", "purpose", "scope", "resources", "roles", "preconditions", "sequence", "itp", "checklist", "risks", "environment", "emergency", "sketch", "gaps"]
}


def draft(activity, key):
    prompt = f"""Draft a generic, comprehensive infrastructure/building construction method statement for this activity: {activity}. The activity name is the only project input. Cover preparation, execution, inspection, completion and handover. Do not claim compliance with project documents or cite specifications/drawings. Do not invent project-specific drawing IDs, dimensions, tolerances, acceptance limits, test frequencies, permits or risk scores. Where site-specific values or approvals are needed, say 'Confirm before work starts' and list the gap. Use broadly applicable good construction practice and practical safety controls, while clearly marking this as a generic draft. ITP points must be H (hold), W (witness), R (review), or S (surveillance); where the project-specific point is unknown, state 'Confirm before work starts'. Risk ratings must say 'Assess using site risk assessment process' rather than assigning unsupported scores. Sketch is 5-9 short sequential labels for an indicative process flow, never a design drawing. All output requires competent engineering, quality and HSE review before use.\n\nActivity: {activity}"""
    client = OpenAI(api_key=key, base_url="https://api.groq.com/openai/v1")
    response = client.chat.completions.create(
        model=MODEL,
        messages=[{"role": "user", "content": prompt}],
        max_completion_tokens=8192,
        reasoning_effort="low",
        response_format={"type": "json_schema", "json_schema": {"name": "method_statement", "strict": True, "schema": SCHEMA}},
    )
    content = response.choices[0].message.content
    if not content:
        raise ValueError("Groq returned an empty response. Try again or check model access.")
    return json.loads(content)


def table(doc, headers, rows):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    t = doc.add_table(rows=1, cols=len(headers)); t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for cell, value in zip(t.rows[0].cells, headers):
        cell.text = str(value)
        for run in cell.paragraphs[0].runs:
            run.bold = True; run.font.color.rgb = RGBColor(255, 255, 255)
        shade = OxmlElement("w:shd"); shade.set(qn("w:fill"), "234A64"); cell._tc.get_or_add_tcPr().append(shade)
    for row in rows:
        for cell, value in zip(t.add_row().cells, row): cell.text = str(value)
    doc.add_paragraph()


def bullets(doc, values):
    for item in values: doc.add_paragraph(str(item), style="List Bullet")


def make_doc(data, activity):
    doc = Document(); sec = doc.sections[0]; sec.top_margin = Inches(.6); sec.bottom_margin = Inches(.6)
    doc.styles["Normal"].font.name = "Aptos"; doc.styles["Normal"].font.size = Pt(9)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("DRAFT — FOR REVIEW AND APPROVAL"); r.bold = True; r.font.color.rgb = RGBColor(190, 50, 45)
    doc.add_heading(data["title"], 0)
    table(doc, ["Document control", "Value"], [["Activity", activity], ["Document number", "Assign per project control procedure"], ["Revision", "A — Draft"], ["Date", date.today().isoformat()], ["Status", "For competent review; not for construction"]])
    doc.add_heading("1. Purpose", 1); doc.add_paragraph(data["purpose"])
    doc.add_heading("2. Scope", 1); doc.add_paragraph(data["scope"])
    doc.add_heading("3. Resources and responsibilities", 1); bullets(doc, data["resources"])
    table(doc, ["Role", "Responsibility"], [[x["role"], x["duty"]] for x in data["roles"]])
    doc.add_heading("4. Preconditions", 1); bullets(doc, data["preconditions"])
    doc.add_heading("5. Work method and sequence", 1)
    table(doc, ["Step", "Method", "Release / hold requirement"], [[x["step"], x["method"], x["release"]] for x in data["sequence"]])
    doc.add_heading("6. Indicative activity sketch", 1)
    doc.add_paragraph("Conceptual process flow only; not to scale and not for construction.")
    table(doc, ["Process flow"], [[f"{i+1}. {node}"] for i, node in enumerate(data["sketch"])])
    doc.add_heading("7. QA/QC inspection and test plan", 1)
    table(doc, ["Stage", "Inspection / test", "Acceptance criteria", "Frequency", "Record", "Point", "Owner"], [[x[k] for k in ("stage", "check", "acceptance", "frequency", "record", "point", "owner")] for x in data["itp"]])
    doc.add_heading("8. Field checklist", 1)
    table(doc, ["No.", "Check", "Status / remarks"], [[i+1, item, "☐ Accept  ☐ Reject  ☐ N/A — ______"] for i, item in enumerate(data["checklist"])])
    doc.add_paragraph("Inspected by: __________________   Date: __________   Signature: __________________")
    doc.add_heading("9. Safety risk assessment", 1)
    table(doc, ["Hazard", "Potential effect", "Initial risk", "Controls", "Residual risk", "Owner"], [[x[k] for k in ("hazard", "effect", "initial", "controls", "residual", "owner")] for x in data["risks"]])
    doc.add_paragraph("Verify risk ratings and controls against site conditions before work starts.")
    doc.add_heading("10. Environmental controls", 1); bullets(doc, data["environment"])
    doc.add_heading("11. Emergency arrangements", 1); bullets(doc, data["emergency"])
    doc.add_heading("12. Assumptions and information gaps", 1); bullets(doc, data["gaps"] or ["Confirm site-specific requirements before work starts."])
    doc.add_heading("13. Review and approval", 1)
    table(doc, ["Prepared by", "Checked by", "Approved by"], [["Name / signature / date", "Name / signature / date", "Name / signature / date"]])
    doc.add_paragraph("This AI-assisted generic document is a draft. Competent construction, engineering, QA/QC and HSE personnel must review it for the actual site conditions, permits, temporary works and applicable requirements before use.")
    stream = io.BytesIO(); doc.save(stream); return stream.getvalue()


st.title("Infrastructure & Buildings Method Statement Builder")
st.caption("Enter an activity name to draft a method statement and download it in Word format. Powered by Groq.")
st.warning("Every generated document is a draft requiring competent engineering, QA/QC and HSE review before approval or construction use.")
st.info("This version generates generic drafts from the activity name only. It does not use or claim compliance with project specifications or drawings.")
activity = st.text_input("Activity name", placeholder="e.g., Reinforced concrete foundation construction")
api_key = st.secrets.get("GROQ_API_KEY", os.getenv("GROQ_API_KEY", ""))
if st.button("Generate method statement", type="primary", disabled=not activity.strip()):
    if not api_key: st.error("Configure GROQ_API_KEY in Streamlit secrets before generating.")
    else:
        with st.spinner("Generating a generic activity draft…"):
            try: st.session_state.generated = (draft(activity.strip(), api_key), activity.strip())
            except Exception as exc: st.error(f"Generation failed: {exc}")
if "generated" in st.session_state:
    result, act = st.session_state.generated
    st.success("Draft generated. Resolve the listed gaps and review all technical content.")
    st.subheader(result["title"])
    st.markdown("**Information gaps to resolve**")
    for gap in result["gaps"]: st.write("• " + gap)
    with st.expander("Preview work sequence"):
        for item in result["sequence"]: st.markdown(f"**{item['step']}** — {item['method']}\n\nRelease: {item['release']}")
    filename = re.sub(r"[^A-Za-z0-9]+", "_", act).strip("_") or "method_statement"
    st.download_button("Download Word method statement", make_doc(result, act), file_name=f"{filename}_method_statement.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
