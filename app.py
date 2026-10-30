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


def draft(activity, key):
    prompt = f"""Draft a generic infrastructure/building method statement for this activity: {activity}. The activity name is the only project input. Cover preparation, work steps, inspection, completion and handover. Do not claim compliance with project documents or cite specifications/drawings. Do not invent project-specific dimensions, tolerances, acceptance limits, test frequencies, permits or risk scores. Mark site-specific unknowns 'Confirm before work starts' and list them in gaps. Keep every text field concise (1-3 sentences) to fit the document. ITP points must be H, W, R or S; where unknown, say 'Confirm before work starts'. Risk ratings must say 'Assess using site risk assessment process'. Sketch must contain 5-9 short process labels, not a design drawing. Require competent engineering, quality and HSE review.

Return one valid JSON object only, with exactly these keys and value types. Do not wrap arrays in objects, and do not include the words 'items' or 'type' as wrapper fields:
{{"title":"string","purpose":"string","scope":"string","resources":["string"],"roles":[{{"role":"string","duty":"string"}}],"preconditions":["string"],"sequence":[{{"step":"string","method":"string","release":"string"}}],"itp":[{{"stage":"string","check":"string","acceptance":"string","frequency":"string","record":"string","point":"string","owner":"string"}}],"checklist":["string"],"risks":[{{"hazard":"string","effect":"string","initial":"string","controls":"string","residual":"string","owner":"string"}}],"environment":["string"],"emergency":["string"],"sketch":["string"],"gaps":["string"]}}

Activity: {activity}"""
    client = OpenAI(api_key=key, base_url="https://api.groq.com/openai/v1")
    response = client.chat.completions.create(
        model=MODEL,
        messages=[{"role": "user", "content": prompt}],
        max_completion_tokens=8192,
        reasoning_effort="low",
        response_format={"type": "json_object"},
    )
    content = response.choices[0].message.content
    if not content:
        raise ValueError("Groq returned an empty response. Try again or check model access.")
    try:
        raw = json.loads(content)
    except json.JSONDecodeError as exc:
        raise ValueError("Groq returned invalid JSON. Please try generating again.") from exc
    return normalize_result(raw)


def normalize_result(raw):
    """Repair common JSON-mode wrappers and fill missing sections safely."""
    if not isinstance(raw, dict):
        raise ValueError("Groq returned a JSON value that was not an object. Please try again.")

    string_fields = ("title", "purpose", "scope")
    string_lists = ("resources", "preconditions", "checklist", "environment", "emergency", "sketch", "gaps")
    object_lists = {
        "roles": ("role", "duty"),
        "sequence": ("step", "method", "release"),
        "itp": ("stage", "check", "acceptance", "frequency", "record", "point", "owner"),
        "risks": ("hazard", "effect", "initial", "controls", "residual", "owner"),
    }

    def unwrap(value):
        if isinstance(value, list):
            return value
        if isinstance(value, dict):
            for key in ("items", "values", "data"):
                if isinstance(value.get(key), list):
                    return value[key]
        return []

    for key in string_fields:
        value = raw.get(key, "")
        raw[key] = value if isinstance(value, str) else str(value)
    for key in string_lists:
        values = unwrap(raw.get(key))
        cleaned = []
        for item in values:
            if isinstance(item, str):
                cleaned.append(item)
            elif isinstance(item, dict):
                text = item.get("text") or item.get("item") or item.get("description")
                if isinstance(text, str):
                    cleaned.append(text)
        raw[key] = cleaned
    for key, fields in object_lists.items():
        values = unwrap(raw.get(key))
        raw[key] = [
            {field: (item.get(field) if isinstance(item.get(field), str) else "Confirm before work starts") for field in fields}
            for item in values if isinstance(item, dict)
        ]
    return raw


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
