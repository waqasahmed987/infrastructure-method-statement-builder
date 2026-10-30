import io
import json
import os
import re
import textwrap
from datetime import date
import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from openai import OpenAI

MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-20b")

st.set_page_config(page_title="Infrastructure Method Statement Builder", page_icon="🏗️", layout="wide")


def draft(activity, key):
    prompt = f"""Prepare a detailed, activity-specific generic construction method statement for: {activity}. The activity name is the only project input. Do not write a short overview. Cover planning and preparation, mobilization, materials/equipment receipt, workface setup, execution in logical sub-steps, inspections/tests, correction of defects, protection, cleanup, completion and handover. Include only steps applicable to this activity. Do not claim compliance with project documents or cite specifications/drawings. Do not invent dimensions, tolerances, acceptance limits, test frequencies, permits or risk scores. Where site-specific requirements are unknown, write 'Confirm before work starts' and list the required information in gaps.

Make the work method thorough and practical: provide 12-16 sequential steps when the activity warrants it; each step's method should explain the action, responsible trade/role, tools/plant or materials, execution controls, quality check, and safety control in about 2-4 sentences. ITP: provide 8-12 activity-relevant inspection stages, clear checks, qualitative acceptance criteria without invented numbers, frequency/record, H/W/R/S point, and responsible role. Checklist: at least 12 field checks. Risk assessment: at least 8 activity-specific hazards with consequences and practical controls; initial/residual ratings must say 'Assess on site using the project risk matrix'. Include at least 5 environmental controls and 4 emergency actions. Preconditions should be specific and cover permits, access, competence, coordination, materials, equipment and work area where relevant. Sketch labels: 5-9 short sequential nodes for a process flow. All content is a generic draft for competent engineering, QA/QC and HSE review.

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


def flow_sketch_png(activity, nodes):
    labels = nodes[:9] or ["Plan and brief", "Prepare work area", "Execute activity", "Inspect", "Protect and hand over"]
    fig, ax = plt.subplots(figsize=(8.0, 6.3))
    fig.patch.set_facecolor("white"); ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
    ax.text(.5, .98, f"ACTIVITY WORK SEQUENCE — {activity[:65]}", ha="center", va="top", fontsize=10, weight="bold", color="#234A64")
    top, gap, box_h = .89, .095, .065
    for i, label in enumerate(labels):
        y = top - i * gap
        ax.text(.5, y, f"{i+1}. {textwrap.fill(label, 56)}", ha="center", va="center", fontsize=8,
                bbox={"boxstyle": "round,pad=0.45", "facecolor": "#EEF4F7", "edgecolor": "#234A64", "linewidth": 1.1})
        if i < len(labels) - 1:
            ax.annotate("", xy=(.5, y-gap+box_h*.6), xytext=(.5, y-box_h*.6), arrowprops={"arrowstyle": "->", "color": "#56788C", "lw": 1.3})
    ax.text(.5, .025, "Indicative process flow — verify the method and sequence for the actual site.", ha="center", fontsize=7, color="#555555")
    image = io.BytesIO(); fig.savefig(image, format="png", dpi=170, bbox_inches="tight"); plt.close(fig); image.seek(0)
    return image


def workface_sketch_png(activity):
    fig, ax = plt.subplots(figsize=(8.2, 3.3))
    fig.patch.set_facecolor("white"); ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
    ax.text(.5, .98, "INDICATIVE WORKFACE ARRANGEMENT — CONCEPT ONLY", ha="center", va="top", fontsize=10, weight="bold", color="#234A64")
    outer = plt.Rectangle((.035, .16), .93, .69, fill=False, linestyle="--", linewidth=1.4, edgecolor="#607D8B")
    ax.add_patch(outer); ax.text(.05, .81, "CONTROLLED WORK ZONE", fontsize=7, color="#455A64", va="bottom")
    labels = ["Materials /\nplant staging", f"Active workface:\n{textwrap.fill(activity, 18)}", "Inspection /\nrelease", "Protection /\nhandover"]
    xs = [.065, .30, .535, .77]
    for i, (x, label) in enumerate(zip(xs, labels)):
        box = plt.Rectangle((x, .43), .17, .26, facecolor="#EEF4F7" if i != 1 else "#DCEAF1", edgecolor="#234A64", linewidth=1.1)
        ax.add_patch(box); ax.text(x+.085, .56, label, ha="center", va="center", fontsize=7)
        if i < len(xs)-1:
            ax.annotate("", xy=(xs[i+1]-.012, .56), xytext=(x+.182, .56), arrowprops={"arrowstyle": "->", "color": "#56788C", "lw": 1.3})
    ax.text(.5, .27, "Maintain safe access · exclusion controls · housekeeping · emergency route", ha="center", fontsize=7, color="#455A64")
    ax.text(.5, .08, "Not to scale. Site team to plan actual access, laydown, segregation and workface locations.", ha="center", fontsize=7, color="#555555")
    image = io.BytesIO(); fig.savefig(image, format="png", dpi=170, bbox_inches="tight"); plt.close(fig); image.seek(0)
    return image


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
    doc.add_paragraph("Carry out the steps in order unless the site supervisor approves a safe, coordinated alternative. Stop and resolve any failed inspection before covering or proceeding past its release point.")
    for i, step in enumerate(data["sequence"], 1):
        doc.add_heading(f"5.{i} {step['step']}", 2)
        doc.add_paragraph(step["method"])
        gate = doc.add_paragraph()
        gate.add_run("Inspection / release requirement: ").bold = True
        gate.add_run(step["release"])
    doc.add_heading("6. Activity sketches", 1)
    doc.add_paragraph("These conceptual sketches explain sequence and workface organization only. They are not to scale, engineered layouts, or construction drawings. The site team must plan actual access, equipment positions, exclusion areas and interfaces.")
    doc.add_heading("6.1 Work sequence flow", 2)
    doc.add_picture(flow_sketch_png(activity, data["sketch"]), width=Inches(6.5))
    doc.add_heading("6.2 Indicative workface arrangement", 2)
    doc.add_picture(workface_sketch_png(activity), width=Inches(6.5))
    doc.add_heading("7. QA/QC inspection and test plan", 1)
    doc.add_paragraph("H = hold; W = witness; R = review; S = surveillance. Confirm notification and release arrangements before work starts.")
    for i, item in enumerate(data["itp"], 1):
        doc.add_heading(f"7.{i} {item['stage']} — {item['point']} point", 2)
        table(doc, ["ITP field", "Requirement / record"], [["Inspection / test", item["check"]], ["Acceptance criteria", item["acceptance"]], ["Frequency", item["frequency"]], ["Record", item["record"]], ["Responsible", item["owner"]]])
    doc.add_heading("8. Field checklist", 1)
    table(doc, ["No.", "Check", "Status / remarks"], [[i+1, item, "☐ Accept  ☐ Reject  ☐ N/A — ______"] for i, item in enumerate(data["checklist"])])
    doc.add_paragraph("Inspected by: __________________   Date: __________   Signature: __________________")
    doc.add_heading("9. Safety risk assessment", 1)
    doc.add_paragraph("Assess initial and residual risk using the site risk matrix and actual conditions. Apply the listed controls, brief the workforce, and stop work if conditions change or controls fail.")
    for i, item in enumerate(data["risks"], 1):
        doc.add_heading(f"9.{i} {item['hazard']}", 2)
        table(doc, ["Risk assessment field", "Assessment / controls"], [["Potential consequence", item["effect"]], ["Initial risk", item["initial"]], ["Preventive / mitigating controls", item["controls"]], ["Residual risk", item["residual"]], ["Responsible owner", item["owner"]]])
    doc.add_paragraph("Verify risk ratings and controls against site conditions before work starts.")
    doc.add_heading("10. Environmental controls", 1); bullets(doc, data["environment"])
    doc.add_heading("11. Emergency arrangements", 1); bullets(doc, data["emergency"])
    doc.add_heading("12. Assumptions and information gaps", 1); bullets(doc, data["gaps"] or ["Confirm site-specific requirements before work starts."])
    doc.add_heading("13. Review and approval", 1)
    table(doc, ["Prepared by", "Checked by", "Approved by"], [["Name / signature / date", "Name / signature / date", "Name / signature / date"]])
    doc.add_paragraph("This AI-assisted generic document is a draft. Competent construction, engineering, QA/QC and HSE personnel must review it for the actual site conditions, permits, temporary works and applicable requirements before use.")
    stream = io.BytesIO(); doc.save(stream); return stream.getvalue()


st.title("Infrastructure & Buildings Method Statement Builder")
st.caption("Enter an activity name to draft a detailed method statement with sketches and download it in Word format. Powered by Groq.")
st.warning("Every generated document is a draft requiring competent engineering, QA/QC and HSE review before approval or construction use.")
st.info("This version generates generic drafts from the activity name only. It does not use or claim compliance with project specifications or drawings.")
activity = st.text_input("Activity name", placeholder="e.g., Reinforced concrete foundation construction")
api_key = st.secrets.get("GROQ_API_KEY", os.getenv("GROQ_API_KEY", ""))
if st.button("Generate method statement", type="primary", disabled=not activity.strip()):
    if not api_key: st.error("Configure GROQ_API_KEY in Streamlit secrets before generating.")
    else:
        with st.spinner("Generating the detailed activity method statement…"):
            try: st.session_state.generated = (draft(activity.strip(), api_key), activity.strip())
            except Exception as exc: st.error(f"Generation failed: {exc}")
if "generated" in st.session_state:
    result, act = st.session_state.generated
    st.success("Detailed draft generated. Resolve the listed gaps and review all technical content.")
    st.subheader(result["title"])
    st.markdown("**Information gaps to resolve**")
    for gap in result["gaps"]: st.write("• " + gap)
    with st.expander("Preview detailed work sequence", expanded=True):
        for item in result["sequence"]: st.markdown(f"**{item['step']}** — {item['method']}\n\nRelease: {item['release']}")
    sketch_col1, sketch_col2 = st.columns(2)
    with sketch_col1: st.image(flow_sketch_png(act, result["sketch"]), caption="Indicative work sequence flow")
    with sketch_col2: st.image(workface_sketch_png(act), caption="Indicative workface arrangement")
    filename = re.sub(r"[^A-Za-z0-9]+", "_", act).strip("_") or "method_statement"
    st.download_button("Download Word method statement", make_doc(result, act), file_name=f"{filename}_method_statement.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
