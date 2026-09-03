"""ARC D4 Bankability Pathway and Concept Note Generator."""
from __future__ import annotations

import base64
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="ARC D4 Programme Generator", page_icon="◆", layout="wide")

DEFAULT_BOOK = Path(__file__).parent / "data" / "ARC_D4_Automation_Matrix.xlsx"
# Supports both the packaged local layout and a simple Streamlit Cloud upload
# where the workbook sits next to app.py.
if not DEFAULT_BOOK.exists():
    DEFAULT_BOOK = Path(__file__).parent / "ARC_D4_Automation_Matrix.xlsx"
EMBEDDED_BOOK = Path(__file__).parent / "ARC_D4_Automation_Matrix.b64"


def read_matrix(file) -> dict[str, pd.DataFrame]:
    """Read every sheet and find its true header row (the row with Stage ID, Field ID etc.)."""
    book = pd.ExcelFile(file)
    result = {}
    for sheet in book.sheet_names:
        raw = pd.read_excel(book, sheet_name=sheet, header=None)
        header_row = 0
        for idx, row in raw.iterrows():
            values = {str(v).strip() for v in row.dropna().tolist()}
            if values & {"Stage ID", "Step ID", "Rule ID", "Field ID", "#", "Template ID", "Check ID", "Set ID"}:
                header_row = idx
                break
        frame = pd.read_excel(book, sheet_name=sheet, header=header_row)
        frame = frame.dropna(how="all").dropna(axis=1, how="all")
        result[sheet] = frame
    return result


@st.cache_data(show_spinner=False)
def load_default():
    if DEFAULT_BOOK.exists():
        return read_matrix(DEFAULT_BOOK)
    return read_matrix(BytesIO(base64.b64decode(EMBEDDED_BOOK.read_text())))


def nonblank(value):
    return value is not None and str(value).strip() not in {"", "nan", "None"}


def status(value):
    return "✓ Complete" if nonblank(value) else "○ Not yet supplied"


def audit(innovation, field, value, decision="COMMIT"):
    st.session_state.audit.append({
        "Entry ID": f"AUD-{len(st.session_state.audit)+1:04}", "Innovation": innovation,
        "Field": field, "Value committed": value, "Decision": decision,
        "Timestamp (UTC)": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    })


def make_export(matrix, profile, audits):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame([profile]).T.rename(columns={0: "Value"}).to_excel(writer, sheet_name="Innovation_Profile")
        pd.DataFrame(audits).to_excel(writer, sheet_name="Audit_Log", index=False)
        matrix["01_Stages"].to_excel(writer, sheet_name="Pathway", index=False)
        matrix["12_Validation_Rules"].to_excel(writer, sheet_name="Validation_Rules", index=False)
    return output.getvalue()


st.title("ARC D4 — Programme Generator")
st.caption("A guided implementation of the D4 bankability pathway. It produces a controlled working record, not a submission-ready proposal.")

with st.sidebar:
    st.header("Programme workbook")
    upload = st.file_uploader("Use an updated automation matrix", type="xlsx")
    st.caption("If no workbook is uploaded, the supplied ARC D4 Automation Matrix is used.")

try:
    if upload:
        matrix = read_matrix(upload)
    elif DEFAULT_BOOK.exists() or EMBEDDED_BOOK.exists():
        matrix = load_default()
    else:
        st.warning("Upload the ARC D4 Automation Matrix (.xlsx) in the sidebar to start the programme workflow.")
        st.info("The hosted app deliberately waits for the governed workbook rather than generating programme content without its source data.")
        st.stop()
except Exception as exc:
    st.error(f"The workbook could not be read: {exc}")
    st.stop()

for key, default in {"profile": {}, "audit": [], "selected_stage": "S1"}.items():
    if key not in st.session_state:
        st.session_state[key] = default

portfolio = matrix["06_Portfolio"].copy()
portfolio["Innovation"] = portfolio["Innovation"].astype(str)
innovation = st.sidebar.selectbox("Innovation to work on", portfolio["Innovation"].tolist())
row = portfolio.loc[portfolio["Innovation"] == innovation].iloc[0]
types = matrix.get("20_Innovation_Type", pd.DataFrame())
match = types.loc[types.get("Innovation", pd.Series(dtype=str)).astype(str) == innovation] if not types.empty else pd.DataFrame()
default_type = match.iloc[0]["Type"] if not match.empty else ""

profile = st.session_state.profile.setdefault(innovation, {
    "innovation_id": f"INV-{int(row['#']):03d}", "innovation_name": innovation,
    "innovation_type": default_type, "delivery_counterpart": "", "basis_risk_treatment": "",
    "premium_transition_pathway": "", "recurrent_cost_custodian": "", "parametric_trigger": "No",
    "subsidised_cost": "No", "template": "TPL-EU"
})

tabs = st.tabs(["1. Pathway", "2. Innovation profile", "3. Evidence & research", "4. Portfolio", "5. Output & audit"])

with tabs[0]:
    stages = matrix["01_Stages"].fillna("—")
    st.subheader("The required sequence")
    lanes = stages.groupby("Lane", sort=False)
    for lane, items in lanes:
        with st.expander(str(lane), expanded=True):
            for _, s in items.iterrows():
                left, right = st.columns([1, 5])
                if left.button(str(s["Stage ID"]), key=f"stage_{s['Stage ID']}"):
                    st.session_state.selected_stage = s["Stage ID"]
                right.markdown(f"**{s['Stage ID']} — {s['Stage name']}**  ")
                right.caption(f"Entry: {s['Entry condition']}  |  Exit: {s['Exit condition']}  |  {s['Automation level']}")
    selected = stages.loc[stages["Stage ID"] == st.session_state.selected_stage].iloc[0]
    st.info(f"Selected: {selected['Stage ID']} — {selected['Stage name']}. Gate: {selected['Decision gate']}")
    steps = matrix["02_Steps"]
    stage_steps = steps.loc[steps["Stage ID"] == selected["Stage ID"]]
    if len(stage_steps): st.dataframe(stage_steps, use_container_width=True, hide_index=True)
    else: st.caption("This stage is governed by the entry and exit conditions above; no separate executable step is listed.")

with tabs[1]:
    st.subheader("Investment-Ready Innovation Profile")
    st.caption("Complete fields as evidence becomes available. Blank required fields remain visible as gaps; the app will not invent text.")
    a, b = st.columns(2)
    with a:
        profile["innovation_type"] = st.selectbox("Innovation type", ["", "Tech", "Non-tech", "Hybrid"], index=["", "Tech", "Non-tech", "Hybrid"].index(profile.get("innovation_type", "") if profile.get("innovation_type", "") in ["", "Tech", "Non-tech", "Hybrid"] else ""))
        profile["delivery_counterpart"] = st.text_input("Named delivery counterpart *", profile.get("delivery_counterpart", ""))
        profile["recurrent_cost_custodian"] = st.text_input("Named recurrent-cost custodian *", profile.get("recurrent_cost_custodian", ""))
    with b:
        profile["parametric_trigger"] = st.selectbox("Does the concept use a parametric trigger?", ["No", "Yes"], index=1 if profile.get("parametric_trigger") == "Yes" else 0)
        profile["subsidised_cost"] = st.selectbox("Does any cost start subsidised?", ["No", "Yes"], index=1 if profile.get("subsidised_cost") == "Yes" else 0)
        profile["template"] = st.selectbox("Preferred financing template", matrix["09_Template_Map"]["Template ID"].tolist(), index=max(0, matrix["09_Template_Map"]["Template ID"].tolist().index(profile.get("template", "TPL-EU"))))
    if profile["parametric_trigger"] == "Yes":
        profile["basis_risk_treatment"] = st.text_area("Basis-risk treatment *", profile.get("basis_risk_treatment", ""))
    if profile["subsidised_cost"] == "Yes":
        profile["premium_transition_pathway"] = st.text_area("Premium-transition pathway *", profile.get("premium_transition_pathway", ""))
    st.divider()
    components = matrix["07_IRIP_Components"]
    for _, c in components.iterrows():
        key = f"component_{c['#']}"
        profile[key] = st.text_area(f"{c['#']}. {c['Component']} — {c['Status if absent']}", profile.get(key, ""), key=f"{innovation}_{key}")

with tabs[2]:
    st.subheader("Evidence sufficiency and research break-out")
    dictionary = matrix["04_Field_Dictionary"]
    required = dictionary[dictionary["Required by"].astype(str).str.contains(profile["template"].replace("TPL-", ""), case=False, na=False)]
    st.caption("Fields below are governed by the workbook. Use a research request only where a permitted evidence gap cannot be filled from the D3 record.")
    st.dataframe(required[["Field ID", "Label", "Source", "Validation", "Research break-out permitted"]], use_container_width=True, hide_index=True)
    st.markdown("#### Available research tasks")
    st.dataframe(matrix["10_Research_Broker"], use_container_width=True, hide_index=True)

with tabs[3]:
    st.subheader("Portfolio context")
    st.dataframe(portfolio, use_container_width=True, hide_index=True)
    st.markdown("#### Gap linkages for the selected innovation")
    links = matrix["17_Gap_Linkages"]
    # Excel sometimes preserves an invisible trailing space in this heading.
    # Locate it defensively so the workflow does not fail because of formatting.
    linkage_name_column = next(
        (column for column in links.columns
         if "intervention" in str(column).strip().lower()
         and "innovation" in str(column).strip().lower()),
        None,
    )
    selected_links = (
        links[links[linkage_name_column].astype(str).str.contains(
            innovation.split("+")[0].strip(), case=False, na=False)]
        if linkage_name_column else pd.DataFrame()
    )
    if len(selected_links): st.dataframe(selected_links, use_container_width=True, hide_index=True)
    else: st.warning("No direct linkage is found by name. Record the relevant governed gap linkage before admission.")
    st.markdown("#### Five-pillar response")
    pillar = matrix["08_Pillar_Response"]
    st.dataframe(pillar[pillar["Innovation"] == innovation], use_container_width=True, hide_index=True)

with tabs[4]:
    st.subheader("Generation gate and controlled export")
    failures = []
    if not nonblank(profile.get("delivery_counterpart")): failures.append("V02 — a delivery counterpart must be named.")
    if not nonblank(profile.get("recurrent_cost_custodian")): failures.append("V05 — a recurrent-cost custodian must be named.")
    if profile.get("parametric_trigger") == "Yes" and not nonblank(profile.get("basis_risk_treatment")): failures.append("V03 — basis-risk treatment is required for a parametric trigger.")
    if profile.get("subsidised_cost") == "Yes" and not nonblank(profile.get("premium_transition_pathway")): failures.append("V04 — premium-transition pathway is required for subsidised costs.")
    absent = [str(c["#"]) for _, c in matrix["07_IRIP_Components"].iterrows() if not nonblank(profile.get(f"component_{c['#']}"))]
    if absent: failures.append("V01 — required IRIP components not yet addressed: " + ", ".join(absent))
    if failures:
        st.error("Generation is blocked. Resolve the following conditions:")
        for failure in failures: st.write("• " + failure)
    else:
        st.success("All current blocking checks pass. The innovation can proceed to the concept-note generation stage.")
        if st.button("Commit profile to audit log"):
            for field, value in profile.items():
                if nonblank(value): audit(innovation, field, value)
            st.success("Profile values committed to the in-session audit trail.")
    st.markdown("#### Audit trail")
    if st.session_state.audit: st.dataframe(pd.DataFrame(st.session_state.audit), use_container_width=True, hide_index=True)
    else: st.caption("No profile values committed yet.")
    st.download_button("Download controlled working record (.xlsx)", make_export(matrix, profile, st.session_state.audit), f"ARC_D4_{profile['innovation_id']}_working_record.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
