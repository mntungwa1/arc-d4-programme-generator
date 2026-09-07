"""ARC D4 Bankability Pathway and Concept Note Generator."""
from __future__ import annotations

import base64
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from supabase import create_client

st.set_page_config(page_title="ARC D4 Programme Generator", page_icon="◆", layout="wide")

DEFAULT_BOOK = Path(__file__).parent / "data" / "ARC_D4_Automation_Matrix.xlsx"
# Supports both the packaged local layout and a simple Streamlit Cloud upload
# where the workbook sits next to app.py.
if not DEFAULT_BOOK.exists():
    DEFAULT_BOOK = Path(__file__).parent / "ARC_D4_Automation_Matrix.xlsx"
EMBEDDED_BOOK = Path(__file__).parent / "ARC_D4_Automation_Matrix.b64"
SUPABASE_URL = "https://wrejrxzgyuxsfbxutezg.supabase.co"
SUPABASE_KEY = "sb_publishable_UfYoG2ZgKP0nLA5KGwEG6w_2rCP9_R8"
APP_URL = "https://arc-d4-programme-generator-c7qdwgesqvnwjafgvpxgat.streamlit.app/"


def shared_client():
    client = create_client(SUPABASE_URL, SUPABASE_KEY)
    session = st.session_state.get("d4_auth_session")
    if session:
        client.auth.set_session(session.access_token, session.refresh_token)
    return client


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
        # Plain-language labels used throughout the interface.
        frame = frame.replace({
            r"(?i)non-blocking": "Non-Mandatory",
            r"(?i)blocking": "Mandatory",
        }, regex=True)
        result[sheet] = frame
    return result


@st.cache_data(show_spinner=False)
def load_default(display_terms_version="mandatory-v1"):
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



D3_CRITERIA = (
    ("BPS", "Measurable impact", 0.20), ("SAS", "Strategic alignment", 0.15),
    ("GRS", "GESI responsiveness", 0.15), ("TFS", "Technical feasibility", 0.10),
    ("IFS", "Institutional feasibility", 0.10), ("CVS", "Cost-effectiveness", 0.10),
    ("RRS", "Scalability and transferability", 0.10), ("SIS", "Sustainability and institutionalisation", 0.10),
)

def c7_description(record):
    values = {str(key).strip(): value for key, value in record.items()}
    return (f"{values.get('Innovation', 'This record')} is a {str(values.get('Type', 'classified')).lower()} innovation in {values.get('Pillar', 'the portfolio')}, "
            f"positioned at {values.get('Value chain', 'the relevant position')} on the DRM value chain. It is classified as "
            f"{str(values.get('DRM cycle', 'cross-cutting')).lower()}, {str(values.get('Maturity', 'unconfirmed')).lower()}, "
            f"with {str(values.get('Hazard reach', 'unconfirmed')).lower()} reach.")

def score_d3_profile(profile):
    transfers = max(1, int(profile.get("bps_transfers", 1)))
    bps = max(0.0, 10 - 2 * (transfers - 1) - (1 if profile.get("bps_latency") == "Yes" else 0))
    sas = 5 * (min(4, int(profile.get("sas_sendai", 0))) / 4) + 5 * (min(3, int(profile.get("sas_priorities", 0))) / 3)
    grs = (10 / 7) * sum(float(profile.get(f"gesi_{n}", 0)) for n in range(1, 8))
    tfs = 10 * ((float(profile.get("tfs_absorptive", 0)) + float(profile.get("tfs_timeline", 0))) / 2)
    ifs = 10 * float(profile.get("ifs_dependency", 0))
    cvs = float(profile.get("cvs_score", 0))
    rrs = min(10.0, (10 / 16) * (float(profile.get("rrs_operational", 0)) + 0.5 * float(profile.get("rrs_pipeline", 0))))
    sis = 10 * ((float(profile.get("sis_budget", 0)) + float(profile.get("sis_custodian", 0)) + float(profile.get("sis_recurrent", 0))) / 3)
    scores = {"BPS": bps, "SAS": sas, "GRS": grs, "TFS": tfs, "IFS": ifs, "CVS": cvs, "RRS": rrs, "SIS": sis}
    ipi = sum(scores[key] * weight for key, _, weight in D3_CRITERIA)
    band = "5 — Very high" if ipi > 8 else "4 — High" if ipi > 6 else "3 — Moderate" if ipi > 4 else "2 — Low" if ipi > 2 else "1 — Very low"
    return scores, round(ipi, 2), round(ipi * 10, 1), band, ("Institutionalisation track" if min(tfs, ifs) == 0 else "Operational track")

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
    workstream = st.radio("Workstream", ["C1 — Create a new innovation", "C2 — Portfolio selection and investment readiness"])
    st.divider()
    st.subheader("Shared innovation register")
    if "d4_auth_session" not in st.session_state:
        auth_email = st.text_input("Email", key="auth_email")
        auth_password = st.text_input("Password", type="password", key="auth_password")
        sign_in, sign_up = st.columns(2)
        if sign_in.button("Sign in") and auth_email and auth_password:
            try:
                result = shared_client().auth.sign_in_with_password({"email": auth_email, "password": auth_password})
                st.session_state.d4_auth_session = result.session
                st.rerun()
            except Exception as exc:
                st.error(f"Sign-in failed: {exc}")
        if sign_up.button("Create account") and auth_email and auth_password:
            try:
                shared_client().auth.sign_up({
                    "email": auth_email,
                    "password": auth_password,
                    "options": {"email_redirect_to": APP_URL},
                })
                st.success("Account created. Confirm the email, then return here and sign in.")
            except Exception as exc:
                st.error(f"Account creation failed: {exc}")
        if st.button("Resend confirmation email") and auth_email:
            try:
                shared_client().auth.resend({
                    "type": "signup",
                    "email": auth_email,
                    "options": {"email_redirect_to": APP_URL},
                })
                st.success("A new confirmation email has been sent to the live app address.")
            except Exception as exc:
                st.error(f"Could not resend the confirmation email: {exc}")
    else:
        if st.button("Sign out"):
            del st.session_state.d4_auth_session
            st.rerun()

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
shared_innovations = []
if "d4_auth_session" in st.session_state:
    try:
        shared_innovations = shared_client().table("d4_innovations").select("innovation_name,innovation_url,innovation_type,d3_profile").order("innovation_name").execute().data or []
    except Exception as exc:
        st.sidebar.warning(f"Shared register unavailable: {exc}")
shared_names = [item["innovation_name"] for item in shared_innovations]
innovation = st.sidebar.selectbox("Innovation to work on", portfolio["Innovation"].tolist() + shared_names)
is_shared_innovation = innovation in shared_names
row = portfolio.loc[portfolio["Innovation"] == innovation].iloc[0] if not is_shared_innovation else None
types = matrix.get("20_Innovation_Type", pd.DataFrame())
match = types.loc[types.get("Innovation", pd.Series(dtype=str)).astype(str) == innovation] if not types.empty else pd.DataFrame()
default_type = (next(item["innovation_type"] for item in shared_innovations if item["innovation_name"] == innovation)
                if is_shared_innovation else (match.iloc[0]["Type"] if not match.empty else ""))

shared_record = next((item for item in shared_innovations if item["innovation_name"] == innovation), {})
saved_d3_profile = shared_record.get("d3_profile") or {} if is_shared_innovation else {}
base_profile = {
    "innovation_id": f"D4-{len(shared_names):03d}" if is_shared_innovation else f"INV-{int(row['#']):03d}",
    "innovation_name": innovation, "innovation_type": default_type, "delivery_counterpart": "",
    "basis_risk_treatment": "", "premium_transition_pathway": "",
    "recurrent_cost_custodian": "", "short_description": "", "parametric_trigger": "No",
    "subsidised_cost": "No", "template": "TPL-EU",
}
profile = st.session_state.profile.setdefault(innovation, {**base_profile, **saved_d3_profile})
profile["innovation_name"] = innovation

tab_labels = (["1. Pathway", "2. New Innovation Profile", "3. Evidence & research", "4. Portfolio", "5. Output & audit"]
              if workstream.startswith("C1") else ["1. Pathway"])
tabs = st.tabs(tab_labels)

with tabs[0]:
    if workstream.startswith("C1"):
        st.info("**Workstream C1** — create and admit a new innovation. Complete the New Innovation Profile, then use S1 and S2 to add it to the shared Innovation to work on list.")
    else:
        st.info("**Workstream C2** — select an admitted innovation, review S1 and S2, then progress it through Workstream D.")
    pathway_view, selected_view = st.columns([3, 1], gap="large")
    stages = matrix["01_Stages"].fillna("—")
    permitted_stages = ["S1", "S2"] if workstream.startswith("C1") else stages["Stage ID"].tolist()
    if st.session_state.selected_stage not in permitted_stages:
        st.session_state.selected_stage = permitted_stages[0]
    visible_stages = stages[stages["Stage ID"].isin(permitted_stages)]
    with pathway_view:
        st.subheader("The required sequence")
        lanes = visible_stages.groupby("Lane", sort=False)
        for lane, items in lanes:
            with st.expander(str(lane), expanded=True):
                for _, s in items.iterrows():
                    left, right = st.columns([1, 5])
                    if left.button(str(s["Stage ID"]), key=f"stage_{s['Stage ID']}"):
                        st.session_state.selected_stage = s["Stage ID"]
                    right.markdown(f"**{s['Stage ID']} — {s['Stage name']}**  ")
                    right.caption(f"Entry: {s['Entry condition']}  |  Exit: {s['Exit condition']}  |  {s['Automation level']}")
    with selected_view:
        selected = stages.loc[stages["Stage ID"] == st.session_state.selected_stage].iloc[0]
        st.subheader("Selected stage")
        st.markdown(f"**{selected['Stage ID']} — {selected['Stage name']}**")
        st.caption(f"Gate: {selected['Decision gate']}")
        st.markdown("**Entry condition**")
        st.write(selected["Entry condition"])
        st.markdown("**Exit condition**")
        st.write(selected["Exit condition"])
        st.markdown("**Accountable**")
        st.write(selected["Accountable"])
        if profile.get("matrix_ipi") is not None:
            st.markdown("**Annex H matrix summary**")
            st.caption(f"IPI {float(profile['matrix_ipi']):.2f}/10 · {profile.get('matrix_band', 'Band pending')} · {profile.get('matrix_track', 'Gate pending')} · Evidence: {profile.get('evidence_confidence', 'pending')}")
            if profile.get("short_description"):
                st.caption(profile["short_description"])
        steps = matrix["02_Steps"]
        stage_steps = steps.loc[steps["Stage ID"] == selected["Stage ID"]]
        if len(stage_steps):
            st.markdown("**Required actions**")
            for _, step in stage_steps.iterrows():
                st.caption(f"{step['Step ID']} — {step['Step description']}")
        if selected["Stage ID"] == "S1":
            st.markdown("**S1 portfolio records**")
            s1_records = portfolio.copy()
            score_column = "IPI v2.0 (computed)"
            s1_records["Record status"] = s1_records[score_column].apply(
                lambda value: "Pending" if str(value).strip().upper() == "PENDING" else "Scored"
            )
            st.dataframe(s1_records[["#", "Innovation", score_column, "Record status"]],
                         use_container_width=True, hide_index=True, height=420)
        elif selected["Stage ID"] == "S2":
            st.markdown("**S2 selected Top 10**")
            top_ten = portfolio.sort_values("#").head(10)
            st.caption("The current workbook records these as the first ten D3 portfolio selections.")
            st.dataframe(top_ten[["#", "Innovation", "IPI v1 / Tier (D3 ref)", "Confidence"]],
                         use_container_width=True, hide_index=True, height=340)

if workstream.startswith("C1"):
    with tabs[1]:
        st.subheader("Investment-Ready Innovation Profile")
        if workstream.startswith("C1"):
            st.info("Workstream C1 creates a new innovation. Once admitted, it is added permanently to the shared list at left and can proceed through Workstream C2.")
        st.caption("Complete fields as evidence becomes available. Blank required fields remain visible as gaps; the app will not invent text.")
        classification = matrix["20_Innovation_Type"].copy()
        c7_row = classification.loc[classification["Innovation"].astype(str) == innovation]
        c7_default = c7_description(c7_row.iloc[0]) if len(c7_row) else ""
        profile["short_description"] = st.text_area("Innovation description (D3 Annex B.2 / Table C.7)", profile.get("short_description") or c7_default, help="Concise description of what the innovation does. The pre-filled statement is generated from the governed Annex C.7 classification.")
        profile["new_innovation_name"] = st.text_input(
            "New Innovation Name",
            profile.get("new_innovation_name", profile.get("innovation_name", "")),
            help="Enter the name to use for this new or adapted innovation profile.",
        )
        profile["innovation_url"] = st.text_input(
            "Innovation URL",
            profile.get("innovation_url", ""),
            placeholder="https://example.org/innovation",
            help="Add the official web page, source record, or supporting information link for the innovation.",
        )
        if "d4_auth_session" not in st.session_state:
            st.warning("Sign in through the Shared innovation register in the sidebar before admitting the completed D3 innovation.")
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

        st.divider()
        st.subheader("Annex H — D3 Prioritisation Matrix")
        st.caption("Eight weighted criteria. Evidence confidence is reported separately and does not change the IPI.")
        with st.expander("Enter criterion determinations and calculate IPI", expanded=False):
            a1, a2 = st.columns(2)
            with a1:
                profile["bps_transfers"] = st.number_input("BPS: institutional transfers to household", 1, 6, int(profile.get("bps_transfers", 1)), key=f"{innovation}_bps_transfers")
                profile["bps_latency"] = st.selectbox("BPS: activation effect exceeds seven days?", ["No", "Yes"], index=1 if profile.get("bps_latency") == "Yes" else 0, key=f"{innovation}_bps_latency")
                profile["tfs_absorptive"] = st.selectbox("TFS: absorptive capacity", [0.0, 0.5, 1.0], key=f"{innovation}_tfs_a")
                profile["tfs_timeline"] = st.selectbox("TFS: timeline fit", [0.0, 0.5, 1.0], key=f"{innovation}_tfs_t")
                profile["ifs_dependency"] = st.selectbox("IFS: dependency resolution", [0.0, 0.5, 1.0], key=f"{innovation}_ifs")
                profile["cvs_score"] = st.number_input("CVS: value-for-money score", 0.0, 10.0, float(profile.get("cvs_score", 0.0)), 0.5, key=f"{innovation}_cvs")
                profile["cvs_basis"] = st.text_input("CVS: stated appraisal basis", profile.get("cvs_basis", ""), key=f"{innovation}_cvs_basis")
            with a2:
                profile["sas_sendai"] = st.number_input("SAS: evidenced Sendai targets", 0, 4, int(profile.get("sas_sendai", 0)), key=f"{innovation}_sas_sendai")
                profile["sas_priorities"] = st.number_input("SAS: evidenced SADC DRM priorities", 0, 3, int(profile.get("sas_priorities", 0)), key=f"{innovation}_sas_priorities")
                profile["rrs_operational"] = st.number_input("RRS: operational Member States", 0, 16, int(profile.get("rrs_operational", 0)), key=f"{innovation}_rrs_operational")
                profile["rrs_pipeline"] = st.number_input("RRS: planned/pilot Member States", 0, 16, int(profile.get("rrs_pipeline", 0)), key=f"{innovation}_rrs_pipeline")
                profile["sis_budget"] = st.selectbox("SIS: domestic-budget pathway", [0.0, 0.5, 1.0], key=f"{innovation}_sis_budget")
                profile["sis_custodian"] = st.selectbox("SIS: named institutional custodian", [0.0, 0.5, 1.0], key=f"{innovation}_sis_custodian")
                profile["sis_recurrent"] = st.selectbox("SIS: recurrent-cost profile", [0.0, 0.5, 1.0], key=f"{innovation}_sis_recurrent")
                profile["evidence_confidence"] = st.selectbox("Evidence confidence", ["High", "Moderate", "Low"], index=["High", "Moderate", "Low"].index(profile.get("evidence_confidence", "Moderate")), key=f"{innovation}_evidence_confidence")
            st.markdown("**GRS — seven inclusion determinations:** 1 = designed; 0.5 = asserted/incidental; 0 = absent.")
            categories = ["Women and girls", "Children and youth", "Older persons", "Persons with disabilities", "Migrants and displaced persons", "Remote communities", "Indigenous and knowledge-holding communities"]
            cols = st.columns(4)
            for n, label in enumerate(categories, 1):
                with cols[(n - 1) % 4]:
                    profile[f"gesi_{n}"] = st.selectbox(label, [0.0, 0.5, 1.0], key=f"{innovation}_gesi_{n}")
            scores, ipi, ipi100, band, track = score_d3_profile(profile)
            profile.update({f"score_{key.lower()}": round(value, 2) for key, value in scores.items()})
            profile.update({"matrix_ipi": ipi, "matrix_ipi_100": ipi100, "matrix_band": band, "matrix_track": track})
            st.dataframe(pd.DataFrame([{"Criterion": key, "Weight": weight, "Score (0–10)": round(scores[key], 2)} for key, _, weight in D3_CRITERIA]), use_container_width=True, hide_index=True)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("IPI", f"{ipi:.2f} / 10")
            m2.metric("IPI₁₀₀", f"{ipi100:.1f}")
            m3.metric("Band", band.split(" — ")[0])
            m4.metric("Gate", track)

        st.divider()
        st.subheader("Admit completed D3 innovation")
        st.caption("This saves the complete C1/D3 profile with the innovation in the protected shared register. It becomes available in Workstream C2.")
        if "d4_auth_session" in st.session_state and st.button("Add innovation and D3 profile to the shared list"):
            name = profile["new_innovation_name"].strip()
            if len(name) < 3:
                st.error("Enter a New Innovation Name of at least three characters.")
            else:
                try:
                    user = shared_client().auth.get_user().user
                    saved_profile = {key: value for key, value in profile.items()}
                    saved_profile["innovation_name"] = name
                    saved_profile["innovation_url"] = profile["innovation_url"].strip()
                    shared_client().table("d4_innovations").insert({
                        "innovation_name": name,
                        "innovation_url": profile["innovation_url"].strip() or None,
                        "innovation_type": profile.get("innovation_type") or None,
                        "d3_profile": saved_profile,
                        "created_by": user.id,
                    }).execute()
                    st.success("Innovation and its D3 profile are now in the shared list. Select it in the sidebar and change to Workstream C2.")
                except Exception as exc:
                    st.error(f"Could not add the innovation: {exc}")
    
    with tabs[2]:
        st.subheader("Evidence sufficiency and research break-out")
        dictionary = matrix["04_Field_Dictionary"]
        required = dictionary[dictionary["Required by"].astype(str).str.contains(profile["template"].replace("TPL-", ""), case=False, na=False)]
        st.caption("Fields below are governed by the workbook. Use a research request only where a permitted evidence gap cannot be filled from the D3 record.")
        st.dataframe(required[["Field ID", "Label", "Source", "Validation", "Research break-out permitted"]], use_container_width=True, hide_index=True)
        st.markdown("#### Available research tasks")
        st.dataframe(matrix["10_Research_Broker"], use_container_width=True, hide_index=True)
    
    with tabs[3]:
        st.subheader("Portfolio context — D3 Annex C.7")
        classification = matrix["20_Innovation_Type"].copy()
        classification["Innovation description"] = classification.apply(c7_description, axis=1)
        st.dataframe(classification, use_container_width=True, hide_index=True)
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
    
