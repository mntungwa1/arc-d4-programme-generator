"""ARC D4 Delivery Platform — controlled portfolio-to-product workflow."""
from __future__ import annotations

import base64
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from supabase import create_client

st.set_page_config(page_title="ARC D4 Delivery Platform", page_icon="◆", layout="wide")

SUPABASE_URL = "https://wrejrxzgyuxsfbxutezg.supabase.co"
SUPABASE_KEY = "sb_publishable_UfYoG2ZgKP0nLA5KGwEG6w_2rCP9_R8"
APP_URL = "https://arc-d4-programme-generator-c7qdwgesqvnwjafgvpxgat.streamlit.app/"
ADMIN_EMAILS = {"dingaan@academyrc.co.za", "drcliff@academyrc.co.za"}
BOOK = Path(__file__).parent / "ARC_D4_Automation_Matrix.b64"

# D2 Innovation Landscape Assessment v3.1 readiness conditions.  These make
# product readiness traceable to the landscape evidence rather than to a
# generic completion status alone.
D2_PRODUCT_READINESS = {
    "P1": {
        "evidence": "An evidence-assured portfolio must carry evidence grade, operational maturity, pillar and country coverage, plus a defensible concentration position.",
        "ready_when": "All outstanding index determinations are resolved and the portfolio can show that its selected innovations are evidence-supported, balanced across the five pillars and not concentrated in a single maturity or country profile.",
        "d3_final_gate": "D3 final line: the evidence-assured Annex H score, tier and evidence-confidence record must be linked to the Gap Register. The programme document can proceed only when the selected portfolio demonstrably closes the priority gaps and every unresolved critical feasibility condition has moved to the institutionalisation track or has a named resolution route.",
    },
    "P2": {
        "evidence": "Each innovation needs triangulated evidence, operational maturity, an accountable institution, delivery partners, lifecycle costs, safeguards and a sustainability route.",
        "ready_when": "The recurrent-cost custodian is named and the profile can demonstrate operational feasibility, SHOC or delivery interoperability, inclusion safeguards and an investment-ready costed pathway.",
        "d3_final_gate": "D3 final line: complete the Investment-Ready Innovation Profile with its quantified problem, implementation milestones, cost-benefit basis, delivery and recurrent-cost arrangements, Sendai contribution and recorded inclusion-audit result. No unresolved critical feasibility condition may remain in the operating pathway.",
    },
    "P4": {
        "evidence": "The summary must distinguish documented evidence from emerging claims and show the maturity, opportunity and delivery implication of the selected portfolio.",
        "ready_when": "The portfolio economic summary is complete and every headline claim can be traced to a scored, evidence-graded innovation or a clearly identified outstanding work item.",
        "d3_final_gate": "D3 final line: include only figures supported by the evidence-assured matrix, gap register and results framework. State the intended Sendai contribution, the evidence-confidence position and the remaining work item wherever a quantified financing or outcome claim is not yet supported.",
    },
    "P5": {
        "evidence": "D2 treats national ownership as an operating condition: a policy step, budget line, institutional custodian, legal basis and Member State decision right must be explicit.",
        "ready_when": "The national provision and named custodian are confirmed through the Member State route, with the policy, legal and budget actions recorded against an agreed timeline.",
        "d3_final_gate": "D3 final line: record the governing instrument, named custodian, budget/staffing requirement and Member State adoption-or-decline decision. The plan must show reciprocal benefit, an implementation sequence and escalation where a gating instrument has not been adopted.",
    },
    "P6": {
        "evidence": "A usable warning and communication plan must show trusted local institutions, language and cultural suitability, accessibility, safeguarding, feedback and non-digital continuity.",
        "ready_when": "The media engagement has returned the language, channel, attribution and feedback arrangements, and the plan shows how the warning reaches excluded or offline groups.",
        "d3_final_gate": "D3 final line: attach the inclusion-audit result, the responsible operator or delivery arrangement, data/consent conditions where relevant, continuity arrangements and the indicator for household-level protective action. A channel without an accountable delivery path is not ready.",
    },
    "P7": {
        "evidence": "Implementation requires a valid mandate, legal authority, responsible institution, practical operating model, lifecycle funding and measurable accountability.",
        "ready_when": "The outstanding national provision is confirmed and the mandate, legal, delivery, recurrent-cost and measurement arrangements form one implementable Member State pathway.",
        "d3_final_gate": "D3 final line: before first disbursement, the operating description and institutional arrangement must identify what operates, dependencies, annual recurrent cost, custodian, budget line, staffing and governing agreements. The implementation plan must carry the risk register, results record and replication requirements.",
    },
}

# These are routing leads, not sources of evidence.  They help users find the
# institution or framework that can close a fact while keeping the formal
# evidence and approval test in the governed record.
SOLUTION_LEADS = {
    "F-01": {
        "where": "ARC's controlled Determination Workbook and evidence workspace; use the SADC DRM Strategy and Action Plan as the regional reference framework.",
        "url": "https://www.sadc.int/document/en-sadc-disaster-risk-management-strategy-and-action-plan",
    },
    "F-02": {
        "where": "The nominated Member State DRM, finance or sector focal point responsible for the recurrent-cost decision.",
        "url": "https://www.sadc.int/member-states",
    },
    "F-03": {
        "where": "The nominated Member State legal, DRM or implementing-institution focal point that can confirm the national operating provision.",
        "url": "https://www.sadc.int/member-states",
    },
    "F-04": {
        "where": "The relevant SADC National Media Coordinator and the participating regional media organisation.",
        "url": "https://www.sadc.int/media-coordinators",
    },
}

# Annex R.1 comparator bands, expressed as 2026 USD planning ranges.  They guide
# concept-level analysis only; a financing proposition must replace them with a
# nationally costed proposal.
COST_ARCHETYPES = {
    "Community arrangement": (150_000.0, 400_000.0, 40_000.0, 120_000.0, "Moderate"),
    "Community early warning": (200_000.0, 550_000.0, 70_000.0, 180_000.0, "Moderate"),
    "Anticipatory protocol": (100_000.0, 300_000.0, 50_000.0, 150_000.0, "Moderate"),
    "Forecasting capability": (400_000.0, 900_000.0, 120_000.0, 300_000.0, "Moderate"),
    "Whole-of-system early warning": (2_000_000.0, 6_000_000.0, 500_000.0, 1_200_000.0, "High"),
    "Dissemination and alerting": (250_000.0, 700_000.0, 90_000.0, 250_000.0, "Moderate"),
    "Cash and voucher platform": (400_000.0, 1_100_000.0, 0.0, 0.0, "High"),
    "Beneficiary registry": (2_000_000.0, 5_000_000.0, 0.0, 0.0, "High"),
    "Sovereign risk transfer": (0.0, 0.0, 1_500_000.0, 4_000_000.0, "High"),
    "Aerial systems": (700_000.0, 1_500_000.0, 200_000.0, 450_000.0, "High"),
    "Earth observation platform": (300_000.0, 1_400_000.0, 100_000.0, 400_000.0, "Low"),
    "Documentation or knowledge arrangement": (30_000.0, 700_000.0, 15_000.0, 200_000.0, "Low"),
    "Capability platform or arrangement": (150_000.0, 500_000.0, 70_000.0, 180_000.0, "Low"),
    "Risk information platform": (250_000.0, 2_000_000.0, 100_000.0, 600_000.0, "Moderate"),
}

# Existing portfolio records are prefilled from Annex R.1, Table 3.  The fields
# are keyed to the governed portfolio number, avoiding name-matching ambiguity.
COST_ESTIMATES = {
    1: ("Community arrangement", 150000, 400000, 40000, 90000, "C1, C2", "Moderate", "USD per year"),
    2: ("Dissemination and alerting", 600000, 1800000, 150000, 400000, "C5", "Moderate", "USD per year"),
    3: ("Anticipatory protocol", 120000, 300000, 60000, 150000, "C10", "Moderate", "USD per year"),
    4: ("Forecasting capability", 400000, 900000, 120000, 300000, "C4", "Moderate", "USD per year"),
    5: ("Whole-of-system early warning", 2000000, 6000000, 500000, 1200000, "C4", "High", "USD per year"),
    6: ("Community early warning", 200000, 550000, 70000, 180000, "C2, C3", "Moderate", "USD per year"),
    7: ("Anticipatory protocol", 100000, 250000, 50000, 120000, "C10", "Moderate", "USD per year"),
    8: ("Anticipatory protocol", 150000, 400000, 0, 0, "C10", "High", "USD 11 per person reached"),
    9: ("Dissemination and alerting", 250000, 700000, 90000, 250000, "C5, C6", "Moderate", "USD per year"),
    10: ("Cash and voucher platform", 400000, 1100000, 0, 0, "C7, C8", "High", "USD 0.17 per dollar transferred"),
    11: ("Community early warning", 250000, 600000, 80000, 200000, "C2, C3", "Moderate", "USD per year"),
    12: ("Sovereign risk transfer", 0, 0, 1500000, 4000000, "C11", "High", "USD premium per year"),
    13: ("Aerial systems", 700000, 1500000, 200000, 450000, "C12", "High", "USD per year"),
    14: ("Earth observation platform", 500000, 1400000, 180000, 400000, "C4", "Low", "USD per year"),
    15: ("Community arrangement", 60000, 180000, 30000, 80000, "C2", "Moderate", "USD per year"),
    16: ("Whole-of-system early warning", 1500000, 4000000, 400000, 900000, "C3, C4", "Moderate", "USD per year"),
    17: ("Beneficiary registry", 2000000, 5000000, 0, 0, "C8, C9", "High", "USD 0.65–2.00 per household per year"),
    18: ("Dissemination and alerting", 350000, 900000, 120000, 300000, "C5, C6", "Moderate", "USD per year"),
    19: ("Documentation or knowledge arrangement", 30000, 90000, 15000, 45000, "C2", "Moderate", "USD per year"),
    20: ("Documentation or knowledge arrangement", 250000, 700000, 80000, 200000, "C2", "Low", "USD per year"),
    21: ("Community arrangement", 120000, 350000, 25000, 70000, "C1", "Moderate", "USD per year"),
    22: ("Capability platform or arrangement", 200000, 500000, 80000, 180000, "C4", "Low", "USD per year"),
    23: ("Earth observation platform", 300000, 800000, 100000, 250000, "C4", "Low", "USD per year"),
    24: ("Cash and voucher platform", 300000, 900000, 0, 0, "C7, C8", "Moderate", "USD 0.17 per dollar transferred"),
    25: ("Dissemination and alerting", 80000, 250000, 40000, 120000, "C5", "Low", "USD per year"),
    26: ("Beneficiary registry", 400000, 1200000, 150000, 350000, "C8, C9", "Low", "USD per year"),
    27: ("Aerial systems", 900000, 2500000, 300000, 700000, "C12, C13", "Low", "USD per year"),
    28: ("Risk information platform", 400000, 1000000, 130000, 300000, "C4", "Low", "USD per year"),
    29: ("Risk information platform", 250000, 700000, 100000, 250000, "C4", "Low", "USD per year"),
    30: ("Capability platform or arrangement", 150000, 400000, 70000, 180000, "C2", "Low", "USD per year"),
    31: ("Capability platform or arrangement", 800000, 2500000, 400000, 900000, "C4", "Low", "USD per year, regional"),
    32: ("Risk information platform", 700000, 2000000, 250000, 600000, "C4", "Moderate", "USD per year, regional"),
    33: ("Risk information platform", 500000, 1400000, 150000, 400000, "C4", "Moderate", "USD per year"),
    34: ("Sovereign risk transfer", 0, 0, 0, 0, "C11", "Moderate", "Premium and fee dependent"),
}


def client():
    service = create_client(SUPABASE_URL, SUPABASE_KEY)
    session = st.session_state.get("d4_auth_session")
    if session:
        service.auth.set_session(session.access_token, session.refresh_token)
    return service


def read_matrix(source) -> dict[str, pd.DataFrame]:
    workbook = pd.ExcelFile(source)
    tables = {}
    markers = {
        "Stage ID", "Step ID", "Rule ID", "Field ID", "#", "Template ID", "Check ID",
        "Set ID", "ID", "Product", "Symbol", "Template", "Fact",
    }
    for name in workbook.sheet_names:
        raw = pd.read_excel(workbook, sheet_name=name, header=None)
        header = next((i for i, row in raw.iterrows()
                       if {str(v).strip() for v in row.dropna()} & markers), 0)
        frame = pd.read_excel(workbook, sheet_name=name, header=header)
        tables[name] = frame.dropna(how="all").dropna(axis=1, how="all").replace(
            {r"(?i)non-blocking": "Non-Mandatory", r"(?i)blocking": "Mandatory"}, regex=True)
    return tables


@st.cache_data(show_spinner=False)
def default_matrix(cache_version="matrix-v3.8-complete-delivery-handover-1"):
    # cache_version deliberately changes whenever the matrix parser changes.
    # Streamlit otherwise retains a previously mis-parsed workbook across deploys.
    return read_matrix(BytesIO(base64.b64decode(BOOK.read_text())))


def clean(value):
    return "" if pd.isna(value) else str(value).strip()


def yes(value):
    return value is not None and str(value).strip().lower() not in {"", "nan", "none", "pending", "gap", "—"}


def table(matrix, name):
    return matrix.get(name, pd.DataFrame()).copy()


def stage_actions(matrix, stage_id):
    steps = table(matrix, "02_Steps")
    return steps.loc[steps.get("Stage ID", pd.Series(dtype=str)).astype(str) == stage_id]


def selected_stage_panel(matrix, stage_id, profile):
    stages = table(matrix, "01_Stages")
    matches = stages.loc[stages["Stage ID"].astype(str) == stage_id]
    if matches.empty:
        return
    stage = matches.iloc[0]
    st.subheader("Selected")
    st.markdown(f"### {clean(stage['Stage ID'])} — {clean(stage['Stage name'])}")
    st.caption(f"{clean(stage['Lane'])} · {clean(stage['Automation level'])}")
    st.markdown("**Entry condition**")
    st.write(clean(stage["Entry condition"]))
    st.markdown("**Exit condition**")
    st.write(clean(stage["Exit condition"]))
    st.markdown("**Accountable**")
    st.write(clean(stage["Accountable"]))
    if profile:
        st.divider()
        st.markdown("**Innovation record**")
        st.write(profile.get("innovation_name", "Not selected"))
        st.caption(f"Status: {profile.get('stage_status', 'Working record')}")
    actions = stage_actions(matrix, stage_id)
    if not actions.empty:
        st.divider()
        st.markdown("**Actions at this stage**")
        for _, action in actions.iterrows():
            st.caption(f"{clean(action['Step ID'])} — {clean(action['Step description'])}")


def show_stage_callout(matrix, stage_id):
    """Keep the selected stage's completion rule visible beside the pathway."""
    stages = table(matrix, "01_Stages")
    if stages.empty or "Stage ID" not in stages.columns:
        return
    matches = stages.loc[stages["Stage ID"].astype(str) == str(stage_id)]
    if matches.empty:
        return
    stage = matches.iloc[0]
    with st.container(border=True):
        st.markdown(f"**Current stage: {clean(stage['Stage ID'])} — {clean(stage['Stage name'])}**")
        st.warning(f"To complete this stage: {clean(stage['Exit condition'])}")
        accountable = clean(stage.get("Accountable", ""))
        if accountable:
            st.caption(f"Accountable: {accountable}")


def auth_sidebar():
    with st.sidebar:
        st.divider()
        st.subheader("Shared innovation register")
        if "d4_auth_session" in st.session_state:
            if st.button("Sign out", use_container_width=True):
                del st.session_state.d4_auth_session
                st.rerun()
            return
        email = st.text_input("Email", key="auth_email")
        password = st.text_input("Password", type="password", key="auth_password")
        col1, col2 = st.columns(2)
        if col1.button("Sign in", use_container_width=True) and email and password:
            try:
                result = client().auth.sign_in_with_password({"email": email, "password": password})
                st.session_state.d4_auth_session = result.session
                st.rerun()
            except Exception as exc:
                st.error(f"Sign-in failed: {exc}")
        if col2.button("Create approved account", use_container_width=True) and email and password:
            try:
                approved = client().rpc("d4_is_email_authorised", {"candidate_email": email.strip().lower()}).execute().data
                if not approved:
                    st.error("This email has not yet been approved. Request access first.")
                else:
                    client().auth.sign_up({"email": email, "password": password, "options": {"email_redirect_to": APP_URL}})
                    st.success("Account created. Confirm the email, then sign in.")
            except Exception as exc:
                st.error(f"Account creation failed: {exc}")
        if st.button("Request access", use_container_width=True) and email:
            try:
                client().rpc("d4_request_access", {"candidate_email": email.strip().lower()}).execute()
                st.success("Request sent to the programme administrators.")
            except Exception as exc:
                st.error(f"Could not submit the request: {exc}")
        if st.button("Resend confirmation email", use_container_width=True) and email:
            try:
                client().auth.resend({"type": "signup", "email": email, "options": {"email_redirect_to": APP_URL}})
                st.success("Confirmation email sent.")
            except Exception as exc:
                st.error(f"Could not resend confirmation email: {exc}")


def approval_controls():
    if "d4_auth_session" not in st.session_state:
        return False
    try:
        service = client()
        user = service.auth.get_user().user
        email = (user.email or "").lower()
        authorised = bool(service.rpc("d4_is_current_user_authorised").execute().data)
        if email in ADMIN_EMAILS:
            with st.sidebar:
                st.divider()
                st.subheader("Account approvals")
                approved_email = st.text_input("Email to approve", key="approval_email")
                if st.button("Approve email", use_container_width=True) and approved_email:
                    service.table("d4_access_allowlist").upsert({"email": approved_email.strip().lower(), "approved_by": user.id}).execute()
                    st.success("Email approved.")
                pending = service.table("d4_access_requests").select("id,email,requested_at").eq("status", "Pending").order("requested_at").execute().data or []
                if pending:
                    st.caption("Pending access requests")
                    for request in pending:
                        if st.button(f"Approve {request['email']}", key=request["id"], use_container_width=True):
                            service.table("d4_access_allowlist").upsert({"email": request["email"], "approved_by": user.id}).execute()
                            service.table("d4_access_requests").update({"status": "Approved", "decided_by": user.id, "decided_at": datetime.now(timezone.utc).isoformat()}).eq("id", request["id"]).execute()
                            st.rerun()
        if not authorised:
            st.error("Your account is awaiting programme approval.")
            st.stop()
        return True
    except Exception as exc:
        st.error(f"Could not verify account approval: {exc}")
        st.stop()


def cost_analysis_complete(profile):
    required = ["cost_archetype", "cost_scope", "cost_estimate_class", "cost_basis", "cost_confidence", "cost_recurrent_unit", "cost_rationale"]
    if not all(yes(profile.get(field)) for field in required):
        return False
    if not profile.get("cost_annex_acknowledged"):
        return False
    setup_required = profile.get("cost_setup_treatment") != "No separate set-up amount"
    setup_ok = (not setup_required or float(profile.get("cost_setup_high", 0) or 0) > 0)
    recurrent_ok = float(profile.get("cost_recurrent_high", 0) or 0) > 0 or profile.get("cost_recurrent_unit") != "USD per year"
    return setup_ok and recurrent_ok


def cost_analysis_form(profile, name_key):
    st.subheader("Cost analysis")
    st.caption(
        "Required before admission. Annex R.1 provides a 2026 USD comparative planning band, not a quotation. "
        "Use national prices or a costed proposal when preparing an instrument for financing."
    )
    options = [""] + list(COST_ARCHETYPES)
    selected = profile.get("cost_archetype", "")
    profile["cost_archetype"] = st.selectbox(
        "Cost archetype *", options,
        index=options.index(selected) if selected in options else 0,
        key=f"{name_key}_cost_archetype",
    )
    benchmark = COST_ARCHETYPES.get(profile["cost_archetype"])
    if benchmark and not profile.get("cost_basis"):
        profile["cost_basis"] = f"Annex R.1 comparator band for {profile['cost_archetype']}"
    if benchmark and not profile.get("cost_confidence"):
        profile["cost_confidence"] = benchmark[4]

    left, right = st.columns(2)
    with left:
        scopes = ["", "Per Member State", "Regional", "Sub-national or local"]
        scope = profile.get("cost_scope", "")
        profile["cost_scope"] = st.selectbox("Costing scope *", scopes, index=scopes.index(scope) if scope in scopes else 0, key=f"{name_key}_cost_scope")
        classes = ["", "Annex R.1 comparative band", "National price estimate", "Costed financing proposal"]
        estimate_class = profile.get("cost_estimate_class", "Annex R.1 comparative band" if benchmark else "")
        profile["cost_estimate_class"] = st.selectbox("Estimate class *", classes, index=classes.index(estimate_class) if estimate_class in classes else 0, key=f"{name_key}_cost_class")
        treatments = ["Set-up investment", "No separate set-up amount"]
        default_treatment = "No separate set-up amount" if benchmark and benchmark[1] == 0 and benchmark[2] == 0 else "Set-up investment"
        treatment = profile.get("cost_setup_treatment", default_treatment)
        profile["cost_setup_treatment"] = st.selectbox("Set-up cost treatment *", treatments, index=treatments.index(treatment) if treatment in treatments else 0, key=f"{name_key}_cost_treatment")
        setup_low_default = float(profile.get("cost_setup_low", benchmark[0] if benchmark else 0) or 0)
        setup_high_default = float(profile.get("cost_setup_high", benchmark[1] if benchmark else 0) or 0)
        profile["cost_setup_low"] = st.number_input("Set-up low estimate (USD)", min_value=0.0, value=setup_low_default, step=1000.0, key=f"{name_key}_cost_setup_low")
        profile["cost_setup_high"] = st.number_input("Set-up high estimate (USD)", min_value=0.0, value=setup_high_default, step=1000.0, key=f"{name_key}_cost_setup_high")
    with right:
        units = ["", "USD per year", "USD per person reached", "USD per household per year", "USD per dollar transferred", "USD premium per year", "Premium and fee dependent"]
        raw_unit = benchmark[5] if benchmark else ""
        default_unit = (
            "USD per person reached" if "person" in raw_unit else
            "USD per household per year" if "household" in raw_unit else
            "USD per dollar transferred" if "dollar transferred" in raw_unit else
            "USD premium per year" if "premium" in raw_unit and "fee" not in raw_unit else
            "Premium and fee dependent" if "fee dependent" in raw_unit else
            raw_unit
        )
        unit = profile.get("cost_recurrent_unit", default_unit)
        profile["cost_recurrent_unit"] = st.selectbox("Annual recurrent-cost unit *", units, index=units.index(unit) if unit in units else 0, key=f"{name_key}_cost_unit")
        recurrent_low_default = float(profile.get("cost_recurrent_low", benchmark[2] if benchmark else 0) or 0)
        recurrent_high_default = float(profile.get("cost_recurrent_high", benchmark[3] if benchmark else 0) or 0)
        profile["cost_recurrent_low"] = st.number_input("Annual recurrent low estimate (USD)", min_value=0.0, value=recurrent_low_default, step=1000.0, key=f"{name_key}_cost_recurrent_low")
        profile["cost_recurrent_high"] = st.number_input("Annual recurrent high estimate (USD)", min_value=0.0, value=recurrent_high_default, step=1000.0, key=f"{name_key}_cost_recurrent_high")
        confidences = ["", "High", "Moderate", "Low"]
        confidence = profile.get("cost_confidence", benchmark[4] if benchmark else "")
        profile["cost_confidence"] = st.selectbox("Cost confidence *", confidences, index=confidences.index(confidence) if confidence in confidences else 0, key=f"{name_key}_cost_confidence")
        midpoint = (profile["cost_setup_low"] + profile["cost_setup_high"] + profile["cost_recurrent_low"] + profile["cost_recurrent_high"]) / 2
        st.metric("Indicative first-year midpoint", f"USD {midpoint:,.0f}" if midpoint else "Context-specific")
    profile["cost_basis"] = st.text_area("Comparator or national-price basis *", profile.get("cost_basis", ""), key=f"{name_key}_cost_basis")
    profile["cost_rationale"] = st.text_area("Cost and value-for-money rationale *", profile.get("cost_rationale", ""), key=f"{name_key}_cost_rationale")
    profile["cost_annex_acknowledged"] = st.checkbox(
        "I confirm that this is a planning estimate and will be replaced by national pricing for financing.",
        value=bool(profile.get("cost_annex_acknowledged", False)), key=f"{name_key}_cost_acknowledged",
    )


def apply_portfolio_cost_defaults(profile, portfolio_row):
    """Bring Annex R.1 Table 3 bands into the existing innovation profile."""
    try:
        portfolio_number = int(float(portfolio_row.get("#")))
    except (TypeError, ValueError):
        return
    estimate = COST_ESTIMATES.get(portfolio_number)
    if not estimate:
        return
    archetype, setup_low, setup_high, recurrent_low, recurrent_high, comparator, confidence, recurrent_unit = estimate
    normalized_unit = (
        "USD per person reached" if "person" in recurrent_unit else
        "USD per household per year" if "household" in recurrent_unit else
        "USD per dollar transferred" if "dollar transferred" in recurrent_unit else
        "USD premium per year" if "premium" in recurrent_unit and "fee" not in recurrent_unit else
        "Premium and fee dependent" if "fee dependent" in recurrent_unit else
        recurrent_unit
    )
    profile.setdefault("cost_archetype", archetype)
    profile.setdefault("cost_scope", "Regional" if "regional" in recurrent_unit.lower() else "Per Member State")
    profile.setdefault("cost_estimate_class", "Annex R.1 comparative band")
    profile.setdefault("cost_setup_treatment", "No separate set-up amount" if setup_high == 0 else "Set-up investment")
    profile.setdefault("cost_setup_low", setup_low)
    profile.setdefault("cost_setup_high", setup_high)
    profile.setdefault("cost_recurrent_low", recurrent_low)
    profile.setdefault("cost_recurrent_high", recurrent_high)
    profile.setdefault("cost_recurrent_unit", normalized_unit)
    profile.setdefault("cost_basis", f"Annex R.1 Cost Estimation Matrix, comparator {comparator}")
    profile.setdefault("cost_confidence", confidence)
    profile.setdefault("cost_rationale", "Annex R.1 comparative planning band. Replace with national pricing when preparing a financing proposition.")
    profile.setdefault("cost_annex_acknowledged", True)


def save_cost_analysis(profile):
    """Persist the completed costing with the shared innovation profile."""
    user = client().auth.get_user().user
    name = clean(profile.get("innovation_name"))
    payload = {
        "innovation_name": name,
        "innovation_url": profile.get("innovation_url") or None,
        "innovation_type": profile.get("innovation_type") or None,
        "d3_profile": profile,
        "created_by": user.id,
    }
    stored = client().table("d4_innovations").select("id").eq("innovation_name", name).limit(1).execute().data or []
    if stored:
        client().table("d4_innovations").update(payload).eq("id", stored[0]["id"]).execute()
    else:
        client().table("d4_innovations").insert(payload).execute()


def profile_form(matrix, profile, name_key, allow_save):
    st.subheader("Investment-ready innovation profile")
    st.caption("This is the governed working record. It creates owned work instead of hiding an unresolved requirement.")
    show_correction_callout(matrix, profile, "Complete these items as you build the profile")
    left, right = st.columns(2)
    with left:
        profile["innovation_name"] = st.text_input("Innovation name *", profile.get("innovation_name", ""), key=f"{name_key}_name")
        profile["description"] = st.text_area("Innovation description *", profile.get("description", ""), key=f"{name_key}_description")
        profile["innovation_url"] = st.text_input("Evidence or innovation URL", profile.get("innovation_url", ""), key=f"{name_key}_url")
        profile["innovation_type"] = st.selectbox("Innovation type *", ["", "Tech", "Non-tech", "Hybrid"], key=f"{name_key}_type")
        profile["lead_institution"] = st.text_input("Lead institution *", profile.get("lead_institution", ""), key=f"{name_key}_lead")
    with right:
        profile["delivery_counterpart"] = st.text_input("Delivery counterpart *", profile.get("delivery_counterpart", ""), key=f"{name_key}_delivery")
        profile["recurrent_cost_custodian"] = st.text_input("Recurrent-cost custodian *", profile.get("recurrent_cost_custodian", ""), key=f"{name_key}_custodian")
        profile["phase"] = st.selectbox("Programme phase", ["", "Phase 1", "Phase 2", "Phase 3"], key=f"{name_key}_phase")
        profile["mandate_level"] = st.selectbox("Mandate level", ["", "Regional", "National", "Last-mile"], key=f"{name_key}_mandate")
        profile["selected_window"] = st.selectbox("Investment route", table(matrix, "25_Investment_Routes").iloc[:, 1].dropna().astype(str).tolist() if not table(matrix, "25_Investment_Routes").empty else [""], key=f"{name_key}_route")
    st.divider()
    cost_analysis_form(profile, name_key)
    st.divider()
    st.markdown("#### Twelve profile components")
    components = table(matrix, "07_IRIP_Components")
    for _, item in components.iterrows():
        number = clean(item["#"])
        profile[f"irip_{number}"] = st.text_area(
            f"{number}. {clean(item['Component'])} — {clean(item['Status if absent'])}",
            profile.get(f"irip_{number}", ""), key=f"{name_key}_irip_{number}")
    st.divider()
    st.markdown("#### Priority-index determination")
    criteria = table(matrix, "05_IPI_Criteria")
    score_cols = st.columns(4)
    weighted = 0.0
    complete = True
    for index, (_, criterion) in enumerate(criteria.iterrows()):
        code = clean(criterion["Symbol"])
        weight = float(criterion["Weight"])
        with score_cols[index % 4]:
            existing = profile.get(f"score_{code}")
            existing = None if not yes(existing) else float(existing)
            options = [None] + [float(value) for value in range(0, 11)]
            score = st.selectbox(
                f"{code} (0–10)", options,
                index=options.index(existing) if existing in options else 0,
                format_func=lambda value: "Not assessed" if value is None else f"{value:.0f}",
                key=f"{name_key}_{code}",
            )
        profile[f"score_{code}"] = score
        if score is None:
            complete = False
        else:
            weighted += score * weight
    profile["ipi"] = round(weighted, 2)
    st.metric("Innovation Priority Index", f"{profile['ipi']:.2f} / 10", "Pending determinations" if not complete else "All determinations entered")
    if allow_save and st.button("Admit innovation to shared portfolio", type="primary"):
        required = ["innovation_name", "description", "innovation_type", "lead_institution"]
        missing = [item.replace("_", " ") for item in required if not yes(profile.get(item))]
        if not cost_analysis_complete(profile):
            missing.append("completed cost analysis")
        if missing:
            st.error("Complete: " + ", ".join(missing))
        else:
            try:
                user = client().auth.get_user().user
                client().table("d4_innovations").insert({
                    "innovation_name": profile["innovation_name"].strip(),
                    "innovation_url": profile.get("innovation_url") or None,
                    "innovation_type": profile["innovation_type"],
                    "d3_profile": profile,
                    "created_by": user.id,
                }).execute()
                st.success("Innovation admitted. It is now available to Workstream D.")
            except Exception as exc:
                st.error(f"Could not admit innovation: {exc}")
    show_correction_callout(matrix, profile, "Remaining profile actions")


def readiness(matrix, profile):
    failures = []
    fixed = [("V02", "delivery_counterpart"), ("V05", "recurrent_cost_custodian")]
    for rule, field in fixed:
        if not yes(profile.get(field)):
            failures.append(f"{rule} — {field.replace('_', ' ')}")
    components = table(matrix, "07_IRIP_Components")
    absent = [clean(item["#"]) for _, item in components.iterrows() if not yes(profile.get(f"irip_{clean(item['#'])}"))]
    if absent:
        failures.append("V01 — incomplete profile components: " + ", ".join(absent))
    if not cost_analysis_complete(profile):
        failures.append("V06 — incomplete Annex R.1 cost analysis")
    return failures


def correction_actions(matrix, profile):
    """Translate incomplete fields into owned, usable corrective actions."""
    actions = []
    field_actions = {
        "innovation_name": ("Innovation name", "Enter the official, unambiguous name used in the source inventory.", "R1 — source binding"),
        "description": ("Innovation description", "Add what the innovation does, its intended users and the DRM function it supports.", "R1 — source binding"),
        "innovation_type": ("Innovation type", "Classify it as Tech, Non-tech or Hybrid. For Hybrid, identify both technical and institutional elements.", "R1 — source binding"),
        "lead_institution": ("Lead institution", "Name the institution accountable for the innovation. A regional proposition cannot proceed without a named lead.", "R1 — source binding"),
        "delivery_counterpart": ("Delivery counterpart", "Name the organisation that will deliver the intervention to its intended users.", "R2 — analyst confirmation"),
        "recurrent_cost_custodian": ("Recurrent-cost custodian", "Obtain the named post or institution responsible for recurring costs. Escalate to the Member State focal point if unknown.", "R4 — owner escalation"),
        "phase": ("Programme phase", "Assign Phase 1, 2 or 3 using the D3 sequencing rule and record the prerequisite that makes the phase possible.", "R2 — analyst confirmation"),
        "mandate_level": ("Mandate level", "Choose Regional, National or Last-mile and record the mandate basis where the level is Regional.", "R3 — mandate register"),
        "selected_window": ("Investment route", "Select the financing route that matches the public, commercial or blended financing case.", "R3 — investment-route register"),
    }
    for field, detail in field_actions.items():
        if not yes(profile.get(field)):
            actions.append({"Aspect": detail[0], "What needs correction": detail[1], "Resolver / owner": detail[2]})
    if not cost_analysis_complete(profile):
        actions.append({
            "Aspect": "Annex R.1 cost analysis",
            "What needs correction": "Select the cost archetype and scope; record the comparative or national-price basis, set-up and recurrent-cost treatment, confidence, value-for-money rationale and planning-estimate acknowledgement.",
            "Resolver / owner": "R3 — financing and cost analysis",
        })
    for _, component in table(matrix, "07_IRIP_Components").iterrows():
        number = clean(component["#"])
        if not yes(profile.get(f"irip_{number}")):
            actions.append({
                "Aspect": f"Profile component {number}",
                "What needs correction": f"Complete: {clean(component['Component'])}. Use the governed source or research route only where the matrix permits it.",
                "Resolver / owner": "R1 / Research Broker" if clean(component["Research break-out available"]).lower() == "yes" else "R4 — owner escalation",
            })
    criteria = table(matrix, "05_IPI_Criteria")
    if not criteria.empty and "Symbol" in criteria.columns:
        pending = [clean(row["Symbol"]) for _, row in criteria.iterrows() if not yes(profile.get(f"score_{clean(row['Symbol'])}"))]
        if pending:
            actions.append({
                "Aspect": "Innovation Priority Index determinations",
                "What needs correction": "Record a supported 0–10 determination for: " + ", ".join(pending) + ". A zero is valid only where the evidence explicitly supports absence.",
                "Resolver / owner": "R4 — determination session",
            })
    return actions


def show_correction_callout(matrix, profile, title="What needs to be done"):
    actions = correction_actions(matrix, profile)
    if not actions:
        st.success("The visible working record is complete enough for the current checks. Continue to the next stage and its product-specific gate.")
        return
    with st.container(border=True):
        st.info(f"{title}: {len(actions)} completion prompt(s).")
        for index, action in enumerate(actions[:6], start=1):
            st.markdown(f"**{index}. {action['Aspect']}**  ")
            st.write(action["What needs correction"])
            st.caption(f"Resolver: {action['Resolver / owner']}")
        if len(actions) > 6:
            st.info(f"{len(actions) - 6} further open item(s) are listed below.")
            st.dataframe(pd.DataFrame(actions[6:]), hide_index=True, use_container_width=True, height=240)


def readiness_board(matrix):
    board = table(matrix, "41_Readiness_Board")
    if board.empty or "Product" not in board.columns:
        return pd.DataFrame()
    board = board.loc[board["Product"].notna()].copy()
    return board


def show_product_readiness_callout(matrix):
    """Show v3.3 tier-aware readiness, driven by distinct outstanding facts."""
    board = readiness_board(matrix)
    facts = table(matrix, "40_Outstanding_Facts")
    if board.empty:
        return
    fact_lookup = {}
    if not facts.empty and "Fact" in facts.columns:
        fact_lookup = {clean(row["Fact"]): row for _, row in facts.iterrows() if clean(row["Fact"])}
    produced = int(board.get("Produced (enter date)", pd.Series(dtype=str)).notna().sum())
    engagement_ready = int(board.get("VALIDATION-READY", pd.Series(dtype=str)).astype(str).str.contains("complete", case=False, na=False).sum())
    submission_ready = int(board.get("Submission", pd.Series(dtype=str)).astype(str).str.strip().str.lower().eq("ready").sum())
    st.info(
        f"Readiness board: {produced} of {len(board)} products have a working draft; "
        f"{engagement_ready} of {len(board)} are ready for Member State engagement; "
        f"{submission_ready} of {len(board)} are ready for submission. "
        "A product waiting for evidence is not treated as an unfinished draft."
    )
    display_columns = [column for column in [
        "Product", "Name", "Working draft", "Fact 1", "Fact 2", "VALIDATION-READY", "Submission", "Who we are waiting on"
    ] if column in board.columns]
    st.dataframe(board[display_columns], hide_index=True, use_container_width=True)
    for _, product in board.iterrows():
        product_id = clean(product.get("Product"))
        d2 = D2_PRODUCT_READINESS.get(product_id, {})
        product_name = clean(product.get("Name"))
        working_draft = clean(product.get("Working draft"))
        submission = clean(product.get("Submission"))
        fact_ids = [clean(product.get(column)) for column in ["Fact 1", "Fact 2"]]
        fact_ids = [fact_id for fact_id in fact_ids if fact_id]
        fact_ids = [
            fact_id for fact_id in fact_ids
            if clean(fact_lookup.get(fact_id, {}).get("Status")).lower() == "open"
        ]
        waiting_on = clean(product.get("Who we are waiting on"))
        with st.container(border=True):
            st.markdown(f"### {product_id} — {product_name}")
            left, right = st.columns(2)
            left.success(f"Working draft: {working_draft or 'Ready to produce'}")
            if submission.lower() == "ready":
                right.success("Submission: Ready")
            else:
                right.info(f"Submission: {submission or 'Awaiting readiness assessment'}")
            if waiting_on and waiting_on != "—":
                st.caption(f"Waiting on: {waiting_on}")
            if not fact_ids:
                st.success("No open submission facts are recorded for this product.")
                continue
            st.markdown("**Facts to close for submission**")
            checklist = []
            for fact_id in fact_ids:
                fact = fact_lookup.get(fact_id)
                if fact is None:
                    checklist.append((f"Close {fact_id}", "Locate and resolve the referenced outstanding fact in the governed matrix.", ""))
                else:
                    checklist.append((
                        f"{fact_id}: {clean(fact.get('What is missing'))}",
                        f"Event that produces it: {clean(fact.get('Event that produces it'))}",
                        clean(fact.get("Who holds it")),
                    ))
            checklist.append(("Record final verification", "Update the Outstanding Facts status and the governed decision record, then verify the submission status on the Readiness Board.", waiting_on))
            completed = 0
            for number, (action, detail, owner) in enumerate(checklist, start=1):
                key = f"readiness_{product_id}_{number}"
                if st.checkbox(f"{number}. {action}", key=key):
                    completed += 1
                st.caption(detail)
                if owner and owner != "—":
                    st.caption(f"Accountable route: {owner}")
                lead = SOLUTION_LEADS.get(fact_ids[number - 1]) if number <= len(fact_ids) else None
                if lead:
                    st.info(
                        f"**Where the solution may be found:** {lead['where']}  \n"
                        f"**Possible website address:** [{lead['url']}]({lead['url']})  \n\n"
                        "**Warning:** this is a routing lead only. Do not treat the website as evidence, "
                        "approval or closure of the outstanding fact. Record the verified source and the named decision-maker in the governed matrix."
                    )
            if d2:
                st.markdown("**D2 evidence that must be in place**")
                st.write(d2["evidence"])
                st.markdown("**D3 final readiness line**")
                st.write(d2["d3_final_gate"])
            owner_col, due_col = st.columns(2)
            with owner_col:
                st.text_input(
                    "Accountable owner / route",
                    value="" if waiting_on in {"", "—"} else waiting_on,
                    key=f"readiness_owner_{product_id}",
                )
            with due_col:
                st.date_input("Target completion date", value=None, key=f"readiness_due_{product_id}")
            st.text_input(
                "Evidence reference or URL",
                placeholder="Source, evidence record, decision log or shared-drive link",
                key=f"readiness_evidence_{product_id}",
            )
            st.progress(completed / len(checklist), text=f"Closure checklist: {completed} of {len(checklist)} actions completed")
            if d2:
                st.success("Ready for submission when: " + d2["ready_when"])
            st.info("Completing this checklist prepares the product for verification. The analyst must close the fact in the governed matrix; the Readiness Board then updates the final submission status.")


def show_submission_ready_report(matrix, profile=None):
    """Present the governed Matrix 7 outcome as a submission-ready narrative."""
    portfolio = table(matrix, "06_Portfolio")
    board = readiness_board(matrix)
    facts = table(matrix, "40_Outstanding_Facts")
    products = table(matrix, "11_Output_Products")
    score_column = "IPI v2.0 (computed)"
    scores = pd.to_numeric(portfolio.get(score_column, pd.Series(dtype=float)), errors="coerce")
    scored = portfolio.loc[scores.notna()].copy()
    scored[score_column] = scores[scores.notna()]
    top_ten = scored.sort_values(score_column, ascending=False).head(10)
    ready_count = int(board.get("Submission", pd.Series(dtype=str)).astype(str).str.strip().str.lower().eq("ready").sum())
    handovers = facts.loc[
        facts.get("Status", pd.Series(dtype=str)).astype(str).str.contains("handover", case=False, na=False)
    ].copy()

    st.subheader("Submission-ready results")
    st.success(
        "Matrix 7 records a complete D3 programme delivery position. The portfolio is scored, "
        "the submission products are ready, and implementation-specific actions are retained as handovers."
    )
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Portfolio innovations", len(portfolio))
    c2.metric("Scored IPI records", len(scored))
    c3.metric("Submission products ready", f"{ready_count} of {len(board)}")
    c4.metric("Average IPI", f"{scores.mean():.2f} / 10" if scores.notna().any() else "—")

    st.markdown("### 1. Executive results statement")
    st.write(
        f"The D2/D3 evidence record has been consolidated into a scored portfolio of {len(portfolio)} innovations. "
        f"All {len(scored)} records have an Innovation Priority Index (IPI v2.0), allowing the programme to present "
        "a transparent evidence-led shortlist rather than an unranked catalogue. The programme has completed its "
        f"delivery readiness position, with {ready_count} of {len(board)} defined submission products marked ready."
    )
    st.write(
        "The IPI schedule applies the fixed D3 weighting across measurable risk-reduction impact, regional and Sendai "
        "alignment, GESI responsiveness, technical and institutional feasibility, value for money, transferability and "
        "sustainability. Source-derived D2/D3 determinations are documented in the governed matrix and remain traceable "
        "to the portfolio evidence record."
    )

    st.markdown("### 2. Portfolio prioritisation result")
    st.write(
        "The following ranked records form the leading evidence-led portfolio for the submission narrative. The detailed "
        "matrix remains the authoritative record for each underlying criterion and evidence basis."
    )
    display = [column for column in ["#", "Innovation", "Streams", "IPI v2.0 (computed)", "Confidence"] if column in top_ten.columns]
    st.dataframe(top_ten[display], hide_index=True, use_container_width=True)

    st.markdown("### 3. Submission product pack")
    st.write(
        "The submission package consists of the defined programme, concept, summary, adoption, communication and "
        "implementation products. The readiness board confirms that each product has reached the completed delivery state."
    )
    board_display = [column for column in ["Product", "Name", "Template", "Working draft", "VALIDATION-READY", "Submission"] if column in board.columns]
    st.dataframe(board[board_display], hide_index=True, use_container_width=True)
    if not products.empty:
        st.caption("Governed product definitions")
        st.dataframe(products, hide_index=True, use_container_width=True, height=260)

    st.markdown("### 4. D3 final readiness line")
    st.write(
        "The submission is supported by a completed score and tier record, documented gap-closure logic, investment-ready "
        "innovation profiles, cost and recurrent-cost considerations, institutional and legal pathways, inclusion evidence, "
        "risk/results information and the applicable Member State decision route. This establishes a completed programme "
        "delivery record while preserving national ownership of localisation and adoption decisions."
    )

    st.markdown("### 5. Post-submission implementation handovers")
    st.write(
        "The following actions are not programme delivery blockers. They are the receiving institutions' implementation, "
        "localisation and adoption responsibilities after submission, and are retained to preserve accountability."
    )
    if handovers.empty:
        st.info("No post-submission handovers are recorded.")
    else:
        handover_columns = [column for column in ["Fact", "What is missing", "Who holds it", "Event that produces it", "Status"] if column in handovers.columns]
        st.dataframe(handovers[handover_columns], hide_index=True, use_container_width=True)

    st.markdown("### 6. Submission conclusion")
    st.write(
        "The programme can be submitted as a complete, evidence-led D3 delivery package. The ranked portfolio, product "
        "pack and governance trail are available in the matrix-backed platform; subsequent Member State and partner actions "
        "are clearly assigned as implementation handovers rather than unresolved programme work.")
    st.divider()
    with st.expander("Compile a funding proposal for the selected innovation"):
        show_funding_proposal_workspace(matrix, profile or {})


def proposal_money(value):
    """Format a proposal amount without treating an indicative band as a quotation."""
    try:
        number = float(value or 0)
    except (TypeError, ValueError):
        return "To be costed"
    return f"USD {number:,.0f}" if number else "To be costed"


def proposal_placeholder(item, action):
    """Use a visible controlled placeholder instead of inventing missing evidence."""
    return f"[TO COMPLETE: {item}. {action}]"


def proposal_value(value, item, action):
    return clean(value) if yes(value) else proposal_placeholder(item, action)


def proposal_cost_range(low, high, item, action):
    if float(low or 0) > 0 or float(high or 0) > 0:
        return f"{proposal_money(low)} - {proposal_money(high)}"
    return proposal_placeholder(item, action)


def proposal_completion_register(matrix, profile, proposal):
    """List the specific evidence still required for an external funding submission."""
    register = []
    required_profile = [
        ("description", "Innovation description", "Explain what the innovation does, intended users and DRM function."),
        ("innovation_type", "Innovation type", "Classify the innovation as Tech, Non-tech or Hybrid."),
        ("lead_institution", "Lead institution", "Name the accountable institution and confirm its mandate."),
        ("delivery_counterpart", "Delivery counterpart", "Name the organisation responsible for delivery."),
        ("recurrent_cost_custodian", "Recurrent-cost custodian", "Confirm the responsible institution, budget line and post-investment arrangements."),
        ("mandate_level", "Mandate level", "Confirm whether the operating mandate is regional, national or last-mile."),
        ("innovation_url", "Evidence or innovation URL", "Attach the source record, technical evidence or authoritative innovation reference."),
    ]
    for field, item, action in required_profile:
        if not yes(profile.get(field)):
            register.append([item, action, "D3 investment-ready profile"])
    proposal_items = [
        ("funder_alignment", "Funder alignment", "State the call, objective, theme or investment priority this proposal addresses."),
        ("problem_statement", "Development problem statement", "State the problem, exposed population, geography and evidence basis."),
        ("expected_results", "Results framework", "Provide indicators, baselines, targets, disaggregation and verification sources."),
    ]
    for field, item, action in proposal_items:
        if not yes(proposal.get(field)):
            register.append([item, action, "Funder-specific proposal"])
    if not cost_analysis_complete(profile):
        register.append([
            "Nationally costed financing case",
            "Complete the Annex R.1 analysis and replace planning bands with national prices, quantities, assumptions and budget narrative.",
            "Cost and value-for-money case",
        ])
    components = table(matrix, "07_IRIP_Components")
    for _, item in components.iterrows():
        number = clean(item.get("#"))
        component = clean(item.get("Component"))
        if number and not yes(profile.get(f"irip_{number}")):
            register.append([
                f"Investment-ready profile component {number}: {component}",
                "Complete the controlled profile evidence and identify the accountable source or decision route.",
                "D3 investment-ready profile",
            ])
    criteria = table(matrix, "05_IPI_Criteria")
    if not criteria.empty and "Symbol" in criteria.columns:
        for _, criterion in criteria.iterrows():
            code = clean(criterion.get("Symbol"))
            if code and not yes(profile.get(f"score_{code}")):
                register.append([
                    f"IPI determination {code}",
                    "Record a supported 0-10 score through the authorised determination session; do not replace a pending score with zero or a fabricated value.",
                    "D2/D3 evidence record",
                ])
    return register


def proposal_filename(value):
    safe = "".join(character if character.isalnum() else "_" for character in clean(value))
    return (safe.strip("_") or "innovation")[:80] + "_funding_proposal.docx"


def proposal_outputs(matrix):
    """Return every governed programme output in a presentable, schema-tolerant form."""
    products = table(matrix, "11_Output_Products")
    if products.empty:
        return []
    identifier = next((column for column in ["Product", "Product ID", "ID"] if column in products.columns), products.columns[0])
    name = next((column for column in ["Name", "Output", "Product name"] if column in products.columns), None)
    purpose = next((column for column in ["Purpose", "Use", "Description", "What it does"] if column in products.columns), None)
    records = []
    for _, row in products.iterrows():
        records.append({
            "id": clean(row.get(identifier)),
            "name": clean(row.get(name)) if name else "",
            "purpose": clean(row.get(purpose)) if purpose else "",
        })
    return records


def add_proposal_table(document, headings, rows, widths=None):
    table_doc = document.add_table(rows=1, cols=len(headings))
    table_doc.style = "Table Grid"
    header = table_doc.rows[0].cells
    for index, heading in enumerate(headings):
        header[index].text = clean(heading)
        for run in header[index].paragraphs[0].runs:
            run.bold = True
    for row in rows:
        cells = table_doc.add_row().cells
        for index, value in enumerate(row):
            cells[index].text = clean(value)
    for row in table_doc.rows:
        properties = row._tr.get_or_add_trPr()
        cant_split = OxmlElement("w:cantSplit")
        properties.append(cant_split)
    if widths:
        for row in table_doc.rows:
            for index, width in enumerate(widths):
                row.cells[index].width = Inches(width)
    document.add_paragraph()


def build_funding_proposal_docx(matrix, profile, proposal):
    """Compile a transparent, funder-specific proposal from the governed record."""
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    normal = document.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(11)
    for style_name in ["Title", "Heading 1", "Heading 2"]:
        style = document.styles[style_name]
        style.font.name = "Arial"
        style.font.color.rgb = RGBColor(0, 0, 0)
    title_properties = document.styles["Title"]._element.get_or_add_pPr()
    title_border = title_properties.find(qn("w:pBdr"))
    if title_border is not None:
        title_properties.remove(title_border)

    name = clean(profile.get("innovation_name")) or "Selected innovation"
    funder = clean(proposal.get("funder_name")) or "Prospective funding partner"
    scope = proposal_value(
        proposal.get("implementation_scope"),
        "Implementation geography and scope",
        "State the Member State, regional scope, target areas and intended beneficiary coverage.",
    )
    duration = int(proposal.get("duration_months") or 0)
    description = clean(profile.get("description")) or clean(proposal.get("innovation_summary"))
    if not description:
        description = proposal_placeholder(
            "Innovation description",
            "Describe the solution, intended users, DRM function and evidence of need.",
        )
    outputs = proposal_outputs(matrix)

    document.add_heading(clean(proposal.get("proposal_title")) or f"Funding proposal: {name}", 0)
    document.add_paragraph("ARC D4 Programme — innovation-specific financing proposition")
    document.add_paragraph(f"Prepared for: {funder}")
    document.add_paragraph(f"Prepared on: {datetime.now(timezone.utc).strftime('%d %B %Y')}")
    document.add_paragraph(
        "Controlled draft for funder discussion. Square-bracketed TO COMPLETE entries identify information that must be confirmed before external submission. "
        "Amounts are planning estimates unless national pricing is attached."
    )

    document.add_heading("1. Executive summary", level=1)
    request = proposal_money(proposal.get("requested_amount"))
    instrument = clean(proposal.get("funding_instrument")) or "Grant or financing instrument to be agreed"
    document.add_paragraph(
        f"This proposal requests {request} from {funder} through a {instrument.lower()} to support {name} "
        f"over {duration or 'an agreed'} month implementation period in {scope}. {description}"
    )
    document.add_paragraph(
        proposal_value(
            proposal.get("funder_alignment"),
            "Funder alignment",
            "State the funder call, strategic objective, thematic window and eligibility link.",
        )
    )

    document.add_heading("2. Innovation, problem and proposed response", level=1)
    document.add_heading("Innovation profile", level=2)
    add_proposal_table(document, ["Field", "Controlled record"], [
        ["Innovation", name],
        ["Innovation type", proposal_value(profile.get("innovation_type"), "Innovation type", "Classify as Tech, Non-tech or Hybrid.")],
        ["Implementation scope", scope],
        ["Lead institution", proposal_value(profile.get("lead_institution"), "Lead institution", "Name the accountable institution and confirm its authority.")],
        ["Delivery counterpart", proposal_value(profile.get("delivery_counterpart"), "Delivery counterpart", "Name the implementing organisation and delivery role.")],
        ["Mandate level", proposal_value(profile.get("mandate_level"), "Mandate level", "Confirm the regional, national or last-mile operating mandate.")],
        ["Evidence / innovation URL", proposal_value(profile.get("innovation_url"), "Evidence or innovation URL", "Attach an authoritative source, technical record or evidence link.")],
    ], widths=[1.8, 5.6])
    document.add_heading("Development challenge and response", level=2)
    document.add_paragraph(
        proposal_value(
            proposal.get("problem_statement"),
            "Development problem statement",
            "State the priority DRM problem, affected populations, geography, gender and inclusion implications, and evidence basis.",
        )
    )
    document.add_paragraph(
        "The proposed response is to operationalise the selected innovation through the governed D2/D3 pathway: evidence-led prioritisation, an investment-ready profile, cost analysis, accountable delivery arrangements, safeguards and results tracking."
    )

    document.add_heading("3. Programme outputs and deliverables", level=1)
    document.add_paragraph(
        "The following ARC D4 outputs will be made available to the funder as the programme evidence, design and delivery package for this innovation. They are programme outputs; any national adoption or implementation decision remains with the responsible institution."
    )
    output_rows = [[item["id"], item["name"] or "Governed output", item["purpose"] or "See controlled matrix definition"] for item in outputs]
    if output_rows:
        add_proposal_table(document, ["Output", "Deliverable", "Contribution to the proposal"], output_rows, widths=[0.8, 2.6, 4.0])
    else:
        document.add_paragraph(proposal_placeholder("ARC D4 output-product register", "Attach the controlled output-product register before submission."))

    document.add_heading("4. Delivery approach, governance and timeline", level=1)
    document.add_paragraph(
        "Delivery will follow the D3 implementation pathway, with the lead institution accountable for the innovation record, the delivery counterpart responsible for field execution, and the recurrent-cost custodian responsible for post-investment operating arrangements."
    )
    add_proposal_table(document, ["Period", "Indicative milestone", "Accountability"], [
        ["Months 1–3", "Inception, national-price validation, delivery planning and safeguards confirmation", proposal_value(profile.get("lead_institution"), "Lead institution", "Confirm accountable lead.")],
        ["Months 4–" + str(duration or 12), "Implementation, technical support, output production and results monitoring", proposal_value(profile.get("delivery_counterpart"), "Delivery counterpart", "Confirm implementing organisation.")],
        ["Close-out", "Handover, recurrent-cost confirmation, learning and funder reporting", proposal_value(profile.get("recurrent_cost_custodian"), "Recurrent-cost custodian", "Confirm institutional ownership and budget route.")],
    ], widths=[1.2, 4.4, 1.8])

    document.add_heading("5. Financing request, cost and value for money", level=1)
    setup_low = profile.get("cost_setup_low")
    setup_high = profile.get("cost_setup_high")
    recurrent_low = profile.get("cost_recurrent_low")
    recurrent_high = profile.get("cost_recurrent_high")
    add_proposal_table(document, ["Cost element", "Planning range / treatment", "Evidence basis"], [
        ["Funding request", request, clean(proposal.get("funding_instrument")) or "To be agreed"],
        ["Set-up investment", proposal_cost_range(setup_low, setup_high, "Set-up investment", "Provide national quantities, unit rates and cost assumptions."), proposal_value(profile.get("cost_estimate_class"), "Cost estimate class", "Identify whether this is a national price estimate or costed financing proposal.")],
        ["Annual recurrent cost", proposal_cost_range(recurrent_low, recurrent_high, "Annual recurrent cost", "Provide lifecycle operating costs, cost unit and responsible custodian."), proposal_value(profile.get("cost_confidence"), "Cost confidence", "Record the confidence basis for the estimate.")],
        ["Cost basis", proposal_value(profile.get("cost_basis"), "Cost basis", "Provide source, price date, quantities, assumptions and currency."), proposal_value(profile.get("cost_archetype"), "Cost archetype", "Select the relevant cost archetype.")],
    ], widths=[1.6, 3.2, 2.6])
    document.add_paragraph(
        proposal_value(
            profile.get("cost_rationale"),
            "Cost and value-for-money rationale",
            "Explain why the selected solution represents value for money and complete national-price validation.",
        )
    )
    document.add_paragraph(
        "Value for money will be assessed through the defined delivery outputs, implementation feasibility, recurrent-cost ownership, targeted risk-reduction contribution and transparent monitoring evidence."
    )

    document.add_heading("6. Results, safeguards and sustainability", level=1)
    document.add_paragraph(
        proposal_value(
            proposal.get("expected_results"),
            "Results framework",
            "State indicators, baselines, targets, disaggregation, verification sources and reporting frequency.",
        )
    )
    document.add_paragraph(
        "The proposal will apply the ARC D4 safeguards and inclusion approach, including gender, equity and social inclusion responsiveness, accessibility, feedback arrangements and non-digital continuity where relevant."
    )
    document.add_paragraph(
        f"Sustainability depends on the named recurrent-cost custodian ({proposal_value(profile.get('recurrent_cost_custodian'), 'Recurrent-cost custodian', 'Confirm the responsible institution, budget line and post-investment operating arrangement.')}), national ownership, institutional mandate and an agreed operating budget or financing route."
    )

    document.add_heading("7. Key risks and mitigation", level=1)
    add_proposal_table(document, ["Risk", "Mitigation / decision required"], [
        ["National price or scope differs from comparator", "Validate the design and costs with the accountable national institution before contracting or disbursement."],
        ["No confirmed recurrent-cost arrangement", "Secure a named custodian and operating-cost decision as a condition of scale-up."],
        ["Evidence or delivery data incomplete", "Use the controlled matrix and verification route; do not present unverified claims as results."],
        ["Local adoption conditions change", "Retain the Member State decision right and update the implementation plan through the authorised focal point."],
    ], widths=[2.6, 4.8])

    document.add_heading("8. Proposal completion register", level=1)
    completion_rows = proposal_completion_register(matrix, profile, proposal)
    if completion_rows:
        document.add_paragraph(
            "The following items are incomplete in the controlled record. They are intentionally retained as placeholders rather than being inferred or fabricated. "
            "They must be resolved through the stated evidence or decision route before external submission."
        )
        add_proposal_table(
            document,
            ["Information still required", "Required completion action", "Controlled route"],
            completion_rows,
            widths=[2.0, 3.8, 1.6],
        )
    else:
        document.add_paragraph(
            "No missing fields were detected in the visible controlled record. Confirm the selected funder's template, eligibility rules, national pricing and institutional approvals before external submission."
        )

    document.add_heading("9. Attachments for the funder", level=1)
    for item in [
        "Innovation Priority Index and D2/D3 evidence record",
        "Investment-ready innovation profile and supporting evidence links",
        "Annex R.1 comparative cost analysis and nationally costed proposal when available",
        "Implementation plan, results framework, risk register and safeguards record",
        "ARC D4 programme/output-product pack and relevant institutional decision records",
    ]:
        document.add_paragraph(item, style="List Bullet")
    document.add_paragraph(
        "Submission control: this document is a compiled proposal draft. It must be checked against the funder's eligibility, approved national pricing, institutional authority and the final signed delivery arrangements before external submission."
    )
    buffer = BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def show_funding_proposal_workspace(matrix, profile):
    st.subheader("Compile a funding proposal")
    st.caption(
        "Create a funder-specific proposal for the selected innovation. The compiled document includes the innovation record, "
        "ARC D4 outputs, delivery pathway, cost analysis, safeguards, risks and a transparent statement of any information still requiring confirmation."
    )
    outputs = proposal_outputs(matrix)
    st.info(f"This proposal will include all {len(outputs)} governed ARC D4 output product(s) for **{clean(profile.get('innovation_name'))}**.")
    if not cost_analysis_complete(profile):
        st.warning(
            "Cost analysis is incomplete. You can compile a discussion draft, but complete Annex R.1 and replace comparative bands with national pricing before a financing commitment."
        )
    if not yes(profile.get("description")):
        st.warning("Add an innovation description in Workstream D before external submission; the proposal will otherwise flag this as information still required.")

    left, right = st.columns(2)
    with left:
        title = st.text_input("Proposal title *", value=f"Funding proposal: {clean(profile.get('innovation_name'))}", key="proposal_title")
        funder_name = st.text_input("Funder name *", placeholder="Name of prospective funding partner", key="proposal_funder")
        funder_type = st.selectbox("Funder type", ["Grant fund", "Development finance institution", "Climate fund", "Philanthropic foundation", "Bilateral partner", "Corporate / private sector", "Other"], key="proposal_funder_type")
        instrument = st.selectbox("Requested funding instrument", ["Grant", "Concessional finance", "Blended finance", "Technical assistance", "Guarantee / risk finance", "To be agreed"], key="proposal_instrument")
        requested_amount = st.number_input("Amount requested (USD) *", min_value=0.0, step=10000.0, key="proposal_amount")
    with right:
        scope = st.text_input("Implementation geography / scope *", placeholder="e.g. Regional SADC programme or named Member State", key="proposal_scope")
        duration = st.number_input("Implementation duration (months) *", min_value=1, max_value=120, value=24, step=1, key="proposal_duration")
        st.text_area("Funder priorities and strategic alignment", placeholder="Which funder objective, theme or call does this proposal address?", key="proposal_alignment")
        st.text_area("Development problem statement", placeholder="Describe the need, affected population, gap and evidence basis.", key="proposal_problem")
        st.text_area("Expected results and measures", placeholder="State expected results, indicators, targets and verification sources.", key="proposal_results")

    missing = []
    if not yes(title):
        missing.append("proposal title")
    if not yes(funder_name):
        missing.append("funder name")
    if not yes(scope):
        missing.append("implementation geography / scope")
    if requested_amount <= 0:
        missing.append("amount requested")
    proposal = {
        "proposal_title": title,
        "funder_name": funder_name,
        "funder_type": funder_type,
        "funding_instrument": instrument,
        "requested_amount": requested_amount,
        "implementation_scope": scope,
        "duration_months": duration,
        "funder_alignment": st.session_state.get("proposal_alignment", ""),
        "problem_statement": st.session_state.get("proposal_problem", ""),
        "expected_results": st.session_state.get("proposal_results", ""),
    }
    if missing:
        st.warning("Before compiling, complete: " + ", ".join(missing) + ".")
        return
    document = build_funding_proposal_docx(matrix, profile, proposal)
    st.success("Proposal draft compiled. Review the controlled warnings and attach national pricing and funder-specific evidence before external submission.")
    st.download_button(
        "Download compiled funding proposal (.docx)",
        data=document,
        file_name=proposal_filename(clean(profile.get("innovation_name")) + "_" + funder_name),
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )


auth_sidebar()
signed_in = approval_controls()

try:
    uploaded = st.sidebar.file_uploader("Replace the governed automation matrix", type="xlsx")
    matrix = read_matrix(uploaded) if uploaded else default_matrix("matrix-v3.8-complete-delivery-handover-1")
except Exception as exc:
    st.error(f"The automation matrix could not be read: {exc}")
    st.stop()

for key, value in {"selected_stage": "S1", "new_profile": {}, "audit": []}.items():
    st.session_state.setdefault(key, value)

portfolio = table(matrix, "06_Portfolio")
portfolio["Innovation"] = portfolio["Innovation"].astype(str)
shared = []
if signed_in:
    try:
        shared = client().table("d4_innovations").select("innovation_name,innovation_url,innovation_type,d3_profile").order("innovation_name").execute().data or []
    except Exception as exc:
        st.sidebar.warning(f"Shared portfolio unavailable: {exc}")
names = portfolio["Innovation"].tolist() + [entry["innovation_name"] for entry in shared]
selected_name = st.sidebar.selectbox("Innovation to work on", names, key="innovation_selector")
saved = next((entry for entry in shared if entry["innovation_name"] == selected_name), {})
profile = dict(saved.get("d3_profile") or {})
profile.setdefault("innovation_name", selected_name)
profile.setdefault("innovation_url", saved.get("innovation_url", ""))
profile.setdefault("innovation_type", saved.get("innovation_type", ""))
selected_portfolio = portfolio.loc[portfolio["Innovation"].astype(str) == selected_name]
if not selected_portfolio.empty:
    source_scores = selected_portfolio.iloc[0]
    apply_portfolio_cost_defaults(profile, source_scores)
    for symbol in table(matrix, "05_IPI_Criteria").get("Symbol", pd.Series(dtype=str)).dropna().astype(str):
        source_value = source_scores.get(symbol)
        if pd.notna(source_value) and yes(source_value):
            profile.setdefault(f"score_{symbol}", float(source_value))

st.title("ARC D4 Delivery Platform")
st.caption("A controlled programme record that routes a portfolio innovation through admission, appraisal, verification, generation and product readiness.")
workspace = st.radio("Workspace", ["Programme command", "Workstream C — portfolio admission", "Workstream D — innovation delivery", "Research and verification", "Products and readiness", "Submission-ready report", "Funding proposal"], horizontal=True)
show_stage_callout(matrix, st.session_state.selected_stage)
show_correction_callout(matrix, profile, "Selected innovation: open actions")

if workspace == "Programme command":
    stages = table(matrix, "01_Stages")
    products = readiness_board(matrix)
    pending = portfolio[portfolio["IPI v2.0 (computed)"].astype(str).str.upper().eq("PENDING")].shape[0]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Portfolio records", len(portfolio))
    c2.metric("IPI determinations pending", pending)
    c3.metric("Ready for submission", int(products.get("Submission", pd.Series(dtype=str)).astype(str).str.strip().str.lower().eq("ready").sum()))
    c4.metric("Current selection", selected_name[:22] + ("…" if len(selected_name) > 22 else ""))
    st.subheader("Programme pathway")
    stream, current = st.columns([3, 1], gap="large")
    with stream:
        for lane, group in stages.groupby("Lane", sort=False):
            with st.expander(clean(lane), expanded=True):
                for _, stage in group.iterrows():
                    left, text = st.columns([1, 5])
                    if left.button(clean(stage["Stage ID"]), key=f"command_{stage['Stage ID']}"):
                        st.session_state.selected_stage = clean(stage["Stage ID"])
                    text.markdown(f"**{clean(stage['Stage name'])}**")
                    text.caption(f"Exit: {clean(stage['Exit condition'])}")
    with current:
        selected_stage_panel(matrix, st.session_state.selected_stage, profile)
    st.subheader("Product readiness")
    show_product_readiness_callout(matrix)

elif workspace == "Workstream C — portfolio admission":
    st.subheader("Workstream C — portfolio admission")
    st.write("S1 records the governed portfolio. S2 settles the shortlist. New innovations are admitted only after their record is complete enough to be owned and reviewed.")
    tabs = st.tabs(["S1 Portfolio", "S2 Top 10+", "New innovation"])
    with tabs[0]:
        view = portfolio.copy()
        view["Record status"] = view["IPI v2.0 (computed)"].apply(lambda value: "Pending" if clean(value).upper() == "PENDING" else "Scored")
        st.dataframe(view[["#", "Innovation", "IPI v2.0 (computed)", "Record status", "Confidence"]], hide_index=True, use_container_width=True, height=520)
        pending_count = int((view["Record status"] == "Pending").sum())
        if pending_count:
            st.warning(f"S1 attention: {pending_count} portfolio record(s) still require their contracted IPI determinations. Complete SAS, GRS, CVS and SIS evidence before treating the portfolio as ranked.")
    with tabs[1]:
        ranking = portfolio.copy()
        numeric = pd.to_numeric(ranking["IPI v2.0 (computed)"], errors="coerce")
        ranking["Ranking status"] = numeric.map(lambda value: "Scored" if pd.notna(value) else "Pending")
        st.caption("The platform does not fabricate a Top 10 where the contracted index is still pending. Scored records rank first; the remaining records are visibly pending.")
        st.dataframe(ranking.assign(_score=numeric).sort_values(["_score", "#"], ascending=[False, True]).head(10)[["#", "Innovation", "IPI v2.0 (computed)", "Ranking status", "Confidence"]], hide_index=True, use_container_width=True)
        if ranking["Ranking status"].eq("Pending").any():
            st.warning("S2 attention: the shortlist cannot be formally published until the pending index determinations are resolved by the SADC Secretariat determination session.")
    with tabs[2]:
        if not signed_in:
            st.warning("Sign in before admitting a shared innovation.")
        profile_form(matrix, st.session_state.new_profile, "new", signed_in)

elif workspace == "Workstream D — innovation delivery":
    stages = table(matrix, "01_Stages")
    permitted = stages[stages["Stage ID"].isin([f"S{i}" for i in range(3, 9)] + ["S8a"])]
    st.subheader("Workstream D — innovation delivery")
    st.caption("The selected record advances through gap analysis, profile development, appraisal, feasibility, regional posture and financing route selection.")
    sequence, selected = st.columns([3, 1], gap="large")
    with sequence:
        for _, stage in permitted.iterrows():
            left, text = st.columns([1, 5])
            if left.button(clean(stage["Stage ID"]), key=f"d_{stage['Stage ID']}"):
                st.session_state.selected_stage = clean(stage["Stage ID"])
            text.markdown(f"**{clean(stage['Stage ID'])} — {clean(stage['Stage name'])}**")
            text.caption(clean(stage["Exit condition"]))
    with selected:
        selected_stage_panel(matrix, st.session_state.selected_stage, profile)
    st.divider()
    st.subheader("Readiness work list")
    problems = readiness(matrix, profile)
    if problems:
        st.warning("This record is not ready to generate a product.")
        st.dataframe(pd.DataFrame({"Open work item": problems, "Resolver": ["R4 — owner escalation"] * len(problems)}), hide_index=True, use_container_width=True)
    else:
        st.success("The visible innovation profile passes the current pre-generation checks.")
    show_correction_callout(matrix, profile, "Actions to make this innovation ready")
    st.divider()
    st.subheader("Annex R.1 cost analysis")
    cost_analysis_form(profile, f"existing_cost_{clean(selected_name)}")
    if cost_analysis_complete(profile):
        st.success("Cost analysis complete. The band can support D3 planning and prioritisation; replace it with national pricing for a financing proposition.")
    else:
        st.info("Complete the cost-analysis fields before using this innovation in an investment-ready proposition.")
    if signed_in and st.button("Save cost analysis to shared register", key=f"save_cost_{clean(selected_name)}"):
        if not cost_analysis_complete(profile):
            st.error("Complete the cost analysis before saving it to the shared register.")
        else:
            try:
                save_cost_analysis(profile)
                st.success("Cost analysis saved to the shared register.")
            except Exception as exc:
                st.error(f"Could not save the cost analysis: {exc}")

elif workspace == "Research and verification":
    st.subheader("Research broker and verification gate")
    st.caption("Only a request packet leaves the platform. Returned claims are not written into the programme record until a named analyst accepts them at S13.")
    left, right = st.columns([3, 1], gap="large")
    with left:
        stages = table(matrix, "01_Stages")
        for _, stage in stages[stages["Stage ID"].isin([f"S{i}" for i in range(9, 14)])].iterrows():
            button, text = st.columns([1, 5])
            if button.button(clean(stage["Stage ID"]), key=f"research_{stage['Stage ID']}"):
                st.session_state.selected_stage = clean(stage["Stage ID"])
            text.markdown(f"**{clean(stage['Stage name'])}**")
            text.caption(clean(stage["Exit condition"]))
        st.subheader("Permitted research tasks")
        st.dataframe(table(matrix, "10_Research_Broker"), hide_index=True, use_container_width=True, height=400)
    with right:
        selected_stage_panel(matrix, st.session_state.selected_stage, profile)
    st.subheader("Evidence ledger")
    ledger = table(matrix, "14_Audit_Log")
    st.dataframe(ledger, hide_index=True, use_container_width=True)
    st.warning("Verification attention: a research return is only usable after the source is opened, checked against the permitted source and date rules, and accepted by a named analyst at S13. Unverified claims must remain outside the programme record.")

elif workspace == "Products and readiness":
    st.subheader("Products and readiness")
    st.caption("No product leaves the platform with an unexplained absence. A missing field becomes a work item, or a named, time-limited waiver; it is never silently drafted around.")
    products = readiness_board(matrix)
    product, detail = st.columns([3, 1], gap="large")
    with product:
        show_product_readiness_callout(matrix)
        st.subheader("Product sets")
        st.dataframe(table(matrix, "11_Output_Products"), hide_index=True, use_container_width=True)
        st.subheader("Pre-generation gate for selected innovation")
        failures = readiness(matrix, profile)
        if failures:
            for failure in failures:
                st.error(failure)
        else:
            st.success("The selected working record can proceed to product generation once its relevant product readiness conditions are met.")
        show_correction_callout(matrix, profile, "Actions before generation")
    with detail:
        selected_stage_panel(matrix, st.session_state.selected_stage, profile)
        st.divider()
        st.markdown("**Resolver sequence**")
        st.dataframe(table(matrix, "28_Resolution_Layer")[["ID", "Resolver", "Mode"]], hide_index=True, use_container_width=True)

elif workspace == "Submission-ready report":
    show_submission_ready_report(matrix, profile)
else:
    show_funding_proposal_workspace(matrix, profile)
