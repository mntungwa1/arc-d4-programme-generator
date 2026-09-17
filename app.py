"""ARC D4 Delivery Platform — controlled portfolio-to-product workflow."""
from __future__ import annotations

import base64
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from supabase import create_client

st.set_page_config(page_title="ARC D4 Delivery Platform", page_icon="◆", layout="wide")

SUPABASE_URL = "https://wrejrxzgyuxsfbxutezg.supabase.co"
SUPABASE_KEY = "sb_publishable_UfYoG2ZgKP0nLA5KGwEG6w_2rCP9_R8"
APP_URL = "https://arc-d4-programme-generator-c7qdwgesqvnwjafgvpxgat.streamlit.app/"
ADMIN_EMAILS = {"dingaan@academyrc.co.za", "drcliff@academyrc.co.za"}
BOOK = Path(__file__).parent / "ARC_D4_Automation_Matrix.b64"


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
        "Set ID", "ID", "Product", "Symbol", "Template",
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
def default_matrix(cache_version="matrix-v3.1-headerfix-2"):
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


def profile_form(matrix, profile, name_key, allow_save):
    st.subheader("Investment-ready innovation profile")
    st.caption("This is the governed working record. It creates owned work instead of hiding an unresolved requirement.")
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
            score = st.number_input(f"{code} (0–10)", 0.0, 10.0, float(profile.get(f"score_{code}", 0.0)), 0.5, key=f"{name_key}_{code}")
        profile[f"score_{code}"] = score
        weighted += score * weight
        complete = complete and score > 0
    profile["ipi"] = round(weighted, 2)
    st.metric("Innovation Priority Index", f"{profile['ipi']:.2f} / 10", "Pending determinations" if not complete else "All determinations entered")
    if allow_save and st.button("Admit innovation to shared portfolio", type="primary"):
        required = ["innovation_name", "description", "innovation_type", "lead_institution"]
        missing = [item.replace("_", " ") for item in required if not yes(profile.get(item))]
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
    show_correction_callout(matrix, profile, "Profile completion actions")


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
        pending = [clean(row["Symbol"]) for _, row in criteria.iterrows() if float(profile.get(f"score_{clean(row['Symbol'])}", 0) or 0) <= 0]
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
    st.warning(f"{title}: {len(actions)} item(s) require action before this record is ready.")
    with st.expander("View corrective actions", expanded=True):
        st.dataframe(pd.DataFrame(actions), hide_index=True, use_container_width=True)


def show_product_readiness_callout(products):
    if products.empty or "Readiness" not in products.columns:
        return
    outstanding = products.loc[~products["Readiness"].astype(str).str.upper().str.startswith("READY")].copy()
    if outstanding.empty:
        st.success("All configured products are ready.")
        return
    st.warning(f"Product delivery attention: {len(outstanding)} product(s) are not ready to generate.")
    columns = [column for column in ["Product", "Name", "Blocking fields and how they resolve", "Resolver", "Readiness"] if column in outstanding.columns]
    with st.expander("View product-specific actions", expanded=False):
        st.dataframe(outstanding[columns], hide_index=True, use_container_width=True)


auth_sidebar()
signed_in = approval_controls()

try:
    uploaded = st.sidebar.file_uploader("Replace the governed automation matrix", type="xlsx")
    matrix = read_matrix(uploaded) if uploaded else default_matrix("matrix-v3.1-headerfix-2")
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

st.title("ARC D4 Delivery Platform")
st.caption("A controlled programme record that routes a portfolio innovation through admission, appraisal, verification, generation and product readiness.")
workspace = st.radio("Workspace", ["Programme command", "Workstream C — portfolio admission", "Workstream D — innovation delivery", "Research and verification", "Products and readiness"], horizontal=True)

if workspace == "Programme command":
    stages = table(matrix, "01_Stages")
    products = table(matrix, "32_Product_Readiness")
    pending = portfolio[portfolio["IPI v2.0 (computed)"].astype(str).str.upper().eq("PENDING")].shape[0]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Portfolio records", len(portfolio))
    c2.metric("IPI determinations pending", pending)
    c3.metric("Product packs ready", int(products["Readiness"].astype(str).str.startswith("READY").sum()))
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
    st.dataframe(products[["Product", "Name", "Readiness"]], hide_index=True, use_container_width=True)
    show_product_readiness_callout(products)

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

else:
    st.subheader("Products and readiness")
    st.caption("No product leaves the platform with an unexplained absence. A missing field becomes a work item, or a named, time-limited waiver; it is never silently drafted around.")
    products = table(matrix, "32_Product_Readiness")
    product, detail = st.columns([3, 1], gap="large")
    with product:
        st.dataframe(products, hide_index=True, use_container_width=True)
        show_product_readiness_callout(products)
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
