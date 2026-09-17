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
def default_matrix(cache_version="matrix-v3.5-three-tier-readiness-1"):
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
        st.warning(f"{title}: {len(actions)} open item(s).")
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
        waiting_on = clean(product.get("Who we are waiting on"))
        with st.container(border=True):
            st.markdown(f"### {product_id} — {product_name}")
            left, right = st.columns(2)
            left.success(f"Working draft: {working_draft or 'Ready to produce'}")
            if submission.lower() == "ready":
                right.success("Submission: Ready")
            else:
                right.warning(f"Submission: {submission or 'Awaiting readiness assessment'}")
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
                    st.error(
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


auth_sidebar()
signed_in = approval_controls()

try:
    uploaded = st.sidebar.file_uploader("Replace the governed automation matrix", type="xlsx")
    matrix = read_matrix(uploaded) if uploaded else default_matrix("matrix-v3.5-three-tier-readiness-1")
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
    for symbol in table(matrix, "05_IPI_Criteria").get("Symbol", pd.Series(dtype=str)).dropna().astype(str):
        source_value = source_scores.get(symbol)
        if pd.notna(source_value) and yes(source_value):
            profile.setdefault(f"score_{symbol}", float(source_value))

st.title("ARC D4 Delivery Platform")
st.caption("A controlled programme record that routes a portfolio innovation through admission, appraisal, verification, generation and product readiness.")
workspace = st.radio("Workspace", ["Programme command", "Workstream C — portfolio admission", "Workstream D — innovation delivery", "Research and verification", "Products and readiness"], horizontal=True)
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
