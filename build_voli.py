"""
Build CITS Volume I — Technical Volume (expanded draft).

RFQ 75D301-26-Q-78845 — CDC Customer Information Technology Support (CITS)
Page limit per RFQ Section E: 35 pages, 11pt Times New Roman equivalent, 0.75" margins.
Target: ~33 pages of body to leave layout safety margin.
"""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_shading(cell, fill_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    tc_pr.append(shd)


def style_doc(doc):
    # Page setup
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
    # Default style
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    # Headings
    for level, size, bold in [(1, 14, True), (2, 12, True), (3, 11, True), (4, 11, True)]:
        st = doc.styles[f"Heading {level}"]
        st.font.name = "Times New Roman"
        st.font.size = Pt(size)
        st.font.bold = bold
        st.font.color.rgb = RGBColor(0x00, 0x00, 0x00)


def add_para(doc, text, style="Normal", bold_lead=None):
    p = doc.add_paragraph(style=style)
    if bold_lead:
        run = p.add_run(bold_lead)
        run.bold = True
        if text:
            p.add_run(" " + text)
    else:
        p.add_run(text)
    return p


def add_h(doc, text, level):
    return doc.add_heading(text, level=level)


def add_table(doc, header, rows, col_widths=None, header_fill="D9E1F2"):
    t = doc.add_table(rows=1 + len(rows), cols=len(header))
    t.style = "Light Grid Accent 1"
    # Header
    for j, h in enumerate(header):
        c = t.cell(0, j)
        c.text = ""
        run = c.paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(10)
        set_cell_shading(c, header_fill)
    # Body
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            c = t.cell(i, j)
            c.text = ""
            r = c.paragraphs[0].add_run(str(val))
            r.font.size = Pt(10)
    if col_widths:
        for j, w in enumerate(col_widths):
            for cell in t.columns[j].cells:
                cell.width = Inches(w)
    return t


def add_placeholder(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(f"[PLACEHOLDER: {text}]")
    r.italic = True
    r.font.color.rgb = RGBColor(0x70, 0x70, 0x70)
    return p


# ======================================================================
doc = Document()
style_doc(doc)

# COVER ----------------------------------------------------------------
title = doc.add_paragraph()
title_run = title.add_run("VOLUME I — TECHNICAL VOLUME")
title_run.bold = True
title_run.font.size = Pt(16)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

sub = doc.add_paragraph()
sr = sub.add_run("RFQ 75D301-26-Q-78845 — CDC Customer Information Technology Support (CITS)")
sr.bold = True
sr.font.size = Pt(12)
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

meta = doc.add_paragraph()
meta.add_run("Submitted by ").italic = True
meta.add_run("[PLACEHOLDER: Prime legal entity as registered for Alliant 2] · "
             "[PLACEHOLDER: CAGE code, UEI] · Submission date: June 1, 2026").italic = True
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_paragraph()  # spacer

# =======================================================================
# 1. TECHNICAL APPROACH
# =======================================================================
add_h(doc, "1. Technical Approach", 1)

add_para(doc,
    "Our governing approach to CITS: keep what works, deflect demand, and modernize where it pays back. "
    "CDC's IT customer support model — a tiered ITSM service-delivery structure built on ServiceNow, ITIL, and "
    "PMBOK — is sound. The opportunity is not to replace it. It is to take excess demand out of the system, "
    "automate the repetitive work that fills technicians' queues today, and direct human attention to the work "
    "where human attention matters: emergency response, scientific computing, global health field deployments, "
    "and OCONUS network restoration. This Technical Volume describes how we will assume performance, sustain "
    "CDC's published service levels from day one, and deliver consistent total cost of ownership (TCO) reduction "
    "across the period of performance. We address each task in the Performance Work Statement (PWS), identify "
    "the operational baseline we will inherit, name the specific automation and workflow improvements we will "
    "pursue, set quantified targets, and identify the risks and dependencies the Government should expect us "
    "to manage.")

add_para(doc,
    "This volume is organized to mirror the evaluation structure in Section E. Section 1 is our Technical "
    "Approach, organized as a reading of the requirement (1.1), a Work Plan with milestones (1.2), task-by-task "
    "approaches (1.3), the cross-cutting service disciplines (1.4) including Human-Centered Design (HCD) as our "
    "operating model, Section 508 accessibility (1.5), cybersecurity and compliance posture (1.6), three "
    "test-and-learn pilots scoped to the demand profile (1.7), and a quantified automation and deflection "
    "roadmap with year-over-year targets (1.8). Section 2 is the Staffing Plan. Section 3 is the Management "
    "Approach including the draft Quality Control Plan, draft Transition-In and Transition-Out Plans, and the "
    "GFE accountability approach. Section 4 is Similar Experience.")

# 1.1 Understanding ------------------------------------------------------
add_h(doc, "1.1 Understanding of the Requirement and Operational Baseline", 2)

add_para(doc,
    "CDC OCIO supports approximately 28,000 personnel globally — including approximately 1,275 OCIO contractors "
    "and 375 federal employees — across CONUS and OCONUS sites, ranging from the Roybal and Chamblee campuses "
    "in Atlanta to roughly 60 international Global Health Center locations and Field Site Service postings. "
    "The published incident and task data in PWS Section 4 are the operational baseline we will inherit on "
    "day one. Tables 3 and 4 of the PWS show that, between October 2023 and December 2024, CDC processed "
    "approximately 42,000 incidents and 27,000 tasks. The shape of that volume is consequential and it directs "
    "where automation pays back.")

# Demand profile table
add_para(doc, "Table 1.1-A. Demand profile (CDC-published, Oct 2023 – Dec 2024).", bold_lead="")
add_table(doc,
    header=["Driver", "Volume", "% of category", "Deflection / automation lever"],
    rows=[
        ["Network password reset", "11,459 tasks", "41.8% of tasks",
         "ServiceNow Virtual Agent self-service + audited reset workflow"],
        ["UserID provisioning (rules met)", "6,137 tasks", "22.4% of tasks",
         "Templated request, conditional approval, automated fulfillment"],
        ["PIV card exception request", "3,087 tasks", "11.3% of tasks",
         "Workflow automation + RPA for repetitive verification steps"],
        ["Outlook incidents", "4,357 incidents", "10.3% of incidents",
         "Plain-language KB articles + Tier-1 macros"],
        ["Laptop incidents", "3,590 incidents", "8.5% of incidents",
         "Zero-touch refresh + Dell Command + MECM compliance"],
        ["Smart Card / MFA / VDI", "~7,200 incidents", "~17% of incidents",
         "KB + Predictive Intelligence assignment + Zscaler/Azure runbooks"],
    ],
    col_widths=[1.6, 1.0, 1.1, 2.7])

add_para(doc,
    "These six categories represent more than three quarters of total inbound volume. Our technical approach "
    "is calibrated to that distribution. We do not attempt uniform automation across all task categories; we "
    "sequence improvements where the operational return is largest and where ServiceNow-native workflow and "
    "knowledge management already provide the platform.")

add_para(doc,
    "Three constraints set the boundary on how we approach this scope. First, ServiceNow is CDC's platform of "
    "record for ITSM, and CDC will continue to provide it; we build on it rather than around it. Second, the "
    "Section E proposal preparation instructions prohibit AI products in this proposal and prohibit the use of "
    "AI products after award except those explicitly provided, approved, and made available by CDC. Our Work "
    "Plan therefore relies on ServiceNow workflow, ServiceNow Virtual Agent, knowledge management, robotic "
    "process automation (RPA), and ServiceNow-native analytics. Where future CDC-approved capabilities expand "
    "the available toolset, our continuous-improvement cadence will incorporate them on CDC's authority. Third, "
    "the Customer Service Survey performance metric (96% or greater per PWS Section 9) and the Customer "
    "Personnel Availability metric (80% minimum daily) discipline how we sequence change. We do not run pilots "
    "or transformations that put either metric at risk.",
    bold_lead="Constraints we are designing around.")

add_para(doc,
    "First, demand surge during public health emergencies — the EOC support model and Field Site Service posture in "
    "PWS Task 1.3 cannot be staffed at peak as a baseline. Our approach uses the ServiceNow Virtual Agent for "
    "password and standard-request deflection so that human technician capacity is preserved for surge response, "
    "plus a defined cross-task surge plan from the optional Task 4.1 rapid surge staffing capability. "
    "Second, OCONUS network reliability — Global Activities under PWS Task 2.1 supports approximately 1,800 staff "
    "in roughly 60 international locations on satellite and SD-WAN links where Priority 1 outages cascade. Our "
    "approach pairs CCNA-certified engineers with regional operating procedures and documented deployment "
    "readiness so that 24-to-48 hour deployment windows are met without ad hoc mobilization. "
    "Third, knowledge fragmentation — the highest leverage in the demand profile is in deflection, and deflection "
    "requires a knowledge base that is current, findable, and maintained. Our approach treats knowledge management "
    "as an explicit, measured discipline under PWS Subsection 3.2.3 with named ownership in Task 3.1, not a "
    "residual artifact of ticket closure.",
    bold_lead="Three challenges we anticipate, with how we will address them.")

add_para(doc,
    "Three guarantees frame everything that follows. (1) No SLA exemptions for the new contractor at "
    "performance start — we will meet Table 5 thresholds beginning September 1, 2026. (2) No new tools introduced "
    "without CDC approval — our automation runs inside ServiceNow and the Microsoft platforms CDC already owns. "
    "(3) No AI products in the proposal or in performance, except those CDC explicitly provides and approves — "
    "in compliance with Section E.",
    bold_lead="Three guarantees we make to CDC.")

# 1.2 Work Plan -----------------------------------------------------------
add_h(doc, "1.2 Work Plan and Milestones", 2)

add_para(doc,
    "The Work Plan correlates to the Volume II hour build and is structured to align with CDC's contractual "
    "periods of performance: a one-month transition (8/1/2026 – 8/31/2026) priced firm-fixed-price, a five-month "
    "base period (9/1/2026 – 1/31/2027) priced time-and-materials, four 12-month options, and a six-month "
    "extension priced for evaluation under FAR 52.217-8. The milestone chart in Figure 1.2-A summarizes the "
    "sequence; Table 1.2-B identifies the contractual milestones, deliverables, and decision gates with "
    "named owners.")

add_placeholder(doc,
    "Insert Figure 1.2-A — Master Milestone Chart. Proposal graphics team to render Gantt-style timeline "
    "showing Transition, Base, and Option Periods 1–4 with vertical milestone markers, decision gates, "
    "key personnel onboarding waves, and pilot windows. Visual replaces 1–2 paragraphs of dense prose and "
    "supports Section E (1)(b) requirement for milestone charts.")

add_para(doc, "Table 1.2-B. Contractual milestones, deliverables, and decision gates.", bold_lead="")
add_table(doc,
    header=["Window", "Milestone / Deliverable", "PWS / Section", "Owner"],
    rows=[
        ["Award + 5 BD", "Post-award kickoff convened with CO/COR and Technical Monitors",
         "PWS 3.1, Subtask 3.1.1", "Program Manager"],
        ["Award + 10 BD", "On-call rotation schedules; Transition Plan v1; QCP draft v1",
         "PWS §8 deliverables", "PM / QC Lead"],
        ["Transition (Aug)", "Knowledge transfer; ServiceNow shadowing; GFE inventory walk-down",
         "PWS 3.1.2; §5 GFE", "Transition-In PM"],
        ["Sep 1, 2026", "Full performance start. SLAs apply at Table 5 thresholds, no exemptions",
         "PWS §9", "PM / Task Leads"],
        ["Award + 30 cal.", "Final QCP submitted to CO",
         "PWS §9.1", "QC Lead"],
        ["Monthly (15th)", "Monthly Progress Report; Monthly Financial Report; In-Progress Review",
         "PWS §8 Table 12", "PM"],
        ["Weekly (Wed 5pm ET)", "Weekly Status Report",
         "PWS §8", "PM"],
        ["Quarterly", "Performance trend review with COR; KB review cycle; pilot exit / extend gates",
         "PWS 3.2.6", "PM / QC Lead"],
        ["Annually", "Wall-to-wall property inventory; cyber & privacy training attestations",
         "PWS Task 1; §12", "ITAM Lead / Cyber Lead"],
        ["End of Option 4", "Transition-Out window (3 mo overlap if exercised under FAR 52.217-8)",
         "PWS 3.1.2 item 3", "Transition-Out PM"],
    ],
    col_widths=[1.0, 2.6, 1.4, 1.4])

# 1.3 Approach by Task ---------------------------------------------------
add_h(doc, "1.3 Approach by Task", 2)

add_para(doc,
    "Each subsection below addresses a PWS task or task family. We address all required tasks and note our "
    "approach to optional tasks the Government may exercise. Tasks share a common operating model — ServiceNow "
    "as the system of record, knowledge management as a measured discipline, defined escalation paths, and "
    "named accountability — so we describe shared elements once in §1.4 and then treat each task as a delta "
    "from the common baseline. Each task subsection follows the same structure: operational baseline, operating "
    "model, automation and deflection plan, staffing and credentials, SLA performance posture, and "
    "task-specific risks with mitigations.")

# ---- 1.3.1 Task 1.1 ITSDS (BIG TASK) ----------------------------------
add_h(doc, "1.3.1 Task 1.1 — IT ServiceDesk Services (ITSDS)", 3)

add_para(doc,
    "Multi-channel Tier 1 customer support for all CDC personnel across phone, ServiceNow portal, email, and "
    "chat. The ITSDS function carries the largest single share of inbound volume on CITS and is the gateway "
    "through which most customer experience judgments about IT support are formed. Our approach treats ITSDS "
    "as both a service operation and the primary channel for demand reduction across the whole contract.",
    bold_lead="Objective.")

add_para(doc,
    "From PWS Section 4 and Table 3, the historical category mix is dominated by network password resets "
    "(approximately 41.8% of tasks), UserID provisioning (22.4%), PIV exception handling (11.3%), and the "
    "M365 / endpoint incident set led by Outlook, Teams, OneDrive, MFA, Zscaler, and CITGO Virtual Desktop. "
    "Roughly two-thirds of inbound task volume is, in principle, deflectable to verified self-service. The "
    "Section E ramp standard requires qualified staff within seven days of award or option start, and PWS "
    "Section 9 sets a Customer Personnel Availability minimum of 80% daily and a Customer Service Survey "
    "minimum of 96%. Those numbers are the floor — not the goal.",
    bold_lead="Operational baseline.")

add_para(doc,
    "We staff Tier 1 with a mixed labor-category model — junior, journeyman, senior, and SME — sized to the "
    "Section E historical level-of-effort table. The contact center runs on the CDC-provided NICE Virtual "
    "Contact Center (VCC) system, with ServiceNow as the single system of record for ticket data, work notes, "
    "knowledge articles, and customer satisfaction surveys. Calls and chats are triaged in ServiceNow; "
    "first-contact resolution (FCR) is the explicit team objective; warm-handoff escalation to Tier 2 follows "
    "a documented script with confirmation back to the customer. We verify ticket closure with the requester "
    "before close, in compliance with PWS Subsection 1, item j. A dedicated Service Desk Lead, supported by "
    "shift supervisors and queue managers, owns adherence, occupancy, and quality coaching as standing weekly "
    "items.",
    bold_lead="Operating model.")

add_para(doc,
    "Network password resets and provisioning together account for approximately 64% of inbound task volume. "
    "We will direct the largest share of our demand-reduction effort here through four workstreams. "
    "(1) ServiceNow Virtual Agent self-service flows for password reset, supported by audited identity "
    "verification and reset workflow, with target deflection rising from a Base-period baseline measure to "
    "30% in OP1 and 45% by OP3 (measured as % of inbound password-reset contacts resolved without human "
    "interaction). "
    "(2) Standardized request templates for UserID provisioning that route to conditional approval and "
    "automated fulfillment where the rules are deterministic, removing manual handoffs from approximately "
    "60% of provisioning paths. "
    "(3) Targeted knowledge articles, written in plain language and indexed for findability, that resolve the "
    "top deflectable incident types — Outlook, Teams, OneDrive, MFA — at the customer-self-service layer. "
    "Articles are owned, dated, version-controlled, and reviewed quarterly under §1.4 knowledge-management "
    "discipline, with a target Knowledge Article Currency Rate of 95% (% of articles reviewed within 90 days). "
    "(4) Predictive Intelligence-style ticket categorization within ServiceNow (using CDC-provided platform "
    "features, not vendor-introduced AI) to reduce mis-routes and shorten time-to-assignment.",
    bold_lead="Demand reduction (highest-leverage area).")

add_para(doc,
    "Tier 1 is staffed across Computer Systems Analyst (SME), Senior / Journeyman / Junior Computer User "
    "Support Specialist, with a SME ServiceNow Analyst trio and a VCC Systems Administrator under Public "
    "Trust Level 5. Hardware Repair functions inside Task 1.1 — staffed across journeyman computer hardware "
    "engineers, junior user support specialists, journeyman document management specialists, and an IT "
    "Project Manager — run intake, repair, and disposition in ServiceNow with warranty status tracking, "
    "parts-and-labor recording, and serial-number-level audit trail. ITIL certification is preferred per "
    "Section E Minimum Qualifications; the team holds ITIL coverage at supervisor and SME levels.",
    bold_lead="Staffing and credentials.")

add_para(doc,
    "We meet PWS Section 9 thresholds from day one: monthly average speed-to-answer under two minutes, "
    "missed-call rate under 3%, error rate under 3%, Customer Personnel Availability at or above 80% daily, "
    "and Customer Service Survey at or above 96%. We propose two additional QCP metrics that go beyond "
    "the PWS thresholds: First-Contact Resolution rate (target 75% by OP1), and a Deflection Rate against "
    "deterministic categories (target 30%/45%/55% across OP1/OP2/OP3). Both are visible to the COR daily on "
    "the Task 10.5 reporting platform.",
    bold_lead="SLA performance posture.")

add_para(doc,
    "Three risks dominate at Task 1.1. "
    "(1) Surge from a public health emergency — mitigated by the deflection program (which preserves human "
    "capacity), the Task 4.1 surge bench, and a documented overflow protocol with the Insight Global staffing "
    "channel. "
    "(2) Knowledge-base decay — mitigated by named article ownership, the 95% Currency Rate metric, and "
    "deflection-rate visibility (a falling deflection rate is an early signal of KB decay). "
    "(3) VCC license / subscription scope ambiguity — flagged under §3.10 conditions and assumptions; we "
    "operate to PWS Task 1.1 item 6 which indicates CDC procures the VCC subscription.",
    bold_lead="Task-specific risks and mitigations.")

# ---- 1.3.2 Task 1.2 DSS (BIG) ------------------------------------------
add_h(doc, "1.3.2 Task 1.2 — Deskside Support Services (DSS)", 3)

add_para(doc,
    "Tier 2 deskside support across the Atlanta campuses (Century Center, Chamblee, Corporate Square, "
    "Lawrenceville, Roybal) and the non-Atlanta Customer Support Center sites (Anchorage, Cincinnati, "
    "Ft. Collins, Hyattsville, Morgantown, Pittsburgh, San Juan, Spokane, Denver, plus DSS-assigned "
    "Washington, DC), with audio-visual support to non-Atlanta locations. We meet the OCIO Incident SLA in "
    "PWS Table 5: 15-minute assignment and 2-hour resolution at Priority 1; 30-minute assignment and 4-hour "
    "resolution at Priority 2; M-F 8 AM – 4 PM Local windows for Priorities 3 and 4, all at the 95% "
    "Achieved-Expected-Results level or higher.",
    bold_lead="Coverage and SLA framing.")

add_para(doc,
    "Each campus has a dedicated Service Desk Liaison (SDL) — a Project Manager-class technician who serves "
    "as lead contractor working with the Campus Lead and Backup Campus Lead. SDLs assign incoming tickets, "
    "manage the on-hold queue with documented justifications under PWS Subsection 1.2.1 item e, and own "
    "daily ticket hygiene. SDLs are technically credentialed, not project-managers-in-name-only; technical "
    "depth is a non-negotiable selection criterion. Audio-visual support to non-Atlanta locations runs on "
    "a documented escalation path with named AV-trained technicians per region.",
    bold_lead="Service Desk Liaisons and on-site model.")

add_para(doc,
    "Image deployment, patching, and update verification are run through MECM, Software Center, and "
    "device-specific update utilities (Dell Command Update Tool or equivalent), with compliance validated on "
    "the SDT Info Tool or Government-designated equivalent before re-issue. Zero-touch deployment is our "
    "objective at refresh; per PWS Subsection 1.2.1 item l, no laptop or desktop is issued without verified "
    "patch and BIOS compliance. Coordination with Task 8.1 Endpoint Management is treated as a single "
    "operating relationship — the same compliance baselines, the same change windows, the same drift-detection "
    "feedback loop.",
    bold_lead="Endpoint discipline.")

add_para(doc,
    "Three workstreams reduce deskside burden. "
    "(1) Zero-touch refresh — pre-imaged, pre-patched, pre-encrypted devices delivered to the user via the "
    "Tier 1 fulfillment workflow, removing per-device touch time at refresh. Target: 80% of refreshes "
    "zero-touch by OP1, 95% by OP3. "
    "(2) Self-service software install — Software Center catalog maintained by Task 9.6 with deskside as "
    "consumer, not assembler. "
    "(3) Asset move/add/change request templates — for telework hardware swaps, AV bookings, and end-of-life "
    "pickups — that pre-stage logistics and remove deskside admin time. The Service Desk Liaisons review the "
    "DSS deflection metrics weekly and surface candidates for KB-driven Tier 1 absorption.",
    bold_lead="Automation and deflection plan.")

add_para(doc,
    "DSS is staffed primarily by Journeyman Computer User Support Specialists with Senior and Journeyman "
    "Document Management Specialists for property and AV documentation duties, plus Junior Network Computer "
    "Systems Administrators for image and connectivity escalations. The historical level-of-effort table "
    "publishes 40 Journeyman Computer User Support Specialists across DSS — we propose an LCAT mix matched "
    "to that envelope, weighted toward Roybal and Chamblee where ticket volume concentrates. Field SDL "
    "rotations preserve campus-specific institutional knowledge.",
    bold_lead="Staffing and credentials.")

add_para(doc,
    "We treat the 95% achieved-expected-results target as the QCP performance line. Above the PWS, we measure "
    "Mean Time to Resolution (MTTR) by priority and report it monthly; we measure Customer-Reported Reopen "
    "Rate as an early indicator of premature closure; and we measure Image Compliance Pre-Issue rate at "
    "100% (no laptop or desktop is issued without verified patch and BIOS compliance — that is a binary "
    "metric for this contract).",
    bold_lead="SLA performance posture.")

add_para(doc,
    "(1) Refresh-cycle pile-up causing throughput strain — mitigated by zero-touch refresh and stagger-by-"
    "campus scheduling. "
    "(2) Property accountability gap during high-tempo move/add/change — mitigated by Document Management "
    "Specialist coverage, ServiceNow asset linkage, and the 95% on-time delivery metric in PWS §9. "
    "(3) AV escalations at non-Atlanta sites — mitigated by named regional AV technicians, vendor playbook, "
    "and quarterly tabletop exercises.",
    bold_lead="Task-specific risks and mitigations.")

# ---- 1.3.3 Task 1.3 EPS/FSS (BIG) --------------------------------------
add_h(doc, "1.3.3 Task 1.3 — Emergency Preparedness and Field Site Services (EPS/FSS)", 3)

add_para(doc,
    "EPS and FSS exist for the CDC scenarios that the rest of OCIO is not designed to absorb at peak: "
    "outbreak response, Emergency Operations Center (EOC) activation, and field deployments to Port Health "
    "Stations and other partner sites. We provide 24×7 on-call coverage, EOC shift coverage during "
    "activations, exercise support, and Field Site Service technicians who can travel and remain on site for "
    "the duration of an event. EPS/FSS handles approximately 400 tickets per month historically against ORR "
    "and DEO service-level expectations that are distinct from the standard OCIO SLA.",
    bold_lead="Mission posture.")

add_para(doc,
    "Three operating elements characterize this task. (1) On-call rotation: published schedules delivered to "
    "the COR within 10 days of award and updated quarterly per PWS Section 8, with named primary and backup "
    "for each shift. (2) EOC shift coverage during declared activations: pre-staffed roster scaled to "
    "activation level, with documented shift handoff using ServiceNow work notes and an EOC daily situation "
    "report. (3) Field Site Service technicians at FSS locations: full Property Custodial Officer (PCO) "
    "duties, annual wall-to-wall inventory, continuous ad-hoc reporting, and OCIO-standard end-of-life "
    "processing.",
    bold_lead="Operating model.")

add_para(doc,
    "Automation in EPS/FSS is deliberately conservative — surge response is not where we run experiments. "
    "Two improvements pay back without putting reliability at risk. (a) An EOC Activation Runbook, embedded "
    "as a ServiceNow workflow that automatically opens the standing tickets, shift roster, escalation tree, "
    "and situation-report template within minutes of activation declaration. (b) A pre-staged FSS deployment "
    "kit checklist (hardware, satellite, comms, identity readiness, travel readiness) maintained as a living "
    "ServiceNow record and audited quarterly.",
    bold_lead="Automation plan.")

add_para(doc,
    "Senior and Journeyman Computer User Support Specialists with embedded Journeyman IT Project Managers, "
    "all with EOC familiarity and travel readiness. Field site personnel are credentialed to perform PCO "
    "duties and maintain currency on CDC Property Management procedures (cross-referenced under Task 1.5). "
    "Travel readiness — passports, immunizations as required, and FSS-specific deployment kits — is a "
    "precondition to assignment, not a downstream administrative task.",
    bold_lead="Staffing and credentials.")

add_para(doc,
    "ORR/DEO SLA tracking in ServiceNow with a separate dashboard from the OCIO SLA. Monthly EPS/FSS performance "
    "report includes ticket volume by activation type, MTTR by priority, on-call response latency, and PCO "
    "inventory variance. Activation after-action reports filed within 14 days of activation closeout.",
    bold_lead="SLA performance posture.")

add_para(doc,
    "(1) Activation-induced surge — mitigated by Task 4.1 cross-task surge plan and by deflection-preserved "
    "Tier 1 capacity. "
    "(2) FSS travel readiness lag — mitigated by quarterly readiness audit and a named FSS Operations Lead "
    "who owns currency. "
    "(3) Property variance during high-tempo deployment — mitigated by ServiceNow asset linkage, PCO "
    "duty assignment, and an explicit reconciliation gate at deployment close.",
    bold_lead="Task-specific risks and mitigations.")

# ---- 1.3.4 Task 2.1 GA (BIG) -------------------------------------------
add_h(doc, "1.3.4 Task 2.1 — Global Activities Services (GA)", 3)

add_para(doc,
    "Level 2 and Level 3 IT support to Global Health Center customers in approximately 60 international "
    "locations, supporting roughly 1,800 staff and approximately 60 LANs that include workstations, network "
    "infrastructure, file servers, satellite communications, and video conferencing. Connectivity to Atlanta "
    "is integrated through SD-WAN and site-to-site VPN to Microsoft Azure, Office 365, and CDC core services.",
    bold_lead="Scope.")

add_para(doc,
    "GA runs on a regional model — by sub-region time-zone — with named Regional Engineers, regional operating "
    "procedures, and documented site-specific topologies. The Atlanta-based GA Engineering Lead owns "
    "cross-regional consistency and deployment readiness; regional engineers own day-to-day operations and "
    "the local ITSM relationship. ServiceNow is the system of record across all regions; tickets carry "
    "site, region, and priority-1 cascade indicators.",
    bold_lead="Operating model.")

add_para(doc,
    "Per Section E Staffing Plan minimum qualifications, GA engineering and technical support personnel hold "
    "active Cisco Certified Network Associate (CCNA) certifications and a combination of three to five years "
    "of experience installing, configuring, and troubleshooting network switches, routers, and firewalls. "
    "International travel readiness — passports, visas, low- and high-threat country travel preparation — is "
    "a precondition to assignment. Each engineer holds an active CCNA at hire and maintains currency under "
    "the contract's training currency obligation. The GA Lead is one of our four named Key Personnel.",
    bold_lead="Engineer credentialing.")

add_para(doc,
    "Twenty-four-to-forty-eight-hour deployment windows in declared emergencies are met through pre-cleared "
    "travel readiness, pre-negotiated travel insurance and medical evacuation coverage per PWS Section 12, "
    "and runbooks for the most likely emergency deployment patterns (network restoration, infrastructure "
    "rebuild, satellite reconnection, identity reset). We comply with the safety, security, medical "
    "clearance, insurance (Traveler Health Insurance and Medical Evacuation), Smart Traveler Enrollment "
    "Program (STEP) registration, and Foreign Affairs Counter Threat (FACT) training requirements specified "
    "in PWS Section 12. An executive-level safety and security point of contact is identified at award.",
    bold_lead="Deployment readiness and OCONUS safety.")

add_para(doc,
    "Two non-disruptive automations apply at OCONUS scale. (1) SD-WAN posture monitoring with predictive "
    "alerts before degradation cascades to a P1 — runbook-linked from alert to incident to engineer, with "
    "ServiceNow tying the regional asset to the engineer's queue. (2) Image baseline drift detection on "
    "site servers and workstations through MECM and Intune coordination, surfaced to the regional engineer "
    "before drift becomes a security or stability issue. Both are CDC-platform native; neither introduces "
    "new tooling.",
    bold_lead="Automation plan.")

add_para(doc,
    "The GA team holds Senior Computer Network Support Specialists, Journeyman Computer Network Systems "
    "Administrators, Senior Information Security Analysts, and Senior Document Management Specialists, sized "
    "to the published level of effort. Global travel costs are cost-reimbursable to the published $198K base-"
    "period cap.",
    bold_lead="Staffing and credentials.")

add_para(doc,
    "GA tickets are tracked against PWS Table 5 thresholds with regional rollup. Above the PWS, we measure "
    "Site Availability Percentage (target 99.0%), Mean Time to Site Restoration (P1, target 4 hours from "
    "engineer dispatch), and Quarterly Travel-Readiness Currency at 100% of GA roster.",
    bold_lead="SLA performance posture.")

add_para(doc,
    "(1) High-threat country deployment risk — mitigated by FACT training currency, safety POC engagement, "
    "and pre-cleared travel posture. "
    "(2) Satellite / SD-WAN single-point dependencies at remote sites — mitigated by predictive monitoring, "
    "regional runbooks, and pre-positioned spare equipment where logistics permit. "
    "(3) CCNA shortage in U.S. labor market — mitigated by retention focus on incumbent GA engineers, "
    "Insight Global pipeline at CCNA-credentialed entry points, and contractor-funded CCNA renewal support.",
    bold_lead="Task-specific risks and mitigations.")

# ---- 1.3.5 Task 3.1 PMS ------------------------------------------------
add_h(doc, "1.3.5 Task 3.1 — Program Management Services (PMS)", 3)

add_para(doc,
    "A dedicated Program Manager is named as primary technical and managerial interface to the Contracting "
    "Officer and COR, with at least one named alternate. The PMO owns Quality Assurance Surveillance Plan "
    "(QASP) alignment, KPI tracking, the Weekly Status Report due Wednesday by 5:00 PM ET, the Monthly "
    "In-Progress Review, the Monthly Financial Report, and the Monthly Progress Report due on the 15th. "
    "The post-award kickoff is convened within five business days of award per PWS Subtask 3.1.1.",
    bold_lead="PMO and governance.")

add_para(doc,
    "The PMO operates a single integrated cadence rather than separate task cadences: weekly all-task stand-up "
    "(Task Leads + QC Lead + Cyber Lead), bi-weekly subcontractor coordination, monthly leadership review with "
    "the OCIO sponsor, and a quarterly Continuous Service Improvement (CSI) review tied to PWS Subsection "
    "3.2.6. The cadence is documented in Appendix C decision-rights matrix and is treated as the default — "
    "we adjust it to the COR's preferences in the first 30 days.",
    bold_lead="Integrated cadence.")

add_para(doc,
    "Subtask 3.1.2 obligates us to plan, implement, and complete Transition-In and Transition-Out activities. "
    "Our Transition-In Plan and Transition-Out Plan are summarized in Section 3 (Management Approach) of this "
    "volume. The Transition Plan uses the same labor categories and positions as the Staffing Plan, in "
    "compliance with PWS Subtask 3.1.2 item 1(b).",
    bold_lead="Transition services within Task 3.1.")

# ---- 1.3.6 Task 4 ------------------------------------------------------
add_h(doc, "1.3.6 Task 4 — Specialized Support Services (Optional)", 3)

add_para(doc,
    "Task 4.1 is activated by contract modification when triggered by a public health emergency, an "
    "unanticipated federal mandate, technology refresh, or building moves. We mobilize qualified surge "
    "personnel within 30 days of trigger, drawing from a designated bench inside the prime and subcontractor "
    "team plus the Insight Global staffing channel. Surge personnel are pre-screened for clearance posture "
    "and credential currency where applicable. The Task 4.1 bench is sized to the published Junior Computer "
    "User Support Specialist optional level of effort and is rebalanced quarterly against the active "
    "demand profile.",
    bold_lead="Task 4.1 — Special IT Projects (Optional).")

add_para(doc,
    "Project Managers hold an industry-recognized PM certification (PMP or equivalent), four-plus years of "
    "federal client experience, and demonstrated Agile delivery experience in two-to-four-week sprint cycles "
    "(daily stand-ups, sprint reviews, backlog management) per Section E Minimum Qualifications. Technical "
    "Writers produce 508-compliant documentation aligned to MS365, EPLC, and OCIO PM standards. The optional "
    "PM and Technical Writer roster maps to the historical Senior IT Project Manager level of effort.",
    bold_lead="Task 4.2 — Project Management and Technical Writer Services (Optional).")

# ---- 1.3.7 Task 5 -----------------------------------------------------
add_h(doc, "1.3.7 Task 5 — DSO Operations Support Services", 3)

add_para(doc,
    "Enterprise project management on Microsoft Project and Visio with EPLC stage-gate compliance, supporting "
    "DSO Project Management Team work in cloud, network access control, and M365. Subtask 5.1.2 technical "
    "writing produces 508-compliant artifacts for cloud, NAC, and M365 efforts. The Senior IT Project "
    "Manager and Senior Management Analyst roles operate to OCIO PM standards and integrate with the Task "
    "3.1 PMO governance cadence.",
    bold_lead="Task 5.1 — IT Project Management Services (Required).")

add_para(doc,
    "FAIR-aligned analytics in SQL, Power BI, and network management tools producing performance metrics and "
    "quarterly and annual IT reporting, stood up if and when CDC exercises this option. The role is a Database "
    "Architect SME under the published optional level of effort.",
    bold_lead="Task 5.2 — Data Analysis and Reporting (Optional).")

add_para(doc,
    "Communications and marketing through CDC Connects and the OCIO web to Section 508 standards; web content "
    "in HTML, JavaScript, and XML; backup webmaster role; outage protocol execution. All deliverables "
    "508-compliant and aligned to the OCIO communications voice and brand.",
    bold_lead="Task 5.3 — IT Communications Services (DSO, Optional).")

# ---- 1.3.8 Task 6.1 ---------------------------------------------------
add_h(doc, "1.3.8 Task 6.1 — Application Web Development Services", 3)

add_para(doc,
    "Web and database application development for OCIO Tools on the ServiceNow App Engine and the .NET stack. "
    "Low-code architecture with reusable components is the design default; code-heavy patterns are reserved "
    "for cases where reuse and platform integration cannot meet the requirement. Configuration management runs "
    "in ServiceNow with ITIL-aligned change and release management. Toolset includes (per the PWS) "
    "ServiceNow App Engine Studio, .NET, SQL Server Management Studio, Azure Data Factory, Azure Synapse, "
    "and Azure SQL.",
    bold_lead="Approach.")

add_para(doc,
    "We treat each new app as either a ServiceNow-native solution (default), a Power Platform solution "
    "(secondary), or a custom .NET solution (only where the prior two cannot meet the requirement). This "
    "ordering reduces long-term operating cost — the ServiceNow and Power Platform options inherit CDC's "
    "platform investments and require less standalone maintenance — and it reduces the surface area for "
    "Section 508 and security review.",
    bold_lead="Build-decision framework.")

# ---- 1.3.9 Task 8 -----------------------------------------------------
add_h(doc, "1.3.9 Task 8 — Identity and Access Management Support Services", 3)

add_para(doc,
    "Patch and configuration management across MECM, Intune, WSUS, BigFix, and Microsoft Defender; secure "
    "baselines aligned to CDC-approved configuration; drift detection with automated remediation where "
    "deterministic rules apply. Coordination with the CDC Security and Privacy Office (CSPO) on patch "
    "testing windows is a defined relationship, not an ad hoc one — windows are calendared, and emergency "
    "patches follow PWS Subtask 8.1.2 emergency-patching guidance.",
    bold_lead="Task 8.1 — Endpoint Management.")

add_para(doc,
    "Zero Trust architecture support across Microsoft Azure, Zscaler, multi-factor authentication, and virtual "
    "desktop infrastructure. Policy-based access enables remote and telework productivity for the CDC "
    "workforce, including the CITGO Virtual Desktop environment that appears in the historical incident "
    "profile (~5.5% of incidents). Engineering posture is aligned to the OCIO Zero Trust roadmap; we do not "
    "propose architectural change without OCIO direction.",
    bold_lead="Task 8.2 — Remote Access Services.")

add_para(doc,
    "Identity lifecycle and brokering across Active Directory, Forefront Identity Manager, SQL, and PowerShell. "
    "Identity-automation patterns reduce manual provisioning load — the second-largest task category in the "
    "historical demand profile (22.4%) — without compromising audit posture. The directory automation work "
    "explicitly intersects with Task 1.1 deflection (UserID provisioning rules) so improvements compound "
    "across both tasks.",
    bold_lead="Task 8.3 — Directory Services.")

# ---- 1.3.10 Task 9 ----------------------------------------------------
add_h(doc, "1.3.10 Task 9 — Infrastructure Support Services", 3)

add_para(doc,
    "Layer 2 and Layer 3 network support across domestic locations, integrated with ServiceNow ITSM tooling "
    "and CDC monitoring. Predictive network monitoring identifies degradation before it becomes a P1 outage; "
    "hybrid and cloud network design supports OCIO modernization efforts. Non-global travel up to $15K in "
    "the base period covers field network engagements. The Net team coordinates with Task 2.1 GA on "
    "OCONUS network reach-back and with Task 8.2 RAS on Zero Trust integration.",
    bold_lead="Task 9.1 — Networking (Required).")

add_para(doc,
    "Data center technician services (9.2), virtual engineer services on VMware vSphere/ESX and other "
    "hypervisors (9.3), 24×7 NOC operations services (9.4), data backup engineer services with FIPS-compliant "
    "media handling (9.5), software installation services for VM provisioning aligned to NIST 800-53 baselines "
    "(9.6), and application hosting and customer interface services running on ServiceNow Configuration "
    "Management and ITIL change and configuration management (9.7). We are positioned to stand up each "
    "optional task if and when CDC exercises the option, with no rip-and-replace of existing CDC tooling. "
    "Each optional task has a designated Task Lead pre-identified so option exercise translates to "
    "performance start without organizational realignment.",
    bold_lead="Tasks 9.2 – 9.7 (Optional).")

# ---- 1.3.11 Task 10 ---------------------------------------------------
add_h(doc, "1.3.11 Task 10 — Workplace Productivity Support Services", 3)

add_para(doc,
    "Patch and antivirus management across the M365 platform with CSPO and Event Notification List (ENL) "
    "coordination; APIs and Group Policy Objects; Security Assessment and Authorization (SA&A) documentation; "
    "24×7 break-fix; vendor case management. The Senior Computer Systems Engineer/Architect is the named "
    "M365 platform lead.",
    bold_lead="Task 10.1 — General Requirements (cross-cutting).")

add_para(doc,
    "Enterprise email and messaging administration across Outlook, Microsoft Stream, and Intune. "
    "High-availability design across cloud and on-prem; coordination with M365 platform engineering. Outlook "
    "is the largest single incident driver in the demand profile (10.3%) — we treat KB and macro investment "
    "in Outlook as a Tier-1 deflection priority.",
    bold_lead="Task 10.2 — Communications Email/Messaging.")

add_para(doc,
    "Site collection administration across Teams, SharePoint, OneDrive, and Azure for both M365 cloud and "
    "on-prem operations. Governance follows CDC standard practice; we do not propose to replace existing "
    "collaboration patterns absent CDC direction.",
    bold_lead="Task 10.3 — Enterprise Collaboration.")

add_para(doc,
    "Low-code development and administration across Dynamics 365, Power Platform, and Azure DevOps. Workflow "
    "automation and Power BI visualization. Data modeling that supports Task 10.5 reporting. The Power "
    "Platform team is the second engine of automation on this contract, complementing the ServiceNow "
    "workflow program in Task 1.1 and Task 8.3.",
    bold_lead="Task 10.4 — Process Automation and Visualization.")

add_para(doc,
    "Reporting infrastructure and dashboards on ServiceNow with event-data integration, real-time alerting, "
    "and connection to performance systems. The reporting platform is the channel through which our SLA "
    "performance, deflection performance, and continuous-improvement results are made visible to the COR "
    "and OCIO leadership. Dashboards are co-designed with the COR in the first 30 days post-award.",
    bold_lead="Task 10.5 — Reporting Services.")

# ---- 1.3.12 Task 11 ---------------------------------------------------
add_h(doc, "1.3.12 Task 11.1 — OCIO Communications Services", 3)

add_para(doc,
    "OCIO communications and marketing across CDC Connects and the OCIO web — 508-compliant content, "
    "stakeholder messaging, web content production, and multimedia. The Weekly Status Report and the call-"
    "volume, productivity, and trend-analysis report (due monthly on the 15th) are produced in this task. "
    "The OCIO Communications role works to the OCIO voice and brand standards and supports the customer "
    "communications backbone of the entire CITS service experience — including transition-period customer "
    "communications described in §3.7.")

# 1.4 Cross-cutting and HCD ---------------------------------------------
add_h(doc, "1.4 Cross-Cutting Service Disciplines (Including HCD as Operating Model)", 2)

add_para(doc,
    "Seven cross-cutting disciplines run across every task. They are summarized here so we do not repeat them "
    "in each task narrative.")

add_para(doc,
    "All contractor personnel handle customer interactions with tact and diplomacy, demonstrate active "
    "listening, and accurately record information. The 96%-or-greater Customer Service Survey score is a "
    "managed metric, not a hopeful outcome — coaching, calibration sessions, and structured QA review of "
    "recorded calls are weekly disciplines.",
    bold_lead="(1) Customer service standard (PWS Subsection 3.2.1).")

add_para(doc,
    "Quality is consistency, accuracy, and customer satisfaction. We staff a defined quality control function "
    "inside Program Management, run weekly review of contact channels in Task 1.1, and document corrective "
    "actions in the QCP we submit and maintain throughout the contract. The QCP is described in §3.6.",
    bold_lead="(2) Quality control discipline (PWS Subsection 3.2.2 and Section 9).")

add_para(doc,
    "Knowledge articles in ServiceNow are owned, dated, indexed, and reviewed on a quarterly cadence. "
    "Findability is measured. The objective is deflection — moving deterministic resolution from Tier 1 "
    "capacity to customer self-service — and that requires the knowledge base to be useful, not merely present. "
    "Article ownership maps to a named SME by topic family; the Knowledge Article Currency Rate (target 95%) "
    "is reported monthly.",
    bold_lead="(3) Knowledge management as a measured discipline (PWS Subsection 3.2.3).")

add_para(doc,
    "Every recurring report in PWS Table 12 is owned, calendared, and version-controlled. Reports are "
    "submitted on time and without gross errors, in compliance with the QASP performance requirement. The "
    "reporting platform under Task 10.5 surfaces report status proactively rather than retrospectively.",
    bold_lead="(4) Reporting discipline (PWS Subsection 3.2.4 and Section 8).")

add_para(doc,
    "CDC's PWS calls for innovative technology to improve operations and customer experience, naming RPA "
    "and ServiceNow automation as examples. Our innovation roadmap is built on those rails. RPA addresses "
    "repetitive workflow steps with deterministic logic — onboarding sub-tasks, ticket routing, status "
    "communications. ServiceNow workflow expansion compounds the deflection program. Capabilities CDC "
    "explicitly provides and approves are absorbed into the operating model on CDC's authority through our "
    "continuous-improvement cadence; we do not introduce or use unapproved AI products in performance, in "
    "compliance with the Section E proposal preparation instruction.",
    bold_lead="(5) Innovation in operations and customer experience (PWS Subsection 3.2.5).")

add_para(doc,
    "We monitor and analyze performance and customer experience, build the structure for incremental "
    "improvement, and execute improvements that reduce TCO consistently over time. The change management "
    "discipline is documented in the QCP. Year-over-year TCO reduction is a contract objective; we report "
    "against it monthly.",
    bold_lead="(6) Continuous improvement and total cost of ownership (PWS Subsection 3.2.6).")

add_para(doc,
    "HCD is how we operate, not a separate work product. We use named methods — journey mapping, contextual "
    "inquiry, service blueprinting, usability testing on the ServiceNow self-service portal, and co-design "
    "with the CDC Customer Experience Office — and we tie each to a specific operational outcome. We "
    "journey-map four CDC populations whose support needs are distinct: laboratory and research scientists "
    "(scientific computing, MFA friction, instrument network integration), public health field responders "
    "(EOC tempo, deployable kit reliability, OCONUS connectivity), program staff (M365 productivity, document "
    "collaboration), and global health partners (low-bandwidth resilience, identity reach-back). The journey "
    "maps directly inform the priority of KB articles, the design of self-service flows in the Virtual Agent, "
    "and the choice of automation candidates. Section 508 / WCAG 2.0 AA is treated as design input from day "
    "one, not a remediation task. Voice-of-customer measurement runs continuously through the Customer "
    "Service Survey and a quarterly journey-specific deep-dive coordinated with the CXO team. HCD outcomes "
    "are reported to the COR and OCIO leadership monthly.",
    bold_lead="(7) Human-Centered Design (HCD) as the operating model.")

# 1.5 508 ---------------------------------------------------------------
add_h(doc, "1.5 Section 508 Accessibility", 2)

add_para(doc,
    "Section 508 accessibility is treated as a delivery requirement, not an afterthought. The DHHS Section 508 "
    "Product Assessment Template accompanies this volume per Section E; it does not count against the page "
    "limit. All web content under Tasks 5.3 and 11.1, all reporting outputs under Task 10.5, and all "
    "customer-facing artifacts (knowledge articles, self-service flows, training materials) are produced to "
    "WCAG 2.0 Level A and AA standards. Remediation cost for any non-conformant deliverable is the "
    "contractor's responsibility post-award, as required. The Task 5.3 / 11.1 communications leads carry "
    "508-compliance accountability for their respective deliverables; the QC Lead carries enterprise "
    "accountability across the contract.")

# 1.6 Cyber -------------------------------------------------------------
add_h(doc, "1.6 Cybersecurity and Compliance Posture", 2)

add_para(doc,
    "Our delivery posture meets PWS Section 12 in full: FISMA compliance; NIST SP 800-60 Volume II data "
    "categorization; NIST SP 800-160 Volumes 1 and 2 systems-security-engineering principles; CSPO "
    "coordination on patching and configuration baselines; immediate Cybersecurity Incident Response Team "
    "(CSIRT) reporting of known, reported, or suspected security incidents; DoD 8570.01-m and 8140 "
    "personnel certification posture; and Security Assessment and Authorization (SA&A) package alignment to "
    "the CDC SA&A SOP. Personnel hold Public Trust Level 5 clearances where required (including the three "
    "ServiceNow Analysts and the VCC Systems Administrator under Task 1.1). Foreign-equipment and product "
    "prohibitions under Section 889, FAR 52.204-23 (Kaspersky), FAR 52.204-25, -26, and -27 (covered "
    "telecommunications and ByteDance) are reflected in our supply chain and procurement controls. "
    "Non-Disclosure Agreements are signed by every contractor employee per HHS instruction. Contractor "
    "Information Security Awareness, Privacy, and Records Management training is completed at onboarding "
    "and renewed annually.")

# 1.7 Pilots ------------------------------------------------------------
add_h(doc, "1.7 Three Test-and-Learn Pilots", 2)

add_para(doc,
    "Three pilots scoped to the published demand profile are proposed for the Base Period. Each is small, "
    "measurable, and reversible; each operates inside CDC's existing platform investments; and each has a "
    "named exit-or-extend gate at end of pilot. Pilots are presented for COR concurrence in the first 30 "
    "days post-award.")

add_para(doc, "Table 1.7-A. Three test-and-learn pilots.", bold_lead="")
add_table(doc,
    header=["Pilot", "Hypothesis", "Baseline / target", "Duration", "Exit / extend gate"],
    rows=[
        ["P1 — Virtual Agent password-reset deflection at Roybal",
         "ServiceNow Virtual Agent + audited identity verification deflects deterministic password reset traffic without compromising security or CSAT",
         "Baseline: 0% deflection. Target: 20% deflection of Roybal-originating password resets in pilot window",
         "90 days",
         "Extend to enterprise if deflection ≥15% with CSAT ≥96% and zero security exceptions"],
        ["P2 — Zero-touch refresh for DSS Atlanta image deployments",
         "Pre-imaged, pre-patched, pre-encrypted devices reduce per-device refresh touch time and improve image-compliance pre-issue rate",
         "Baseline: ~30 minutes per refresh. Target: <8 minutes per refresh; 100% pre-issue compliance",
         "60 days",
         "Extend to all Atlanta campuses if compliance is 100% and per-device time meets target"],
        ["P3 — RPA-based PIV exception fulfillment",
         "RPA on Power Automate Desktop fulfills the deterministic verification steps of PIV Card Exception Requests, freeing Tier-1 capacity",
         "Baseline: 3,087 PIV exception tickets / 14 mo. Target: 40% of cycle time automated; FCR rate ≥85%",
         "120 days",
         "Extend to enterprise if cycle-time reduction ≥30% and zero audit exceptions"],
    ],
    col_widths=[1.4, 1.7, 1.6, 0.7, 1.6])

add_para(doc,
    "All three pilots are run inside CDC-provided platforms (ServiceNow, MECM, Power Automate Desktop). None "
    "introduces a new tool, a new vendor, or a new data flow. None depends on AI products. Each pilot has a "
    "named owner, a documented baseline pulled from the Task 10.5 reporting platform, a weekly read-out to "
    "the COR, and an explicit decision gate at pilot end.")

# 1.8 Roadmap -----------------------------------------------------------
add_h(doc, "1.8 Automation and Deflection Roadmap (Year-Over-Year Targets)", 2)

add_para(doc,
    "Cost discipline on CITS is delivered through three compounding levers — deflection, automation, and "
    "delivery-center consolidation — under PWS Subsection 3.2.6 and contract objective 4. Table 1.8-A sets "
    "year-over-year targets we commit to as the operating envelope. Targets are adjusted in coordination with "
    "the COR at the start of each option period if material baseline shifts (e.g., a new application "
    "rollout or a major HHS reorganization) require recalibration.")

add_para(doc, "Table 1.8-A. Automation and deflection targets across the period of performance.", bold_lead="")
add_table(doc,
    header=["Metric", "Base", "OP1", "OP2", "OP3", "OP4 (6 mo)"],
    rows=[
        ["Tier-1 deflection rate (% of deterministic categories self-served)", "Baseline", "30%", "40%", "45%", "50%"],
        ["Knowledge Article Currency Rate (% reviewed within 90 days)", "≥85%", "≥90%", "≥95%", "≥95%", "≥95%"],
        ["First-Contact Resolution Rate (Tier 1)", "Baseline", "70%", "75%", "78%", "78%"],
        ["Zero-touch refresh rate (DSS)", "Baseline", "80%", "90%", "95%", "95%"],
        ["RPA bots in production (count)", "1–2", "5", "8", "10", "10"],
        ["Customer Service Survey ≥96% (PWS §9 floor)", "≥96%", "≥96%", "≥97%", "≥97%", "≥97%"],
        ["Year-over-year TCO trend (vs. prior FY)", "Baseline", "↓", "↓", "↓", "↓"],
    ],
    col_widths=[2.5, 0.9, 0.7, 0.7, 0.7, 0.9])

add_para(doc,
    "Targets are reported monthly to the COR through the Task 10.5 reporting platform. A target miss triggers "
    "the QCP corrective-action protocol described in §3.6.")

# =======================================================================
# 2. STAFFING PLAN
# =======================================================================
add_h(doc, "2. Staffing Plan", 1)

add_para(doc,
    "Our staffing principle: stable people, stable service. CITS is a service-delivery contract before it is "
    "a transformation contract. The single biggest predictor of service stability through transition and "
    "across the period of performance is whether the people doing the work — at the service desk, at "
    "deskside, on the field site teams, in OCONUS network engineering, in the OCIO PMO — are credentialed, "
    "supported, and retained. Our staffing plan is built around that principle. We propose labor categories "
    "sized to the historical level of effort published in Section E; we keep credentialing and clearance "
    "posture aligned to the Minimum Qualifications; we name a Key Personnel cadre that is accountable to "
    "CDC; and we hold a retention strategy that begins on day one of transition and runs across all four "
    "option periods.")

# 2.1 ---------------------------------------------------------------
add_h(doc, "2.1 Staffing Strategy", 2)

add_para(doc,
    "Section E's Estimated Level of Effort table publishes historical hours by labor category — for example, "
    "20 Junior Computer User Support Specialists at Task 1.1 ITSDS, 40 Journeyman Computer User Support "
    "Specialists at Task 1.2 DSS, and seven Senior Computer Systems Analysts at Task 8.1 Endpoint "
    "Management. We have proposed our LCAT mix to align with that historical envelope across the Transition, "
    "Base, and four Option periods. Where our experience suggests the work can be performed more cost-"
    "effectively at a different LCAT mix without compromising service quality, we have noted that delta in "
    "the staffing matrix and explained the basis.",
    bold_lead="Right-sized labor categories.")

add_para(doc,
    "CITS service delivery requires on-site presence at multiple Atlanta campuses and the published non-"
    "Atlanta CSC locations, plus FSS deployable capacity, plus OCONUS-ready GA engineering. The remainder — "
    "Tier 1 service desk, program management, application web development, reporting, and several specialized "
    "functions — supports a delivery-center-based model where eligible work is consolidated into U.S. "
    "delivery centers for cost discipline and service consistency. Our staffing matrix identifies the "
    "on-site, on-campus, and remote/delivery-center designation for each role.",
    bold_lead="Geographic posture.")

add_para(doc,
    "We propose a teamed delivery model with a primary teaming partner and a staffing partner. "
    "[PLACEHOLDER: Insert subcontractor identities and roles. Capture lead to confirm whether subs are named "
    "explicitly in Volume I or referenced generically in Volume I and named in Volume II Business; current "
    "orientation indicates teaming is in progress with task-by-task workshare alignment underway. If subs "
    "are named, briefly describe each sub's role, the labor categories they staff, the percentage of total "
    "labor hours they contribute, and the accountability chain from sub to prime to CO/COR.] Our prime, "
    "[PLACEHOLDER: Prime legal entity], retains program management, key personnel accountability, and "
    "subcontract management responsibility throughout the period of performance. The Subcontract Management "
    "Lead reports to the Program Manager, runs the bi-weekly subcontractor coordination meeting, and is the "
    "named escalation path for any subcontractor performance issue.",
    bold_lead="Subcontractor team.")

# 2.2 ---------------------------------------------------------------
add_h(doc, "2.2 Staffing Matrix", 2)

add_para(doc,
    "The staffing matrix in Appendix A of this volume identifies, for each proposed labor category and role: "
    "(1) the task and subtask supported, (2) the number of full-time equivalents and total proposed labor "
    "hours by period of performance, (3) skill set, (4) certifications and licenses required and held, "
    "(5) security clearance posture, (6) professional experience requirement, (7) education, and (8) "
    "percent of time on the program. The matrix correlates directly to the hour build in Volume II and to "
    "the Section E Estimated Level of Effort table.")

add_placeholder(doc,
    "Appendix A — Staffing Matrix. Pricing/staffing team to populate from the Volume II hour build. Format "
    "requirements from Section E Staffing Plan: number and types of personnel, assigned roles and "
    "responsibilities, skill sets, certifications, security clearances, professional experience, education, "
    "and overall qualifications. Confirm whether the matrix is excluded from the 35-page Volume I limit or "
    "counts toward it — submit as a clarification question by May 12.")

add_para(doc,
    "Where Section E specifies a Minimum Qualification for a labor family — Project Management, IT Service "
    "Desk Services (ITSDS), Global Customer Support Services, IT Asset Management Services — every proposed "
    "individual in that family meets or exceeds the stated standard. Specifically:",
    bold_lead="Labor categories aligned to Section E.")

bullets = [
    "Project Managers hold an industry-recognized project management certification (PMP or equivalent), "
    "have four-plus years of federal client experience, and bring demonstrated Agile delivery experience in "
    "two-to-four-week sprints — including stand-ups, sprint reviews, and backlog management.",
    "IT Service Desk Services personnel hold ITIL certification where preferred by CDC; the team holds "
    "ITIL coverage at the leadership level.",
    "Global Customer Support Services engineering personnel hold active Cisco Certified Network Associate "
    "certifications, have a combination of three to five years of network switch, router, and firewall "
    "experience, and have prior international travel readiness (passports, visas, low- and high-threat "
    "country preparation).",
    "IT Asset Management Services includes certified forklift coverage for Depot operations as required.",
]
for b in bullets:
    p = doc.add_paragraph(b, style="List Bullet")

# 2.3 Key Personnel ---------------------------------------------------
add_h(doc, "2.3 Key Personnel", 2)

add_para(doc,
    "Key Personnel are proposed in accordance with CDCH.04 Key Personnel (January 2026). The four named Key "
    "Personnel positions are summarized below; full resumes are provided in Appendix B and do not count "
    "against the 35-page limit.")

add_para(doc, "Table 2.3-A. Named Key Personnel.", bold_lead="")
add_table(doc,
    header=["Role", "Primary accountability", "Minimum qualification posture"],
    rows=[
        ["Program Manager (PM)",
         "Primary technical and managerial interface to the CO and COR; QASP alignment; PMO governance",
         "PMP or equivalent; 10+ years federal IT services PM; Agile sprint delivery; Public Trust"],
        ["Deputy PM / Quality Control Lead",
         "QCP ownership; QASP performance reporting; corrective-action protocol; PM alternate per PWS 3.1",
         "PMP or equivalent; 7+ years federal QA; ITIL Foundation"],
        ["Service Desk Lead (Task 1.1)",
         "Tier-1 service operations; deflection program; KB & Virtual Agent ownership",
         "ITIL Foundation; 7+ years federal Tier-1 leadership; experience with NICE VCC or equivalent"],
        ["Global Activities Lead (Task 2.1)",
         "OCONUS engineering; deployment readiness; regional engineer accountability",
         "Active CCNA; 8+ years OCONUS network engineering; FACT-current; international travel readiness"],
    ],
    col_widths=[1.4, 2.7, 2.7])

add_placeholder(doc,
    "Appendix B — Key Personnel Resumes. Capture lead to confirm named individuals; resumes inserted in the "
    "appendix and excluded from the 35-page count consistent with prior practice. Each resume to address: "
    "education, certifications, security clearance, federal CDC/HHS experience, demonstrated performance "
    "against comparable SLAs, and references.")

# 2.4 ---------------------------------------------------------------
add_h(doc, "2.4 Approach to Recruit, Hire, and Retain", 2)

add_para(doc,
    "Per PWS Section 9, qualified staff are available within seven days of a task order or option period "
    "start. We meet this through three mechanisms: (1) a pre-cleared and pre-credentialed bench inside the "
    "prime and subcontractor team that is available for direct assignment; (2) a documented hiring funnel "
    "with named recruiting partners that targets ITIL-certified, CCNA-certified, and federal-experienced "
    "candidates; (3) for urgent or surge demand under Task 4.1, an Insight Global-led "
    "[PLACEHOLDER: confirm whether to name staffing partner in Volume I or treat generically as 'staffing "
    "partner'] staffing channel that lets us scale by city, by labor category, and by clearance posture "
    "without ad hoc mobilization.",
    bold_lead="Recruit and hire.")

add_para(doc,
    "If a Key Personnel position becomes vacant, we identify a qualified replacement and propose to CDC "
    "within 10 business days. For non-Key positions, the replacement target is consistent with the Section "
    "9 seven-day standard for new task orders, with backfill from the pre-cleared bench during recruitment. "
    "Fingerprinting timeliness is treated as a managed metric, not an administrative checkpoint — we "
    "coordinate with HHS to schedule within the published windows to avoid invoice deductions for delay "
    "(PWS Section 9 disincentive structure).",
    bold_lead="Replacement velocity.")

add_para(doc,
    "Three retention levers carry the most weight on this contract. (1) Role clarity and career progression — "
    "staff who see a path advance and stay. (2) Training currency — every contractor employee has access to "
    "trainings and certifications relevant to their role; certification renewal is tracked and "
    "contractor-funded for required certifications (ITIL, CCNA, PMP). (3) Performance recognition tied to "
    "operational outcomes — service desk technicians and SDLs whose teams deliver against SLA see that "
    "performance reflected in evaluations and growth opportunities. We sustain 360-degree feedback at "
    "annual cycle and pulse-survey at quarterly cadence to read engagement before it reads as turnover.",
    bold_lead="Retention.")

add_para(doc,
    "Continuity of trusted talent is one of our executive win themes. During transition, we engage exclusively "
    "through CDC-coordinated channels and the staffing-partner pathway to avoid any conduct that could be "
    "characterized as direct outreach for solicitation of incumbent staff. Where incumbent personnel are "
    "interested in continued CITS engagement under our team, we route those expressions of interest through "
    "the staffing channel rather than through prime-to-individual contact, in keeping with sound procurement-"
    "integrity practice. This protocol is documented in §3.7 transition plan.",
    bold_lead="Continuity of trusted talent.")

# 2.5 turnover ---------------------------------------------------------
add_h(doc, "2.5 Employee Turnover Rates", 2)

add_para(doc,
    "Per Section E Staffing Plan instruction, prime and named subcontractor turnover rates for the past two "
    "years are shown below. Turnover is reported as percentage of separations against average headcount on "
    "comparable federal IT services portfolios.")

add_table(doc,
    header=["Entity", "FY24 turnover", "FY25 turnover", "Notes"],
    rows=[
        ["[Prime]", "[PLACEHOLDER %]", "[PLACEHOLDER %]", "Prime federal IT services portfolio (comparable)"],
        ["[Primary teaming partner]", "[PLACEHOLDER %]", "[PLACEHOLDER %]", "Federal IT services book of business"],
        ["Industry benchmark (federal IT services)", "[PLACEHOLDER %]", "[PLACEHOLDER %]", "Reference benchmark only"],
    ],
    col_widths=[2.4, 1.2, 1.2, 2.0])

add_placeholder(doc,
    "Capture lead and HR analytics to provide actual two-year turnover percentages for prime and named subs "
    "before submission. Industry benchmark figure to come from a citable public source (e.g., Computer "
    "Economics or BLS).")

# 2.6 ------------------------------------------------------------------
add_h(doc, "2.6 Subcontracting Plan", 2)

add_para(doc,
    "Per Section E item (c)(5) of the Volume II business volume instructions, a Small Business "
    "Subcontracting Plan accompanies this proposal as required by FAR 52.219-9 (where applicable). "
    "[PLACEHOLDER: Insert summary statement of subcontracting plan posture — confirm whether the prime is a "
    "small business (in which case 52.219-9 does not apply), or whether we are submitting the DHHS "
    "subcontracting plan template. Reference the subcontracting goals by socioeconomic category.] The full "
    "plan is in Volume II.")

# 2.7 staffing risks ---------------------------------------------------
add_h(doc, "2.7 Staffing Risks and Mitigations", 2)

add_para(doc,
    "Three staffing risks dominate. We surface them here so the Government can see how our hiring, retention, "
    "and surge approach is calibrated to the realities of the federal IT services labor market.")

add_table(doc,
    header=["Risk", "Likelihood", "Impact", "Mitigation"],
    rows=[
        ["Clearance throughput lag at transition (fingerprinting / Public Trust 5)",
         "Medium", "High",
         "Pre-cleared bench inside prime and primary teaming partner; documented HHS scheduling protocol "
         "to avoid §9 invoice deductions; staffing-partner pipeline pre-screened for clearance posture."],
        ["CCNA-credentialed engineer shortage in OCONUS-deployable roles",
         "Medium", "High",
         "Retention focus on incumbent GA engineers; contractor-funded CCNA renewal; Insight Global pipeline "
         "filtered by CCNA + travel readiness; cross-training of Task 9.1 networking staff for CCNA path."],
        ["Surge bench depth during a public health emergency",
         "Medium", "Medium",
         "Task 4.1 designated bench rebalanced quarterly; cross-task surge plan; deflection-preserved Tier-1 "
         "capacity; Insight Global rapid-deploy LCAT mapping by city and clearance."],
        ["Voluntary attrition of named SDLs / Tier-2 leads at Atlanta campuses",
         "Low–Medium", "Medium",
         "SDL career progression path; rotation and shadow programs; quarterly retention pulse; "
         "campus-specific knowledge backup roster maintained as Task 1.2 standing artifact."],
        ["Subcontractor performance variance across task families",
         "Low", "Medium",
         "Subcontract Management Lead under PM; bi-weekly performance review; SLA flowdown identical to "
         "prime obligations; named escalation path to CO."],
    ],
    col_widths=[2.4, 0.9, 0.7, 2.7])

# =======================================================================
# 3. MANAGEMENT APPROACH
# =======================================================================
add_h(doc, "3. Management Approach", 1)

add_para(doc,
    "Our management principle: a single accountable interface, a disciplined cadence, and decisions that get "
    "made before they reach the COR. CITS spans 10 PWS tasks, dozens of subtasks, multiple subcontractors, "
    "and a delivery footprint that includes Atlanta campuses, non-Atlanta CSC sites, FSS deployments, and "
    "approximately 60 OCONUS Global Activities locations. Without a clear management spine, that complexity "
    "surfaces as friction at the COR's desk. Our approach pushes accountability and decision-making down to "
    "the lowest level where they can be made well, and reserves CDC's leadership attention for the decisions "
    "that genuinely require it.")

# 3.1 -----
add_h(doc, "3.1 Organizational Structure and Reporting Relationships", 2)
add_para(doc,
    "Our organizational structure has four layers and is designed for clarity, not headcount. The Program "
    "Manager is the named, primary technical and managerial interface to the CO and the COR per PWS Task 3.1. "
    "Reporting to the PM are the named Task Leads for each PWS task family — Service Desk, Deskside, "
    "Emergency Preparedness, Global Activities, IAM, Infrastructure, Workplace Productivity, OCIO "
    "Communications, Application Development, and DSO Operations — together with the Quality Control Lead, "
    "the Cybersecurity and Compliance Lead, and the Subcontract Management Lead. Below the Task Leads are "
    "the Service Desk Liaisons, the Campus Leads, the engineering supervisors, and the operating teams. "
    "The full org chart and decision-rights matrix is in Appendix C.")

add_placeholder(doc,
    "Appendix C — Organizational Chart and Decision-Rights Matrix. Proposal graphics team to render the org "
    "chart aligned to the technical approach in Section 1. Decision-rights matrix should specify, for each "
    "decision class (operational, change, financial, personnel, escalation), whether authority sits with the "
    "COR, the PM, the Task Lead, or the operating team — and the escalation path when authority is "
    "exceeded.")

# 3.2 Performance ------------
add_h(doc, "3.2 Performance Measurement and Metrics Monitoring", 2)
add_para(doc,
    "Performance measurement runs on a single ServiceNow-based reporting platform under Task 10.5 that "
    "integrates ticket data, SLA performance, customer satisfaction survey results, deflection metrics, and "
    "the QASP performance requirements in PWS Section 9. The platform produces:")

bullet_list = [
    "Daily operational dashboards for the PM and Task Leads, reflecting real-time SLA position against PWS "
    "Table 5 thresholds, Tier-1 contact center performance against Task 1.1 standards, and active P1/P2 "
    "incidents.",
    "Weekly Status Reports submitted to the COR and Technical Monitor by 5:00 PM ET each Wednesday per PWS "
    "Section 8 deliverables — covering tasks assigned, tasks completed, reasons for delays, work planned "
    "for the next period, problems identified and corrective actions, and significant meetings.",
    "Monthly In-Progress Reviews and Monthly Progress Reports submitted by the 15th of each month, presented "
    "to OCIO leadership.",
    "Monthly Financial Reports on contract financial status, run rate against budget, and travel against "
    "the Section 7 caps ($12K Task 1.3 non-global, $15K Task 5 non-global, $15K Task 9 non-global, $198K "
    "Task 2.1 global, base period).",
    "Quarterly performance trend reviews with the COR, framing year-over-year improvement against the "
    "contract objective of consistent TCO reduction.",
    "Annual performance retrospective tied to option-period exercise gate; prepares OCIO leadership for "
    "informed option-exercise decisions.",
]
for b in bullet_list:
    doc.add_paragraph(b, style="List Bullet")

# 3.3 Risk -----------------
add_h(doc, "3.3 Risk Management", 2)
add_para(doc,
    "Risk is managed in three loops. The operational risk loop runs daily through the dashboards and is the "
    "responsibility of the Task Leads. The program risk loop runs weekly through PMO governance and is the "
    "responsibility of the PM, who maintains a single program risk register with named owners, current "
    "likelihood and impact assessments, mitigation status, and decision deadlines. The strategic risk loop "
    "runs monthly with the COR and surfaces risks that require Government decisions — for example, changes "
    "in HHS policy, OCIO modernization direction, or budget posture.")

add_para(doc,
    "Table 3.3-A presents the program-level risk register at proposal time. We update the register weekly "
    "and review it monthly with the COR.",
    bold_lead="Program risk register.")

add_para(doc, "Table 3.3-A. Program-level risk register at proposal time.", bold_lead="")
add_table(doc,
    header=["Risk", "L", "I", "Mitigation", "Owner"],
    rows=[
        ["OCONUS network outage cascades during emergency deployment",
         "M", "H",
         "Predictive SD-WAN monitoring; pre-positioned regional spares where logistics permit; CCNA-current "
         "regional engineers; 24-48 hour deployment readiness; FACT currency.",
         "GA Lead"],
        ["Public-health emergency surge exceeds Tier-1 capacity",
         "M", "H",
         "Deflection program preserves Tier-1 human capacity; Task 4.1 surge bench; Insight Global rapid-deploy "
         "channel; cross-task surge runbook activated through PMO.",
         "Service Desk Lead / PM"],
        ["Knowledge-base decay erodes deflection over time",
         "M", "M",
         "Named article ownership; quarterly review cadence; 95% Currency Rate metric; deflection-rate "
         "visibility surfaces decay early.",
         "QC Lead"],
        ["Pricing pressure and TCO commitments require service-quality discipline",
         "H", "M",
         "Year-over-year TCO targets in §1.8; deflection and zero-touch refresh as primary cost levers; "
         "delivery-center consolidation for eligible work; transparent monthly financial reporting.",
         "PM"],
        ["HHS budget continuity / org disruption",
         "M", "M",
         "Contractual flexibility (T&M structure absorbs scope shifts); strategic risk loop with COR; "
         "documented assumptions in §3.10 surface scope-impact items early.",
         "PM"],
        ["Protest defensibility — clean transition, named subs, named individuals",
         "L", "H",
         "Zero direct outreach to incumbent staff; staffing-partner-only channel; documented procurement-"
         "integrity protocol in §3.7; subcontractor naming approved before submission.",
         "PM / Subcontract Lead"],
        ["AI-products restriction interpretation in performance",
         "L", "M",
         "ServiceNow-native automation only; no vendor-introduced AI; absorbed CDC-approved capabilities "
         "via CSI cadence; documented change protocol prevents inadvertent introduction.",
         "PM / Cyber Lead"],
        ["GFE reconciliation gap during transition",
         "L", "H",
         "Continuous-discipline approach (§3.9); Property Custodial Officer coverage; signed walk-down at "
         "transition; ServiceNow asset linkage; 95% on-time delivery metric.",
         "ITAM Lead / PM"],
    ],
    col_widths=[2.6, 0.4, 0.4, 2.7, 1.0])

add_para(doc,
    "Three additional risks are identified during the operational risk loop in PWS Subsection 3.2.6 "
    "continuous-improvement cadence and may not appear in the program register at any single time: "
    "tooling-license expiration risk, training-currency risk for credentialed roles (CCNA, ITIL, PMP), and "
    "vendor case-management dependency for M365 and ServiceNow platform issues. Each is owned by the "
    "relevant Task Lead and reported through the standing weekly cadence.",
    bold_lead="Operational risks managed below the program register.")

# 3.4 ----
add_h(doc, "3.4 Communications, Telework Oversight, Conflict Management", 2)
add_para(doc,
    "In addition to the recurring deliverables in PWS Section 8, the PM hosts a weekly status touchpoint with "
    "the COR (or as the COR prefers), a monthly leadership touchpoint with the OCIO sponsor, and ad hoc "
    "progress meetings as requested. The Continuous Service Improvement (CSI) cadence is published quarterly. "
    "All recurring meetings have a documented agenda template, a designated note-taker, and action-item "
    "tracking in ServiceNow.",
    bold_lead="Communications cadence.")

add_para(doc,
    "For positions designated as remote or hybrid in the PWS Section 6 Place of Performance table, oversight "
    "is built into the daily operating cadence: ServiceNow ticket-level activity logging, Personnel "
    "Availability reporting per PWS Section 9, and Task Lead-level oversight of throughput and quality. "
    "Telework does not loosen accountability — the Personnel Availability calculation under Section 9 "
    "(Talk Time + Hold Time + ACW + Wait Time minus Not Ready Time + Break Time + Lunch Time, divided by "
    "Staffed Time) applies to remote contact center personnel identically to on-site staff.",
    bold_lead="Telework oversight.")

add_para(doc,
    "Conflicts among team members, between subcontractors, or between the contractor team and CDC "
    "stakeholders are surfaced at the lowest appropriate level first. The Task Lead resolves operational "
    "disagreements; the PM resolves cross-task or subcontractor-prime disagreements; the CO and contractor "
    "executive sponsor resolve contractual or scope disagreements. Mediation through HR or executive "
    "escalation is available where required. The full protocol is documented in the QCP.",
    bold_lead="Conflict management.")

# 3.5 ---
add_h(doc, "3.5 Fiscal Responsibility and Cost Discipline", 2)
add_para(doc,
    "Fiscal responsibility is measured monthly in the Monthly Financial Report and reviewed quarterly at the "
    "In-Progress Review. Burn-rate against the labor hour build, ODC and travel against the published caps "
    "(VCC license $320,826 base period; CEO non-global travel $18,326 base period; CEO global travel "
    "$183,326 base period; DSO ISB travel $30,000 base period; hardware/spare parts $137,500 base period), "
    "and option-period readiness are tracked as standing items. The PM raises any anticipated cost-impact "
    "change with the CO before the change manifests in the financials.",
    bold_lead="Fiscal responsibility.")

add_para(doc,
    "Three drivers reduce TCO across the period of performance, consistent with PWS Subsection 3.2.6 and "
    "contract objective 4. (1) Deflection — moving deterministic Tier-1 demand to ServiceNow self-service "
    "and the Virtual Agent reduces inbound volume and supports right-sized Tier-1 staffing. (2) Automation — "
    "RPA and ServiceNow workflow extension reduce per-ticket handling time on standard requests. (3) "
    "Delivery-center consolidation — eligible work consolidated to lower-cost U.S. delivery centers improves "
    "consistency and reduces per-hour cost without sacrificing service quality. Year-over-year TCO targets "
    "are agreed with the COR at the start of each option period and reported against monthly. The targets "
    "in §1.8 are the operating envelope.",
    bold_lead="Cost discipline drivers.")

# 3.6 ---
add_h(doc, "3.6 Draft Quality Control Plan (QCP)", 2)
add_para(doc,
    "A draft Quality Control Plan is summarized below; the full QCP will be submitted to the CO within 30 "
    "calendar days after award per PWS Subsection 9.1, refined in coordination with the COR, and maintained "
    "throughout all periods of performance. Per Section E Management Approach item (iii)(d), the QCP "
    "includes performance measures beyond those in the PWS — meaningful, measurable, and challenging.")

add_para(doc, "Table 3.6-A. QCP performance measures (PWS thresholds plus QCP additions).", bold_lead="")
add_table(doc,
    header=["Measure", "Definition", "Target", "Surveillance method", "Cadence"],
    rows=[
        ["Customer Service Survey (PWS)", "% positive customer survey responses", "≥96%",
         "100% sampling of completed surveys", "Monthly"],
        ["Personnel Availability (PWS)", "Per §9 formula", "≥80% daily",
         "100% inspection (system-derived)", "Daily"],
        ["Speed-to-Answer (PWS)", "Monthly avg time to live agent", "<2 min",
         "100% inspection (VCC-derived)", "Monthly"],
        ["Missed-Call Rate (PWS)", "% calls abandoned before answer", "<3%",
         "100% inspection (VCC-derived)", "Monthly"],
        ["Error Rate (PWS)", "Mis-routed/incorrect dispositions", "<3%",
         "Random sampling QA review", "Weekly"],
        ["IT Asset On-Time Delivery (PWS)", "% shipments on time", "≥95%",
         "100% inspection of shipment records", "Monthly"],
        ["Tier-1 Deflection Rate (QCP +)", "% deterministic resolved without human", "30% OP1 → 50% OP4",
         "ServiceNow-derived report", "Monthly"],
        ["First-Contact Resolution (QCP +)", "Tier-1 close-on-first-touch %", "70% OP1 → 78% OP3",
         "ServiceNow-derived report", "Monthly"],
        ["KB Currency Rate (QCP +)", "% articles reviewed within 90 days", "≥95%",
         "Periodic inspection", "Monthly"],
        ["MTTR by Priority (QCP +)", "Mean time to resolution", "Trend ↓",
         "ServiceNow-derived report", "Monthly"],
        ["Image Compliance Pre-Issue (QCP +)", "% laptops/desktops issued with verified patch + BIOS",
         "100%", "100% inspection", "Continuous"],
        ["Reopen Rate (QCP +)", "% tickets reopened within 7 days", "<5%",
         "ServiceNow-derived report", "Monthly"],
    ],
    col_widths=[1.7, 1.7, 0.9, 1.6, 0.8])

add_para(doc,
    "Surveillance methods follow PWS Section 9 — 100% inspection, periodic inspection, and random sampling — "
    "calibrated to the criticality and frequency of each performance requirement. A target miss triggers "
    "the corrective-action protocol: immediate corrective action for SLA threshold breach, root-cause "
    "analysis for repeat issues, and trend tracking that feeds the continuous-improvement cadence. Change "
    "management is disciplined, documented, with named approvers, and aligned to PWS Subsection 3.2.6.",
    bold_lead="Surveillance and corrective action.")

add_placeholder(doc,
    "Appendix D — Draft Quality Control Plan. Full QCP submitted within 30 calendar days post-award per "
    "PWS §9.1; this volume includes the framework, measures, and corrective-action protocol.")

# 3.7 Transition-In ----------
add_h(doc, "3.7 Draft Transition-In Plan", 2)
add_para(doc,
    "The transition window is one month — August 1, 2026 through August 31, 2026 — under firm-fixed-price "
    "terms. The objective is full performance responsibility on September 1, 2026 with no interruption or "
    "degradation of service levels, no SLA exemptions for the new period, and a closed Government-Furnished "
    "Equipment inventory. The Transition-In Project Manager, identified per PWS Subtask 3.1.2 item 2(a), "
    "is the single point of accountability for that outcome.",
    bold_lead="Transition framing.")

add_para(doc,
    "We have planned the transition in four overlapping phases.",
    bold_lead="Transition phases.")

add_table(doc,
    header=["Phase", "Window", "Primary outcomes", "Exit criteria"],
    rows=[
        ["Phase 1 — Mobilize and Mirror", "Days 1–10",
         "Key Personnel onboarded; security paperwork in motion; ServiceNow access; shadowing begins; "
         "incumbent coordination established",
         "Key Personnel cleared for system access; shadow rotation in place at all critical roles"],
        ["Phase 2 — Knowledge Capture", "Days 5–20",
         "Operational runbooks captured; KB inventory; ticket-trend history; institutional context; "
         "GFE inventory walk-down",
         "Captured artifacts stored in secure environment; GFE walk-down signed by CDC PAO"],
        ["Phase 3 — Operate Under Supervision", "Days 15–28",
         "We operate under incumbent supervision at increasing share of volume; CSAT and SLA performance "
         "tracked daily; corrective actions documented",
         "Trailing-week SLA performance at PWS Table 5 thresholds; CSAT ≥96%"],
        ["Phase 4 — Assume and Stabilize", "Days 25–31",
         "Full handoff; SLA assumed at 100% volume; final knowledge transfer artifacts; transition "
         "after-action review",
         "Sept 1 performance start with no SLA exemption; transition after-action complete"],
    ],
    col_widths=[1.4, 0.9, 2.7, 1.7])

add_para(doc,
    "Knowledge is captured against five categories: (1) operational runbooks per task and subtask; "
    "(2) the escalation matrix and Tier-1/2/3 handoff scripts; (3) the ServiceNow configuration baseline "
    "including knowledge articles, workflows, and Virtual Agent flows; (4) ticket-trend history and known "
    "recurring incidents; (5) institutional context — relationships, governance norms, and CDC-specific "
    "operating practices that do not appear in formal documentation. Captured artifacts are stored in the "
    "contractor's secure environment, transferred to CDC at request, and used as the operational baseline "
    "for the Base Period.",
    bold_lead="Knowledge transfer strategy.")

add_para(doc,
    "Per PWS Subtask 3.1.2 item 2(a) and Section E Transition-In Plan item 1, our approach includes "
    "structured coordination with the incumbent contractor. Our intent throughout is professional and "
    "procedurally clean — we treat incumbent staff with respect, engage exclusively through CDC-coordinated "
    "channels, and avoid any interaction that could be characterized as direct outreach for hiring or "
    "solicitation. Where incumbent staff are interested in continued CITS engagement under our team, those "
    "expressions of interest are routed through staffing-partner channels rather than through prime-to-"
    "individual contact, in keeping with sound procurement-integrity practice.",
    bold_lead="Coordination with the incumbent.")

add_para(doc,
    "Three audiences receive structured transition communications. (1) CDC end users — the approximately "
    "28,000 CDC personnel whose service experience cannot degrade. We coordinate with the OCIO Communications "
    "team to issue clear, plain-language updates on what is changing, what is staying the same, and how to "
    "reach support during the transition. (2) OCIO leadership — weekly transition steering committee "
    "touchpoints with the COR and identified leadership stakeholders. (3) CDC Centers, Institutes, and "
    "Offices — a structured outreach to each major CIO during transition to confirm continuity of support "
    "and surface any local concerns.",
    bold_lead="Communication plan with CDC stakeholders.")

add_para(doc, "Table 3.7-A. Top transition risks and mitigations.", bold_lead="")
add_table(doc,
    header=["Risk", "L", "I", "Mitigation"],
    rows=[
        ["Clearance throughput lag delaying performance start", "M", "H",
         "Pre-cleared bench; prioritized HHS scheduling; staffing-partner clearance pipeline."],
        ["Knowledge gaps in ServiceNow configuration / Virtual Agent flows", "M", "M",
         "Five-category knowledge capture; ServiceNow configuration export; named SME shadows."],
        ["Incumbent staff retention sentiment shifts during transition", "M", "M",
         "Staffing-partner-only outreach; protocol documented; OCIO Comms-coordinated messaging."],
        ["GFE inventory variance at handoff", "L", "H",
         "Continuous-discipline approach (§3.9); signed walk-down; PCO coverage; ServiceNow asset linkage."],
        ["Service-level dip in first weeks of performance start", "M", "M",
         "Phase 3 supervised operation; daily SLA tracking; corrective actions documented; weekly steering."],
    ],
    col_widths=[3.1, 0.5, 0.5, 2.7])

add_placeholder(doc,
    "Appendix E — Full Transition Risk Register. PMO to populate full register including specific named "
    "mitigations, risk owners, decision dates, and contingency triggers. Transition Plan uses the same labor "
    "categories as the Staffing Plan, per PWS Subtask 3.1.2 item 1(b).")

# 3.8 ---
add_h(doc, "3.8 Draft Transition-Out Plan", 2)
add_para(doc,
    "Per PWS Subtask 3.1.2 item 3 and Section E Management Approach item (vi), our approach to the orderly "
    "and efficient transfer of contract responsibilities at conclusion or termination ensures continuity of "
    "operations, preservation of institutional knowledge, and minimal disruption to Government activities. "
    "The Transition-Out Plan, presented in summary here and in full in Appendix F, addresses each element "
    "required by Section E.")

bullets = [
    "Knowledge transfer to the successor — documented processes, procedures, and operational information; "
    "runbooks; escalation matrix; ServiceNow knowledge base; ticket trend history; transition coordination "
    "with the incoming contractor as directed by CDC.",
    "Staffing transition strategy — support for incoming staff onboarding; clear protocol for our employees "
    "who may continue under the successor; reassignment of others within our broader portfolio.",
    "Detailed schedule — transition activities, milestones, and dependencies sequenced over the closeout "
    "window prescribed at the time.",
    "GFE return, transfer, and accountability — full property inventory reconciliation; chain-of-custody "
    "documentation; signed acceptance from the receiving party (Government or successor) per Section 5.",
    "Data management and transfer — all contract-related data, files, and records delivered in accordance "
    "with contract requirements and security protocols; ServiceNow records, knowledge articles, and "
    "operational documentation transferred per Government direction.",
    "Risk identification and mitigation — closeout risk register with specific transition risks (data "
    "continuity, staff retention through closeout, GFE accountability) and named mitigations.",
    "Coordination protocols — defined points of contact, communication cadence with the Government and "
    "successor, and escalation paths.",
]
for b in bullets:
    doc.add_paragraph(b, style="List Bullet")

add_placeholder(doc,
    "Appendix F — Draft Transition-Out Plan. PMO to populate full plan with closeout schedule, GFE "
    "reconciliation procedure, and successor coordination protocol.")

# 3.9 ---
add_h(doc, "3.9 Government-Furnished Equipment Accountability", 2)
add_para(doc,
    "Per PWS Section 5 and Section E Management Approach item (vii), GFE accountability is operated as a "
    "continuous discipline, not a transition event. We maintain a current GFE inventory in the appropriate "
    "property management system, reconciled to the annual wall-to-wall inventory required as a Task 1 "
    "deliverable in PWS Section 8. Property Custodial Officers under PWS Task 1.3 carry day-to-day "
    "accountability at FSS sites; the IT Asset Management function within Task 1 carries enterprise-level "
    "reconciliation. The 95%-or-greater on-time delivery rate for shipped IT equipment per PWS Section 9 is "
    "treated as a managed performance metric, with corrective action triggered if performance falls below "
    "threshold.")

add_para(doc,
    "Receipt and inventory at transition: documented walkthrough with the CDC Property Accountability "
    "Officer (PAO) or designated representative; signed inventory acceptance; serialization captured in the "
    "property management system. In-life maintenance: configuration management baseline; patch and "
    "anti-virus compliance verified before re-issue per PWS Subsection 1.2.1 item l; tracked maintenance "
    "history. End-of-life disposition: pickup and processing per OCIO standard; data sanitization per NIST "
    "800-88; Report of Survey produced where required. Quarterly audit cadence is calendared and reported "
    "to the COR.",
    bold_lead="Accountability protocol.")

# 3.10 ---
add_h(doc, "3.10 Conditions and Assumptions", 2)
add_para(doc,
    "Our proposal is responsive to the RFQ as issued and is predicated on the terms and conditions of the "
    "RFQ, in compliance with the Volume II requirement at item (c)(1). Where assumptions are necessary to "
    "interpret an ambiguous requirement, we have framed them so that they do not adversely impact the "
    "Government's requirement, consistent with the Section E direction that the Government reserves the "
    "right to reject proposals with adverse assumptions.")

add_table(doc,
    header=["#", "Assumption", "Basis", "Impact if violated"],
    rows=[
        ["A1", "CDC procures and provides the NICE VCC subscription for Task 1.1",
         "PWS Task 1.1 item 6 indicates VCC subscription is procured by CDC; ODC table includes VCC "
         "license cost separately",
         "Cost reallocation; conversation with CO before manifesting in financials"],
        ["A2", "ServiceNow license access for contractor personnel is provided by CDC at no contractor cost",
         "ServiceNow is CDC's platform of record; PWS does not require contractor-procured ServiceNow",
         "Cost reallocation; conversation with CO"],
        ["A3", "ServiceNow Virtual Agent and Predictive Intelligence are CDC-provided platform features and "
         "are not 'AI products' under Section E definition",
         "These are ServiceNow-platform-native capabilities CDC may license; clarification submitted to "
         "CO May 12",
         "Solution narrative for Task 1.1 deflection adjusted to KB-only path"],
        ["A4", "OCONUS travel currency exchange and ME insurance treatment follow PWS Section 12 norms",
         "PWS Section 12 specifies insurance and STEP requirements; treatment is standard federal practice",
         "Travel-cost variance reported in Monthly Financial Report"],
        ["A5", "CDC personnel access during transition is sufficient for shadow rotation and knowledge capture",
         "PWS Subtask 3.1.2 item 2(b) provides for incoming PM transition lead access",
         "Transition timeline extends; corrective discussion with CO"],
        ["A6", "Tools not explicitly named as CDC-furnished are the contractor's responsibility unless "
         "approved by CDC",
         "PWS does not enumerate every tool; standard federal contracting practice",
         "Contractor absorbs cost or seeks CDC approval before introduction"],
    ],
    col_widths=[0.4, 1.9, 2.4, 2.0])

# =======================================================================
# 4. SIMILAR EXPERIENCE
# =======================================================================
add_h(doc, "4. Similar Experience", 1)
add_para(doc,
    "Per Section E Evaluation Factor 4 and the corresponding evaluation criteria, we present three references "
    "of relevant contracts similar in scope, size, and complexity, performed within the last three years. "
    "Each reference demonstrates capability against the operating model required by CITS — federal IT "
    "customer support at scale, multi-site delivery, ServiceNow-based ITSM operations, ITIL discipline, and "
    "measurable performance against published service levels and cost-of-ownership targets.")

add_para(doc,
    "Our three references collectively address: (1) federal health agency IT customer support at a population "
    "scale comparable to CDC's approximately 28,000 personnel; (2) tiered ITSM operations on ServiceNow at "
    "federal civilian or DoD-health scale; (3) distributed-site or OCONUS service delivery comparable to "
    "CITS Global Activities. Each reference is current or has concluded within the last three years, with "
    "two customer contacts available for verification.",
    bold_lead="Selection rationale.")

# 4.1 ---
add_h(doc, "4.1 Reference 1 — Federal Health Agency Enterprise IT Service Desk", 2)
add_para(doc, "Contract number, customer, and title.",
    bold_lead="(a)")
add_placeholder(doc,
    "Capture lead to provide: contract number, customer/agency name, contract title.")
add_para(doc,
    "Tiered ITSM service delivery to a federal health agency comparable in scale to CDC. Service Desk on "
    "ServiceNow with multi-channel intake (phone, portal, email, chat); deskside support across multiple "
    "campuses; distributed-site coverage; integration with Microsoft 365 and identity platforms. The "
    "contract included a transition phase coordinated with the predecessor incumbent, a knowledge-management "
    "discipline tied to deflection metrics, and a continuous-improvement cadence reporting to a federal CIO. "
    "Significant accomplishments included deflection-rate increases over baseline within the first option "
    "period and sustained Customer Service Survey performance above the contractual threshold.",
    bold_lead="(b) Narrative of work.")
add_placeholder(doc,
    "Capture lead to add specifics on problems encountered, corrective actions, and significant "
    "accomplishments — quantified where defensible (e.g., deflection % achieved, CSAT %, MTTR delta).")
add_para(doc, "Dollar value, contract type, period of performance, place of performance, personnel.",
    bold_lead="(c)")
add_placeholder(doc,
    "Capture lead to provide dollar value, contract type (T&M / FFP / hybrid), POP, place of performance, "
    "and number/types of personnel used.")
add_para(doc, "Two customer POCs.", bold_lead="(d)")
add_placeholder(doc,
    "Capture lead to provide two customer contacts (COR/CO/Program POC/Project Officer) including name, "
    "address, phone.")

# 4.2 ---
add_h(doc, "4.2 Reference 2 — Federal Civilian or DoD-Health ServiceNow ITSM at Scale", 2)
add_para(doc, "Contract number, customer, and title.", bold_lead="(a)")
add_placeholder(doc, "Capture lead to provide.")
add_para(doc,
    "ServiceNow-based ITSM operations at federal civilian or DoD-health scale, including ServiceNow App "
    "Engine and workflow extension, knowledge management, change and release management, and reporting. "
    "The contract scope included multi-site delivery, integration with identity platforms, and a quality "
    "control plan with measures beyond the published SLAs. Significant accomplishments included reduced "
    "MTTR by priority and demonstrated year-over-year cost-of-ownership reduction.",
    bold_lead="(b) Narrative of work.")
add_placeholder(doc, "Capture lead to add specifics — quantified where defensible.")
add_para(doc, "Dollar value, contract type, period of performance, place of performance, personnel.",
    bold_lead="(c)")
add_placeholder(doc, "Capture lead to provide.")
add_para(doc, "Two customer POCs.", bold_lead="(d)")
add_placeholder(doc, "Capture lead to provide.")

# 4.3 ---
add_h(doc, "4.3 Reference 3 — Distributed-Site / OCONUS Service Delivery", 2)
add_para(doc, "Contract number, customer, and title.", bold_lead="(a)")
add_placeholder(doc, "Capture lead to provide.")
add_para(doc,
    "Distributed-site or OCONUS service delivery comparable to CITS Global Activities. Scope included "
    "multi-region engineering coverage, SD-WAN and site-to-site VPN integration, satellite-link operating "
    "norms, regional runbooks, and travel-readiness discipline (FACT, STEP, medical clearance). Significant "
    "accomplishments included demonstrated 24-48 hour deployment readiness and high site-availability "
    "performance across the OCONUS footprint.",
    bold_lead="(b) Narrative of work.")
add_placeholder(doc, "Capture lead to add specifics — quantified where defensible.")
add_para(doc, "Dollar value, contract type, period of performance, place of performance, personnel.",
    bold_lead="(c)")
add_placeholder(doc, "Capture lead to provide.")
add_para(doc, "Two customer POCs.", bold_lead="(d)")
add_placeholder(doc, "Capture lead to provide.")

# 4.4 ---
add_h(doc, "4.4 Relevance Crosswalk to CITS", 2)
add_para(doc,
    "The crosswalk below identifies, for each PWS task family, which reference demonstrates relevant "
    "performance. This crosswalk is intended to make the suitability of our experience to CITS easy for "
    "the evaluator to verify.")

add_table(doc,
    header=["CITS task family", "Ref 1", "Ref 2", "Ref 3"],
    rows=[
        ["Task 1.1 ITSDS — Tier-1 multi-channel service desk on ServiceNow", "✔ primary", "✔ supporting", ""],
        ["Task 1.2 DSS — Multi-campus deskside / endpoint discipline", "✔ primary", "", ""],
        ["Task 1.3 EPS/FSS — Emergency response / field deployable", "✔ supporting", "", "✔ supporting"],
        ["Task 2.1 GA — OCONUS engineering and travel readiness", "", "", "✔ primary"],
        ["Task 3.1 PMS — Federal IT services PMO and QASP", "✔ supporting", "✔ primary", "✔ supporting"],
        ["Task 6.1 / 8.1–8.3 / 9.1 / 10.1–10.5 — Platform engineering & automation", "", "✔ primary", ""],
        ["TCO discipline / deflection / continuous improvement", "✔ supporting", "✔ primary", "✔ supporting"],
    ],
    col_widths=[3.6, 1.0, 1.0, 1.1])

# =======================================================================
# INTERNAL — clarifying questions (remove before submission)
# =======================================================================
add_h(doc, "Clarifying Questions for Internal Review", 1)
add_para(doc,
    "Internal note to capture team — to be removed before submission. The following items require capture-"
    "lead decision or further information before this draft is finalized for color-team review.")

questions = [
    ("Q1 — Set-aside status.",
     "Orientation flags ambiguity in the SF-18 cover-page checkbox versus body language treatment of "
     "set-aside. Submit clarification to CO by May 12."),
    ("Q2 — Staffing matrix and resume page count.",
     "Section E states the technical volume shall not exceed 35 pages and excludes the DHHS Section 508 "
     "Product Assessment Template. It does not explicitly address whether the staffing matrix and Key "
     "Personnel resumes count toward the 35-page limit. Submit clarification to CO by May 12. (RFQ p.145.)"),
    ("Q3 — AI products definition.",
     "Section E prohibits AI products in the proposal and in performance after award (except CDC-"
     "provided/approved). PWS Subsection 3.2.5 names AI as an example of innovative technology. Submit "
     "clarification: Does ServiceNow Virtual Agent qualify as a CDC-provided/approved capability that can "
     "be referenced in our solution narrative? Does ServiceNow Predictive Intelligence?"),
    ("Q4 — Naming subcontractors in Volume I.",
     "Per protest-defensibility guidance and Deloitte standard practice, named subs require written "
     "approval before appearing in the technical narrative. Confirm whether subcontractors are named "
     "explicitly in Volume I or referenced generically in Volume I and named in Volume II."),
    ("Q5 — Past Performance volume.",
     "Section E identifies four evaluation factors. Past Performance is evaluated separately on a pass/fail "
     "basis via CPARS. Confirm: Past Performance is satisfied via CPARS lookup with no separately-submitted "
     "past performance volume from us, while Similar Experience (the three references in §4) lives within "
     "Volume I."),
    ("Q6 — Win-theme threading review.",
     "Confirm: this draft threads the orientation's governing POV ('keep what works, deflect demand, "
     "modernize capabilities') and the executive win themes (fresh leadership without disruption, "
     "operational stability first, embedded CIO engagement, flexibility amid HHS uncertainty, continuity of "
     "trusted talent, ServiceNow-native automation, delivery-center-based teams, HCD embedded). HCD is now "
     "embedded in §1.4 and threaded through §1.3.1 / §1.3.2 — confirm whether Mary Peck's CXO elevation "
     "should be referenced more explicitly anywhere."),
    ("Q7 — Persona stress-test.",
     "Recommend running this draft through Jane Doe (CXO/HCD) and John Doe (ops/metrics) personas before "
     "pink-team. Personas reference 2025 RFP context and need refresh for 2026 details (Dynanet not GDIT; "
     "$110M–$150M ceiling not $200M)."),
    ("Q8 — Page-budget realism.",
     "Current expanded draft targets ~33 pages of body at 11pt TNR with 0.75-inch margins. Final layout "
     "depends on Q2 (matrix/resumes excluded?) and how much of §4 is in-line versus appendix. Recommend a "
     "layout pass once Q2 is resolved."),
    ("Q9 — Buzzword and Deloitte-first audit.",
     "Recommend running the project's buzzword kill-list audit and the 'Deloitte-first' lead-sentence audit "
     "before pink-team, per Master Context Pack §7."),
    ("Q10 — Cross-volume consistency.",
     "Volume II hour build, Appendix A staffing matrix, and the LCAT-aligned narrative in §2 must be "
     "cross-checked for consistency before submission. Inconsistencies are protest fodder."),
    ("Q11 — Master Context Pack stale facts.",
     "Pack states 40-page limit and references GDIT as continuity sub. Actual RFQ Section E sets 35-page "
     "limit; orientation names Dynanet as primary teaming partner. Pack should be updated and re-baselined "
     "before next color team."),
]
for h, b in questions:
    add_para(doc, b, bold_lead=h)

# Save
doc.save("CITS_VolumeI_Technical_DRAFT.docx")
print("Saved CITS_VolumeI_Technical_DRAFT.docx")
