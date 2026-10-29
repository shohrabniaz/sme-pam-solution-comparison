"""Build Project 40 PAM comparison pack: framework, evaluation, cost, decision guide, report, slides."""
from __future__ import annotations

from datetime import date
from pathlib import Path
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from pptx import Presentation
from pptx.dml.color import RGBColor as PRgb
from pptx.util import Inches, Pt as PPt

ROOT = Path(__file__).resolve().parent
FRAME = ROOT / "01_Framework"
EVAL = ROOT / "02_Evaluation"
OUT = ROOT / "Final_Submission"
TODAY = date.today().isoformat()
NAVY = RGBColor(0x1F, 0x4E, 0x79)
HEADER = "Md Shohrab Hossen Niaz | Project 40 | Laboratory / desk evaluation"


def set_run(p, text, size=11, bold=False, color=None):
    p.clear()
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(size)
    r.bold = bold
    if color:
        r.font.color.rgb = color


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = NAVY


def save(doc, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


# Candidates evaluated for a ~35-person fictional MSP (Coastal Managed IT)
# A = Free-tier Graph JIT (Project 38) — cloud-native light
# B = Entra PIM — cloud-native paid
# C = Open-source PAM pattern (Teleport / Vault-class) — self-hosted

CRITERIA = [
    ("Credential vaulting", "Store and check out shared/admin secrets"),
    ("Session control / recording", "Who used what, when; optional recording"),
    ("Just-in-time elevation", "Time-boxed privilege, not always-on"),
    ("Approval workflow", "Second person before elevation"),
    ("Audit evidence", "Exportable records for review"),
    ("Licence cost (AUD/yr model)", "Cash cost for a 35-person SME"),
    ("Ops burden", "Who runs it day to day"),
    ("Fits Entra ID Free lab", "Can we prove it without paid Azure"),
    ("Multi-tenant MSP fit", "Useful when one firm supports many clients"),
    ("Time to first control", "Days to a working pilot"),
]

# Scores 1-5; higher is better for SME fit on that criterion
# Cost inverted later in narrative: lower cost = better for SME
SCORES = {
    "A_Free_JIT": {
        "Credential vaulting": 1,
        "Session control / recording": 2,
        "Just-in-time elevation": 4,
        "Approval workflow": 2,  # name in JSON only
        "Audit evidence": 4,
        "Licence cost (AUD/yr model)": 5,  # $0
        "Ops burden": 3,
        "Fits Entra ID Free lab": 5,
        "Multi-tenant MSP fit": 3,
        "Time to first control": 5,
    },
    "B_Entra_PIM": {
        "Credential vaulting": 2,
        "Session control / recording": 3,
        "Just-in-time elevation": 5,
        "Approval workflow": 5,
        "Audit evidence": 5,
        "Licence cost (AUD/yr model)": 2,  # paid governance
        "Ops burden": 4,
        "Fits Entra ID Free lab": 1,
        "Multi-tenant MSP fit": 4,
        "Time to first control": 3,
    },
    "C_OpenSource_PAM": {
        "Credential vaulting": 5,
        "Session control / recording": 5,
        "Just-in-time elevation": 4,
        "Approval workflow": 4,
        "Audit evidence": 4,
        "Licence cost (AUD/yr model)": 4,  # software free; hosting/time not free
        "Ops burden": 2,
        "Fits Entra ID Free lab": 2,
        "Multi-tenant MSP fit": 4,
        "Time to first control": 2,
    },
}

COST = [
    ("A Free-tier Graph JIT", 0, 2, "PowerShell + Entra Free; 2 hours/week review grants"),
    ("B Entra PIM", 12000, 4, "Model: identity-governance licences for admins + setup; vendor quotes vary"),
    ("C Open-source PAM", 1800, 12, "Model: small VM/hosting + 12 hours/month care; software licence $0"),
]


def build_framework():
    doc = Document()
    doc.add_heading("PAM requirements and criteria framework", 0)
    set_run(doc.add_paragraph(), HEADER, 11)
    set_run(doc.add_paragraph(), f"Draft date: {TODAY}", 11)
    heading(doc, "SME context", 1)
    doc.add_paragraph(
        "Coastal Managed IT is a fictional ~35-person MSP. Budget is tight. The firm "
        "already uses Microsoft Entra ID Free for the laboratory. Enterprise PAM licences "
        "are hard to justify until a control is proven."
    )
    heading(doc, "Core SME PAM requirements", 1)
    for name, desc in CRITERIA:
        doc.add_paragraph(f"{name}: {desc}", style="List Bullet")
    heading(doc, "Candidates", 1)
    doc.add_paragraph(
        "A — Free-tier Graph JIT (Project 38 scripts): time-boxed group membership, "
        "JSON audit, no paid licence."
    )
    doc.add_paragraph(
        "B — Entra Privileged Identity Management: portal approval, time-limited roles, "
        "paid identity-governance tier."
    )
    doc.add_paragraph(
        "C — Open-source PAM pattern (Teleport / HashiCorp Vault class): vaulting and "
        "session control, self-hosted or managed open-source stack."
    )
    heading(doc, "Scoring", 1)
    doc.add_paragraph(
        "Each criterion is scored 1 to 5 for SME fit. Higher is better. Cost and ops "
        "burden are scored so that cheaper and lighter options score higher. This is a "
        "desk-plus-lab comparison: A was piloted as scripts; B and C are evaluated from "
        "public capability and cost models, not a paid tenant deploy."
    )
    save(doc, FRAME / "PAM_Requirements_Criteria_Framework.docx")


def build_workbook():
    EVAL.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    fill = PatternFill("solid", fgColor="1F4E79")
    white = Font(color="FFFFFF", bold=True)

    ws = wb.active
    ws.title = "CriteriaScores"
    headers = ["Criterion", "A Free JIT", "B Entra PIM", "C Open-source PAM"]
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = white
    for name, _ in CRITERIA:
        ws.append(
            [
                name,
                SCORES["A_Free_JIT"][name],
                SCORES["B_Entra_PIM"][name],
                SCORES["C_OpenSource_PAM"][name],
            ]
        )
    totals = [
        "TOTAL",
        sum(SCORES["A_Free_JIT"].values()),
        sum(SCORES["B_Entra_PIM"].values()),
        sum(SCORES["C_OpenSource_PAM"].values()),
    ]
    ws.append(totals)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 28

    ws2 = wb.create_sheet("CostModel")
    ws2.append(["Option", "AUD/year model", "Hours/month ops", "Notes"])
    for cell in ws2[1]:
        cell.fill = fill
        cell.font = white
    for row in COST:
        ws2.append(list(row))
    for col in ws2.columns:
        ws2.column_dimensions[col[0].column_letter].width = 32

    ws3 = wb.create_sheet("DecisionMatrix")
    ws3.append(["If the SME needs…", "Prefer", "Why"])
    for cell in ws3[1]:
        cell.fill = fill
        cell.font = white
    for row in [
        ["A first control this month on Entra Free", "A Free JIT", "Fast, $0 licence, proves time-boxing"],
        ["Portal approval and Entra role PIM", "B Entra PIM", "Native Microsoft path when licence exists"],
        ["Secret vaulting and session recording", "C Open-source PAM", "Strongest vault/session capability"],
        ["MSP multi-client secrets at scale", "C then B", "Vault first; PIM for Entra roles"],
        ["Lowest cash cost for directory elevation only", "A Free JIT", "Group membership time-box is enough to start"],
    ]:
        ws3.append(row)
    for col in ws3.columns:
        ws3.column_dimensions[col[0].column_letter].width = 40

    path = EVAL / "PAM_Evaluation_Workbook.xlsx"
    wb.save(path)
    return path


def build_comparison_doc():
    doc = Document()
    doc.add_heading("Comparative evaluation results", 0)
    set_run(doc.add_paragraph(), HEADER, 11)
    heading(doc, "Method", 1)
    doc.add_paragraph(
        "I scored three options against the SME criteria framework. Option A was "
        "exercised as the Project 38 PowerShell pilot on Entra ID Free. Options B and C "
        "were scored from documented capabilities and a cost model. I did not buy PIM "
        "and I did not stand up a full Teleport/Vault cluster in production."
    )
    heading(doc, "Totals (higher is better for SME fit)", 1)
    a = sum(SCORES["A_Free_JIT"].values())
    b = sum(SCORES["B_Entra_PIM"].values())
    c = sum(SCORES["C_OpenSource_PAM"].values())
    doc.add_paragraph(f"A Free-tier Graph JIT: {a}/50")
    doc.add_paragraph(f"B Entra PIM: {b}/50")
    doc.add_paragraph(f"C Open-source PAM pattern: {c}/50")
    heading(doc, "What the scores mean", 1)
    doc.add_paragraph(
        "A wins on cost, speed, and fit to the free laboratory. It loses on vaulting and "
        "real second-person approval. B wins on approval workflow and Entra-native JIT. "
        "It loses on licence cost and cannot be proven on Entra Free. C wins on vaulting "
        "and session control. It loses on ops burden and time to first control."
    )
    heading(doc, "Honest limits", 1)
    doc.add_paragraph(
        "The AUD figures for B and C are models for a 35-person firm, not a vendor quote. "
        "Live Graph sign-in for a fresh Project 38 grant was not available in this "
        "session; the free JIT path still rests on the Project 38 design and the earlier "
        "privileged-group laboratory evidence."
    )
    save(doc, EVAL / "Comparative_Evaluation_Results.docx")


def build_cost_doc():
    doc = Document()
    doc.add_heading("Cost and operability analysis", 0)
    set_run(doc.add_paragraph(), HEADER, 11)
    heading(doc, "Cash and care", 1)
    for name, aud, hours, note in COST:
        doc.add_paragraph(f"{name}: about AUD {aud}/year model, {hours} hours/month. {note}")
    heading(doc, "Reading the numbers", 1)
    doc.add_paragraph(
        "Free-tier JIT is not free of people time. Someone must approve by name and run "
        "or schedule revoke. PIM moves that into the portal and the licence bill. "
        "Open-source PAM moves the bill into hosting and whoever patches the PAM box."
    )
    heading(doc, "Operability for an MSP", 1)
    doc.add_paragraph(
        "For Coastal Managed IT, the first operational win is stopping standing "
        "SG-CMIT-Privileged-Admins membership. That does not need a vault. It needs a "
        "short elevation window and a file you can show. Vaulting shared client "
        "passwords is a second project once the directory elevation path is stable."
    )
    save(doc, EVAL / "Cost_Operability_Analysis.docx")


def build_decision_guide():
    doc = Document()
    doc.add_heading("SME PAM decision guide", 0)
    set_run(doc.add_paragraph(), HEADER, 11)
    heading(doc, "Start here", 1)
    doc.add_paragraph(
        "1. Inventory standing privileged groups and disabled orphans.",
        style="List Number",
    )
    doc.add_paragraph(
        "2. Put day-to-day work on role packs. Keep admin out of the joiner pack.",
        style="List Number",
    )
    doc.add_paragraph(
        "3. Pilot free-tier time-boxed elevation (Project 38 path) for directory admin.",
        style="List Number",
    )
    doc.add_paragraph(
        "4. Buy Entra PIM when portal approval and Entra role JIT are required.",
        style="List Number",
    )
    doc.add_paragraph(
        "5. Add open-source or commercial vaulting when shared secrets and session "
        "recording matter more than directory group membership.",
        style="List Number",
    )
    heading(doc, "Do not", 1)
    doc.add_paragraph(
        "Do not buy enterprise PAM only because the catalogue recommended it. Do not "
        "leave standing privileged groups in place while you evaluate tools. Do not "
        "treat a disabled account as deprovisioned if it is still in an admin group."
    )
    save(doc, EVAL / "SME_PAM_Decision_Guide.docx")


def build_report():
    OUT.mkdir(parents=True, exist_ok=True)
    doc = Document()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(
        title,
        "Open-source versus cloud PAM for a small managed service provider\n"
        "A criteria-based comparison for a constrained SME budget",
        16,
        True,
        NAVY,
    )
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(
        meta,
        "Md Shohrab Hossen Niaz\n"
        "Project 40 — Privileged Access Management\n"
        "Fictional case: Coastal Managed IT Pty Ltd | Desk evaluation + free-tier lab path\n"
        "Word count (excluding title block and reference list): PLACEHOLDER",
        11,
    )

    a = sum(SCORES["A_Free_JIT"].values())
    b = sum(SCORES["B_Entra_PIM"].values())
    c = sum(SCORES["C_OpenSource_PAM"].values())

    sections = [
        (
            "Abstract",
            [
                "I compared three privileged-access paths for a fictional 35-person MSP: a "
                "free-tier Microsoft Graph just-in-time group elevation, Entra Privileged "
                "Identity Management, and an open-source PAM pattern in the Teleport or "
                "Vault class. I scored them on vaulting, session control, JIT, approval, "
                "audit, cost, ops burden, free-lab fit, MSP multi-tenant fit, and time to "
                f"first control. Totals were {a}/50 for free JIT, {b}/50 for PIM, and {c}/50 "
                "for open-source PAM. The free path is the right first control. PIM is the "
                "right buy when portal approval is required. Open-source wins when secret "
                "vaulting and session recording matter. The AUD figures for paid and "
                "self-hosted options are models, not vendor quotes."
            ],
        ),
        (
            "1. Introduction",
            [
                "Privileged access management matters for an MSP because one admin account "
                "can reach many customer environments. Enterprise PAM licences are hard to "
                "justify for a firm that still has standing membership of an admin group. "
                "Project 40 asks how open-source and cloud-native options compare for core "
                "SME needs, what they cost to run, and which fits firms at different maturity.",
                "This work follows Project 38. That project designed a free-tier time-boxed "
                "elevation into SG-CMIT-Privileged-Admins. Project 40 places that path next "
                "to PIM and an open-source PAM pattern so the firm can choose without "
                "guessing.",
            ],
        ),
        (
            "2. Related work and standards",
            [
                "The Essential Eight control on restricting administrative privileges sets "
                "the outcome. ISO/IEC 27001:2022 A.8.2 says the same in management-system "
                "language. Microsoft documents PIM as the Entra-native product for "
                "approval-based, time-limited roles. Open-source projects such as Teleport "
                "and HashiCorp Vault publish capability for vaulting and session control. "
                "Tuononen (2023) treated MSP privilege as a liability problem. I treated "
                "the buying decision as a criteria problem: what does the SME need first, "
                "and what can it prove without a paid tenant.",
            ],
        ),
        (
            "3. Method",
            [
                "I built a ten-criterion framework for a ~35-person MSP. I scored three "
                "candidates from 1 to 5 on each criterion. Option A is the Project 38 "
                "free-tier Graph JIT path. Option B is Entra PIM. Option C is an "
                "open-source PAM pattern. I also built a simple annual cost and monthly "
                "hours model. Option A was tied to the laboratory scripts. Options B and C "
                "were not deployed on paid infrastructure in this pass. Live Graph "
                "authentication for a fresh grant was unavailable, so A rests on the "
                "Project 38 design and the earlier privileged-group laboratory evidence "
                "rather than a new interactive Graph run.",
            ],
        ),
        (
            "4. Results",
            [
                f"Totals: A {a}/50, B {b}/50, C {c}/50. A leads on licence cost, free-lab "
                "fit, and time to first control. B leads on approval workflow and "
                "Entra-native JIT quality. C leads on credential vaulting and session "
                "control. The cost model put A at about AUD 0 licence plus a few hours a "
                "week, B at roughly AUD 12,000 a year for a small governance footprint, "
                "and C at roughly AUD 1,800 hosting plus about 12 hours a month of care. "
                "Those dollar figures are planning models.",
                "The practical reading for Coastal Managed IT is ordered, not absolute. "
                "Stop standing privileged groups with free JIT first. Buy PIM when the "
                "firm needs portal approval and Entra role elevation. Add open-source or "
                "commercial vaulting when shared secrets and session recording become the "
                "main gap.",
            ],
        ),
        (
            "5. Discussion",
            [
                "A single winner is the wrong deliverable. Maturity matters. A firm that "
                "still has disabled accounts in admin groups does not need a vault first. "
                "It needs membership that expires. A firm that already time-boxes Entra "
                "roles and still shares root passwords needs vaulting more than another "
                "directory script.",
                "Limits are explicit. I did not purchase PIM. I did not run a production "
                "Teleport or Vault cluster. Cost numbers will move with vendor quotes and "
                "hosting choices. The comparison is still useful because the criteria and "
                "the decision order stay stable when the quotes change.",
            ],
        ),
        (
            "6. Conclusion",
            [
                "For a small MSP budget, open-source and cloud PAM are not substitutes for "
                "each other. Free-tier time-boxed directory elevation is the affordable "
                "first control. Entra PIM is the cloud-native upgrade when approval and "
                "role JIT are required. Open-source PAM is the path when vaulting and "
                "session control dominate. Choose by the gap you still have, not by the "
                "largest brochure.",
            ],
        ),
        (
            "References",
            [
                "Australian Signals Directorate. (n.d.). Essential Eight. https://www.cyber.gov.au/resources-business-and-government/essential-cyber-security/essential-eight",
                "HashiCorp. (n.d.). Vault documentation. https://developer.hashicorp.com/vault/docs",
                "International Organization for Standardization. (2022). ISO/IEC 27001:2022.",
                "Microsoft. (n.d.). What is Privileged Identity Management? https://learn.microsoft.com/entra/id-governance/privileged-identity-management/pim-configure",
                "Teleport. (n.d.). Teleport documentation. https://goteleport.com/docs/",
                "Tuononen, H. (2023). Privileged access management model for a managed service provider (Master's thesis, JAMK). http://theseus.fi/handle/10024/800216",
            ],
        ),
    ]

    body_words = 0
    for title_text, paras in sections:
        heading(doc, title_text, 1)
        for para in paras:
            doc.add_paragraph(para)
            if title_text != "References":
                body_words += len(para.split())

    for p in doc.paragraphs:
        if "PLACEHOLDER" in p.text:
            set_run(
                p,
                "Md Shohrab Hossen Niaz\n"
                "Project 40 — Privileged Access Management\n"
                "Fictional case: Coastal Managed IT Pty Ltd | Desk evaluation + free-tier lab path\n"
                f"Word count (excluding title block and reference list): {body_words}",
                11,
            )
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            break

    path = OUT / "01_Project40_Report.docx"
    doc.save(path)
    print("BODY_WORDS", body_words)
    return path, body_words


def build_slides():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    navy = PRgb(0x1F, 0x4E, 0x79)

    def add_slide(title, bullets, notes):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12), Inches(1))
        r = box.text_frame.paragraphs[0].add_run()
        r.text = title
        r.font.size = PPt(28)
        r.font.bold = True
        r.font.color.rgb = navy
        body = slide.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12), Inches(5))
        tf = body.text_frame
        tf.word_wrap = True
        first = True
        for b in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            run = p.add_run()
            run.text = b
            run.font.size = PPt(18)
            run.font.color.rgb = PRgb(0x33, 0x33, 0x33)
            p.space_after = PPt(10)
        slide.notes_slide.notes_text_frame.text = notes

    a = sum(SCORES["A_Free_JIT"].values())
    b = sum(SCORES["B_Entra_PIM"].values())
    c = sum(SCORES["C_OpenSource_PAM"].values())

    add_slide(
        "Open-source vs cloud PAM",
        [
            "Fictional MSP, ~35 staff, constrained budget.",
            "Compare free JIT, Entra PIM, and open-source PAM.",
            "Choose by the gap, not the brochure.",
        ],
        "This is Project 40. It sits after the free-tier JIT pilot.",
    )
    add_slide(
        "Three options",
        [
            "A: Free-tier Graph JIT — time-boxed admin group.",
            "B: Entra PIM — portal approval, paid licence.",
            "C: Open-source PAM — vaulting and sessions, you host it.",
        ],
        "A was piloted as scripts. B and C are scored from capability and cost models.",
    )
    add_slide(
        "Scores out of 50",
        [
            f"A Free JIT: {a}",
            f"B Entra PIM: {b}",
            f"C Open-source PAM: {c}",
            "A wins cost and speed. B wins approval. C wins vaulting.",
        ],
        "Treat AUD figures as models. Do not quote them as a vendor price.",
    )
    add_slide(
        "Decision order",
        [
            "Stop standing privilege with free JIT first.",
            "Buy PIM when portal approval is required.",
            "Add vaulting when shared secrets are the main gap.",
        ],
        "That order matches Essential Eight admin restriction without forcing a big licence on day one.",
    )
    add_slide(
        "Where I landed",
        [
            "No single winner for every SME.",
            "Maturity decides the next buy.",
            "Measure leftover privileged hours, not tool logos.",
        ],
        "Questions?",
    )
    path = OUT / "02_Project40_Presentation.pptx"
    prs.save(path)
    return path


def build_readme():
    doc = Document()
    doc.add_heading("Project submission", 0)
    doc.add_paragraph(
        "Md Shohrab Hossen Niaz\n"
        "Project 40 — open-source versus cloud PAM for SME budgets\n"
        "Coastal Managed IT is fictional. Free-tier lab path plus desk evaluation."
    )
    doc.add_heading("In this folder", 1)
    for line in [
        "01 report.",
        "02 presentation.",
        "Appendix A. Evaluation workbook.",
        "Appendix B. Framework, comparison, cost analysis, decision guide.",
        "Appendix C. Link note to Project 38 free JIT scripts.",
    ]:
        doc.add_paragraph(line)
    save(doc, OUT / "00_Read_Me_First.docx")


def build_link_note():
    doc = Document()
    doc.add_heading("Laboratory path used for option A", 0)
    doc.add_paragraph(
        "Option A in this comparison is the Project 38 free-tier JIT pack at "
        "../Project38_JIT_Privileged_Access/. Scripts 01–04 inventory standing privilege, "
        "grant time-boxed elevation, revoke on expiry, and measure exposure hours."
    )
    doc.add_paragraph(
        "Live Graph interactive sign-in was not completed in this session. The "
        "comparison still uses that design as the affordable cloud-native candidate."
    )
    save(doc, OUT / "Appendix" / "C_Lab_Link" / "Project38_Free_JIT_Link.docx")


def main():
    FRAME.mkdir(parents=True, exist_ok=True)
    EVAL.mkdir(parents=True, exist_ok=True)
    build_framework()
    build_workbook()
    build_comparison_doc()
    build_cost_doc()
    build_decision_guide()
    report, words = build_report()
    slides = build_slides()
    build_readme()
    build_link_note()

    app_a = OUT / "Appendix" / "A_Evaluation"
    app_b = OUT / "Appendix" / "B_Documents"
    app_a.mkdir(parents=True, exist_ok=True)
    app_b.mkdir(parents=True, exist_ok=True)
    shutil.copy2(EVAL / "PAM_Evaluation_Workbook.xlsx", app_a / "PAM_Evaluation_Workbook.xlsx")
    for name in [
        "Comparative_Evaluation_Results.docx",
        "Cost_Operability_Analysis.docx",
        "SME_PAM_Decision_Guide.docx",
    ]:
        shutil.copy2(EVAL / name, app_b / name)
    shutil.copy2(FRAME / "PAM_Requirements_Criteria_Framework.docx", app_b / "PAM_Requirements_Criteria_Framework.docx")
    print("Wrote", report, words)
    print("Wrote", slides)


if __name__ == "__main__":
    main()
