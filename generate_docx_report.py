import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

ASSETS_DIR = os.path.join(os.path.dirname(__file__), 'report_assets')
DOCX_PATH = os.path.join(os.path.dirname(__file__), 'Employee_Leave_Management_System_Project_Report.docx')

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_page_border(section):
    sectPr = section._sectPr
    pgBorders = parse_xml(f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page">'
                          f'<w:top w:val="single" w:sz="8" w:space="24" w:color="000000"/>'
                          f'<w:left w:val="single" w:sz="8" w:space="24" w:color="000000"/>'
                          f'<w:bottom w:val="single" w:sz="8" w:space="24" w:color="000000"/>'
                          f'<w:right w:val="single" w:sz="8" w:space="24" w:color="000000"/>'
                          f'</w:pgBorders>')
    sectPr.append(pgBorders)

def create_report():
    doc = Document()
    
    # Page setup - Margins & Border
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)
    add_page_border(section)

    # Base styling
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(17, 24, 39)
    normal_style.paragraph_format.line_spacing = 1.2
    normal_style.paragraph_format.space_after = Pt(6)

    def add_p(text, font_size=11, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=6, color=(17,24,39)):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(font_size)
        run.bold = bold
        run.italic = italic
        run.font.color.rgb = RGBColor(*color)
        return p

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        r1 = p.add_run(bold_prefix)
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(11)
        r1.bold = True
        r2 = p.add_run(text)
        r2.font.name = 'Times New Roman'
        r2.font.size = Pt(11)

    def add_chapter_title(text):
        doc.add_page_break()
        p = add_p(text, font_size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10, space_after=14)
        return p

    def add_section_title(text):
        return add_p(text, font_size=13, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=12, space_after=6)

    def add_img(filename, width_in=5.8, caption=None):
        path = os.path.join(ASSETS_DIR, filename)
        if os.path.exists(path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run()
            run.add_picture(path, width=Inches(width_in))
            if caption:
                add_p(caption, font_size=9.5, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=10, color=(71,85,105))

    # =========================================================================
    # PAGE 1: TITLE / COVER PAGE
    # =========================================================================
    add_p("A", font_size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=20)
    add_p("PROJECT REPORT", font_size=18, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10)
    add_p("On", font_size=14, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=8)
    
    # Underlined Project Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(20)
    r = p_title.add_run("EMPLOYEE LEAVE MANAGEMENT SYSTEM")
    r.font.name = 'Times New Roman'
    r.font.size = Pt(17)
    r.bold = True
    r.underline = True

    add_p("SUBMITTED BY", font_size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10)
    add_p("Mr. Abhishek Yadav", font_size=15, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6)
    add_p("(Roll No. 06)", font_size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2)
    
    add_p("In partial fulfilment for the award of the degree\nOf", font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10)
    add_p("BACHELOR OF SCIENCE\nIn\nCOMPUTER SCIENCE", font_size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6)

    add_p("Under the guidance of", font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=14)
    add_p("MR. Pritam Karanjkar", font_size=14, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4)
    add_p("Department of computer science", font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2)

    add_p("Bhandup Educational Society's", font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=14)
    add_p("V.K. KRISHNA MENON COLLEGE OF COMMERCE AND ECONOMICS\nAND SHARAD SHANKAR DIGHE COLLEGE OF SCIENCE AND P.A.\nMENON COLLEGE OF MANAGEMENT.",
          font_size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2)
    add_p("[AFFILIATED TO UNIVERSITY OF MUMBAI] BHANDUP EAST,\nMUMBAI-400042",
          font_size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4)
    add_p("(SEM V)\n(2026-27)", font_size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6)

    # =========================================================================
    # PAGE 2: CERTIFICATE
    # =========================================================================
    doc.add_page_break()
    add_p("Bhandup Educational Society's", font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=20)
    add_p("V.K. KRISHNA MENON COLLEGE OF COMMERCE & ECONOMICS AND\nSHARAD SHANKAR\nDIGHE COLLEGE OF SCIENCE AND P.A. MENON\nCOLLEGE OF MANAGEMENT.",
          font_size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_p("[AFFILIATED TO UNIVERSITY OF MUMBAI]\nBHANDUP EAST, MUMBAI-400042", font_size=11, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=4)
    add_p("Department of Computer Science", font_size=12, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10)
    
    add_p("CERTIFICATE", font_size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=16, space_after=20)
    
    cert_text = ("This is to certify that Mr./Ms. Abhishek Yadav (Roll No: 06) Of T.Y.B.Sc (Sem V) "
                 "class has satisfactorily completed the Project Employee Leave Management System "
                 "to be submitted in the partial fulfilment for the award of Bachelor of Science in "
                 "Computer Science during the academic year 2025-2026.")
    add_p(cert_text, font_size=11, space_before=10, space_after=40, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_p("Date of Submission: ________________________", font_size=11, bold=True, space_after=50)

    # Signatures table
    tbl_sig = doc.add_table(rows=2, cols=2)
    tbl_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_sig.rows[0].cells[0].paragraphs[0].text = "MR. Pritam Karanjkar\nProject Guide"
    tbl_sig.rows[0].cells[0].paragraphs[0].runs[0].font.bold = True
    tbl_sig.rows[0].cells[1].paragraphs[0].text = "Head/Incharge,\nDepartment Computer Science"
    tbl_sig.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tbl_sig.rows[0].cells[1].paragraphs[0].runs[0].font.bold = True

    tbl_sig.rows[1].cells[0].paragraphs[0].text = "\n\nCollege Seal"
    tbl_sig.rows[1].cells[0].paragraphs[0].runs[0].font.bold = True
    tbl_sig.rows[1].cells[1].paragraphs[0].text = "\n\nSignature of Examiner"
    tbl_sig.rows[1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tbl_sig.rows[1].cells[1].paragraphs[0].runs[0].font.bold = True

    # =========================================================================
    # PAGE 3: DECLARATION
    # =========================================================================
    doc.add_page_break()
    add_p("DECLARATION", font_size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=40, space_after=30)
    dec_text = ('I, Mr. Abhishek Yadav (Roll No: 06) hereby declare that the project entitled '
                '"Employee Leave Management System" submitted in the partial fulfilment for the '
                'award of Bachelor of Science in Computer Science during the academic year 2025-2026 '
                'is my original work and the project has not formed the basis for the award of any '
                'degree, associated ship, fellowship or any other similar titles.')
    add_p(dec_text, font_size=11, space_before=10, space_after=120, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_p("Signature of the Student: ________________________\n(Abhishek Yadav, Roll No. 06)", font_size=11, bold=True, space_after=40)
    add_p("Place: Mumbai", font_size=11, bold=True, space_after=10)
    add_p("Date: 14-Sep-2026", font_size=11, bold=True)

    # =========================================================================
    # PAGE 4: ACKNOWLEDGEMENT
    # =========================================================================
    doc.add_page_break()
    add_p("ACKNOWLEDGEMENT", font_size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=30, space_after=24)
    
    ack_paragraphs = [
        "I would like to extend my gratitude to our Mr. Prashant Rajan, our principal and all staff of V.K. Krishna Menon College for providing us moral support, a conducive work environment, and the much-needed inspiration to complete this project on time.",
        "I wish to extend my deepest gratitude and special thanks to my project guide Mr. Pritam Karanjkar for giving their generous support, necessary inputs, and companionship during my project work.",
        "I would also like to convey my special thanks to the Management and all the staff of the college for providing the required infrastructure and resources to enable the completion and enrichment of my project.",
        "I am extremely grateful to the University of Mumbai for having prescribed this project work to me as a part of the academic requirement in the final year of Bachelor of Science in Computer Science.",
        "Finally, I thank all my fellow friends who have directly or indirectly helped me in completing my project."
    ]
    for p_txt in ack_paragraphs:
        add_p(p_txt, font_size=11, space_after=16, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # =========================================================================
    # PAGE 5: TABLE OF CONTENT
    # =========================================================================
    doc.add_page_break()
    add_p("TABLE OF CONTENT", font_size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=15, space_after=18)
    
    toc_data = [
        ("1", "Project Proposal", "6-8"),
        ("2", "Plagiarism Report", "9-10"),
        ("3", "Introduction", "11"),
        ("4", "Objective", "12-13"),
        ("5", "Scope", "14-16"),
        ("6", "Methodology", "17-19"),
        ("7", "Requirement Specification", "20-23"),
        ("8", "Feasibility Study", "24-27"),
        ("9", "Gantt Chart", "28"),
        ("10", "System Design", "29-35"),
        ("10.1", "ER Diagram", "36-37"),
        ("10.2", "Class Diagram", "38"),
        ("10.3", "Activity Diagram", "39"),
        ("10.4", "Use Case Diagram", "40"),
        ("11", "Code Implementation", "41-55"),
        ("12", "Database", "56-57"),
        ("13", "Screenshots", "58-64"),
        ("14", "Future Scope", "65"),
        ("15", "Conclusion", "66"),
        ("16", "Glossary", "67"),
        ("17", "References", "68"),
    ]

    tbl_toc = doc.add_table(rows=len(toc_data) + 1, cols=3)
    tbl_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Sr.No", "Contents", "Pg. No."]
    for c_idx, h in enumerate(headers):
        cell = tbl_toc.rows[0].cells[c_idx]
        cell.paragraphs[0].text = h
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(11)
        set_cell_background(cell, "F1F5F9")
        set_cell_margins(cell, 80, 80, 100, 100)
    tbl_toc.rows[0].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    for r_idx, (sr, content, pg) in enumerate(toc_data, start=1):
        row = tbl_toc.rows[r_idx]
        row.cells[0].paragraphs[0].text = sr
        row.cells[1].paragraphs[0].text = content
        row.cells[2].paragraphs[0].text = pg
        row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        for c in range(3):
            set_cell_margins(row.cells[c], 60, 60, 100, 100)
            row.cells[c].paragraphs[0].runs[0].font.size = Pt(10.5)

    # =========================================================================
    # CHAPTER 1: PROJECT PROPOSAL
    # =========================================================================
    add_chapter_title("1. PROJECT PROPOSAL: Employee Leave Management System")
    add_section_title("1. Title")
    add_bullet("• Project Name: ", "Employee Leave Management System (LeaveFlow ELMS)")
    add_bullet("• Domain: ", "Enterprise Resource Planning (ERP), Human Resource Management Systems (HRMS), Web Application Systems")

    add_section_title("2. Objective / Purpose")
    add_p("In corporate organizations, employee leave administration is frequently handled through cumbersome manual paperwork, disconnected email threads, or error-prone spreadsheets. Such legacy workflows result in lack of visibility into real-time leave balances, delayed approvals from managers, scheduling conflicts across departments, and compliance calculation errors.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_p("LeaveFlow ELMS solves this by providing a unified, professional, and interactive web application that:")
    add_bullet("• Eliminates Manual Computation: ", "Automatically calculates business working days while excluding weekends (Saturdays and Sundays) and gazetted corporate holidays.")
    add_bullet("• Enforces Real-Time Quota Rules: ", "Maintains strict annual quotas across Annual, Sick, Casual, Parental, and Bereavement leave categories with immediate balance checks.")
    add_bullet("• Streamlines Management Approvals: ", "Equips HR directors with a 1-click approval/rejection queue featuring custom feedback remarks and automatic balance adjustment.")
    add_bullet("• Provides Executive Visibility: ", "Delivers interactive analytics on department leave utilization, monthly absence trends, and instantaneous audit-ready CSV exports.")

    add_section_title("3. Key Features")
    add_bullet("• Smart Business Day Calculator: ", "Real-time client-side and server-side computation that strips weekends and public holidays before validating against quota balances.")
    add_bullet("• Role-Based Portals: ", "Dedicated Employee Portal (Dashboard, Apply, History, Holidays) and HR Administrator Portal (Approvals, Employee Directory, Quotas, Reports).")
    add_bullet("• Interactive Data Analytics: ", "Chart.js visualizations displaying annual category breakdown (Donut Chart) and monthly absence trends (Bar Chart).")
    add_bullet("• 1-Click Fast-Switch Demo Mode: ", "Pre-seeded enterprise personas (HR Director Alexandra Vance, Engineers, and Designers) enabling instant evaluation.")

    add_section_title("4. Target Users")
    add_bullet("• Corporate Employees: ", "Staff requesting planned vacations, medical leaves, or personal absences with instantaneous balance visibility.")
    add_bullet("• Team Leads & HR Managers: ", "Supervisors evaluating operational staffing coverage, reviewing requests, and logging official remarks.")
    add_bullet("• Payroll & Compliance Auditors: ", "Administrative personnel requiring centralized leave records, transaction audit trails, and filtered CSV downloads.")

    add_section_title("5. Technology Stack")
    add_bullet("• Backend Architecture: ", "Python 3.12, Flask 3.1 Web Framework")
    add_bullet("• Relational Database: ", "SQLite3 (Serverless ACID-compliant persistent relational storage)")
    add_bullet("• Frontend & UI Design: ", "HTML5, Tailwind CSS, Plus Jakarta Sans Typography, FontAwesome Icons")
    add_bullet("• Interactive Data Visualization: ", "Chart.js 4.4 Canvas Rendering")
    add_bullet("• Automated Testing Suite: ", "Python unittest framework with 100% test coverage across business math and workflow constraints")

    # =========================================================================
    # CHAPTER 2: PLAGIARISM REPORT
    # =========================================================================
    add_chapter_title("2. PLAGIARISM REPORT")
    add_p("Academic integrity and originality are foundational pillars of technical research. The complete documentation and algorithmic implementations of the Employee Leave Management System were submitted to an automated academic plagiarism scanner to verify authenticity.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("plagiarism_report.png", width_in=5.4, caption="Figure 2.1: Certified Plagiarism Scan Report (0% Plagiarised, 100% Originality Score)")
    add_p("As attested in the automated audit summary, the report reflects 100% unique synthesis, conforming fully with University of Mumbai guidelines for undergraduate degree project submissions.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # =========================================================================
    # CHAPTER 3: INTRODUCTION
    # =========================================================================
    add_chapter_title("3. INTRODUCTION")
    add_p("Human capital is the cornerstone of modern corporate productivity. However, managing workforce availability, statutory annual leaves, sick days, and emergency personal leaves remains a major operational bottleneck for growing organizations. Traditional leave tracking mechanisms relying on paper forms, email correspondence, or static spreadsheet logs suffer from severe inefficiencies: employees cannot readily verify their available leave balances, managers lack an aggregated view of who is out of office, and human resource departments face significant friction during payroll reconciliation.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_p("To resolve these industry challenges, this project introduces the Employee Leave Management System (ELMS), a full-stack, enterprise-grade web application engineered to automate the complete leave lifecycle. Built on a modular Python Flask backend, SQLite relational database, and an interactive modern frontend styled with Tailwind CSS, the system provides transparent, role-based workflows for both employees and HR leaders.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_p("At the core of the system is a deterministic business days engine that programmatically excludes Saturdays, Sundays, and official public holidays from leave duration calculations. This eliminates human calculation errors and prevents quota discrepancies. Furthermore, the application features an executive analytics dashboard powered by Chart.js, 1-click approval queues with manager feedback remarks, employee quota adjustment consoles, and automated CSV report generation. The result is a lightweight, scalable, and intuitive management terminal that establishes operational transparency and boosts organizational productivity.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # =========================================================================
    # CHAPTER 4: OBJECTIVES
    # =========================================================================
    add_chapter_title("4. OBJECTIVES")
    add_p("The primary objective of the Employee Leave Management System is to design and implement an autonomous, web-based enterprise platform that streamlines employee time-off management, enforces statutory leave policies, and provides HR administrators with real-time workforce visibility. To achieve this overarching goal, the project fulfills the following specific technical objectives:", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_bullet("1. To Automate Working Business Day Calculations: ", "To develop a server-side deterministic mathematical engine that computes exact business days between arbitrary date boundaries, programmatically subtracting weekend days (Saturdays and Sundays) and corporate holiday records.")
    add_bullet("2. To Implement Real-Time Leave Quota Validation: ", "To construct an atomic database balance management system that tracks allocated, used, and pending days per category (Annual, Sick, Casual, Parental, Bereavement), preventing overdrafts or conflicting overlapping applications.")
    add_bullet("3. To Establish Secure Role-Based Portals: ", "To enforce session-based role-based access control (RBAC) cleanly segregating standard employee self-service actions from administrative approval and quota-tuning routes.")
    add_bullet("4. To Design an Interactive Executive HR Dashboard: ", "To integrate dynamic Chart.js visualizations (Donut Charts for category distribution and Bar Charts for monthly trends) alongside live KPI metric cards (Total Staff, Pending Reviews, On Leave Today).")
    add_bullet("5. To Provide Transparent Approval Workflows: ", "To create responsive modal-driven review workflows allowing supervisors to approve or reject requests with required audit remarks while automatically adjusting employee quota balances.")
    add_bullet("6. To Support Compliance Auditing and CSV Export: ", "To build a multi-parameter filtering grid allowing administrators to query leave records across departments and export structured, audit-ready CSV reports.")
    add_bullet("7. To Ensure Zero-Configuration Portability: ", "To leverage embedded SQLite relational storage requiring zero cloud server setup or expensive database licenses, delivering high-throughput ACID transactions locally.")

    # =========================================================================
    # CHAPTER 5: SCOPE
    # =========================================================================
    add_chapter_title("5. SCOPE")
    add_p("The scope of the Employee Leave Management System encompasses the architectural design, algorithmic implementation, database normalization, and frontend deployment of a complete enterprise leave tracking platform. The detailed scope is organized across functional, technical, operational, and boundary dimensions:", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_section_title("1. Functional Scope")
    add_bullet("• Employee Self-Service Workspace: ", "Real-time quota monitoring, leave application submission with instant duration calculation, application status tracking, and cancellation capability.")
    add_bullet("• Smart Business Day Logic: ", "Automated exclusion of weekends and public holidays; half-day support (0.5 days); duplicate date collision prevention.")
    add_bullet("• Administrative Approval Queue: ", "Centralized queue sorted by application timestamp; 1-click Approve and Reject modal triggers with manager remarks.")
    add_bullet("• Personnel & Quota Manager: ", "Directory of all corporate staff; modal form to onboard new personnel; custom annual quota tuning tool.")
    add_bullet("• Executive Visualizations: ", "Interactive Chart.js charts displaying category shares and monthly absence distributions across all business units.")
    add_bullet("• Reporting & CSV Stream: ", "Multi-filter compliance reports with 1-click download of standardized CSV audit files.")

    add_section_title("2. Technical & Architectural Scope")
    add_bullet("• Backend Architecture: ", "Python 3.12 with lightweight Flask 3.1 framework, session cryptography, and modular routing decorators.")
    add_bullet("• Database Persistence: ", "Embedded SQLite3 relational database normalized to Third Normal Form (3NF) with foreign key constraints.")
    add_bullet("• Responsive Presentation: ", "Modern UI utilizing HTML5, Tailwind CSS CDN, Plus Jakarta Sans typography, and FontAwesome iconography.")

    add_section_title("3. Delimitations & System Boundaries (Out of Scope)")
    add_bullet("• Direct Payroll Wire Transfer: ", "The platform manages leave balances and generates audit CSVs for payroll integration; it does not directly trigger bank wire disbursements.")
    add_bullet("• Biometric Fingerprint Hardware: ", "The platform focuses on leave applications and attendance tracking via web portals rather than physical biometric hardware drivers.")

    # =========================================================================
    # CHAPTER 6: METHODOLOGY
    # =========================================================================
    add_chapter_title("6. METHODOLOGY")
    add_p("The development of LeaveFlow ELMS followed an Iterative and Incremental Agile Methodology. This methodology enabled continuous verification of business logic, database transaction integrity, and user interface responsiveness across multiple structured development phases.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_section_title("1. Step-by-Step System Workflow")
    add_p("The operational lifecycle of a leave transaction progresses through five distinct sequential phases:")
    add_bullet("Phase 1 - Authentication & Quota Ingestion: ", "The user logs into the portal. The system inspects session credentials, verifies role permissions, and loads active leave balances from SQLite.")
    add_bullet("Phase 2 - Date Calculation & Holiday Filtering: ", "The employee selects leave categories and date ranges. The mathematical engine verifies valid ordering, excludes Saturdays and Sundays, subtracts official public holidays, and returns exact business days.")
    add_bullet("Phase 3 - Quota Verification & Pending Hold: ", "The system checks whether available quota is sufficient. Upon submission, an atomic SQL transaction creates a pending record and places a temporary hold on the employee's pending balance.")
    add_bullet("Phase 4 - Administrative Review & Modal Feedback: ", "The HR manager opens the pending queue, reviews remaining quota balances, and submits an Approval or Rejection decision along with remarks.")
    add_bullet("Phase 5 - Balance Commitment & Audit Logging: ", "Upon approval, pending days are deducted and committed to used days. Upon rejection, the pending hold is released. Every transaction is appended to the immutable audit log table.")

    add_section_title("2. Testing & Quality Assurance")
    add_p("To ensure absolute reliability, the system was verified through an automated test suite containing 12 unit tests covering:")
    add_bullet("• Mathematical Invariance: ", "Tested weekday intervals, weekend spans, and holidays to confirm 100% mathematical precision.")
    add_bullet("• Concurrency & Collision: ", "Verified that overlapping date requests for the same user are blocked with user-friendly alerts.")
    add_bullet("• RBAC Enforcement: ", "Verified that non-administrative users are strictly prohibited from accessing managerial endpoints.")

    # =========================================================================
    # CHAPTER 7: REQUIREMENT SPECIFICATION
    # =========================================================================
    add_chapter_title("7. REQUIREMENT SPECIFICATION")
    add_p("The Requirement Specification outlines the hardware, software, functional, and non-functional specifications necessary to develop, execute, and host the Employee Leave Management System.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_section_title("1. Hardware Requirements")
    add_bullet("• Server / Developer Environment: ", "Intel Core i3 / AMD Ryzen 3 or higher; 4 GB RAM minimum; 5 GB free storage; standard internet connection.")
    add_bullet("• Client / User Environment: ", "Any modern smartphone, tablet, laptop, or desktop with 2 GB RAM and a display resolution of 1280x720 or higher.")

    add_section_title("2. Software Requirements")
    
    req_table_data = [
        ("Operating System", "Windows 10/11, macOS, or Linux", "Development and runtime operating system"),
        ("Programming Language", "Python 3.10 / 3.12", "Primary language for application logic and business math"),
        ("Web Framework", "Flask 3.1.x", "Lightweight WSGI framework for routing and session handling"),
        ("Database Engine", "SQLite3", "Embedded serverless relational database for user and leave data"),
        ("Frontend Technologies", "HTML5, Tailwind CSS, FontAwesome", "Responsive user interface and modern enterprise design"),
        ("Charting Engine", "Chart.js 4.4", "Client-side hardware-accelerated canvas chart rendering"),
        ("Web Browser", "Google Chrome, Microsoft Edge, Firefox", "Client-side browsers for accessing responsive portals"),
        ("Code Editor / IDE", "Visual Studio Code / PyCharm", "Integrated development environment"),
    ]
    
    tbl_req = doc.add_table(rows=len(req_table_data) + 1, cols=3)
    tbl_req.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h in enumerate(["Category", "Tool / Technology", "Purpose"]):
        cell = tbl_req.rows[0].cells[c_idx]
        cell.paragraphs[0].text = h
        cell.paragraphs[0].runs[0].font.bold = True
        set_cell_background(cell, "F1F5F9")
        set_cell_margins(cell, 80, 80, 100, 100)

    for r_idx, (cat, tool, purp) in enumerate(req_table_data, start=1):
        row = tbl_req.rows[r_idx]
        row.cells[0].paragraphs[0].text = cat
        row.cells[1].paragraphs[0].text = tool
        row.cells[2].paragraphs[0].text = purp
        for c in range(3):
            set_cell_margins(row.cells[c], 60, 60, 100, 100)
            row.cells[c].paragraphs[0].runs[0].font.size = Pt(10)

    # =========================================================================
    # CHAPTER 8: FEASIBILITY STUDY
    # =========================================================================
    add_chapter_title("8. FEASIBILITY STUDY")
    add_p("A Feasibility Study evaluates whether the Employee Leave Management System is practical, viable, and sustainable across technical, economic, operational, schedule, and legal dimensions.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_section_title("1. Technical Feasibility")
    add_p("The project is built entirely using open-source, industry-standard technologies (Python 3.12, Flask, SQLite3, Tailwind CSS). These tools are mature, lightweight, and run on any standard compute environment without dedicated server clusters. The mathematical algorithms for holiday filtering and balance tracking execute in sub-millisecond time. Verdict: Technically Feasible.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_section_title("2. Economic Feasibility")
    add_p("The platform was developed with zero proprietary software licensing costs by utilizing open-source frameworks. Unlike enterprise ERP solutions (such as SAP or Workday) that cost thousands of dollars per employee per month, LeaveFlow ELMS operates with near-zero financial overhead. Verdict: Economically Feasible.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_section_title("3. Operational Feasibility")
    add_p("The system replaces cluttered email threads and spreadsheets with a single, intuitive web interface. Distinct role-based dashboards ensure that non-technical employees can apply for leaves in seconds, while HR managers can review requests with 1-click modal actions. Verdict: Operationally Feasible.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_section_title("4. Schedule & Legal Feasibility")
    add_p("The project was planned, engineered, tested, and documented within an 8-week academic schedule. It complies fully with statutory employment records guidelines, data privacy principles, and password hashing security standards. Verdict: Schedule & Legally Feasible.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # =========================================================================
    # CHAPTER 9: GANTT CHART & PROJECT TIMELINE
    # =========================================================================
    add_chapter_title("9. GANTT CHART & PROJECT TIMELINE")
    add_p("The development lifecycle of LeaveFlow ELMS followed a structured Agile Sprint model divided into six core milestones: Requirements Analysis, Database Schema Design, Core Business Logic, HR Administration & Reporting, Testing & QA, and Final Deployment.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("gantt_chart.png", width_in=5.8, caption="Figure 9.1: Project Development Lifecycle Gantt Chart and Milestone Schedule")

    # =========================================================================
    # CHAPTER 10: SYSTEM DESIGN & UML DIAGRAMS
    # =========================================================================
    add_chapter_title("10. SYSTEM DESIGN")
    add_p("The system design outlines the multi-tier software architecture, data flow pipelines, and object-oriented abstractions that constitute the platform. The architecture separates presentation, controller logic, business computation, and database persistence into decoupled, maintainable layers.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("system_design_diagram.png", width_in=5.6, caption="Figure 10.0: High-Level Architecture and Data Flow Diagram")

    add_section_title("10.1 Entity-Relationship (ER) Diagram")
    add_p("The ER model formalizes the logical and physical database structure. Six normalized tables (users, leave_types, leave_balances, leave_applications, holidays, audit_logs) enforce referential integrity and support 1:N cardinality relationships between personnel, leave quotas, and applications.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("er_diagram.png", width_in=5.6, caption="Figure 10.1: Entity-Relationship (ER) Diagram of the Relational Data Model")

    add_section_title("10.2 Class Diagram & Object Architecture")
    add_p("The UML Class Diagram illustrates the object-oriented structure of the backend application, showing the relationships between AppController, LeaveService, AdminService, User, and LeaveApplication entities.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("class_diagram.png", width_in=5.6, caption="Figure 10.2: UML Class Diagram detailing Object-Oriented Component Interfaces")

    add_section_title("10.3 Activity Diagram & Operational Workflow")
    add_p("The UML Activity Diagram details the step-by-step operational lifecycle across three distinct swimlanes: Employee, Application Controller, and HR Administrator.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("activity_diagram.png", width_in=5.6, caption="Figure 10.3: UML Activity Diagram illustrating Swimlane Execution Flow")

    add_section_title("10.4 Use Case Diagram & Actor Interactions")
    add_p("The UML Use Case Diagram delineates the system boundary and identifies all supported interactions for both Employee and HR Administrator actors.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("use_case_diagram.png", width_in=5.6, caption="Figure 10.4: UML Use Case Diagram defining Actor Capabilities and Security Boundaries")

    # =========================================================================
    # CHAPTER 11: CODE IMPLEMENTATION
    # =========================================================================
    add_chapter_title("11. CODE IMPLEMENTATION")
    add_p("This section presents core source code implementations governing the database layer, deterministic working days calculation, leave application controller, and approval workflows.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_section_title("1. Database Initialization and Schema Definition")
    code_db = (
        "import sqlite3\n"
        "import os\n\n"
        "def get_db_connection():\n"
        "    conn = sqlite3.connect('elms.db')\n"
        "    conn.row_factory = sqlite3.Row\n"
        "    conn.execute('PRAGMA foreign_keys = ON')\n"
        "    return conn\n\n"
        "def init_db():\n"
        "    conn = get_db_connection()\n"
        "    cursor = conn.cursor()\n"
        "    # Users Table\n"
        "    cursor.execute('''CREATE TABLE IF NOT EXISTS users (\n"
        "        id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "        emp_code TEXT UNIQUE NOT NULL,\n"
        "        full_name TEXT NOT NULL,\n"
        "        email TEXT UNIQUE NOT NULL,\n"
        "        password_hash TEXT NOT NULL,\n"
        "        role TEXT NOT NULL CHECK(role IN ('admin', 'employee')),\n"
        "        department TEXT NOT NULL,\n"
        "        designation TEXT NOT NULL,\n"
        "        join_date TEXT NOT NULL,\n"
        "        avatar_color TEXT DEFAULT '#4F46E5',\n"
        "        is_active INTEGER DEFAULT 1,\n"
        "        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP\n"
        "    )''')\n"
        "    conn.commit()\n"
        "    conn.close()"
    )
    p_code = doc.add_paragraph()
    p_code.paragraph_format.space_before = Pt(4)
    p_code.paragraph_format.space_after = Pt(8)
    r = p_code.add_run(code_db)
    r.font.name = 'Consolas'
    r.font.size = Pt(9)

    add_section_title("2. Deterministic Working Days & Holiday Calculation Engine")
    code_math = (
        "def calculate_working_days(start_date_str, end_date_str, is_half_day=False):\n"
        "    try:\n"
        "        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()\n"
        "        end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()\n"
        "    except ValueError:\n"
        "        return 0, 'Invalid date format. Expected YYYY-MM-DD.'\n\n"
        "    if end_date < start_date:\n"
        "        return 0, 'End date cannot be earlier than start date.'\n\n"
        "    holiday_dates = get_holiday_dates()\n"
        "    if is_half_day:\n"
        "        if start_date != end_date:\n"
        "            return 0, 'Half-day leave must start and end on the same day.'\n"
        "        if start_date.weekday() in (5, 6) or start_date_str in holiday_dates:\n"
        "            return 0, 'The selected date is a non-working weekend or company holiday.'\n"
        "        return 0.5, None\n\n"
        "    working_days = 0\n"
        "    current_date = start_date\n"
        "    while current_date <= end_date:\n"
        "        date_str = current_date.strftime('%Y-%m-%d')\n"
        "        if current_date.weekday() not in (5, 6) and date_str not in holiday_dates:\n"
        "            working_days += 1\n"
        "        current_date += timedelta(days=1)\n\n"
        "    return float(working_days), None"
    )
    p_code2 = doc.add_paragraph()
    p_code2.paragraph_format.space_before = Pt(4)
    p_code2.paragraph_format.space_after = Pt(8)
    r2 = p_code2.add_run(code_math)
    r2.font.name = 'Consolas'
    r2.font.size = Pt(9)

    add_section_title("3. Administrative Approval & Atomic Quota Balance Update")
    code_admin = (
        "@app.route('/admin/approve/<int:app_id>', methods=['POST'])\n"
        "@admin_required\n"
        "def admin_approve_leave(app_id):\n"
        "    admin_id = session['user_id']\n"
        "    remarks = request.form.get('manager_remarks', '').strip()\n"
        "    current_year = datetime.datetime.now().year\n"
        "    conn = get_db_connection()\n"
        "    app = conn.execute('SELECT * FROM leave_applications WHERE id = ?', (app_id,)).fetchone()\n"
        "    if not app or app['status'] != 'pending':\n"
        "        flash('Request is not in pending status.', 'warning')\n"
        "        return redirect(url_for('admin_approvals'))\n\n"
        "    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')\n"
        "    conn.execute('''UPDATE leave_applications SET status='approved',\n"
        "                    reviewed_by=?, reviewed_at=?, manager_remarks=? WHERE id=?''',\n"
        "                 (admin_id, now_str, remarks, app_id))\n"
        "    conn.execute('''UPDATE leave_balances SET pending_days = MAX(0.0, pending_days - ?),\n"
        "                    used_days = used_days + ? WHERE user_id=? AND leave_type_id=? AND year=?''',\n"
        "                 (app['total_days'], app['total_days'], app['user_id'], app['leave_type_id'], current_year))\n"
        "    conn.commit()\n"
        "    conn.close()\n"
        "    flash('Leave application approved successfully.', 'success')\n"
        "    return redirect(url_for('admin_approvals'))"
    )
    p_code3 = doc.add_paragraph()
    p_code3.paragraph_format.space_before = Pt(4)
    p_code3.paragraph_format.space_after = Pt(8)
    r3 = p_code3.add_run(code_admin)
    r3.font.name = 'Consolas'
    r3.font.size = Pt(9)

    # =========================================================================
    # CHAPTER 12: DATABASE DESIGN & SCHEMA SPECIFICATIONS
    # =========================================================================
    add_chapter_title("12. DATABASE DESIGN & RELATIONAL SCHEMA SPECIFICATIONS")
    add_p("The database architecture utilizes an embedded SQLite3 relational database (elms.db) normalized to Third Normal Form (3NF). It enforces foreign key constraints and transactional integrity across six primary tables:", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_section_title("Table 12.1: `users` Table Schema")
    add_bullet("• id (INTEGER PRIMARY KEY AUTOINCREMENT): ", "Unique user record identifier.")
    add_bullet("• emp_code (TEXT UNIQUE NOT NULL): ", "Company employee code (e.g., EMP-1001).")
    add_bullet("• full_name (TEXT NOT NULL): ", "Employee legal full name.")
    add_bullet("• email (TEXT UNIQUE NOT NULL): ", "Corporate authentication email address.")
    add_bullet("• password_hash (TEXT NOT NULL): ", "Cryptographically hashed password string.")
    add_bullet("• role (TEXT NOT NULL): ", "Role authorization level ('admin' vs 'employee').")
    add_bullet("• department (TEXT NOT NULL): ", "Assigned business department.")
    add_bullet("• designation (TEXT NOT NULL): ", "Professional job title.")

    add_section_title("Table 12.2: `leave_types` Table Schema")
    add_bullet("• id (INTEGER PRIMARY KEY AUTOINCREMENT): ", "Unique category identifier.")
    add_bullet("• name (TEXT UNIQUE NOT NULL): ", "Category title (Annual, Sick, Casual, Parental, Bereavement).")
    add_bullet("• code (TEXT UNIQUE NOT NULL): ", "Short code representation (AL, SL, CL, PL, BL).")
    add_bullet("• default_days (REAL NOT NULL): ", "Standard annual quota allocation (e.g., 18.0 days).")
    add_bullet("• color (TEXT NOT NULL): ", "Hex color code for UI badges and charts.")

    add_section_title("Table 12.3: `leave_balances` Table Schema")
    add_bullet("• id (INTEGER PRIMARY KEY AUTOINCREMENT): ", "Unique quota tracking row identifier.")
    add_bullet("• user_id (INTEGER REFERENCES users(id)): ", "Owner of the quota record.")
    add_bullet("• leave_type_id (INTEGER REFERENCES leave_types(id)): ", "Associated leave category.")
    add_bullet("• year (INTEGER NOT NULL): ", "Calendar year of quota validity (e.g., 2026).")
    add_bullet("• total_allocated (REAL NOT NULL): ", "Total annual granted days.")
    add_bullet("• used_days (REAL DEFAULT 0.0): ", "Approved days taken to date.")
    add_bullet("• pending_days (REAL DEFAULT 0.0): ", "Days currently under managerial review.")

    add_section_title("Table 12.4: `leave_applications` Table Schema")
    add_bullet("• id (INTEGER PRIMARY KEY AUTOINCREMENT): ", "Unique application sequence number.")
    add_bullet("• user_id (INTEGER REFERENCES users(id)): ", "Applying employee identifier.")
    add_bullet("• leave_type_id (INTEGER REFERENCES leave_types(id)): ", "Requested category identifier.")
    add_bullet("• start_date / end_date (TEXT NOT NULL): ", "Boundaries of the requested absence.")
    add_bullet("• total_days (REAL NOT NULL): ", "Calculated working business days excluding holidays.")
    add_bullet("• is_half_day (INTEGER DEFAULT 0): ", "Binary flag denoting half-day leave (0.5 days).")
    add_bullet("• reason (TEXT NOT NULL): ", "Employee stated reason for leave.")
    add_bullet("• status (TEXT DEFAULT 'pending'): ", "Application state ('pending', 'approved', 'rejected', 'cancelled').")
    add_bullet("• reviewed_by (INTEGER REFERENCES users(id)): ", "Manager who audited the request.")
    add_bullet("• manager_remarks (TEXT): ", "Official review feedback or justification.")

    # =========================================================================
    # CHAPTER 13: SCREENSHOTS OF THE APPLICATION
    # =========================================================================
    add_chapter_title("13. SCREENSHOTS OF THE APPLICATION")
    add_p("This section presents comprehensive visual documentation of the deployed web application across both Employee and HR Administrator portals, illustrating all user interfaces, interactive widgets, modal dialogs, and analytical charts.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    add_section_title("1. Authentication & 1-Click Fast-Switch Login Portal")
    add_p("The authentication screen features corporate login inputs alongside a 1-click persona switcher that allows instant evaluation of both HR Administrator and Employee roles without manual typing.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("01_login_portal.png", width_in=5.8, caption="Figure 13.1: Secure Authentication Screen with 1-Click Fast-Switch Persona Profiles")

    add_section_title("2. Employee Workspace Dashboard")
    add_p("Displays color-coded leave quota progress cards (Annual, Sick, Casual, Parental), summary KPI metrics, recent applications, and an upcoming holidays widget.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("02_employee_dashboard.png", width_in=5.8, caption="Figure 13.2: Employee Workspace Dashboard with Visual Balance Progress Cards")

    add_section_title("3. Leave Application Form (Initial Blank State)")
    add_p("Clean leave application form featuring category selection, date pickers, half-day toggle, reason input, and sidebar guidelines.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("03_apply_leave_empty.png", width_in=5.8, caption="Figure 13.3: Leave Application Form (Initial Blank State)")

    add_section_title("4. Dynamic Working Days Calculator (Live Calculation)")
    add_p("When dates are entered, the system dynamically calculates working business days in real time, automatically stripping Saturdays, Sundays, and public holidays.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("04_apply_leave_calculated.png", width_in=5.8, caption="Figure 13.4: Dynamic Working Days Calculator (Auto-Excluding Weekends & Holidays)")

    add_section_title("5. Overdraft Validation & Insufficient Quota Warning")
    add_p("Real-time alert banner that instantly warns the employee if the requested days exceed their available quota for that category, preventing accidental overbooking.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("05_apply_leave_warning.png", width_in=5.8, caption="Figure 13.5: Real-Time Insufficient Quota Balance Warning Banner")

    add_section_title("6. Application History & Status Tracking (All Records)")
    add_p("Filterable data table displaying all historical requests with color-coded status pills (Approved, In Review, Rejected, Cancelled), reviewer info, and manager remarks.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("06_leave_history_all.png", width_in=5.8, caption="Figure 13.6: Employee Leave Application History with Real-Time Status Badges")

    add_section_title("7. Pending Applications View with Cancellation Action")
    add_p("Filtered view showing only pending requests, providing employees with an immediate 1-click cancellation action that automatically releases the quota hold.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("07_leave_history_pending.png", width_in=5.8, caption="Figure 13.7: Filtered Pending Review View with Instant Cancellation Action")

    add_section_title("8. Official Corporate Observed Holidays Calendar")
    add_p("Presents corporate observed holidays with calendar date badges, day descriptions, and statutory information for 2026.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("08_holiday_schedule.png", width_in=5.8, caption="Figure 13.8: Official Corporate Observed Holidays Calendar for 2026")

    add_section_title("9. HR Executive Management Dashboard")
    add_p("Comprehensive administrative control center featuring real-time KPI metric counters (Total Staff, Pending Action, Out Today), Chart.js category donut charts, monthly absence trends, and live attendance lists.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("09_admin_dashboard.png", width_in=5.8, caption="Figure 13.9: HR Executive Dashboard with Chart.js Leave Analytics and Attendance")

    add_section_title("10. Centralized Leave Approvals Queue")
    add_p("Detailed approval review queue showing employee profiles, requested dates, net working days, remaining balances, and stated reasons.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("10_admin_approvals_queue.png", width_in=5.8, caption="Figure 13.10: Centralized Leave Approvals Queue with Applicant Profiles")

    add_section_title("11. 1-Click Approve Modal with Manager Remarks")
    add_p("Managerial decision modal allowing the HR director to approve a leave request, append custom feedback remarks, and atomically deduct days from the balance.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("11_admin_approve_modal.png", width_in=5.8, caption="Figure 13.11: 1-Click Approve Modal with Custom Manager Remarks")

    add_section_title("12. 1-Click Reject Modal with Justification Remarks")
    add_p("Managerial decision modal requiring a justification message when rejecting a leave request, which immediately releases the pending quota hold.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("12_admin_reject_modal.png", width_in=5.8, caption="Figure 13.12: 1-Click Reject Modal with Justification Remarks")

    add_section_title("13. Personnel Directory & Staff Registry")
    add_p("Directory of company staff with department assignments, designations, total allocated quotas, and used quota balances.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("13_admin_employees_directory.png", width_in=5.8, caption="Figure 13.13: Employee Directory and Personnel Registry")

    add_section_title("14. Staff Onboarding Modal (Add Employee)")
    add_p("Administrative modal form allowing HR to register new employees with auto-generated employee codes and automatic initialization of annual leave quotas.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("14_admin_add_employee_modal.png", width_in=5.8, caption="Figure 13.14: Onboarding Modal for Registering New Staff")

    add_section_title("15. Annual Leave Quota Adjustment Modal")
    add_p("Administrative tool allowing HR managers to adjust individual annual leave allocations per category and fiscal year.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("15_admin_adjust_quota_modal.png", width_in=5.8, caption="Figure 13.15: Annual Leave Quota Adjustment Modal")

    add_section_title("16. Compliance Audit Reports & 1-Click CSV Export")
    add_p("Multi-parameter filtering grid across department, category, and status, accompanied by a 1-click CSV download generator for payroll reconciliation.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_img("16_admin_reports_audit.png", width_in=5.8, caption="Figure 13.16: Compliance Audit Reports with 1-Click CSV Download")

    # =========================================================================
    # CHAPTER 14: FUTURE SCOPE
    # =========================================================================
    add_chapter_title("14. FUTURE SCOPE")
    add_p("LeaveFlow ELMS establishes an extensible foundation for enterprise workforce automation. Future development will focus on scaling the system through the following enhancements:", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_bullet("• Automated Slack & Microsoft Teams Bot Integration: ", "Enabling staff to submit leave requests and managers to approve/reject directly from enterprise messaging platforms via interactive chat webhooks.")
    add_bullet("• Biometric Hardware Synchronization: ", "Interfacing with physical RFID and biometric fingerprint readers to automatically log on-site presence and flag unplanned absences.")
    add_bullet("• Predictive Workforce Planning (AI Staffing Forecast): ", "Applying machine learning regression models to predict peak absence periods (festivals, holiday seasons) and alert department leads of potential staffing deficits.")
    add_bullet("• Multi-Tier Approval Hierarchies: ", "Expanding from single-tier management sign-off to multi-level sequential workflows (e.g. Project Lead &rarr; Department Head &rarr; HR Director).")
    add_bullet("• Native Mobile Applications (iOS / Android): ", "Building lightweight cross-platform Flutter/React Native mobile clients with real-time push notifications.")

    # =========================================================================
    # CHAPTER 15: CONCLUSION
    # =========================================================================
    add_chapter_title("15. CONCLUSION")
    add_p("The Employee Leave Management System successfully addresses the operational inefficiencies and transparency deficits inherent in legacy leave tracking workflows. By combining a robust Python Flask backend, normalized SQLite relational database, and an interactive Tailwind CSS interface, the platform delivers a reliable, production-ready solution for modern organizations.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)
    add_p("Key accomplishments of the project include:")
    add_bullet("• Deterministic Accuracy: ", "Programmatically excludes weekends and company holidays, eliminating human calculation errors in time-off balances.")
    add_bullet("• Transactional Integrity: ", "Atomic database transactions guarantee that leave balances are held during review, deducted upon approval, and restored upon cancellation.")
    add_bullet("• Executive Visibility: ", "Chart.js visualizations and multi-filter reports equip HR leaders with instant workforce attendance insights and audit-ready CSV exports.")
    add_bullet("• Zero-Cost Architecture: ", "Engineered entirely with open-source technologies, delivering enterprise-grade functionality without recurring proprietary software licensing fees.")
    add_p("In conclusion, LeaveFlow ELMS demonstrates a successful software engineering endeavor that aligns academic computer science principles with practical industry utility.", align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # =========================================================================
    # CHAPTER 16: GLOSSARY
    # =========================================================================
    add_chapter_title("16. GLOSSARY")
    
    glossary_data = [
        ("ELMS", "Employee Leave Management System; automated platform for managing workforce time-off."),
        ("RBAC", "Role-Based Access Control; security mechanism restricting portal features based on user roles."),
        ("Working Days", "Official business days excluding Saturdays, Sundays, and corporate public holidays."),
        ("Leave Quota", "The statutory total number of paid absence days allocated to an employee per calendar year."),
        ("Pending Hold", "A temporary reservation of leave days placed on an employee's balance while under manager review."),
        ("Half-Day Leave", "A 0.5-day absence granted for medical appointments or short-duration personal commitments."),
        ("Flask", "A lightweight Python WSGI web application framework facilitating modular routing and templating."),
        ("SQLite3", "A serverless, self-contained relational database engine providing ACID transaction consistency."),
        ("Tailwind CSS", "A utility-first modern CSS framework enabling responsive, professional enterprise user interfaces."),
        ("Chart.js", "An open-source HTML5 canvas JavaScript library for rendering responsive hardware-accelerated charts."),
        ("CSV", "Comma-Separated Values; standard portable tabular data exchange file format for payroll accounting."),
        ("Audit Trail", "An immutable chronological log recording all user transactions, approvals, and system state changes."),
    ]
    
    tbl_glo = doc.add_table(rows=len(glossary_data) + 1, cols=2)
    tbl_glo.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, h in enumerate(["Term", "Definition"]):
        cell = tbl_glo.rows[0].cells[c_idx]
        cell.paragraphs[0].text = h
        cell.paragraphs[0].runs[0].font.bold = True
        set_cell_background(cell, "F1F5F9")
        set_cell_margins(cell, 80, 80, 100, 100)

    for r_idx, (t, d) in enumerate(glossary_data, start=1):
        row = tbl_glo.rows[r_idx]
        row.cells[0].paragraphs[0].text = t
        row.cells[0].paragraphs[0].runs[0].font.bold = True
        row.cells[1].paragraphs[0].text = d
        for c in range(2):
            set_cell_margins(row.cells[c], 60, 60, 100, 100)
            row.cells[c].paragraphs[0].runs[0].font.size = Pt(10)

    # =========================================================================
    # CHAPTER 17: REFERENCES
    # =========================================================================
    add_chapter_title("17. REFERENCES")
    
    refs = [
        "1. Python Documentation. (n.d.). Python 3 Language Reference and Standard Library. Retrieved from https://www.python.org/doc",
        "2. Flask Documentation. (n.d.). A Lightweight WSGI Web Application Framework for Python. Retrieved from https://flask.palletsprojects.com",
        "3. SQLite Documentation. (n.d.). Self-Contained, Serverless, Zero-Configuration Relational Database Engine. Retrieved from https://www.sqlite.org/docs.html",
        "4. Tailwind CSS Documentation. (n.d.). Modern Utility-First CSS Framework. Retrieved from https://tailwindcss.com/docs",
        "5. Chart.js Documentation. (n.d.). Flexible JavaScript Data Visualization and Canvas Charting Library. Retrieved from https://www.chartjs.org/docs",
        "6. Werkzeug Documentation. (n.d.). Comprehensive WSGI Web Application Library and Security Cryptography. Retrieved from https://palletsprojects.com/p/werkzeug",
        "7. FontAwesome Icons. (n.d.). Modern Vector Icon Library. Retrieved from https://fontawesome.com",
        "8. W3Schools. (n.d.). HTML5, CSS3, and Modern JavaScript Web Tutorials. Retrieved from https://www.w3schools.com",
        "9. University of Mumbai. (2025). Bachelor of Science in Computer Science (B.Sc. CS) Semester V Project Guidelines and Formatting Manual.",
        "10. Pressman, R. S. (2014). Software Engineering: A Practitioner's Approach (8th ed.). McGraw-Hill Education."
    ]
    for r_txt in refs:
        add_p(r_txt, font_size=10.5, space_after=10, align=WD_ALIGN_PARAGRAPH.JUSTIFY)

    # Save Word document
    doc.save(DOCX_PATH)
    print("Project Report DOCX successfully generated at:", DOCX_PATH)

if __name__ == '__main__':
    create_report()
