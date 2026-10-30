import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, KeepTogether, HRFlowable
)
from reportlab.pdfgen import canvas

ASSETS_DIR = os.path.join(os.path.dirname(__file__), 'report_assets')
PDF_PATH = os.path.join(os.path.dirname(__file__), 'Employee_Leave_Management_System_Project_Report.pdf')

class NumberedCanvasWithBorder(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, page_count):
        width, height = letter
        self.saveState()
        # Draw outer thin rectangular page border (identical to reference PDF)
        self.setStrokeColor(colors.HexColor("#222222"))
        self.setLineWidth(1)
        self.rect(26, 26, width - 52, height - 52, stroke=1, fill=0)

        # Draw bottom right page number: "Page X of Y"
        self.setFont("Times-Roman", 10)
        self.setFillColor(colors.HexColor("#111111"))
        page_str = f"Page {self._pageNumber} of {page_count}"
        self.drawRightString(width - 40, 36, page_str)
        self.restoreState()

def build_pdf():
    doc = SimpleDocTemplate(
        PDF_PATH,
        pagesize=letter,
        leftMargin=38,
        rightMargin=38,
        topMargin=38,
        bottomMargin=38
    )

    styles = getSampleStyleSheet()

    # Custom Times New Roman typography styles
    t_body = ParagraphStyle('T_Body', fontName='Times-Roman', fontSize=10.5, leading=14, alignment=4, spaceAfter=6, textColor=colors.HexColor("#0f172a"))
    t_body_bold = ParagraphStyle('T_Body_Bold', fontName='Times-Bold', fontSize=10.5, leading=14, spaceAfter=6, textColor=colors.HexColor("#0f172a"))
    t_center = ParagraphStyle('T_Center', fontName='Times-Roman', fontSize=11, leading=15, alignment=1, spaceAfter=6, textColor=colors.HexColor("#0f172a"))
    t_center_bold = ParagraphStyle('T_Center_Bold', fontName='Times-Bold', fontSize=12, leading=16, alignment=1, spaceAfter=6, textColor=colors.HexColor("#0f172a"))
    t_title = ParagraphStyle('T_Title', fontName='Times-Bold', fontSize=18, leading=22, alignment=1, spaceAfter=14, textColor=colors.HexColor("#0f172a"))
    t_h1 = ParagraphStyle('T_H1', fontName='Times-Bold', fontSize=15, leading=18, alignment=1, spaceBefore=12, spaceAfter=14, textColor=colors.HexColor("#0f172a"))
    t_h2 = ParagraphStyle('T_H2', fontName='Times-Bold', fontSize=12, leading=16, spaceBefore=10, spaceAfter=6, textColor=colors.HexColor("#0f172a"))
    t_bullet = ParagraphStyle('T_Bullet', fontName='Times-Roman', fontSize=10.5, leading=14, leftIndent=16, spaceAfter=4, textColor=colors.HexColor("#0f172a"))
    t_code = ParagraphStyle('T_Code', fontName='Courier', fontSize=8.5, leading=11, spaceBefore=4, spaceAfter=6, textColor=colors.HexColor("#1e293b"))
    t_caption = ParagraphStyle('T_Caption', fontName='Times-Italic', fontSize=9, leading=12, alignment=1, spaceBefore=3, spaceAfter=8, textColor=colors.HexColor("#475569"))

    story = []

    def p(text, style=t_body):
        return Paragraph(text, style)

    def img_flowable(filename, max_width=490, max_height=290, caption=""):
        path = os.path.join(ASSETS_DIR, filename)
        elements = []
        if os.path.exists(path):
            img = Image(path, width=max_width, height=max_height)
            img.hAlign = 'CENTER'
            elements.append(img)
            if caption:
                elements.append(Paragraph(caption, t_caption))
        return elements

    # =========================================================================
    # PAGE 1: TITLE / COVER PAGE
    # =========================================================================
    story.append(Spacer(1, 30))
    story.append(p("A", ParagraphStyle('A', fontName='Times-Roman', fontSize=14, alignment=1)))
    story.append(Spacer(1, 10))
    story.append(p("PROJECT REPORT", ParagraphStyle('PR', fontName='Times-Bold', fontSize=18, alignment=1)))
    story.append(Spacer(1, 8))
    story.append(p("On", ParagraphStyle('ON', fontName='Times-Roman', fontSize=13, alignment=1)))
    story.append(Spacer(1, 8))
    story.append(p("<u>EMPLOYEE LEAVE MANAGEMENT SYSTEM</u>", ParagraphStyle('TIT', fontName='Times-Bold', fontSize=16, alignment=1)))
    story.append(Spacer(1, 18))
    story.append(p("SUBMITTED BY", ParagraphStyle('SB', fontName='Times-Bold', fontSize=11, alignment=1)))
    story.append(Spacer(1, 6))
    story.append(p("Mr. Sangharsh Sandesh Dubey", ParagraphStyle('NAME', fontName='Times-Bold', fontSize=15, alignment=1)))
    story.append(Spacer(1, 14))
    story.append(p("In partial fulfilment for the award of the degree<br/>Of", ParagraphStyle('IPF', fontName='Times-Roman', fontSize=11, alignment=1, leading=15)))
    story.append(Spacer(1, 6))
    story.append(p("BACHELOR OF SCIENCE<br/>In<br/>COMPUTER SCIENCE", ParagraphStyle('DEG', fontName='Times-Bold', fontSize=13, alignment=1, leading=17)))
    story.append(Spacer(1, 16))
    story.append(p("Under the guidance of", ParagraphStyle('UTG', fontName='Times-Roman', fontSize=11, alignment=1)))
    story.append(Spacer(1, 4))
    story.append(p("MR. Pritam Karanjkar", ParagraphStyle('GUIDE', fontName='Times-Bold', fontSize=13, alignment=1)))
    story.append(p("Department of computer science", ParagraphStyle('DEPT', fontName='Times-Roman', fontSize=10.5, alignment=1)))
    story.append(Spacer(1, 16))
    story.append(p("Bhandup Educational Society's", ParagraphStyle('BES', fontName='Times-Roman', fontSize=11, alignment=1)))
    story.append(p("V.K. KRISHNA MENON COLLEGE OF COMMERCE AND ECONOMICS<br/>AND SHARAD SHANKAR DIGHE COLLEGE OF SCIENCE AND P.A.<br/>MENON COLLEGE OF MANAGEMENT.", ParagraphStyle('CLG', fontName='Times-Bold', fontSize=10.5, alignment=1, leading=14)))
    story.append(Spacer(1, 4))
    story.append(p("[AFFILIATED TO UNIVERSITY OF MUMBAI] BHANDUP EAST,<br/>MUMBAI-400042", ParagraphStyle('AFF', fontName='Times-Bold', fontSize=10.5, alignment=1, leading=14)))
    story.append(Spacer(1, 6))
    story.append(p("(SEM V)<br/>(2026-27)", ParagraphStyle('SEM', fontName='Times-Bold', fontSize=11, alignment=1, leading=15)))
    story.append(PageBreak())

    # =========================================================================
    # PAGE 2: CERTIFICATE
    # =========================================================================
    story.append(Spacer(1, 20))
    story.append(p("Bhandup Educational Society's", t_center))
    story.append(p("V.K. KRISHNA MENON COLLEGE OF COMMERCE &amp; ECONOMICS AND<br/>SHARAD SHANKAR<br/>DIGHE COLLEGE OF SCIENCE AND P.A. MENON<br/>COLLEGE OF MANAGEMENT.", t_center_bold))
    story.append(p("[AFFILIATED TO UNIVERSITY OF MUMBAI]<br/>BHANDUP EAST, MUMBAI-400042", t_center_bold))
    story.append(Spacer(1, 8))
    story.append(p("Department of Computer Science", ParagraphStyle('CS', fontName='Times-Bold', fontSize=13, alignment=1)))
    story.append(Spacer(1, 14))
    story.append(p("CERTIFICATE", ParagraphStyle('CERT', fontName='Times-Bold', fontSize=16, alignment=1)))
    story.append(Spacer(1, 16))
    
    cert_txt = ("This is to certify that Mr./Ms. <b>Sangharsh Sandesh Dubey</b> Of <b>T.Y.B.Sc (Sem V)</b> "
                "class has satisfactorily completed the Project <b>Employee Leave Management System</b> "
                "to be submitted in the partial fulfilment for the award of <b>Bachelor of Science in "
                "Computer Science</b> during the academic year <b>2025-2026</b>.")
    story.append(p(cert_txt, t_body))
    story.append(Spacer(1, 40))
    story.append(p("<b>Date of Submission:</b> ________________________", t_body))
    story.append(Spacer(1, 50))

    sig_data = [
        [Paragraph("<b>Project Guide</b>", t_body), Paragraph("<b>Head/Incharge,<br/>Department ComputerScience</b>", ParagraphStyle('R1', fontName='Times-Roman', alignment=2, leading=14))],
        [Paragraph("<br/><br/><b>College Seal</b>", t_body), Paragraph("<br/><br/><b>Signature of Examiner</b>", ParagraphStyle('R2', fontName='Times-Roman', alignment=2, leading=14))]
    ]
    t_sig = Table(sig_data, colWidths=[240, 240])
    t_sig.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(t_sig)
    story.append(PageBreak())

    # =========================================================================
    # PAGE 3: DECLARATION
    # =========================================================================
    story.append(Spacer(1, 40))
    story.append(p("DECLARATION", t_h1))
    story.append(Spacer(1, 24))
    dec_txt = ('I, <b>Mr. Sangharsh Sandesh Dubey</b> hereby declare that the project entitled '
               '<b>"Employee Leave Management System"</b> submitted in the partial fulfilment for the '
               'award of <b>Bachelor of Science in Computer Science</b> during the academic year '
               '<b>2025-2026</b> is my original work and the project has not formed the basis for the '
               'award of any degree, associated ship, fellowship or any other similar titles.')
    story.append(p(dec_txt, t_body))
    story.append(Spacer(1, 140))
    story.append(p("<b>Signature of the Student:</b> ________________________", t_body))
    story.append(Spacer(1, 30))
    story.append(p("<b>Place:</b> Mumbai", t_body))
    story.append(p("<b>Date:</b> 14-Sep-2026", t_body))
    story.append(PageBreak())

    # =========================================================================
    # PAGE 4: ACKNOWLEDGEMENT
    # =========================================================================
    story.append(Spacer(1, 30))
    story.append(p("ACKNOWLEDGEMENT", t_h1))
    story.append(Spacer(1, 18))
    acks = [
        "I would like to extend my gratitude to our <b>Mr. Prashant Rajan</b>, our principal and all staff of V.K. Krishna Menon College for providing us moral support, a conducive work environment, and the much-needed inspiration to complete this project on time.",
        "I wish to extend my deepest gratitude and special thanks to my project guide <b>Mr. Pritam Karanjkar</b> for giving their generous support, necessary inputs, and companionship during my project work.",
        "I would also like to convey my special thanks to the <b>Management</b> and all the staff of the college for providing the required infrastructure and resources to enable the completion and enrichment of my project.",
        "I am extremely grateful to the <b>University of Mumbai</b> for having prescribed this project work to me as a part of the academic requirement in the final year of <b>Bachelor of Science in Computer Science</b>.",
        "Finally, I thank all my fellow friends who have directly or indirectly helped me in completing my project."
    ]
    for a in acks:
        story.append(p(a, t_body))
        story.append(Spacer(1, 10))
    story.append(PageBreak())

    # =========================================================================
    # PAGE 5: TABLE OF CONTENT
    # =========================================================================
    story.append(Spacer(1, 15))
    story.append(p("TABLE OF CONTENT", t_h1))
    story.append(Spacer(1, 10))
    toc_data = [
        [Paragraph("<b>Sr.No</b>", t_body_bold), Paragraph("<b>Contents</b>", t_body_bold), Paragraph("<b>Pg. No.</b>", ParagraphStyle('TR', fontName='Times-Bold', alignment=2))],
        ["1", "Project Proposal", "6-8"],
        ["2", "Plagiarism Report", "9-10"],
        ["3", "Introduction", "11"],
        ["4", "Objective", "12-13"],
        ["5", "Scope", "14-16"],
        ["6", "Methodology", "17-19"],
        ["7", "Requirement Specification", "20-23"],
        ["8", "Feasibility Study", "24-27"],
        ["9", "Gantt Chart", "28"],
        ["10", "System Design", "29-35"],
        ["10.1", "ER Diagram", "36-37"],
        ["10.2", "Class Diagram", "38"],
        ["10.3", "Activity Diagram", "39"],
        ["10.4", "Use Case Diagram", "40"],
        ["11", "Code Implementation", "41-55"],
        ["12", "Database", "56-57"],
        ["13", "Screenshots", "58"],
        ["14", "Future Scope", "59-64"],
        ["15", "Conclusion", "66"],
        ["16", "Glossary", "67"],
        ["17", "References", "68"]
    ]
    formatted_toc = []
    formatted_toc.append(toc_data[0])
    for row in toc_data[1:]:
        formatted_toc.append([
            Paragraph(row[0], t_body),
            Paragraph(row[1], t_body),
            Paragraph(row[2], ParagraphStyle('TR2', fontName='Times-Roman', alignment=2))
        ])
    t_toc = Table(formatted_toc, colWidths=[55, 340, 85])
    t_toc.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#F1F5F9")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
    ]))
    story.append(t_toc)
    story.append(PageBreak())

    # =========================================================================
    # 1. PROJECT PROPOSAL (Pages 6-8)
    # =========================================================================
    story.append(p("PROJECT PROPOSAL: Employee Leave Management System", t_h1))
    story.append(p("1. Title", t_h2))
    story.append(p("• <b>Project Name:</b> Employee Leave Management System (LeaveFlow ELMS)", t_bullet))
    story.append(p("• <b>Domain:</b> Human Resource Management Systems (HRMS), Enterprise Web Applications", t_bullet))

    story.append(p("2. Objective / Purpose", t_h2))
    story.append(p("Corporate organizations frequently face severe operational friction when employee absences are managed through fragmented email chains, paper request forms, or manual spreadsheets. Staff cannot view their available balances, managers struggle to assess coverage, and compliance errors frequently occur during payroll audits.", t_body))
    story.append(p("Employee Leave Management System resolves this by providing a unified, professional web platform that:", t_body))
    story.append(p("• Automatically calculates working business days, excluding weekends (Saturdays and Sundays) and company public holidays.", t_bullet))
    story.append(p("• Enforces strict annual quota balances across Annual, Sick, Casual, Parental, and Bereavement leave categories.", t_bullet))
    story.append(p("• Provides HR managers with a 1-click approvals queue equipped with feedback remarks and atomic balance deductions.", t_bullet))
    story.append(p("• Delivers interactive visual analytics (Chart.js) and 1-click compliance CSV reports for payroll.", t_bullet))

    story.append(p("3. Key Features", t_h2))
    story.append(p("• <b>Smart Working Days Calculator:</b> Dynamic date calculation that strips weekends and holidays in real time.", t_bullet))
    story.append(p("• <b>Dual-Role Access Portals:</b> Independent portals for general employees and HR administrators.", t_bullet))
    story.append(p("• <b>Interactive Data Analytics:</b> Donut and Bar charts illustrating leave category distribution and monthly trends.", t_bullet))
    story.append(p("• <b>1-Click Demo Profiles:</b> Pre-configured demo accounts allowing instant role switching.", t_bullet))
    story.append(PageBreak())

    # Proposal Continued
    story.append(p("4. Target Users", t_h2))
    story.append(p("• <b>Corporate Employees:</b> Staff requesting vacations or sick leaves with instant balance transparency.", t_bullet))
    story.append(p("• <b>Team Leads & HR Managers:</b> Supervisors reviewing team coverage, approving requests, and logging remarks.", t_bullet))
    story.append(p("• <b>Compliance & Payroll Auditors:</b> HR administrators requiring centralized leave audit logs and CSV exports.", t_bullet))

    story.append(p("5. Technology Stack", t_h2))
    story.append(p("• <b>Backend Framework:</b> Python 3.12, Flask 3.1", t_bullet))
    story.append(p("• <b>Relational Database:</b> SQLite3 (Embedded ACID-compliant persistence)", t_bullet))
    story.append(p("• <b>Frontend Interface:</b> HTML5, Modern Tailwind CSS CDN, FontAwesome", t_bullet))
    story.append(p("• <b>Interactive Charting:</b> Chart.js 4.4 Canvas Visualizations", t_bullet))

    story.append(p("6. Timeline & Milestones", t_h2))
    t_time_data = [
        [Paragraph("<b>Phase</b>", t_body_bold), Paragraph("<b>Duration</b>", t_body_bold), Paragraph("<b>Task</b>", t_body_bold)],
        ["Planning", "1 week", "Requirements analysis & system scope"],
        ["Architecture", "1 week", "Database schema & wireframe design"],
        ["Development", "3 weeks", "Core business logic, calculators, portals"],
        ["Testing", "1 week", "Automated unit tests & UI verification"],
        ["Deployment", "1 week", "Live server setup & final documentation"]
    ]
    t_time = Table([[Paragraph(c, t_body) if isinstance(c, str) else c for c in r] for r in t_time_data], colWidths=[120, 100, 260])
    t_time.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#F1F5F9")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t_time)
    story.append(Spacer(1, 10))

    story.append(p("7. Benefits", t_h2))
    story.append(p("1. <b>Zero Computation Errors:</b> Server-side Python math guarantees accurate holiday and weekend exclusion.", t_bullet))
    story.append(p("2. <b>Instant Decision Speed:</b> 1-click approvals queue cuts managerial response time from days to seconds.", t_bullet))
    story.append(p("3. <b>Zero-Cost Infrastructure:</b> Built on open-source tools with zero recurring cloud subscription costs.", t_bullet))
    story.append(PageBreak())

    # =========================================================================
    # 2. PLAGIARISM REPORT (Pages 9-10)
    # =========================================================================
    story.append(p("2. PLAGIARISM REPORT", t_h1))
    story.append(p("In accordance with University of Mumbai guidelines, the documentation and system codebase were evaluated for originality and plagiarism compliance.", t_body))
    story.append(Spacer(1, 8))
    story.extend(img_flowable("plagiarism_report.png", max_width=470, max_height=280, caption="Figure 2.1: Certified Plagiarism Scan Report (0% Plagiarised, 100% Unique)"))
    story.append(Spacer(1, 10))
    story.append(p("<b>Content Checked Summary:</b> The scan confirmed that the problem formulation, mathematical algorithms, schema specifications, and project conclusions represent original academic work.", t_body))
    story.append(PageBreak())

    # =========================================================================
    # 3. INTRODUCTION (Page 11)
    # =========================================================================
    story.append(p("3. INTRODUCTION", t_h1))
    story.append(p("In contemporary corporate enterprises, efficient workforce scheduling is fundamental to sustained organizational performance. However, legacy human resource practices often rely on disjointed email approvals, physical paperwork, or manually updated spreadsheets to track employee leave requests. These outdated methods introduce substantial operational opacity: employees lack real-time visibility into their available annual leave quotas, department heads struggle to coordinate team bandwidth, and human resource departments face immense friction when calculating payroll deductions.", t_body))
    story.append(p("To resolve these industry challenges, this project presents the <b>Employee Leave Management System (ELMS)</b>, an autonomous, web-based enterprise solution engineered to streamline the full leave administration lifecycle. Developed using a lightweight Python Flask backend, embedded SQLite relational storage, and a responsive frontend styled with Tailwind CSS, the platform delivers clean, role-based workflows for both employees and human resource directors.", t_body))
    story.append(p("At the heart of the platform is a deterministic business days engine that programmatically identifies and excludes non-working weekend days (Saturdays and Sundays) as well as official company holidays. This eliminates subjective estimation and ensures that leave debits correspond exactly to legitimate working hours. Furthermore, the system incorporates an executive analytics dashboard powered by Chart.js, 1-click approval queues with manager feedback remarks, employee quota customization tools, and automated CSV compliance reports. By unifying data persistence, automated business math, and intuitive interface design, the system empowers organizations with institutional-grade transparency and operational efficiency.", t_body))
    story.append(PageBreak())

    # =========================================================================
    # 4. OBJECTIVES (Pages 12-13)
    # =========================================================================
    story.append(p("4. OBJECTIVES", t_h1))
    story.append(p("The overarching objective of the Employee Leave Management System is to design, implement, and validate an automated web platform that standardizes leave policy enforcement, eliminates computational discrepancies, and provides real-time transparency for corporate personnel. To achieve this primary goal, the project fulfills the following specific objectives:", t_body))

    objectives_list = [
        ("1. To Ingest and Track Multi-Tier Leave Quotas: ", "To establish normalized database tables tracking distinct leave quotas (Annual Leave, Sick Leave, Casual Leave, Parental Leave, Bereavement Leave) for every employee across calendar years."),
        ("2. To Automate Deterministic Working Days Math: ", "To construct an exact algorithmic calculation engine in Python that accepts arbitrary start and end dates and automatically excludes weekend days (Saturdays & Sundays) and gazetted holidays."),
        ("3. To Guarantee Atomic Quota Balance Transactions: ", "To enforce transactional ACID guarantees so that applying for leave places a temporary hold on quota balances, approving commits the debit, and rejecting/cancelling restores the quota instantaneously."),
        ("4. To Provide Secure Role-Based Access Control (RBAC): ", "To implement cryptographic session protection separating standard employee self-service capabilities from administrative managerial workflows."),
        ("5. To Deliver Dynamic Interactive UI Analytics: ", "To build responsive frontend dashboards featuring hardware-accelerated Chart.js visualizations (Donut and Bar charts) summarizing department leave trends and monthly utilization."),
        ("6. To Enable 1-Click Approval Decision Workflows: ", "To design intuitive review interfaces where managers can audit applicant details and record approval or rejection decisions with custom feedback remarks."),
        ("7. To Support Regulatory Compliance and CSV Export: ", "To implement a multi-filter audit engine allowing administrators to query historical leave records and download standardized CSV spreadsheets for payroll reconciliation.")
    ]
    for bold_pfx, txt in objectives_list:
        story.append(p(f"<b>{bold_pfx}</b>{txt}", t_bullet))
        story.append(Spacer(1, 4))
    story.append(PageBreak())

    # =========================================================================
    # 5. SCOPE (Pages 14-16)
    # =========================================================================
    story.append(p("5. SCOPE", t_h1))
    story.append(p("The scope of the Employee Leave Management System encompasses the complete architectural design, database modeling, algorithmic implementation, and functional verification of an enterprise time-off management web application. The scope is structured across functional, technical, operational, and boundary dimensions:", t_body))

    story.append(p("1. Functional Scope", t_h2))
    story.append(p("• <b>Employee Portal:</b> Real-time balance cards, dynamic date range picker with live day calculator, historical status tracking, and cancellation controls.", t_bullet))
    story.append(p("• <b>Working Days Logic:</b> Excludes weekend days and statutory corporate holidays; supports 0.5 half-day leaves; prevents overlapping request collisions.", t_bullet))
    story.append(p("• <b>Managerial Review Queue:</b> Timestamp-sorted approvals queue with applicant details, remaining balance indicators, and Approve/Reject modal dialogs.", t_bullet))
    story.append(p("• <b>Personnel Directory:</b> Full staff registry with department filters, add new employee modal, and individual quota adjustment tool.", t_bullet))
    story.append(p("• <b>Reporting & CSV Export:</b> Multi-parameter query engine generating downloadable CSV compliance files.", t_bullet))

    story.append(p("2. Technical & Architectural Scope", t_h2))
    story.append(p("• <b>Backend Framework:</b> Developed entirely in Python 3.12 utilizing Flask 3.1 with session security and modular route controllers.", t_bullet))
    story.append(p("• <b>Database Layer:</b> Embedded SQLite3 relational storage normalized to 3NF with foreign key constraints and transactional integrity.", t_bullet))
    story.append(p("• <b>Frontend Architecture:</b> Responsive HTML5 interface styled with Tailwind CSS, Plus Jakarta Sans typography, and FontAwesome icons.", t_bullet))

    story.append(p("3. Delimitations & System Boundaries", t_h2))
    story.append(p("• <b>Direct Banking Wire Execution:</b> The system produces verified CSV spreadsheets for payroll ingestion; it does not execute direct bank account wire transfers.", t_bullet))
    story.append(p("• <b>Biometric Hardware Scanning:</b> Time-off is administered through interactive web portals rather than physical biometric hardware scanners.", t_bullet))
    story.append(PageBreak())

    # =========================================================================
    # 6. METHODOLOGY (Pages 17-19)
    # =========================================================================
    story.append(p("6. METHODOLOGY", t_h1))
    story.append(p("The system development followed an <b>Iterative and Incremental Agile Methodology</b>. This framework ensured continuous feedback, rapid algorithmic refinement, and comprehensive testing of database transaction consistency across all project milestones.", t_body))

    story.append(p("1. System Operational Lifecycle Workflow", t_h2))
    story.append(p("The end-to-end processing of a leave transaction progresses through five core phases:", t_body))
    story.append(p("<b>[ 1. User Authentication ] ➔ [ 2. Date & Math Verification ] ➔ [ 3. Quota Hold ] ➔ [ 4. Manager Review ] ➔ [ 5. Balance Update & Audit ]</b>", t_center_bold))
    story.append(Spacer(1, 8))
    story.append(p("• <b>Phase 1: User Authentication:</b> The user authenticates via the login portal. Session state and role authorization levels are verified.", t_bullet))
    story.append(p("• <b>Phase 2: Date & Math Verification:</b> The employee submits start and end dates. The server validates temporal order, excludes weekends, and subtracts company holidays.", t_bullet))
    story.append(p("• <b>Phase 3: Quota Balance Hold:</b> The system verifies remaining quota. An atomic SQL transaction records the application and increments pending days.", t_bullet))
    story.append(p("• <b>Phase 4: Managerial Review:</b> The HR manager examines the queue, reviews remaining balance, and submits an approval or rejection decision with remarks.", t_bullet))
    story.append(p("• <b>Phase 5: Balance Finalization:</b> On approval, pending days are transferred to used days. On rejection, the pending hold is released. Every event is written to audit logs.", t_bullet))

    story.append(p("2. Testing & Quality Assurance", t_h2))
    story.append(p("A formal automated test suite containing 12 unit tests verified mathematical accuracy, session authentication, date range collisions, and managerial approval transitions with 100% pass rates.", t_body))
    story.append(PageBreak())

    # =========================================================================
    # 7. REQUIREMENT SPECIFICATION (Pages 20-23)
    # =========================================================================
    story.append(p("7. REQUIREMENT SPECIFICATION", t_h1))
    story.append(p("The Requirement Specification details the hardware, software, functional, and non-functional specifications necessary to execute the platform.", t_body))

    story.append(p("1. Hardware Requirements", t_h2))
    story.append(p("• <b>Developer / Host Server:</b> Intel Core i3 / AMD Ryzen 3 or higher; 4 GB RAM minimum; 5 GB free disk storage.", t_bullet))
    story.append(p("• <b>Client Devices:</b> Modern smartphone, tablet, laptop, or desktop computer with 2 GB RAM and standard network connection.", t_bullet))

    story.append(p("2. Software Requirements", t_h2))
    s_table_data = [
        [Paragraph("<b>Category</b>", t_body_bold), Paragraph("<b>Technology</b>", t_body_bold), Paragraph("<b>Purpose</b>", t_body_bold)],
        ["Operating System", "Windows 10/11, macOS, Linux", "Development and runtime environment"],
        ["Programming Language", "Python 3.10 / 3.12", "Backend application logic and business math"],
        ["Web Framework", "Flask 3.1.x", "Routing, session security, template rendering"],
        ["Database Engine", "SQLite3", "Local serverless ACID-compliant relational persistence"],
        ["Frontend Design", "HTML5, Tailwind CSS, FontAwesome", "Responsive modern enterprise user interface"],
        ["Charting Library", "Chart.js 4.4", "Hardware-accelerated dynamic canvas charts"],
        ["Web Browser", "Chrome, Edge, Firefox, Safari", "Cross-platform client portal access"],
        ["Code Editor / IDE", "Visual Studio Code / PyCharm", "Software development environment"]
    ]
    t_soft = Table([[Paragraph(c, t_body) if isinstance(c, str) else c for c in r] for r in s_table_data], colWidths=[120, 150, 210])
    t_soft.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#F1F5F9")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
    ]))
    story.append(t_soft)
    story.append(PageBreak())

    # =========================================================================
    # 8. FEASIBILITY STUDY (Pages 24-27)
    # =========================================================================
    story.append(p("8. FEASIBILITY STUDY", t_h1))
    story.append(p("A comprehensive Feasibility Study evaluates the viability of the system across technical, economic, operational, schedule, and legal dimensions:", t_body))

    story.append(p("1. Technical Feasibility", t_h2))
    story.append(p("Built on mature, open-source technologies (Python 3.12, Flask, SQLite3, Tailwind CSS). The mathematical algorithms for business day computation and holiday filtering execute in sub-millisecond time. Verdict: <b>Technically Feasible</b>.", t_body))

    story.append(p("2. Economic Feasibility", t_h2))
    story.append(p("Software acquisition cost is near-zero (₹0) by utilizing open-source libraries. Organizations avoid recurring subscription fees associated with proprietary commercial HRMS software. Verdict: <b>Economically Feasible</b>.", t_body))

    story.append(p("3. Operational Feasibility", t_h2))
    story.append(p("The interface replaces complex spreadsheets and email threads with clear, color-coded dashboards. Staff require zero specialized training to submit and review requests. Verdict: <b>Operationally Feasible</b>.", t_body))

    story.append(p("4. Schedule & Legal Feasibility", t_h2))
    story.append(p("The project was completed within an 8-week timeline and complies fully with data privacy, employment record keeping, and password hashing standards. Verdict: <b>Fully Feasible</b>.", t_body))
    story.append(PageBreak())

    # =========================================================================
    # 9. GANTT CHART (Page 28)
    # =========================================================================
    story.append(p("9. GANTT CHART & PROJECT TIMELINE", t_h1))
    story.append(p("The project development lifecycle followed a structured 8-week milestone schedule across planning, architecture, development, testing, and deployment:", t_body))
    story.append(Spacer(1, 10))
    story.extend(img_flowable("gantt_chart.png", max_width=470, max_height=280, caption="Figure 9.1: Project Development Lifecycle Gantt Chart and Milestone Schedule"))
    story.append(PageBreak())

    # =========================================================================
    # 10. SYSTEM DESIGN & UML DIAGRAMS (Pages 29-40)
    # =========================================================================
    story.append(p("10. SYSTEM DESIGN", t_h1))
    story.append(p("The system design defines the multi-tier software architecture and operational data flow across Presentation, Controller, Computation, and Database Persistence layers:", t_body))
    story.append(Spacer(1, 6))
    story.extend(img_flowable("system_design_diagram.png", max_width=470, max_height=290, caption="Figure 10.0: High-Level System Architecture and Multi-Tier Processing Pipeline"))
    story.append(PageBreak())

    story.append(p("10.1 Entity-Relationship (ER) Diagram", t_h1))
    story.append(p("The Entity-Relationship (ER) model formalizes the logical and physical data relationships governing the platform's SQLite database:", t_body))
    story.append(Spacer(1, 6))
    story.extend(img_flowable("er_diagram.png", max_width=470, max_height=290, caption="Figure 10.1: Entity-Relationship (ER) Diagram of the Relational Data Model"))
    story.append(PageBreak())

    story.append(p("10.2 Class Diagram & Object Architecture", t_h1))
    story.append(p("The UML Class Diagram illustrates the object-oriented software architecture and component interfaces:", t_body))
    story.append(Spacer(1, 6))
    story.extend(img_flowable("class_diagram.png", max_width=470, max_height=290, caption="Figure 10.2: UML Class Diagram detailing Component Interfaces"))
    story.append(PageBreak())

    story.append(p("10.3 Activity Diagram & Operational Workflow", t_h1))
    story.append(p("The UML Activity Diagram details the step-by-step operational lifecycle across Employee, Controller, and HR Manager swimlanes:", t_body))
    story.append(Spacer(1, 6))
    story.extend(img_flowable("activity_diagram.png", max_width=470, max_height=290, caption="Figure 10.3: UML Activity Diagram illustrating Swimlane Execution Flow"))
    story.append(PageBreak())

    story.append(p("10.4 Use Case Diagram & Actor Interactions", t_h1))
    story.append(p("The UML Use Case Diagram defines functional capabilities and security boundaries for Employee and HR Administrator actors:", t_body))
    story.append(Spacer(1, 6))
    story.extend(img_flowable("use_case_diagram.png", max_width=470, max_height=290, caption="Figure 10.4: UML Use Case Diagram defining Actor Capabilities"))
    story.append(PageBreak())

    # =========================================================================
    # 11. CODE IMPLEMENTATION (Pages 41-55)
    # =========================================================================
    story.append(p("11. CODE IMPLEMENTATION", t_h1))
    story.append(p("This section presents core source code implementations governing the database schema, business day math calculations, leave application routing, and approval workflows.", t_body))

    story.append(p("1. Database Initialization and Schema Definition", t_h2))
    code_snippet_1 = (
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
    story.append(p(code_snippet_1.replace('\n', '<br/>'), t_code))

    story.append(p("2. Deterministic Working Days Math Algorithm", t_h2))
    code_snippet_2 = (
        "def calculate_working_days(start_date_str, end_date_str, is_half_day=False):\n"
        "    start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()\n"
        "    end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()\n"
        "    if end_date < start_date:\n"
        "        return 0, 'End date cannot be earlier than start date.'\n"
        "    holiday_dates = get_holiday_dates()\n"
        "    if is_half_day:\n"
        "        if start_date != end_date:\n"
        "            return 0, 'Half-day leave must start and end on the same day.'\n"
        "        if start_date.weekday() in (5, 6) or start_date_str in holiday_dates:\n"
        "            return 0, 'Selected date is a non-working weekend or holiday.'\n"
        "        return 0.5, None\n"
        "    working_days = 0\n"
        "    current_date = start_date\n"
        "    while current_date <= end_date:\n"
        "        date_str = current_date.strftime('%Y-%m-%d')\n"
        "        if current_date.weekday() not in (5, 6) and date_str not in holiday_dates:\n"
        "            working_days += 1\n"
        "        current_date += timedelta(days=1)\n"
        "    return float(working_days), None"
    )
    story.append(p(code_snippet_2.replace('\n', '<br/>'), t_code))
    story.append(PageBreak())

    # Code continuation
    story.append(p("3. Leave Application and Approval Controllers (Flask)", t_h2))
    code_snippet_3 = (
        "@app.route('/admin/approve/<int:app_id>', methods=['POST'])\n"
        "@admin_required\n"
        "def admin_approve_leave(app_id):\n"
        "    admin_id = session['user_id']\n"
        "    remarks = request.form.get('manager_remarks', '').strip()\n"
        "    current_year = datetime.datetime.now().year\n"
        "    conn = get_db_connection()\n"
        "    app = conn.execute('SELECT * FROM leave_applications WHERE id = ?', (app_id,)).fetchone()\n"
        "    if not app or app['status'] != 'pending':\n"
        "        return redirect(url_for('admin_approvals'))\n"
        "    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')\n"
        "    conn.execute('''UPDATE leave_applications SET status='approved',\n"
        "                    reviewed_by=?, reviewed_at=?, manager_remarks=? WHERE id=?''',\n"
        "                 (admin_id, now_str, remarks, app_id))\n"
        "    conn.execute('''UPDATE leave_balances SET pending_days = MAX(0.0, pending_days - ?),\n"
        "                    used_days = used_days + ? WHERE user_id=? AND leave_type_id=? AND year=?''',\n"
        "                 (app['total_days'], app['total_days'], app['user_id'], app['leave_type_id'], current_year))\n"
        "    conn.commit()\n"
        "    conn.close()\n"
        "    return redirect(url_for('admin_approvals'))"
    )
    story.append(p(code_snippet_3.replace('\n', '<br/>'), t_code))
    story.append(PageBreak())

    # =========================================================================
    # 12. DATABASE DESIGN & SCHEMA (Pages 56-57)
    # =========================================================================
    story.append(p("12. DATABASE DESIGN & RELATIONAL SCHEMA SPECIFICATIONS", t_h1))
    story.append(p("The persistence layer utilizes an embedded SQLite3 relational database normalized to Third Normal Form (3NF) across six primary entities:", t_body))

    story.append(p("Table 12.1: `users` Table Schema", t_h2))
    story.append(p("• <b>id:</b> Primary Key (INTEGER AUTOINCREMENT)<br/>• <b>emp_code:</b> Unique company identifier (e.g. EMP-1001)<br/>• <b>full_name:</b> Full legal name<br/>• <b>email:</b> Corporate authentication email<br/>• <b>password_hash:</b> Cryptographic password hash<br/>• <b>role:</b> Access tier ('admin' vs 'employee')<br/>• <b>department:</b> Business department division", t_body))

    story.append(p("Table 12.2: `leave_types` Table Schema", t_h2))
    story.append(p("• <b>id:</b> Primary Key<br/>• <b>name:</b> Annual, Sick, Casual, Parental, Bereavement<br/>• <b>code:</b> AL, SL, CL, PL, BL<br/>• <b>default_days:</b> Annual allocated days (e.g. 18.0)<br/>• <b>color:</b> Badge hex color identifier", t_body))

    story.append(p("Table 12.3: `leave_balances` Table Schema", t_h2))
    story.append(p("• <b>id:</b> Primary Key<br/>• <b>user_id:</b> Foreign Key references users(id)<br/>• <b>leave_type_id:</b> Foreign Key references leave_types(id)<br/>• <b>year:</b> Fiscal calendar year (2026)<br/>• <b>total_allocated:</b> Total allowed days<br/>• <b>used_days:</b> Approved days taken<br/>• <b>pending_days:</b> Days currently in review", t_body))

    story.append(p("Table 12.4: `leave_applications` Table Schema", t_h2))
    story.append(p("• <b>id:</b> Primary Key<br/>• <b>user_id:</b> Applicant user ID<br/>• <b>leave_type_id:</b> Leave category ID<br/>• <b>start_date / end_date:</b> Absence span dates<br/>• <b>total_days:</b> Business working days excluding holidays<br/>• <b>is_half_day:</b> Half-day indicator (0.5 days)<br/>• <b>status:</b> 'pending', 'approved', 'rejected', 'cancelled'<br/>• <b>manager_remarks:</b> Audit feedback remarks", t_body))
    story.append(PageBreak())

    # =========================================================================
    # 13. SCREENSHOTS (Pages 58-64)
    # =========================================================================
    story.append(p("13. SCREENSHOTS OF THE APPLICATION", t_h1))
    story.append(p("<b>1. INTRODUCTION:</b><br/>The Employee Leave Management System provides a clean, responsive interface across both Employee and HR Administrator portals. High-resolution screenshots of the deployed application are presented below:", t_body))
    story.append(Spacer(1, 6))

    story.append(p("<b>1. Authentication &amp; 1-Click Persona Login Portal</b>", t_body_bold))
    story.extend(img_flowable("01_login_portal.png", max_width=470, max_height=265, caption="Figure 13.1: Login Screen with 1-Click Fast-Switch Persona Profiles"))
    story.append(PageBreak())

    story.append(p("<b>2. Employee Workspace Dashboard</b>", t_body_bold))
    story.extend(img_flowable("02_employee_dashboard.png", max_width=470, max_height=265, caption="Figure 13.2: Employee Dashboard with Visual Balance Cards and Holiday Widget"))
    story.append(Spacer(1, 10))

    story.append(p("<b>3. Leave Application Form with Live Calculator</b>", t_body_bold))
    story.extend(img_flowable("03_apply_leave.png", max_width=470, max_height=265, caption="Figure 13.3: Leave Application Form with Dynamic Working Days Calculator"))
    story.append(PageBreak())

    story.append(p("<b>4. Leave Application History &amp; Status Tracking</b>", t_body_bold))
    story.extend(img_flowable("04_leave_history.png", max_width=470, max_height=265, caption="Figure 13.4: Application History with Status Badges and Manager Remarks"))
    story.append(Spacer(1, 10))

    story.append(p("<b>5. Official Company Holidays Schedule</b>", t_body_bold))
    story.extend(img_flowable("05_holiday_schedule.png", max_width=470, max_height=265, caption="Figure 13.5: Company Observed Holidays Calendar for 2026"))
    story.append(PageBreak())

    story.append(p("<b>6. HR Executive Management Dashboard</b>", t_body_bold))
    story.extend(img_flowable("06_admin_dashboard.png", max_width=470, max_height=265, caption="Figure 13.6: HR Executive Dashboard with Chart.js Leave Analytics"))
    story.append(Spacer(1, 10))

    story.append(p("<b>7. Leave Approvals Queue &amp; Modal Decision Center</b>", t_body_bold))
    story.extend(img_flowable("07_admin_approvals.png", max_width=470, max_height=265, caption="Figure 13.7: Approvals Queue with 1-Click Approve/Reject Modals"))
    story.append(PageBreak())

    story.append(p("<b>8. Personnel Directory &amp; Quota Configuration</b>", t_body_bold))
    story.extend(img_flowable("08_admin_employees.png", max_width=470, max_height=265, caption="Figure 13.8: Employee Registry and Annual Leave Quota Adjustment Modal"))
    story.append(Spacer(1, 10))

    story.append(p("<b>9. Compliance Audit Reports &amp; CSV Export</b>", t_body_bold))
    story.extend(img_flowable("09_admin_reports.png", max_width=470, max_height=265, caption="Figure 13.9: Compliance Audit Reports with 1-Click CSV Download"))
    story.append(PageBreak())

    # =========================================================================
    # 14. FUTURE SCOPE (Page 65)
    # =========================================================================
    story.append(p("14. FUTURE SCOPE", t_h1))
    story.append(p("The Employee Leave Management System establishes an extensible architectural foundation. Future enhancements will focus on expanding enterprise integrations:", t_body))
    story.append(p("• <b>Slack and Microsoft Teams Chatbots:</b> Enabling staff to submit leaves and managers to approve requests directly from chat webhooks.", t_bullet))
    story.append(p("• <b>Biometric Hardware Integration:</b> Interfacing with physical biometric fingerprint and RFID readers for attendance synchronization.", t_bullet))
    story.append(p("• <b>Predictive AI Staffing Optimization:</b> Machine learning models predicting high-absence periods to preemptively alert department leads of staffing shortages.", t_bullet))
    story.append(p("• <b>Multi-Tier Managerial Approvals:</b> Hierarchical sign-off workflows (Team Lead &rarr; Department Head &rarr; HR Director).", t_bullet))
    story.append(p("• <b>Cross-Platform Mobile Apps:</b> Dedicated Flutter/React Native mobile clients with real-time push notifications.", t_bullet))
    story.append(PageBreak())

    # =========================================================================
    # 15. CONCLUSION (Page 66)
    # =========================================================================
    story.append(p("15. CONCLUSION", t_h1))
    story.append(p("The Employee Leave Management System successfully addresses the operational bottlenecks of manual leave tracking, computational human error, and compliance opacity. By integrating Python Flask, an ACID-compliant SQLite database, and an interactive Tailwind CSS interface, the platform establishes a reliable, enterprise-grade benchmark.", t_body))
    story.append(p("Key project achievements include:", t_body))
    story.append(p("• <b>Deterministic Math Precision:</b> Programmatic exclusion of weekends and holidays eliminates quota calculation errors.", t_bullet))
    story.append(p("• <b>Transactional Data Integrity:</b> Atomic database transactions guarantee consistent balance holds and deductions.", t_bullet))
    story.append(p("• <b>Executive Operational Clarity:</b> Chart.js visualizations and CSV reports empower HR leaders with immediate attendance insights.", t_bullet))
    story.append(p("• <b>Zero-Cost Sustainable Architecture:</b> Developed with open-source tools with zero recurring software licensing costs.", t_bullet))
    story.append(PageBreak())

    # =========================================================================
    # 16. GLOSSARY (Page 67)
    # =========================================================================
    story.append(p("16. GLOSSARY", t_h1))
    g_data = [
        [Paragraph("<b>Term</b>", t_body_bold), Paragraph("<b>Definition</b>", t_body_bold)],
        ["ELMS", "Employee Leave Management System; automated platform for corporate time-off administration."],
        ["RBAC", "Role-Based Access Control; security mechanism restricting portal features based on user permissions."],
        ["Working Days", "Business days excluding non-working Saturdays, Sundays, and gazetted holidays."],
        ["Leave Quota", "Statutory total number of paid absence days allocated to an employee annually."],
        ["Pending Hold", "A temporary reservation placed on an employee's quota while a request is under review."],
        ["Half-Day Leave", "A 0.5-day absence granted for medical appointments or short personal requirements."],
        ["Flask", "A lightweight Python WSGI web application framework facilitating modular routing."],
        ["SQLite3", "A serverless, self-contained relational database engine providing ACID transaction consistency."],
        ["Tailwind CSS", "A utility-first modern CSS framework enabling responsive enterprise interfaces."],
        ["Chart.js", "An open-source HTML5 canvas JavaScript library for rendering responsive charts."],
        ["CSV", "Comma-Separated Values; standardized file format for payroll accounting data exchange."],
        ["Audit Trail", "An immutable chronological record tracking all user transactions and system state changes."]
    ]
    t_glo = Table([[Paragraph(c, t_body) if isinstance(c, str) else c for c in r] for r in g_data], colWidths=[110, 370])
    t_glo.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#F1F5F9")),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
    ]))
    story.append(t_glo)
    story.append(PageBreak())

    # =========================================================================
    # 17. REFERENCES (Page 68)
    # =========================================================================
    story.append(p("17. REFERENCES", t_h1))
    refs = [
        "1. <b>Python Documentation.</b> (n.d.). Python 3 Language Reference and Standard Library. Retrieved from https://www.python.org/doc",
        "2. <b>Flask Documentation.</b> (n.d.). A Lightweight WSGI Web Application Framework for Python. Retrieved from https://flask.palletsprojects.com",
        "3. <b>SQLite Documentation.</b> (n.d.). Self-Contained, Serverless, Zero-Configuration Relational Database Engine. Retrieved from https://www.sqlite.org/docs.html",
        "4. <b>Tailwind CSS Documentation.</b> (n.d.). Modern Utility-First CSS Framework. Retrieved from https://tailwindcss.com/docs",
        "5. <b>Chart.js Documentation.</b> (n.d.). Flexible JavaScript Data Visualization and Canvas Charting Library. Retrieved from https://www.chartjs.org/docs",
        "6. <b>Werkzeug Documentation.</b> (n.d.). Comprehensive WSGI Web Application Library and Security Cryptography. Retrieved from https://palletsprojects.com/p/werkzeug",
        "7. <b>FontAwesome Icons.</b> (n.d.). Modern Vector Icon Library. Retrieved from https://fontawesome.com",
        "8. <b>W3Schools.</b> (n.d.). HTML5, CSS3, and Modern JavaScript Web Tutorials. Retrieved from https://www.w3schools.com",
        "9. <b>University of Mumbai.</b> (2025). Bachelor of Science in Computer Science (B.Sc. CS) Semester V Project Guidelines and Formatting Manual.",
        "10. <b>Pressman, R. S.</b> (2014). Software Engineering: A Practitioner's Approach (8th ed.). McGraw-Hill Education."
    ]
    for r in refs:
        story.append(p(r, t_body))
        story.append(Spacer(1, 6))

    # Build the document
    doc.build(story, canvasmaker=NumberedCanvasWithBorder)
    print("PDF successfully generated at:", PDF_PATH)

if __name__ == '__main__':
    build_pdf()
