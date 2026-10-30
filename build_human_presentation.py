import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_human_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Academic / Human Corporate Palette
    NAVY_DARK = RGBColor(26, 54, 93)       # #1A365D (Deep Academic Navy)
    NAVY_LIGHT = RGBColor(43, 108, 176)    # #2B6CB0 (Corporate Blue)
    TEXT_DARK = RGBColor(45, 55, 72)       # #2D3748 (Body text)
    TEXT_MUTED = RGBColor(113, 128, 150)   # #718096 (Sub-text)
    BG_LIGHT = RGBColor(247, 250, 252)     # #F7FAFC (Light slate background)
    BOX_BG = RGBColor(255, 255, 255)       # White
    BOX_BORDER = RGBColor(226, 232, 240)   # Light border
    WHITE = RGBColor(255, 255, 255)
    GREEN = RGBColor(40, 167, 69)          # Professional green
    FONT_NAME = "Arial"

    def apply_footer(slide, current_page, total_pages=10):
        # Footer rule line
        rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.8), Inches(11.733), Inches(0.015))
        rule.fill.solid()
        rule.fill.fore_color.rgb = RGBColor(226, 232, 240)
        rule.line.fill.background()

        # Footer text left
        tb_left = slide.shapes.add_textbox(Inches(0.8), Inches(6.85), Inches(8.0), Inches(0.4))
        tf_l = tb_left.text_frame
        tf_l.margin_left = tf_l.margin_top = tf_l.margin_right = tf_l.margin_bottom = 0
        p_l = tf_l.paragraphs[0]
        p_l.text = "Employee Leave Management System | Project Presentation"
        p_l.font.name = FONT_NAME
        p_l.font.size = Pt(9.5)
        p_l.font.color.rgb = TEXT_MUTED

        # Footer text right (Slide Number)
        tb_right = slide.shapes.add_textbox(Inches(10.5), Inches(6.85), Inches(2.0), Inches(0.4))
        tf_r = tb_right.text_frame
        tf_r.margin_left = tf_r.margin_top = tf_r.margin_right = tf_r.margin_bottom = 0
        p_r = tf_r.paragraphs[0]
        p_r.text = f"Slide {current_page} of {total_pages}"
        p_r.font.name = FONT_NAME
        p_r.font.size = Pt(9.5)
        p_r.font.color.rgb = TEXT_MUTED
        p_r.alignment = PP_ALIGN.RIGHT

    def add_standard_header(slide, title_text, section_text):
        # Section Category Header
        sec_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(11.7), Inches(0.3))
        sec_tf = sec_box.text_frame
        sec_tf.margin_left = sec_tf.margin_top = sec_tf.margin_right = sec_tf.margin_bottom = 0
        sec_p = sec_tf.paragraphs[0]
        sec_p.text = section_text.upper()
        sec_p.font.name = FONT_NAME
        sec_p.font.size = Pt(10)
        sec_p.font.bold = True
        sec_p.font.color.rgb = NAVY_LIGHT

        # Main Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(11.7), Inches(0.6))
        tf = title_box.text_frame
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.name = FONT_NAME
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = NAVY_DARK

        # Top Accent Line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.45), Inches(11.733), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = NAVY_LIGHT
        line.line.fill.background()

    def create_box(slide, left, top, width, height, bg=BOX_BG, border=BOX_BORDER):
        box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        box.fill.solid()
        box.fill.fore_color.rgb = bg
        box.line.color.rgb = border
        box.line.width = Pt(1)
        return box

    # =========================================================================
    # SLIDE 1: Title Slide (Clean Human Academic Design)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY_DARK
    bg1.line.fill.background()

    # Top thin accent strip
    top_strip = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.08))
    top_strip.fill.solid()
    top_strip.fill.fore_color.rgb = RGBColor(66, 153, 225)
    top_strip.line.fill.background()

    # Main Title Block
    t_box = s1.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.3), Inches(2.2))
    tf1 = t_box.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.text = "FINAL YEAR PROJECT PRESENTATION"
    p0.font.name = FONT_NAME
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = RGBColor(190, 227, 248)

    p1 = tf1.add_paragraph()
    p1.text = "Employee Leave Management System"
    p1.font.name = FONT_NAME
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.space_before = Pt(8)

    p2 = tf1.add_paragraph()
    p2.text = "A Web-Based Portal for Automating Corporate Leave Applications and Approval Workflows"
    p2.font.name = FONT_NAME
    p2.font.size = Pt(16)
    p2.font.color.rgb = RGBColor(226, 232, 240)
    p2.space_before = Pt(8)

    # Dividing line on title slide
    d_line = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(3.9), Inches(11.3), Inches(0.02))
    d_line.fill.solid()
    d_line.fill.fore_color.rgb = RGBColor(74, 85, 104)
    d_line.line.fill.background()

    # Student / Project Details Container
    det_box = s1.shapes.add_textbox(Inches(1.0), Inches(4.2), Inches(11.3), Inches(2.3))
    dtf = det_box.text_frame
    dtf.word_wrap = True

    dp1 = dtf.paragraphs[0]
    dp1.text = "Submitted By:"
    dp1.font.name = FONT_NAME
    dp1.font.size = Pt(13)
    dp1.font.bold = True
    dp1.font.color.rgb = RGBColor(190, 227, 248)

    dp2 = dtf.add_paragraph()
    dp2.text = "Abhishek Yadav"
    dp2.font.name = FONT_NAME
    dp2.font.size = Pt(16)
    dp2.font.bold = True
    dp2.font.color.rgb = WHITE
    dp2.space_before = Pt(4)

    dp3 = dtf.add_paragraph()
    dp3.text = "Roll Number: 06    |    Degree Program: Bachelor of Technology (B.Tech)"
    dp3.font.name = FONT_NAME
    dp3.font.size = Pt(13)
    dp3.font.color.rgb = RGBColor(226, 232, 240)
    dp3.space_before = Pt(3)

    dp4 = dtf.add_paragraph()
    dp4.text = "Department of Computer Science & Engineering"
    dp4.font.name = FONT_NAME
    dp4.font.size = Pt(13)
    dp4.font.color.rgb = RGBColor(203, 213, 225)
    dp4.space_before = Pt(3)

    # =========================================================================
    # SLIDE 2: Introduction & Problem Statement
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_standard_header(s2, "Introduction & Problem Statement", "Project Background")
    apply_footer(s2, 2)

    # Left Box: Traditional System Challenges
    create_box(s2, Inches(0.8), Inches(1.7), Inches(5.6), Inches(4.8), bg=RGBColor(255, 255, 255))
    tb_l = s2.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(5.2), Inches(4.4))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True

    p = tf_l.paragraphs[0]
    p.text = "Limitations of Manual Leave Systems"
    p.font.name = FONT_NAME
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    points_l = [
        ("Paper-Based Approvals:", "Leave applications submitted on physical forms or unorganized emails result in delayed responses and lost records."),
        ("Inaccurate Day Calculations:", "Weekends and official holidays are frequently miscounted, leading to incorrect deductions from employee allowances."),
        ("Quota Discrepancies:", "Without centralized balance tracking, employees may overdraw their quotas before approval."),
        ("Lack of Audit Trail:", "Difficult for HR departments to maintain transparent records and generate compliance reports for management audits.")
    ]
    for title, desc in points_l:
        p_t = tf_l.add_paragraph()
        p_t.text = f"•  {title} {desc}"
        p_t.font.name = FONT_NAME
        p_t.font.size = Pt(12)
        p_t.font.color.rgb = TEXT_DARK
        p_t.space_before = Pt(10)

    # Right Box: Proposed Solution
    create_box(s2, Inches(6.8), Inches(1.7), Inches(5.7), Inches(4.8), bg=RGBColor(247, 250, 252))
    tb_r = s2.shapes.add_textbox(Inches(7.0), Inches(1.9), Inches(5.3), Inches(4.4))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True

    p = tf_r.paragraphs[0]
    p.text = "The Proposed Web Solution"
    p.font.name = FONT_NAME
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY_LIGHT

    points_r = [
        ("Centralized Web Portal:", "A single responsive portal accessible by employees and managers from any modern browser."),
        ("Automatic Calendar Engine:", "Working days are calculated automatically by excluding Saturdays, Sundays, and gazetted holidays in real-time."),
        ("Role-Based Dashboards:", "Dedicated interfaces for employees to track balances and for HR administrators to manage approvals."),
        ("Complete History & Export:", "Instant self-service cancellation, audit logging, and one-click export of leave records to CSV format.")
    ]
    for title, desc in points_r:
        p_t = tf_r.add_paragraph()
        p_t.text = f"•  {title} {desc}"
        p_t.font.name = FONT_NAME
        p_t.font.size = Pt(12)
        p_t.font.color.rgb = TEXT_DARK
        p_t.space_before = Pt(10)

    # =========================================================================
    # SLIDE 3: Objectives of the Project
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_standard_header(s3, "Objectives of the Project", "Project Goals")
    apply_footer(s3, 3)

    objectives = [
        ("1. End-to-End Workflow Automation",
         "Eliminate paper forms by providing a digital lifecycle: submission, balance verification, manager review, quota deduction, and historical archiving."),
        ("2. Accurate Working-Days Calculation",
         "Develop an intelligent calendar calculation algorithm that accurately skips weekends (Saturdays and Sundays) and gazetted corporate holidays."),
        ("3. Strict Role-Based Access Control",
         "Implement secure authentication separating standard employee actions (apply, view, cancel) from HR administrator rights (approve, reject, add staff, calibrate quotas)."),
        ("4. Real-Time Quota & Attendance Tracking",
         "Display live leave balance cards with dynamic progress meters, preventing duplicate applications or requests exceeding available balances."),
        ("5. Compliance Reporting and Analytics",
         "Provide management analytics on leave utilization trends and enable one-click export of full audit records to CSV files for HR compliance.")
    ]

    for i, (title, desc) in enumerate(objectives):
        top_pos = Inches(1.7 + i * 0.98)
        create_box(s3, Inches(0.8), top_pos, Inches(11.733), Inches(0.86), bg=WHITE)

        # Blue indicator bar on the left
        ind = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), top_pos, Inches(0.12), Inches(0.86))
        ind.fill.solid()
        ind.fill.fore_color.rgb = NAVY_LIGHT
        ind.line.fill.background()

        tb = s3.shapes.add_textbox(Inches(1.1), top_pos + Inches(0.08), Inches(11.2), Inches(0.7))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_NAME
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_DARK

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = FONT_NAME
        p2.font.size = Pt(11)
        p2.font.color.rgb = TEXT_DARK
        p2.space_before = Pt(2)

    # =========================================================================
    # SLIDE 4: System Architecture
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_standard_header(s4, "System Architecture & Data Flow", "System Design")
    apply_footer(s4, 4)

    # 3 Horizontal Cards representing the 3 Tiers
    tiers = [
        ("Presentation Layer (Client)", [
            "HTML5 & Tailwind CSS for modern responsive interface",
            "Jinja2 Template Engine rendering dynamic views",
            "Client-side JavaScript for live duration calculation",
            "Chart.js for interactive analytics graphs",
            "Compatible with desktop and mobile web browsers"
        ]),
        ("Application Layer (Flask Web Server)", [
            "Python 3.12 running Flask micro-framework",
            "Authentication & Session Management module",
            "Working days business logic engine (excluding holidays)",
            "Leave application validation & quota checks",
            "REST API endpoints (/api/calculate-days, /api/chart-data)"
        ]),
        ("Database Layer (Persistence)", [
            "SQLite3 relational database engine",
            "Referential integrity enforced with Foreign Key PRAGMA",
            "Transactional safety (commit / rollback)",
            "Encrypted password storage using PBKDF2 SHA-256",
            "Automated database seeding on initial launch"
        ])
    ]

    for i, (tier_title, items) in enumerate(tiers):
        left_pos = Inches(0.8 + i * 4.0)
        create_box(s4, left_pos, Inches(1.7), Inches(3.8), Inches(4.8), bg=WHITE)

        # Header bar for each tier
        hbar = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, Inches(1.7), Inches(3.8), Inches(0.6))
        hbar.fill.solid()
        hbar.fill.fore_color.rgb = NAVY_DARK
        hbar.line.fill.background()

        h_tb = s4.shapes.add_textbox(left_pos, Inches(1.75), Inches(3.8), Inches(0.5))
        h_tf = h_tb.text_frame
        h_p = h_tf.paragraphs[0]
        h_p.text = tier_title
        h_p.font.name = FONT_NAME
        h_p.font.size = Pt(13)
        h_p.font.bold = True
        h_p.font.color.rgb = WHITE
        h_p.alignment = PP_ALIGN.CENTER

        tb = s4.shapes.add_textbox(left_pos + Inches(0.2), Inches(2.45), Inches(3.4), Inches(3.9))
        tf = tb.text_frame
        tf.word_wrap = True

        for k, item in enumerate(items):
            p = tf.add_paragraph() if k > 0 else tf.paragraphs[0]
            p.text = f"•  {item}"
            p.font.name = FONT_NAME
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_DARK
            p.space_before = Pt(8) if k > 0 else Pt(0)

    # =========================================================================
    # SLIDE 5: Technology Stack
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_standard_header(s5, "Technology Stack & Tools Used", "Technical Specifications")
    apply_footer(s5, 5)

    tech_categories = [
        ("Backend Technologies", [
            ("Python 3.12", "Core server programming language chosen for stability and extensive libraries."),
            ("Flask 3.0", "WSGI web application framework handling routes, sessions, and request dispatching."),
            ("Werkzeug Security", "Provides cryptographic password hashing using PBKDF2 SHA-256 algorithm."),
            ("Gunicorn", "Production-grade UNIX WSGI HTTP server used for cloud deployment on Render.")
        ]),
        ("Frontend & User Interface", [
            ("HTML5 & Jinja2", "Semantic markup and dynamic server-side template rendering."),
            ("Tailwind CSS", "Utility-first CSS framework ensuring responsive, clean, and modern layout."),
            ("JavaScript (ES6)", "Asynchronous DOM manipulation, live calculation calls, and modal interactions."),
            ("Chart.js", "Canvas-based charting library for visualizing leave trends and category usage.")
        ]),
        ("Database & DevOps", [
            ("SQLite3 Engine", "Lightweight, zero-configuration relational database with full ACID compliance."),
            ("Python unittest", "Automated unit test framework used for comprehensive quality assurance."),
            ("Git & GitHub", "Distributed version control system for tracking changes and repository hosting."),
            ("Render Cloud", "Cloud hosting platform providing continuous deployment directly from GitHub.")
        ])
    ]

    for i, (cat_title, items) in enumerate(tech_categories):
        left_pos = Inches(0.8 + i * 4.0)
        create_box(s5, left_pos, Inches(1.7), Inches(3.8), Inches(4.8), bg=WHITE)

        # Header box
        hdr = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, Inches(1.7), Inches(3.8), Inches(0.55))
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = NAVY_LIGHT
        hdr.line.fill.background()

        h_tb = s5.shapes.add_textbox(left_pos, Inches(1.75), Inches(3.8), Inches(0.45))
        h_tf = h_tb.text_frame
        h_p = h_tf.paragraphs[0]
        h_p.text = cat_title
        h_p.font.name = FONT_NAME
        h_p.font.size = Pt(13)
        h_p.font.bold = True
        h_p.font.color.rgb = WHITE
        h_p.alignment = PP_ALIGN.CENTER

        tb = s5.shapes.add_textbox(left_pos + Inches(0.2), Inches(2.4), Inches(3.4), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True

        for k, (name, desc) in enumerate(items):
            p = tf.add_paragraph() if k > 0 else tf.paragraphs[0]
            p.text = name
            p.font.name = FONT_NAME
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = NAVY_DARK
            p.space_before = Pt(8) if k > 0 else Pt(0)

            p_sub = tf.add_paragraph()
            p_sub.text = desc
            p_sub.font.name = FONT_NAME
            p_sub.font.size = Pt(10.5)
            p_sub.font.color.rgb = TEXT_DARK
            p_sub.space_before = Pt(2)

    # =========================================================================
    # SLIDE 6: Database Design (Schema)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_standard_header(s6, "Database Schema & Entity Relationships", "Data Architecture")
    apply_footer(s6, 6)

    tables_info = [
        ("users", "Stores staff information", [
            "id (INTEGER PK), emp_code (UNIQUE), full_name, email (UNIQUE)",
            "password_hash (PBKDF2), role ('admin' / 'employee')",
            "department, designation, join_date, is_active"
        ]),
        ("leave_types", "Defines leave categories", [
            "id (INTEGER PK), name (Annual, Sick, Casual, etc.)",
            "code (AL, SL, CL, ML, PL, UL), default_days, color_code"
        ]),
        ("leave_balances", "Tracks annual allowances per employee", [
            "id (INTEGER PK), user_id (FK -> users), leave_type_id (FK -> leave_types)",
            "year (INTEGER), total_allocated, used_days, pending_days"
        ]),
        ("leave_applications", "Stores submitted leave requests", [
            "id (INTEGER PK), user_id (FK), leave_type_id (FK), start_date, end_date",
            "total_days, is_half_day, reason, status ('pending', 'approved', 'rejected', 'cancelled')",
            "reviewed_by (FK), reviewed_at, manager_remarks"
        ]),
        ("holidays", "Official company calendar", [
            "id (INTEGER PK), name, holiday_date (YYYY-MM-DD), description"
        ])
    ]

    for i, (t_name, t_purpose, t_cols) in enumerate(tables_info):
        top_pos = Inches(1.7 + i * 0.98)
        create_box(s6, Inches(0.8), top_pos, Inches(11.733), Inches(0.88), bg=WHITE)

        # Table badge
        t_badge = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.9), top_pos + Inches(0.12), Inches(2.2), Inches(0.62))
        t_badge.fill.solid()
        t_badge.fill.fore_color.rgb = RGBColor(237, 242, 247)
        t_badge.line.color.rgb = RGBColor(203, 213, 225)
        tb_tf = t_badge.text_frame
        tb_p = tb_tf.paragraphs[0]
        tb_p.text = t_name
        tb_p.font.name = FONT_NAME
        tb_p.font.size = Pt(12)
        tb_p.font.bold = True
        tb_p.font.color.rgb = NAVY_DARK
        tb_p.alignment = PP_ALIGN.CENTER

        tb = s6.shapes.add_textbox(Inches(3.3), top_pos + Inches(0.08), Inches(9.0), Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = f"Purpose: {t_purpose}"
        p1.font.name = FONT_NAME
        p1.font.size = Pt(11.5)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_LIGHT

        for col_line in t_cols:
            p2 = tf.add_paragraph()
            p2.text = f"Fields: {col_line}"
            p2.font.name = FONT_NAME
            p2.font.size = Pt(10.5)
            p2.font.color.rgb = TEXT_DARK
            p2.space_before = Pt(2)

    # =========================================================================
    # SLIDE 7: Core Modules & Functionalities
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_standard_header(s7, "Core Modules & System Functionalities", "System Features")
    apply_footer(s7, 7)

    # Left: Employee Module
    create_box(s7, Inches(0.8), Inches(1.7), Inches(5.6), Inches(4.8), bg=WHITE)
    tb_m1 = s7.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.2), Inches(4.5))
    tf_m1 = tb_m1.text_frame
    tf_m1.word_wrap = True

    p = tf_m1.paragraphs[0]
    p.text = "1. Employee Self-Service Module"
    p.font.name = FONT_NAME
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    emp_features = [
        ("Personal Dashboard:", "Overview of annual quota, total used days, pending review days, and available balance."),
        ("Leave Application Form:", "Interactive form with dropdown categories, date pickers, and reason text area."),
        ("Dynamic Day Calculator:", "Calculates working days instantly via AJAX before submission."),
        ("Half-Day Option:", "Checkbox sets duration to 0.5 days on start date with single date constraint."),
        ("Application History:", "Tabular display of all submitted requests with status badges and manager remarks."),
        ("Self-Service Cancellation:", "Allows withdrawing pending requests, immediately releasing the quota hold.")
    ]
    for title, desc in emp_features:
        p_t = tf_m1.add_paragraph()
        p_t.text = f"•  {title} {desc}"
        p_t.font.name = FONT_NAME
        p_t.font.size = Pt(11)
        p_t.font.color.rgb = TEXT_DARK
        p_t.space_before = Pt(6)

    # Right: Admin Module
    create_box(s7, Inches(6.8), Inches(1.7), Inches(5.7), Inches(4.8), bg=WHITE)
    tb_m2 = s7.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.5))
    tf_m2 = tb_m2.text_frame
    tf_m2.word_wrap = True

    p = tf_m2.paragraphs[0]
    p.text = "2. HR Administrator & Manager Module"
    p.font.name = FONT_NAME
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY_LIGHT

    admin_features = [
        ("Executive KPI Dashboard:", "Real-time indicators showing active employees, pending review count, and employees on leave today."),
        ("Approvals Queue:", "List of pending applications showing applicant profile, requested dates, and current balance."),
        ("Modal Decision Workflows:", "Approve or reject requests with mandatory manager feedback remarks."),
        ("Employee Directory:", "Register new employees with automated initial leave balance creation."),
        ("Quota Calibration:", "Adjust annual quota allocation per employee and category as needed."),
        ("Reporting & CSV Export:", "Filter leave history by department, status, or type; export compliance records to CSV.")
    ]
    for title, desc in admin_features:
        p_t = tf_m2.add_paragraph()
        p_t.text = f"•  {title} {desc}"
        p_t.font.name = FONT_NAME
        p_t.font.size = Pt(11)
        p_t.font.color.rgb = TEXT_DARK
        p_t.space_before = Pt(6)

    # =========================================================================
    # SLIDE 8: Key Business Logic & Validation Rules
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_standard_header(s8, "Key Business Logic & Validation Rules", "System Implementation")
    apply_footer(s8, 8)

    rules = [
        ("1. Business Day Calculation Algorithm",
         "The system iterates through each calendar day in the requested date range. If a day falls on Saturday (weekday 5) or Sunday (weekday 6) or matches an official public holiday date in the database, it is automatically excluded. Only net business working days are counted."),
        ("2. Pending Hold Mechanism",
         "When an employee applies for leave, the requested days are immediately added to 'pending_days' in their balance record. This temporarily locks the quota to prevent concurrent double-booking. Upon approval, pending days move to 'used_days'; if rejected or cancelled, pending days are released."),
        ("3. Date Overlap and Conflict Detection",
         "Before saving an application, the system queries existing pending and approved records for the same employee. If the requested date range overlaps with any existing booking, the application is rejected with a descriptive warning message."),
        ("4. Balance Sufficiency Verification",
         "Applications requesting more business days than the remaining available quota (total_allocated - used_days - pending_days) are automatically rejected prior to database insertion."),
        ("5. Security & Input Sanitization",
         "Strict role-based authorization decorators (@login_required, @admin_required) prevent privilege escalation. Form inputs and parameters use SQL query parameter binding to eliminate SQL injection vulnerabilities.")
    ]

    for i, (title, desc) in enumerate(rules):
        top_pos = Inches(1.7 + i * 0.98)
        create_box(s8, Inches(0.8), top_pos, Inches(11.733), Inches(0.88), bg=WHITE)

        tb = s8.shapes.add_textbox(Inches(1.0), top_pos + Inches(0.08), Inches(11.3), Inches(0.75))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = FONT_NAME
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = NAVY_DARK

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.name = FONT_NAME
        p2.font.size = Pt(11)
        p2.font.color.rgb = TEXT_DARK
        p2.space_before = Pt(3)

    # =========================================================================
    # SLIDE 9: Testing & Quality Assurance
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_standard_header(s9, "Testing Methodology & Verification Results", "Quality Assurance")
    apply_footer(s9, 9)

    # Summary Result Box
    res_box = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.7), Inches(11.733), Inches(1.0))
    res_box.fill.solid()
    res_box.fill.fore_color.rgb = RGBColor(240, 253, 244)
    res_box.line.color.rgb = RGBColor(187, 247, 208)

    r_tb = s9.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.3), Inches(0.8))
    r_tf = r_tb.text_frame
    rp1 = r_tf.paragraphs[0]
    rp1.text = "Automated Unit Test Results: 17 Passed / 0 Failed (100% Pass Rate)"
    rp1.font.name = FONT_NAME
    rp1.font.size = Pt(16)
    rp1.font.bold = True
    rp1.font.color.rgb = GREEN

    rp2 = r_tf.add_paragraph()
    rp2.text = "Executed using Python's standard unittest test runner against isolated test databases. Execution time: 4.44 seconds."
    rp2.font.name = FONT_NAME
    rp2.font.size = Pt(11.5)
    rp2.font.color.rgb = TEXT_DARK
    rp2.space_before = Pt(3)

    # 3 Test Groups
    test_groups = [
        ("Calculation Unit Tests", [
            ("Regular Day Counting:", "Mon-Wed correctly computes 3.0 days."),
            ("Weekend Exclusion:", "Fri-Mon correctly computes 2.0 days."),
            ("Holiday Exclusion:", "Public holidays excluded from total count."),
            ("Half-Day Calculation:", "Single date half-day verified as 0.5 days."),
            ("Date Sequence Guard:", "End date earlier than start date rejected.")
        ]),
        ("Workflow Integration Tests", [
            ("Apply & Pending Hold:", "Application places instant hold on quota."),
            ("Manager Approval:", "Transfers pending hold to used days accurately."),
            ("Manager Rejection:", "Releases pending hold with zero quota deduction."),
            ("Self-Service Cancel:", "Cancelled pending leaves restore balance."),
            ("Insufficient Balance:", "Rejects request when quota is exceeded.")
        ]),
        ("Security & Validation Tests", [
            ("Authentication Check:", "Valid login session creation verified."),
            ("Role Separation:", "Prevents standard users accessing admin portal."),
            ("Overlap Prevention:", "Blocks concurrent overlapping leave dates."),
            ("Duplicate Email Check:", "Prevents multiple accounts with same email."),
            ("Quota Threshold Guard:", "Rejects negative quota allocation values.")
        ])
    ]

    for i, (g_title, g_items) in enumerate(test_groups):
        left_pos = Inches(0.8 + i * 4.0)
        create_box(s9, left_pos, Inches(2.9), Inches(3.8), Inches(3.6), bg=WHITE)

        h = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_pos, Inches(2.9), Inches(3.8), Inches(0.5))
        h.fill.solid()
        h.fill.fore_color.rgb = NAVY_LIGHT
        h.line.fill.background()

        h_tb = s9.shapes.add_textbox(left_pos, Inches(2.95), Inches(3.8), Inches(0.4))
        h_tf = h_tb.text_frame
        hp = h_tf.paragraphs[0]
        hp.text = g_title
        hp.font.name = FONT_NAME
        hp.font.size = Pt(12)
        hp.font.bold = True
        hp.font.color.rgb = WHITE
        hp.alignment = PP_ALIGN.CENTER

        tb = s9.shapes.add_textbox(left_pos + Inches(0.15), Inches(3.5), Inches(3.5), Inches(2.9))
        tf = tb.text_frame
        tf.word_wrap = True

        for k, (iname, idesc) in enumerate(g_items):
            p = tf.add_paragraph() if k > 0 else tf.paragraphs[0]
            p.text = f"• {iname} {idesc}"
            p.font.name = FONT_NAME
            p.font.size = Pt(9.8)
            p.font.color.rgb = TEXT_DARK
            p.space_before = Pt(4) if k > 0 else Pt(0)

    # =========================================================================
    # SLIDE 10: Conclusion & Future Scope
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_standard_header(s10, "Conclusion & Future Enhancements", "Project Summary")
    apply_footer(s10, 10)

    # Left: Conclusion Summary
    create_box(s10, Inches(0.8), Inches(1.7), Inches(5.6), Inches(4.8), bg=WHITE)
    tb_c = s10.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.2), Inches(4.5))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True

    p = tf_c.paragraphs[0]
    p.text = "Conclusion & Summary"
    p.font.name = FONT_NAME
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY_DARK

    conclusions = [
        ("Automated Administrative Process:", "The system successfully replaces paper-based and email-based leave tracking with an automated, responsive web application."),
        ("Elimination of Calculation Errors:", "Smart business-days engine accurately excludes non-working days, preventing human counting errors."),
        ("Improved Transparency:", "Both employees and managers have real-time visibility over balances, request statuses, and historical records."),
        ("Production Cloud Readiness:", "Configured with Gunicorn WSGI and deployed to cloud hosting with continuous GitHub deployment.")
    ]
    for title, desc in conclusions:
        p_t = tf_c.add_paragraph()
        p_t.text = f"•  {title} {desc}"
        p_t.font.name = FONT_NAME
        p_t.font.size = Pt(11.5)
        p_t.font.color.rgb = TEXT_DARK
        p_t.space_before = Pt(10)

    # Right: Future Scope
    create_box(s10, Inches(6.8), Inches(1.7), Inches(5.7), Inches(4.8), bg=WHITE)
    tb_fs = s10.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.5))
    tf_fs = tb_fs.text_frame
    tf_fs.word_wrap = True

    p = tf_fs.paragraphs[0]
    p.text = "Future Scope & Enhancements"
    p.font.name = FONT_NAME
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY_LIGHT

    future_items = [
        ("Email and SMS Notifications:", "Automated email alerts sent to managers upon new requests and to employees upon status updates."),
        ("Calendar Synchronization:", "Integration with Google Calendar and Microsoft Outlook for synchronized team availability."),
        ("Biometric Attendance Integration:", "Connecting leave records with physical office punch clocks or biometric attendance systems."),
        ("Mobile Application:", "Developing dedicated Android and iOS companion apps using Flutter or React Native.")
    ]
    for title, desc in future_items:
        p_t = tf_fs.add_paragraph()
        p_t.text = f"•  {title} {desc}"
        p_t.font.name = FONT_NAME
        p_t.font.size = Pt(11.5)
        p_t.font.color.rgb = TEXT_DARK
        p_t.space_before = Pt(10)

    # Save presentation
    output_filename = "Employee_Leave_Management_System_Presentation.pptx"
    prs.save(output_filename)
    print(f"Human-styled presentation saved successfully as {output_filename}")

if __name__ == '__main__':
    create_human_presentation()
