# LeaveFlow — Enterprise Employee Leave Management System (ELMS)

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
![Flask](https://img.shields.io/badge/Flask-3.0%2B-black?logo=flask)
![SQLite](https://img.shields.io/badge/Database-SQLite-003B57?logo=sqlite)
![TailwindCSS](https://img.shields.io/badge/Frontend-TailwindCSS-06B6D4?logo=tailwindcss)
![Tests](https://img.shields.io/badge/Tests-17%20Passed%20(100%25)-brightgreen)
![Status](https://img.shields.io/badge/Status-Production%20Ready-success)

A modern, full-stack, enterprise-grade Employee Leave Management System built with Python, Flask, SQLite, and Tailwind CSS. **LeaveFlow** automates the entire lifecycle of corporate time-off requests, approval workflows, dynamic quota calculations, real-time analytics, and compliance audit tracking.

---

## 🌟 Key Features

### 🏢 Employee Self-Service Portal
- **Smart Date Calculation Engine:** Automatically calculates business working days while excluding weekends (Saturdays & Sundays) and official gazetted company holidays.
- **Half-Day Support:** Built-in 0.5-day leave support with automatic date synchronization.
- **Real-Time Quota Snapshot:** Live progress bars and remaining balance alerts before submitting requests to prevent over-allocation.
- **Application History & Cancellation:** Full visibility over pending, approved, and rejected requests with self-service cancellation that instantly restores quotas.
- **Holidays Calendar:** Integrated view of upcoming company and public holidays.

### 🛡️ HR & Management Admin Portal
- **Approvals Queue:** Fast-track approval/rejection modal workflows with custom manager feedback remarks.
- **Employee Directory:** Manage corporate staff, departments, designations, and roles.
- **Quota Adjustments:** Custom allocation of annual leave allowances per employee and leave category.
- **Interactive Analytics:** Real-time Chart.js visual charts displaying leave distribution by type and monthly utilization trends.
- **Audit Compliance & CSV Export:** Multi-filter query engine (by department, status, and category) with one-click export to audit-ready CSV reports.

---

## 🛠️ Technology Stack

- **Backend:** Python, Flask, Werkzeug Security
- **Database:** SQLite3 with foreign key enforcement and transactional integrity
- **Frontend:** Jinja2 Templates, HTML5, Tailwind CSS, FontAwesome 6
- **Data Visualization:** Chart.js
- **Testing:** Python `unittest` (17 automated unit and integration tests)

---

## 🚀 Quick Start Guide

### 1. Clone the Repository
```bash
git clone https://github.com/YOUR_USERNAME/employeeleavemanagementsystem.git
cd employeeleavemanagementsystem
```

### 2. Set Up a Virtual Environment (Optional but Recommended)
```bash
# On Windows
python -m venv venv
venv\Scripts\activate

# On macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the Application
```bash
python app.py
```

Open your browser and navigate to:
```
http://127.0.0.1:5000
```
> **Note:** The SQLite database (`elms.db`) is automatically initialized and pre-seeded with enterprise demo employees, leave balances, and holidays on first startup!

---

## 🔑 Pre-Configured Demo Credentials

The login screen includes **1-Click Demo Profile buttons** for instant access:

| Role | Employee Name | Email | Password |
|---|---|---|---|
| **HR Administrator** | Alexandra Vance | `admin@company.com` | `admin123` |
| **Software Engineer** | John Doe | `john.doe@company.com` | `employee123` |
| **UI/UX Designer** | Sarah Jenkins | `sarah.j@company.com` | `employee123` |
| **DevOps Specialist** | Michael Chang | `michael.c@company.com` | `employee123` |

---

## 🧪 Running Automated Tests

Run the comprehensive unit test suite:
```bash
python -m unittest discover tests
```

Expected output:
```text
----------------------------------------------------------------------
Ran 17 tests in 4.440s

OK
```

---

## 📂 Project Structure

```text
employeeleavemanagementsystem/
├── app.py                     # Core Flask application & API route handlers
├── database.py                # Database connection, schemas, and working-day logic
├── seed_data.py               # Enterprise pre-seed data generator
├── requirements.txt           # Python package dependencies
├── README.md                  # Project documentation
├── .gitignore                 # Git ignore configuration
├── static/
│   ├── css/custom.css         # Styling enhancements & custom scrollbars
│   └── js/
│       ├── app.js             # Date duration calculator & modal handlers
│       └── charts.js          # Chart.js analytics configuration
├── templates/
│   ├── base.html              # Base layout with sidebar navigation & header
│   ├── login.html             # Login portal with 1-click demo profiles
│   ├── employee/
│   │   ├── dashboard.html     # Employee summary & quota balances
│   │   ├── apply.html         # Leave request submission form
│   │   ├── history.html       # Application history & cancellation
│   │   └── holidays.html      # Public holiday schedule
│   └── admin/
│       ├── dashboard.html     # Admin KPIs & attendance summary
│       ├── approvals.html     # Leave approval/rejection queue
│       ├── employees.html     # Employee directory & quota management
│       └── reports.html       # Compliance search & CSV export
├── tests/
│   └── test_system.py         # Automated unit & integration test suite
└── report_assets/             # UI screenshots and UML design diagrams
```

---

## 📄 License
This project is open-source and available under the [MIT License](LICENSE).
