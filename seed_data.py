import datetime
from werkzeug.security import generate_password_hash
from database import get_db_connection, init_db

def seed_database():
    init_db()
    conn = get_db_connection()
    cursor = conn.cursor()

    # Clear existing demo data cleanly if requested or if users table is empty
    user_count = cursor.execute("SELECT COUNT(*) as count FROM users").fetchone()['count']
    if user_count > 0:
        print("Database already contains data. Skipping re-seed.")
        conn.close()
        return

    print("Seeding database with professional enterprise demo data...")

    # 1. Insert Leave Types
    leave_types_data = [
        ("Annual Leave", "AL", 18.0, "#3B82F6", "Standard paid annual vacation leave", 0),
        ("Sick Leave", "SL", 12.0, "#EF4444", "Medical leave for health recovery and medical appointments", 1),
        ("Casual Leave", "CL", 10.0, "#10B981", "Short-notice leaves for urgent personal and family affairs", 0),
        ("Parental Leave", "PL", 15.0, "#8B5CF6", "Paid parental care and family bonding leave", 1),
        ("Bereavement Leave", "BL", 5.0, "#64748B", "Compassionate leave for family emergencies", 0),
    ]
    cursor.executemany('''
        INSERT INTO leave_types (name, code, default_days, color, description, requires_document)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', leave_types_data)

    # 2. Insert Holidays for current and upcoming seasons
    holidays_data = [
        ("New Year's Day", "2026-01-01", "Global Public Holiday"),
        ("Republic Day", "2026-01-26", "National Holiday"),
        ("Holi Festival", "2026-03-04", "Spring Cultural Festival"),
        ("Good Friday", "2026-04-03", "Public Holiday"),
        ("Labor Day", "2026-05-01", "International Workers' Day"),
        ("Independence Day", "2026-08-15", "National Holiday"),
        ("Gandhi Jayanti", "2026-10-02", "National Holiday"),
        ("Diwali / Festival of Lights", "2026-11-08", "Cultural Holiday"),
        ("Christmas Day", "2026-12-25", "Public Holiday"),
    ]
    cursor.executemany('''
        INSERT INTO holidays (name, holiday_date, description)
        VALUES (?, ?, ?)
    ''', holidays_data)

    # 3. Insert Users (Admin + 3 Employees)
    current_year = datetime.datetime.now().year
    users_data = [
        (
            "EMP-1000",
            "Alexandra Vance",
            "admin@company.com",
            generate_password_hash("admin123"),
            "admin",
            "People Operations & HR",
            "Head of Human Resources",
            "2023-01-10",
            "#6366F1"
        ),
        (
            "EMP-1001",
            "John Doe",
            "john.doe@company.com",
            generate_password_hash("employee123"),
            "employee",
            "Engineering",
            "Senior Full Stack Engineer",
            "2023-03-15",
            "#2563EB"
        ),
        (
            "EMP-1002",
            "Sarah Jenkins",
            "sarah.j@company.com",
            generate_password_hash("employee123"),
            "employee",
            "Product & Design",
            "Lead UI/UX Designer",
            "2023-07-01",
            "#EC4899"
        ),
        (
            "EMP-1003",
            "Michael Chang",
            "michael.c@company.com",
            generate_password_hash("employee123"),
            "employee",
            "Marketing & Growth",
            "Senior Growth Manager",
            "2024-01-12",
            "#F59E0B"
        ),
    ]
    cursor.executemany('''
        INSERT INTO users (emp_code, full_name, email, password_hash, role, department, designation, join_date, avatar_color)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', users_data)

    # 4. Fetch created users and leave types to create balances
    users = cursor.execute("SELECT id, role, full_name FROM users").fetchall()
    leave_types = cursor.execute("SELECT id, code, default_days FROM leave_types").fetchall()

    for u in users:
        for lt in leave_types:
            cursor.execute('''
                INSERT INTO leave_balances (user_id, leave_type_id, year, total_allocated, used_days, pending_days)
                VALUES (?, ?, ?, ?, 0.0, 0.0)
            ''', (u['id'], lt['id'], current_year, lt['default_days']))

    # 5. Insert Sample Leave Applications with realistic states
    # Find user IDs
    john_id = cursor.execute("SELECT id FROM users WHERE email='john.doe@company.com'").fetchone()['id']
    sarah_id = cursor.execute("SELECT id FROM users WHERE email='sarah.j@company.com'").fetchone()['id']
    michael_id = cursor.execute("SELECT id FROM users WHERE email='michael.c@company.com'").fetchone()['id']
    admin_id = cursor.execute("SELECT id FROM users WHERE email='admin@company.com'").fetchone()['id']

    al_id = cursor.execute("SELECT id FROM leave_types WHERE code='AL'").fetchone()['id']
    sl_id = cursor.execute("SELECT id FROM leave_types WHERE code='SL'").fetchone()['id']
    cl_id = cursor.execute("SELECT id FROM leave_types WHERE code='CL'").fetchone()['id']

    # Past Approved Leave for John Doe (Annual Leave 3 days)
    cursor.execute('''
        INSERT INTO leave_applications (user_id, leave_type_id, start_date, end_date, total_days, is_half_day, reason, status, applied_at, reviewed_by, reviewed_at, manager_remarks)
        VALUES (?, ?, '2026-08-10', '2026-08-12', 3.0, 0, 'Family vacation trip', 'approved', '2026-08-01 10:30:00', ?, '2026-08-02 11:00:00', 'Approved. Enjoy your break!')
    ''', (john_id, al_id, admin_id))
    cursor.execute('UPDATE leave_balances SET used_days = used_days + 3.0 WHERE user_id = ? AND leave_type_id = ? AND year = ?', (john_id, al_id, current_year))

    # Past Approved Sick Leave for John Doe (1 day)
    cursor.execute('''
        INSERT INTO leave_applications (user_id, leave_type_id, start_date, end_date, total_days, is_half_day, reason, status, applied_at, reviewed_by, reviewed_at, manager_remarks)
        VALUES (?, ?, '2026-09-07', '2026-09-07', 1.0, 0, 'Acute viral fever and doctor visit', 'approved', '2026-09-07 08:15:00', ?, '2026-09-07 09:00:00', 'Get well soon.')
    ''', (john_id, sl_id, admin_id))
    cursor.execute('UPDATE leave_balances SET used_days = used_days + 1.0 WHERE user_id = ? AND leave_type_id = ? AND year = ?', (john_id, sl_id, current_year))

    # Pending Casual Leave for John Doe (2 days)
    cursor.execute('''
        INSERT INTO leave_applications (user_id, leave_type_id, start_date, end_date, total_days, is_half_day, reason, status, applied_at)
        VALUES (?, ?, '2026-10-15', '2026-10-16', 2.0, 0, 'Attending brother wedding reception', 'pending', datetime('now', '-2 hours'))
    ''', (john_id, cl_id))
    cursor.execute('UPDATE leave_balances SET pending_days = pending_days + 2.0 WHERE user_id = ? AND leave_type_id = ? AND year = ?', (john_id, cl_id, current_year))

    # Pending Sick Leave for Sarah Jenkins (1 day, Half-day)
    cursor.execute('''
        INSERT INTO leave_applications (user_id, leave_type_id, start_date, end_date, total_days, is_half_day, reason, status, applied_at)
        VALUES (?, ?, '2026-10-14', '2026-10-14', 0.5, 1, 'Dental clinic appointment in afternoon', 'pending', datetime('now', '-4 hours'))
    ''', (sarah_id, sl_id))
    cursor.execute('UPDATE leave_balances SET pending_days = pending_days + 0.5 WHERE user_id = ? AND leave_type_id = ? AND year = ?', (sarah_id, sl_id, current_year))

    # Past Rejected Leave for Michael Chang (Casual Leave)
    cursor.execute('''
        INSERT INTO leave_applications (user_id, leave_type_id, start_date, end_date, total_days, is_half_day, reason, status, applied_at, reviewed_by, reviewed_at, manager_remarks)
        VALUES (?, ?, '2026-09-20', '2026-09-22', 2.0, 0, 'Personal road trip', 'rejected', '2026-09-18 09:00:00', ?, '2026-09-18 16:30:00', 'Sprint release milestone scheduled during these dates. Please reschedule.')
    ''', (michael_id, cl_id, admin_id))

    # Pending Annual Leave for Michael Chang (3 days)
    cursor.execute('''
        INSERT INTO leave_applications (user_id, leave_type_id, start_date, end_date, total_days, is_half_day, reason, status, applied_at)
        VALUES (?, ?, '2026-11-02', '2026-11-04', 3.0, 0, 'Post-campaign rest and travel', 'pending', datetime('now', '-1 day'))
    ''', (michael_id, al_id))
    cursor.execute('UPDATE leave_balances SET pending_days = pending_days + 3.0 WHERE user_id = ? AND leave_type_id = ? AND year = ?', (michael_id, al_id, current_year))

    # Audit logs
    cursor.execute('''
        INSERT INTO audit_logs (user_id, action, details)
        VALUES (?, 'SYSTEM_INIT', 'System seeded with default departments, leave types, and demo personnel.')
    ''', (admin_id,))

    conn.commit()
    conn.close()
    print("Database seeding completed successfully!")

if __name__ == '__main__':
    seed_database()
