import sqlite3
import os
from datetime import datetime, timedelta

def get_db_path():
    return os.environ.get('ELMS_DB_PATH', os.path.join(os.path.dirname(__file__), 'elms.db'))

def get_db_connection():
    conn = sqlite3.connect(get_db_path())
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()

    # Users table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_code TEXT UNIQUE NOT NULL,
            full_name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin', 'employee')),
            department TEXT NOT NULL,
            designation TEXT NOT NULL,
            join_date TEXT NOT NULL,
            avatar_color TEXT DEFAULT '#4F46E5',
            is_active INTEGER DEFAULT 1,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Leave types table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS leave_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            code TEXT UNIQUE NOT NULL,
            default_days REAL NOT NULL,
            color TEXT NOT NULL,
            description TEXT,
            requires_document INTEGER DEFAULT 0
        )
    ''')

    # Leave balances table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS leave_balances (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            leave_type_id INTEGER NOT NULL REFERENCES leave_types(id) ON DELETE CASCADE,
            year INTEGER NOT NULL,
            total_allocated REAL NOT NULL,
            used_days REAL DEFAULT 0.0,
            pending_days REAL DEFAULT 0.0,
            UNIQUE(user_id, leave_type_id, year)
        )
    ''')

    # Leave applications table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS leave_applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            leave_type_id INTEGER NOT NULL REFERENCES leave_types(id),
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            total_days REAL NOT NULL,
            is_half_day INTEGER DEFAULT 0,
            reason TEXT NOT NULL,
            status TEXT DEFAULT 'pending' CHECK(status IN ('pending', 'approved', 'rejected', 'cancelled')),
            applied_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            reviewed_by INTEGER REFERENCES users(id),
            reviewed_at TIMESTAMP,
            manager_remarks TEXT
        )
    ''')

    # Holidays table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS holidays (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            holiday_date TEXT NOT NULL,
            description TEXT
        )
    ''')

    # Audit logs table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER REFERENCES users(id),
            action TEXT NOT NULL,
            details TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    conn.commit()
    conn.close()

def log_audit(user_id, action, details):
    try:
        conn = get_db_connection()
        conn.execute(
            'INSERT INTO audit_logs (user_id, action, details) VALUES (?, ?, ?)',
            (user_id, action, details)
        )
        conn.commit()
        conn.close()
    except Exception:
        pass

def get_holiday_dates():
    conn = get_db_connection()
    holidays = conn.execute('SELECT holiday_date FROM holidays').fetchall()
    conn.close()
    return {h['holiday_date'] for h in holidays}

def calculate_working_days(start_date_str, end_date_str, is_half_day=False):
    """
    Calculates number of working days between start_date and end_date (inclusive).
    Excludes Saturdays (5), Sundays (6), and public holidays in the database.
    """
    try:
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
    except ValueError:
        return 0, "Invalid date format. Expected YYYY-MM-DD."

    if end_date < start_date:
        return 0, "End date cannot be earlier than start date."

    if is_half_day:
        if start_date != end_date:
            return 0, "Half-day leave must start and end on the same day."
        # If the half day is a weekend or holiday
        holiday_dates = get_holiday_dates()
        if start_date.weekday() in (5, 6):
            return 0, "The selected date falls on a weekend."
        if start_date_str in holiday_dates:
            return 0, "The selected date is a company holiday."
        return 0.5, None

    holiday_dates = get_holiday_dates()
    working_days = 0
    current_date = start_date

    while current_date <= end_date:
        date_str = current_date.strftime('%Y-%m-%d')
        # Skip Saturday (5) and Sunday (6)
        if current_date.weekday() not in (5, 6) and date_str not in holiday_dates:
            working_days += 1
        current_date += timedelta(days=1)

    if working_days == 0:
        return 0, "No business working days in the selected date range (all days are weekends or holidays)."

    return float(working_days), None

if __name__ == '__main__':
    init_db()
    print("Database initialized successfully at", get_db_path())
