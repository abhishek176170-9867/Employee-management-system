import os
import csv
import io
import datetime
from functools import wraps
from flask import (
    Flask, render_template, request, redirect, url_for, session,
    flash, jsonify, Response, send_file
)
from werkzeug.security import check_password_hash, generate_password_hash
from database import (
    get_db_connection, init_db, log_audit, calculate_working_days, get_holiday_dates
)
from seed_data import seed_database

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'elms_enterprise_secret_key_2026_secure')
app.config['TEMPLATES_AUTO_RELOAD'] = True

# Ensure DB initialized and seeded on startup
with app.app_context():
    init_db()
    seed_database()

# -------------------------------------------------------------
# Context Processors & Decorators
# -------------------------------------------------------------
@app.context_processor
def inject_global_data():
    current_year = datetime.datetime.now().year
    pending_count = 0
    if session.get('role') == 'admin':
        try:
            conn = get_db_connection()
            pending_count = conn.execute(
                "SELECT COUNT(*) as count FROM leave_applications WHERE status = 'pending'"
            ).fetchone()['count']
            conn.close()
        except Exception:
            pending_count = 0
    return {
        'current_year': current_year,
        'global_pending_count': pending_count,
        'today_date': datetime.date.today().strftime('%Y-%m-%d')
    }

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please log in to access this portal.', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session or session.get('role') != 'admin':
            flash('Access restricted. Administrator privileges required.', 'danger')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

# -------------------------------------------------------------
# Authentication Routes
# -------------------------------------------------------------
@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id' in session:
        if session.get('role') == 'admin':
            return redirect(url_for('admin_dashboard'))
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        email = request.form.get('email', '').strip().lower()
        password = request.form.get('password', '')

        conn = get_db_connection()
        user = conn.execute('SELECT * FROM users WHERE email = ? AND is_active = 1', (email,)).fetchone()
        conn.close()

        if user and check_password_hash(user['password_hash'], password):
            session['user_id'] = user['id']
            session['emp_code'] = user['emp_code']
            session['full_name'] = user['full_name']
            session['email'] = user['email']
            session['role'] = user['role']
            session['department'] = user['department']
            session['designation'] = user['designation']
            session['avatar_color'] = user['avatar_color']

            log_audit(user['id'], 'LOGIN', f"User logged in from portal.")
            flash(f"Welcome back, {user['full_name']}!", 'success')

            if user['role'] == 'admin':
                return redirect(url_for('admin_dashboard'))
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid email address or password. Please try again.', 'danger')

    return render_template('login.html')

@app.route('/logout')
def logout():
    user_id = session.get('user_id')
    if user_id:
        log_audit(user_id, 'LOGOUT', 'User logged out.')
    session.clear()
    flash('You have been securely signed out.', 'info')
    return redirect(url_for('login'))

# -------------------------------------------------------------
# Employee Portal Routes
# -------------------------------------------------------------
@app.route('/')
@login_required
def root():
    if session.get('role') == 'admin':
        return redirect(url_for('admin_dashboard'))
    return redirect(url_for('dashboard'))

@app.route('/dashboard')
@login_required
def dashboard():
    user_id = session['user_id']
    current_year = datetime.datetime.now().year
    conn = get_db_connection()

    # Fetch leave balances with type details
    balances = conn.execute('''
        SELECT lb.*, lt.name as leave_name, lt.code as leave_code, lt.color as leave_color, lt.description
        FROM leave_balances lb
        JOIN leave_types lt ON lb.leave_type_id = lt.id
        WHERE lb.user_id = ? AND lb.year = ?
        ORDER BY lt.id ASC
    ''', (user_id, current_year)).fetchall()

    # Fetch recent applications
    recent_applications = conn.execute('''
        SELECT la.*, lt.name as leave_name, lt.code as leave_code, lt.color as leave_color
        FROM leave_applications la
        JOIN leave_types lt ON la.leave_type_id = lt.id
        WHERE la.user_id = ?
        ORDER BY la.applied_at DESC
        LIMIT 5
    ''', (user_id,)).fetchall()

    # Upcoming company holidays (next 4)
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    upcoming_holidays = conn.execute('''
        SELECT * FROM holidays
        WHERE holiday_date >= ?
        ORDER BY holiday_date ASC
        LIMIT 4
    ''', (today_str,)).fetchall()

    # Calculate summary counts
    total_allowed = sum(b['total_allocated'] for b in balances)
    total_used = sum(b['used_days'] for b in balances)
    total_pending = sum(b['pending_days'] for b in balances)
    total_remaining = max(0.0, total_allowed - total_used - total_pending)

    conn.close()

    return render_template(
        'employee/dashboard.html',
        balances=balances,
        recent_applications=recent_applications,
        upcoming_holidays=upcoming_holidays,
        total_allowed=total_allowed,
        total_used=total_used,
        total_pending=total_pending,
        total_remaining=total_remaining
    )

@app.route('/leave/apply', methods=['GET', 'POST'])
@login_required
def apply_leave():
    user_id = session['user_id']
    current_year = datetime.datetime.now().year
    conn = get_db_connection()

    if request.method == 'POST':
        leave_type_id = request.form.get('leave_type_id', type=int)
        start_date_str = request.form.get('start_date', '').strip()
        end_date_str = request.form.get('end_date', '').strip()
        is_half_day = 1 if request.form.get('is_half_day') == 'on' else 0
        reason = request.form.get('reason', '').strip()

        # Fallback for half-day requests if end_date was omitted or read-only
        if is_half_day and not end_date_str:
            end_date_str = start_date_str

        if not leave_type_id or not start_date_str or not end_date_str or not reason:
            flash('Please fill in all mandatory fields.', 'danger')
            return redirect(url_for('apply_leave'))

        # Calculate working business days
        days, err = calculate_working_days(start_date_str, end_date_str, bool(is_half_day))
        if err:
            flash(err, 'danger')
            return redirect(url_for('apply_leave'))

        # Check existing user balance
        balance = conn.execute('''
            SELECT * FROM leave_balances
            WHERE user_id = ? AND leave_type_id = ? AND year = ?
        ''', (user_id, leave_type_id, current_year)).fetchone()

        if not balance:
            flash('Leave quota not configured for this category.', 'danger')
            return redirect(url_for('apply_leave'))

        remaining_available = balance['total_allocated'] - balance['used_days'] - balance['pending_days']
        if days > remaining_available:
            flash(
                f"Insufficient leave balance. You requested {days} day(s), but have only {remaining_available:.1f} day(s) available.",
                'danger'
            )
            return redirect(url_for('apply_leave'))

        # Check for date overlap with existing pending or approved leaves
        overlap = conn.execute('''
            SELECT id FROM leave_applications
            WHERE user_id = ?
              AND status IN ('pending', 'approved')
              AND NOT (end_date < ? OR start_date > ?)
        ''', (user_id, start_date_str, end_date_str)).fetchone()

        if overlap:
            flash('You already have an existing leave application covering one or more of these dates.', 'warning')
            return redirect(url_for('apply_leave'))

        # Insert application & update pending days hold
        conn.execute('''
            INSERT INTO leave_applications
            (user_id, leave_type_id, start_date, end_date, total_days, is_half_day, reason, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, 'pending')
        ''', (user_id, leave_type_id, start_date_str, end_date_str, days, is_half_day, reason))

        conn.execute('''
            UPDATE leave_balances
            SET pending_days = pending_days + ?
            WHERE id = ?
        ''', (days, balance['id']))

        conn.commit()
        log_audit(user_id, 'LEAVE_APPLIED', f"Applied for {days} days of leave from {start_date_str} to {end_date_str}.")
        conn.close()

        flash(f'Your leave request for {days} day(s) has been submitted successfully for manager approval.', 'success')
        return redirect(url_for('leave_history'))

    # GET request: load available leave types with current user balances
    leave_types = conn.execute('''
        SELECT lt.*, lb.total_allocated, lb.used_days, lb.pending_days,
               (lb.total_allocated - lb.used_days - lb.pending_days) as available_days
        FROM leave_types lt
        JOIN leave_balances lb ON lt.id = lb.leave_type_id
        WHERE lb.user_id = ? AND lb.year = ?
        ORDER BY lt.id ASC
    ''', (user_id, current_year)).fetchall()

    conn.close()
    return render_template('employee/apply.html', leave_types=leave_types)

@app.route('/leave/history')
@login_required
def leave_history():
    user_id = session['user_id']
    status_filter = request.args.get('status', 'all')
    conn = get_db_connection()

    query = '''
        SELECT la.*, lt.name as leave_name, lt.code as leave_code, lt.color as leave_color,
               u_rev.full_name as reviewer_name
        FROM leave_applications la
        JOIN leave_types lt ON la.leave_type_id = lt.id
        LEFT JOIN users u_rev ON la.reviewed_by = u_rev.id
        WHERE la.user_id = ?
    '''
    params = [user_id]

    if status_filter in ['pending', 'approved', 'rejected', 'cancelled']:
        query += ' AND la.status = ?'
        params.append(status_filter)

    query += ' ORDER BY la.applied_at DESC'
    applications = conn.execute(query, params).fetchall()
    conn.close()

    return render_template('employee/history.html', applications=applications, current_filter=status_filter)

@app.route('/leave/cancel/<int:app_id>', methods=['POST'])
@login_required
def cancel_leave(app_id):
    user_id = session['user_id']
    current_year = datetime.datetime.now().year
    conn = get_db_connection()

    application = conn.execute('''
        SELECT * FROM leave_applications
        WHERE id = ? AND user_id = ?
    ''', (app_id, user_id)).fetchone()

    if not application:
        flash('Leave application not found.', 'danger')
        conn.close()
        return redirect(url_for('leave_history'))

    if application['status'] != 'pending':
        flash('Only pending leave requests can be cancelled.', 'warning')
        conn.close()
        return redirect(url_for('leave_history'))

    # Update status to cancelled
    conn.execute('''
        UPDATE leave_applications
        SET status = 'cancelled'
        WHERE id = ?
    ''', (app_id,))

    # Release pending days in leave_balances
    app_year = int(application['start_date'][:4]) if application['start_date'] else current_year
    conn.execute('''
        UPDATE leave_balances
        SET pending_days = MAX(0.0, pending_days - ?)
        WHERE user_id = ? AND leave_type_id = ? AND year = ?
    ''', (application['total_days'], user_id, application['leave_type_id'], app_year))

    conn.commit()
    log_audit(user_id, 'LEAVE_CANCELLED', f"Cancelled pending leave request #{app_id}.")
    conn.close()

    flash('Leave application has been cancelled.', 'info')
    return redirect(url_for('leave_history'))

@app.route('/holidays')
@login_required
def holidays():
    conn = get_db_connection()
    all_holidays = conn.execute('SELECT * FROM holidays ORDER BY holiday_date ASC').fetchall()
    conn.close()
    return render_template('employee/holidays.html', holidays=all_holidays)

# -------------------------------------------------------------
# Admin & HR Portal Routes
# -------------------------------------------------------------
@app.route('/admin/dashboard')
@admin_required
def admin_dashboard():
    conn = get_db_connection()
    today_str = datetime.date.today().strftime('%Y-%m-%d')
    current_year = datetime.datetime.now().year

    # KPI 1: Active Employees
    total_employees = conn.execute('SELECT COUNT(*) as count FROM users WHERE is_active = 1').fetchone()['count']

    # KPI 2: Pending Applications
    pending_count = conn.execute("SELECT COUNT(*) as count FROM leave_applications WHERE status = 'pending'").fetchone()['count']

    # KPI 3: On Leave Today
    on_leave_today_count = conn.execute('''
        SELECT COUNT(DISTINCT user_id) as count FROM leave_applications
        WHERE status = 'approved' AND start_date <= ? AND end_date >= ?
    ''', (today_str, today_str)).fetchone()['count']

    # KPI 4: Total Days Approved This Year
    approved_days_year = conn.execute('''
        SELECT COALESCE(SUM(total_days), 0.0) as total FROM leave_applications
        WHERE status = 'approved' AND strftime('%Y', start_date) = ?
    ''', (str(current_year),)).fetchone()['total']

    # Recent pending applications (up to 6)
    pending_applications = conn.execute('''
        SELECT la.*, u.full_name, u.emp_code, u.department, u.avatar_color,
               lt.name as leave_name, lt.code as leave_code, lt.color as leave_color
        FROM leave_applications la
        JOIN users u ON la.user_id = u.id
        JOIN leave_types lt ON la.leave_type_id = lt.id
        WHERE la.status = 'pending'
        ORDER BY la.applied_at ASC
        LIMIT 6
    ''').fetchall()

    # Employees on leave today list
    on_leave_today_list = conn.execute('''
        SELECT la.*, u.full_name, u.department, u.avatar_color, lt.name as leave_name
        FROM leave_applications la
        JOIN users u ON la.user_id = u.id
        JOIN leave_types lt ON la.leave_type_id = lt.id
        WHERE la.status = 'approved' AND start_date <= ? AND end_date >= ?
    ''', (today_str, today_str)).fetchall()

    conn.close()

    return render_template(
        'admin/dashboard.html',
        total_employees=total_employees,
        pending_count=pending_count,
        on_leave_today_count=on_leave_today_count,
        approved_days_year=approved_days_year,
        pending_applications=pending_applications,
        on_leave_today_list=on_leave_today_list
    )

@app.route('/admin/approvals')
@admin_required
def admin_approvals():
    conn = get_db_connection()
    current_year = datetime.datetime.now().year

    pending_list = conn.execute('''
        SELECT la.*, u.full_name, u.emp_code, u.email, u.department, u.designation, u.avatar_color,
               lt.name as leave_name, lt.code as leave_code, lt.color as leave_color,
               (lb.total_allocated - lb.used_days) as current_remaining
        FROM leave_applications la
        JOIN users u ON la.user_id = u.id
        JOIN leave_types lt ON la.leave_type_id = lt.id
        LEFT JOIN leave_balances lb ON (lb.user_id = u.id AND lb.leave_type_id = lt.id AND lb.year = ?)
        WHERE la.status = 'pending'
        ORDER BY la.applied_at ASC
    ''', (current_year,)).fetchall()

    conn.close()
    return render_template('admin/approvals.html', applications=pending_list)

@app.route('/admin/approve/<int:app_id>', methods=['POST'])
@admin_required
def admin_approve_leave(app_id):
    admin_id = session['user_id']
    remarks = request.form.get('manager_remarks', '').strip()
    current_year = datetime.datetime.now().year
    conn = get_db_connection()

    application = conn.execute('''
        SELECT la.*, u.full_name
        FROM leave_applications la
        JOIN users u ON la.user_id = u.id
        WHERE la.id = ?
    ''', (app_id,)).fetchone()

    if not application or application['status'] != 'pending':
        flash('Request is no longer pending or does not exist.', 'warning')
        conn.close()
        return redirect(url_for('admin_approvals'))

    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Update application status
    conn.execute('''
        UPDATE leave_applications
        SET status = 'approved',
            reviewed_by = ?,
            reviewed_at = ?,
            manager_remarks = ?
        WHERE id = ?
    ''', (admin_id, now_str, remarks, app_id))

    # Move from pending_days to used_days in leave_balances
    app_year = int(application['start_date'][:4]) if application['start_date'] else current_year
    conn.execute('''
        UPDATE leave_balances
        SET pending_days = MAX(0.0, pending_days - ?),
            used_days = used_days + ?
        WHERE user_id = ? AND leave_type_id = ? AND year = ?
    ''', (application['total_days'], application['total_days'], application['user_id'], application['leave_type_id'], app_year))

    conn.commit()
    log_audit(admin_id, 'LEAVE_APPROVED', f"Approved leave request #{app_id} for {application['full_name']}.")
    conn.close()

    flash(f"Leave request for {application['full_name']} has been approved.", 'success')
    return redirect(url_for('admin_approvals'))

@app.route('/admin/reject/<int:app_id>', methods=['POST'])
@admin_required
def admin_reject_leave(app_id):
    admin_id = session['user_id']
    remarks = request.form.get('manager_remarks', '').strip()
    current_year = datetime.datetime.now().year
    conn = get_db_connection()

    application = conn.execute('''
        SELECT la.*, u.full_name
        FROM leave_applications la
        JOIN users u ON la.user_id = u.id
        WHERE la.id = ?
    ''', (app_id,)).fetchone()

    if not application or application['status'] != 'pending':
        flash('Request is no longer pending or does not exist.', 'warning')
        conn.close()
        return redirect(url_for('admin_approvals'))

    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Update application status
    conn.execute('''
        UPDATE leave_applications
        SET status = 'rejected',
            reviewed_by = ?,
            reviewed_at = ?,
            manager_remarks = ?
        WHERE id = ?
    ''', (admin_id, now_str, remarks, app_id))

    # Release pending hold without adding to used_days
    app_year = int(application['start_date'][:4]) if application['start_date'] else current_year
    conn.execute('''
        UPDATE leave_balances
        SET pending_days = MAX(0.0, pending_days - ?)
        WHERE user_id = ? AND leave_type_id = ? AND year = ?
    ''', (application['total_days'], application['user_id'], application['leave_type_id'], app_year))

    conn.commit()
    log_audit(admin_id, 'LEAVE_REJECTED', f"Rejected leave request #{app_id} for {application['full_name']}.")
    conn.close()

    flash(f"Leave request for {application['full_name']} has been rejected.", 'info')
    return redirect(url_for('admin_approvals'))

@app.route('/admin/employees')
@admin_required
def admin_employees():
    current_year = datetime.datetime.now().year
    conn = get_db_connection()

    employees = conn.execute('''
        SELECT u.*,
               COALESCE((SELECT SUM(total_allocated) FROM leave_balances WHERE user_id = u.id AND year = ?), 0) as total_quota,
               COALESCE((SELECT SUM(used_days) FROM leave_balances WHERE user_id = u.id AND year = ?), 0) as used_quota
        FROM users u
        ORDER BY u.role DESC, u.full_name ASC
    ''', (current_year, current_year)).fetchall()

    leave_types = conn.execute('SELECT * FROM leave_types ORDER BY id ASC').fetchall()
    departments = [
        "Engineering", "Product & Design", "Marketing & Growth",
        "People Operations & HR", "Sales & Accounts", "Finance & Legal"
    ]

    conn.close()
    return render_template('admin/employees.html', employees=employees, leave_types=leave_types, departments=departments)

@app.route('/admin/employees/add', methods=['POST'])
@admin_required
def admin_add_employee():
    admin_id = session['user_id']
    full_name = request.form.get('full_name', '').strip()
    email = request.form.get('email', '').strip().lower()
    department = request.form.get('department', '').strip()
    designation = request.form.get('designation', '').strip()
    role = request.form.get('role', 'employee')
    if role not in ['admin', 'employee']:
        role = 'employee'
    password = request.form.get('password', 'employee123')
    current_year = datetime.datetime.now().year

    if not full_name or not email or not department or not designation:
        flash('Please fill in all employee fields.', 'danger')
        return redirect(url_for('admin_employees'))

    conn = get_db_connection()
    # Check duplicate email
    existing = conn.execute('SELECT id FROM users WHERE email = ?', (email,)).fetchone()
    if existing:
        flash(f'An employee with email {email} already exists.', 'danger')
        conn.close()
        return redirect(url_for('admin_employees'))

    # Generate emp_code
    max_id = conn.execute('SELECT MAX(id) as max_id FROM users').fetchone()['max_id'] or 1000
    emp_code = f"EMP-{max_id + 1}"
    colors = ['#2563EB', '#7C3AED', '#DB2777', '#059669', '#D97706', '#4B5563']
    avatar_color = colors[(max_id + 1) % len(colors)]
    today_str = datetime.date.today().strftime('%Y-%m-%d')

    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO users (emp_code, full_name, email, password_hash, role, department, designation, join_date, avatar_color)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (emp_code, full_name, email, generate_password_hash(password), role, department, designation, today_str, avatar_color))

    new_user_id = cursor.lastrowid

    # Initialize leave balances
    leave_types = conn.execute('SELECT id, default_days FROM leave_types').fetchall()
    for lt in leave_types:
        cursor.execute('''
            INSERT INTO leave_balances (user_id, leave_type_id, year, total_allocated, used_days, pending_days)
            VALUES (?, ?, ?, ?, 0.0, 0.0)
        ''', (new_user_id, lt['id'], current_year, lt['default_days']))

    conn.commit()
    log_audit(admin_id, 'USER_CREATED', f"Created user {full_name} ({emp_code}).")
    conn.close()

    flash(f"New employee {full_name} ({emp_code}) added successfully with default leave quotas.", 'success')
    return redirect(url_for('admin_employees'))

@app.route('/admin/employees/quota', methods=['POST'])
@admin_required
def admin_update_quota():
    admin_id = session['user_id']
    user_id = request.form.get('user_id', type=int)
    leave_type_id = request.form.get('leave_type_id', type=int)
    new_quota = request.form.get('total_allocated', type=float)
    current_year = datetime.datetime.now().year

    if not user_id or not leave_type_id or new_quota is None or new_quota < 0:
        flash('Invalid quota values provided.', 'danger')
        return redirect(url_for('admin_employees'))

    conn = get_db_connection()
    user_exists = conn.execute('SELECT id FROM users WHERE id = ?', (user_id,)).fetchone()
    lt_exists = conn.execute('SELECT id FROM leave_types WHERE id = ?', (leave_type_id,)).fetchone()
    if not user_exists or not lt_exists:
        flash('Target employee or leave category does not exist.', 'danger')
        conn.close()
        return redirect(url_for('admin_employees'))
    conn.execute('''
        INSERT INTO leave_balances (user_id, leave_type_id, year, total_allocated, used_days, pending_days)
        VALUES (?, ?, ?, ?, 0.0, 0.0)
        ON CONFLICT(user_id, leave_type_id, year)
        DO UPDATE SET total_allocated = excluded.total_allocated
    ''', (user_id, leave_type_id, current_year, new_quota))

    conn.commit()
    log_audit(admin_id, 'QUOTA_UPDATED', f"Updated leave quota for user #{user_id}, type #{leave_type_id} to {new_quota} days.")
    conn.close()

    flash('Employee leave quota updated successfully.', 'success')
    return redirect(url_for('admin_employees'))

@app.route('/admin/reports')
@admin_required
def admin_reports():
    conn = get_db_connection()
    status_filter = request.args.get('status', '')
    dept_filter = request.args.get('department', '')
    type_filter = request.args.get('leave_type_id', '')

    query = '''
        SELECT la.*, u.full_name, u.emp_code, u.department,
               lt.name as leave_name, lt.code as leave_code, lt.color as leave_color,
               u_rev.full_name as reviewer_name
        FROM leave_applications la
        JOIN users u ON la.user_id = u.id
        JOIN leave_types lt ON la.leave_type_id = lt.id
        LEFT JOIN users u_rev ON la.reviewed_by = u_rev.id
        WHERE 1=1
    '''
    params = []

    if status_filter:
        query += ' AND la.status = ?'
        params.append(status_filter)
    if dept_filter:
        query += ' AND u.department = ?'
        params.append(dept_filter)
    if type_filter:
        query += ' AND la.leave_type_id = ?'
        params.append(type_filter)

    query += ' ORDER BY la.applied_at DESC'
    applications = conn.execute(query, params).fetchall()

    leave_types = conn.execute('SELECT * FROM leave_types ORDER BY id ASC').fetchall()
    departments = conn.execute('SELECT DISTINCT department FROM users ORDER BY department ASC').fetchall()

    # Aggregate summaries
    total_records = len(applications)
    total_approved_days = sum(a['total_days'] for a in applications if a['status'] == 'approved')
    total_pending_days = sum(a['total_days'] for a in applications if a['status'] == 'pending')

    conn.close()

    return render_template(
        'admin/reports.html',
        applications=applications,
        leave_types=leave_types,
        departments=[d['department'] for d in departments],
        status_filter=status_filter,
        dept_filter=dept_filter,
        type_filter=type_filter,
        total_records=total_records,
        total_approved_days=total_approved_days,
        total_pending_days=total_pending_days
    )

@app.route('/admin/reports/export')
@admin_required
def admin_reports_export():
    conn = get_db_connection()
    status_filter = request.args.get('status', '')
    dept_filter = request.args.get('department', '')
    type_filter = request.args.get('leave_type_id', '')

    query = '''
        SELECT la.id, u.emp_code, u.full_name, u.department, lt.name as leave_type,
               la.start_date, la.end_date, la.total_days, la.is_half_day, la.reason,
               la.status, la.applied_at, u_rev.full_name as reviewer, la.reviewed_at, la.manager_remarks
        FROM leave_applications la
        JOIN users u ON la.user_id = u.id
        JOIN leave_types lt ON la.leave_type_id = lt.id
        LEFT JOIN users u_rev ON la.reviewed_by = u_rev.id
        WHERE 1=1
    '''
    params = []
    if status_filter:
        query += ' AND la.status = ?'
        params.append(status_filter)
    if dept_filter:
        query += ' AND u.department = ?'
        params.append(dept_filter)
    if type_filter:
        query += ' AND la.leave_type_id = ?'
        params.append(type_filter)

    query += ' ORDER BY la.applied_at DESC'
    rows = conn.execute(query, params).fetchall()
    conn.close()

    si = io.StringIO()
    cw = csv.writer(si)
    cw.writerow([
        'Application ID', 'Employee ID', 'Full Name', 'Department', 'Leave Category',
        'Start Date', 'End Date', 'Days', 'Is Half Day', 'Reason', 'Status',
        'Applied At', 'Reviewed By', 'Reviewed At', 'Manager Remarks'
    ])
    for r in rows:
        cw.writerow([
            r['id'], r['emp_code'], r['full_name'], r['department'], r['leave_type'],
            r['start_date'], r['end_date'], r['total_days'],
            'Yes' if r['is_half_day'] else 'No', r['reason'], r['status'].capitalize(),
            r['applied_at'], r['reviewer'] or 'N/A', r['reviewed_at'] or 'N/A', r['manager_remarks'] or ''
        ])

    output = io.BytesIO(si.getvalue().encode('utf-8'))
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(
        output,
        mimetype='text/csv',
        as_attachment=True,
        download_name=f'leave_records_report_{timestamp}.csv'
    )

# -------------------------------------------------------------
# Dynamic JSON APIs for Frontend Interactivity & Charts
# -------------------------------------------------------------
@app.route('/api/calculate-days', methods=['POST'])
@login_required
def api_calculate_days():
    data = request.get_json() or {}
    start_date = data.get('start_date')
    end_date = data.get('end_date')
    is_half_day = bool(data.get('is_half_day', False))

    if is_half_day and start_date and not end_date:
        end_date = start_date

    if not start_date or not end_date:
        return jsonify({'valid': False, 'message': 'Both start and end dates are required.'})

    days, err = calculate_working_days(start_date, end_date, is_half_day)
    if err:
        return jsonify({'valid': False, 'message': err, 'days': 0})

    return jsonify({
        'valid': True,
        'days': days,
        'message': f"{days} working business day(s) calculated (weekends and company holidays excluded)."
    })

@app.route('/api/chart-data')
@admin_required
def api_chart_data():
    conn = get_db_connection()
    current_year = datetime.datetime.now().year

    # 1. Leave breakdown by Type (Donut Chart)
    type_distribution = conn.execute('''
        SELECT lt.name, lt.color, COALESCE(SUM(la.total_days), 0) as total_days
        FROM leave_types lt
        LEFT JOIN leave_applications la ON (lt.id = la.leave_type_id AND la.status = 'approved' AND strftime('%Y', la.start_date) = ?)
        GROUP BY lt.id
    ''', (str(current_year),)).fetchall()

    # 2. Monthly Leave Distribution (Bar Chart)
    # Months 1 to 12
    monthly_data = []
    month_labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for m in range(1, 13):
        m_str = f"{m:02d}"
        days = conn.execute('''
            SELECT COALESCE(SUM(total_days), 0) as total
            FROM leave_applications
            WHERE status = 'approved' AND strftime('%Y-%m', start_date) = ?
        ''', (f"{current_year}-{m_str}",)).fetchone()['total']
        monthly_data.append(days)

    # 3. Department Leave Utilization
    dept_distribution = conn.execute('''
        SELECT u.department, COALESCE(SUM(la.total_days), 0) as total_days
        FROM users u
        LEFT JOIN leave_applications la ON (u.id = la.user_id AND la.status = 'approved' AND strftime('%Y', la.start_date) = ?)
        GROUP BY u.department
    ''', (str(current_year),)).fetchall()

    conn.close()

    return jsonify({
        'type_labels': [t['name'] for t in type_distribution],
        'type_data': [t['total_days'] for t in type_distribution],
        'type_colors': [t['color'] for t in type_distribution],
        'month_labels': month_labels,
        'month_data': monthly_data,
        'dept_labels': [d['department'] for d in dept_distribution],
        'dept_data': [d['total_days'] for d in dept_distribution]
    })

if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5000)
