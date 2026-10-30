import unittest
import os
import tempfile
import sqlite3

import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Set testing environment and dedicated test DB path before importing app
TEST_DB_PATH = os.path.join(tempfile.gettempdir(), 'elms_test_suite.db')
if os.path.exists(TEST_DB_PATH):
    try:
        os.remove(TEST_DB_PATH)
    except Exception:
        pass

os.environ['ELMS_DB_PATH'] = TEST_DB_PATH
os.environ['SECRET_KEY'] = 'test_secret_key_12345'

from database import calculate_working_days, get_db_connection, init_db
from seed_data import seed_database
from app import app

class TestLeaveManagementSystem(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        init_db()
        seed_database()

    @classmethod
    def tearDownClass(cls):
        if os.path.exists(TEST_DB_PATH):
            try:
                os.remove(TEST_DB_PATH)
            except Exception:
                pass

    def setUp(self):
        self.app = app
        self.app.config['TESTING'] = True
        self.client = self.app.test_client()

    def test_working_days_calculation_regular(self):
        # 2026-10-12 (Mon) to 2026-10-14 (Wed) -> 3 business days
        days, err = calculate_working_days("2026-10-12", "2026-10-14", False)
        self.assertIsNone(err)
        self.assertEqual(days, 3.0)

    def test_working_days_calculation_with_weekend(self):
        # 2026-10-16 (Fri) to 2026-10-19 (Mon) -> Fri, Mon (2 working days; Sat & Sun excluded)
        days, err = calculate_working_days("2026-10-16", "2026-10-19", False)
        self.assertIsNone(err)
        self.assertEqual(days, 2.0)

    def test_working_days_calculation_with_holiday(self):
        # 2026-01-01 is seeded as New Year's Day (Thursday)
        # 2026-01-01 to 2026-01-02 (Fri) -> only Friday should count (1 day)
        days, err = calculate_working_days("2026-01-01", "2026-01-02", False)
        self.assertIsNone(err)
        self.assertEqual(days, 1.0)

    def test_working_days_half_day(self):
        # 2026-10-13 (Tue) half day -> 0.5
        days, err = calculate_working_days("2026-10-13", "2026-10-13", True)
        self.assertIsNone(err)
        self.assertEqual(days, 0.5)

        # Half day across multiple days should fail
        days, err = calculate_working_days("2026-10-13", "2026-10-14", True)
        self.assertIsNotNone(err)

    def test_working_days_invalid_order(self):
        # End date before start date
        days, err = calculate_working_days("2026-10-20", "2026-10-10", False)
        self.assertIsNotNone(err)
        self.assertIn("cannot be earlier", err)

    def test_authentication_login_success_and_failure(self):
        # 1. Test invalid password when not logged in
        res_fail = self.client.post('/login', data={
            'email': 'admin@company.com',
            'password': 'wrongpassword'
        }, follow_redirects=True)
        self.assertIn(b'Invalid email address or password', res_fail.data)

        # 2. Test valid admin login
        res = self.client.post('/login', data={
            'email': 'admin@company.com',
            'password': 'admin123'
        }, follow_redirects=False)
        self.assertEqual(res.status_code, 302)
        self.assertIn('/admin/dashboard', res.headers['Location'])

        # Clean up session
        self.client.get('/logout')

    def test_role_based_access_control(self):
        # Log in as normal employee
        self.client.post('/login', data={
            'email': 'john.doe@company.com',
            'password': 'employee123'
        })

        # Attempt to access admin dashboard
        res = self.client.get('/admin/dashboard', follow_redirects=False)
        self.assertEqual(res.status_code, 302)
        self.assertIn('/dashboard', res.headers['Location'])

    def test_api_calculate_days(self):
        # Log in first
        self.client.post('/login', data={
            'email': 'john.doe@company.com',
            'password': 'employee123'
        })

        res = self.client.post('/api/calculate-days', json={
            'start_date': '2026-10-12',
            'end_date': '2026-10-16',
            'is_half_day': False
        })
        self.assertEqual(res.status_code, 200)
        json_data = res.get_json()
        self.assertTrue(json_data['valid'])
        self.assertEqual(json_data['days'], 5.0)

    def test_leave_application_and_approval_flow(self):
        # 1. Log in as employee (Sarah Jenkins)
        self.client.post('/login', data={
            'email': 'sarah.j@company.com',
            'password': 'employee123'
        })

        conn = get_db_connection()
        sarah = conn.execute("SELECT id FROM users WHERE email='sarah.j@company.com'").fetchone()
        al_type = conn.execute("SELECT id FROM leave_types WHERE code='AL'").fetchone()

        bal_before = conn.execute(
            "SELECT * FROM leave_balances WHERE user_id=? AND leave_type_id=? AND year=2026",
            (sarah['id'], al_type['id'])
        ).fetchone()
        pending_before = bal_before['pending_days']
        used_before = bal_before['used_days']
        conn.close()

        # 2. Apply for 2 days leave in December
        post_res = self.client.post('/leave/apply', data={
            'leave_type_id': al_type['id'],
            'start_date': '2026-12-07',
            'end_date': '2026-12-08',
            'reason': 'Attending design conference in December'
        }, follow_redirects=True)
        self.assertEqual(post_res.status_code, 200)
        self.assertIn(b'submitted successfully', post_res.data)

        # Check pending balance increased by 2
        conn = get_db_connection()
        bal_after_apply = conn.execute(
            "SELECT * FROM leave_balances WHERE user_id=? AND leave_type_id=? AND year=2026",
            (sarah['id'], al_type['id'])
        ).fetchone()
        self.assertEqual(bal_after_apply['pending_days'], pending_before + 2.0)

        # Find the newly created application ID
        new_app = conn.execute(
            "SELECT id FROM leave_applications WHERE user_id=? AND start_date='2026-12-07' AND status='pending'",
            (sarah['id'],)
        ).fetchone()
        self.assertIsNotNone(new_app)
        app_id = new_app['id']
        conn.close()

        # 3. Log out Sarah, log in as Admin
        self.client.get('/logout')
        self.client.post('/login', data={
            'email': 'admin@company.com',
            'password': 'admin123'
        })

        # 4. Approve the application
        approve_res = self.client.post(f'/admin/approve/{app_id}', data={
            'manager_remarks': 'Approved, represent us well at the conference!'
        }, follow_redirects=True)
        self.assertEqual(approve_res.status_code, 200)

        # Check balance: pending_days released, used_days incremented by 2
        conn = get_db_connection()
        bal_after_approve = conn.execute(
            "SELECT * FROM leave_balances WHERE user_id=? AND leave_type_id=? AND year=2026",
            (sarah['id'], al_type['id'])
        ).fetchone()
        self.assertEqual(bal_after_approve['pending_days'], pending_before)
        self.assertEqual(bal_after_approve['used_days'], used_before + 2.0)

        # Check status is approved with remarks
        app_row = conn.execute("SELECT * FROM leave_applications WHERE id=?", (app_id,)).fetchone()
        self.assertEqual(app_row['status'], 'approved')
        self.assertIn('conference', app_row['manager_remarks'])
        conn.close()

    def test_leave_cancellation(self):
        # 1. Log in as employee (Michael Chang)
        self.client.post('/login', data={
            'email': 'michael.c@company.com',
            'password': 'employee123'
        })

        # Apply for leave in late December
        conn = get_db_connection()
        user = conn.execute("SELECT id FROM users WHERE email='michael.c@company.com'").fetchone()
        al_type = conn.execute("SELECT id FROM leave_types WHERE code='AL'").fetchone()
        conn.close()

        self.client.post('/leave/apply', data={
            'leave_type_id': al_type['id'],
            'start_date': '2026-12-21',
            'end_date': '2026-12-22',
            'reason': 'Pre-holiday rest'
        }, follow_redirects=True)

        conn = get_db_connection()
        app_row = conn.execute(
            "SELECT id FROM leave_applications WHERE user_id=? AND start_date='2026-12-21' AND status='pending'",
            (user['id'],)
        ).fetchone()
        self.assertIsNotNone(app_row)
        app_id = app_row['id']
        conn.close()

        # Cancel the request
        cancel_res = self.client.post(f'/leave/cancel/{app_id}', follow_redirects=True)
        self.assertEqual(cancel_res.status_code, 200)

        conn = get_db_connection()
        cancelled_app = conn.execute("SELECT status FROM leave_applications WHERE id=?", (app_id,)).fetchone()
        self.assertEqual(cancelled_app['status'], 'cancelled')
        conn.close()

    def test_admin_add_employee_and_adjust_quota(self):
        # Log in as Admin
        self.client.post('/login', data={
            'email': 'admin@company.com',
            'password': 'admin123'
        })

        # Add new employee
        add_res = self.client.post('/admin/employees/add', data={
            'full_name': 'Rachel Green',
            'email': 'rachel.g@company.com',
            'department': 'Marketing & Growth',
            'designation': 'Social Media Specialist',
            'role': 'employee',
            'password': 'employee123'
        }, follow_redirects=True)
        self.assertEqual(add_res.status_code, 200)

        conn = get_db_connection()
        new_emp = conn.execute("SELECT * FROM users WHERE email='rachel.g@company.com'").fetchone()
        self.assertIsNotNone(new_emp)
        self.assertEqual(new_emp['full_name'], 'Rachel Green')

        # Check balances auto-initialized
        balances = conn.execute("SELECT * FROM leave_balances WHERE user_id=?", (new_emp['id'],)).fetchall()
        self.assertGreater(len(balances), 0)

        # Adjust quota
        first_type = balances[0]['leave_type_id']
        quota_res = self.client.post('/admin/employees/quota', data={
            'user_id': new_emp['id'],
            'leave_type_id': first_type,
            'total_allocated': 25.0
        }, follow_redirects=True)
        self.assertEqual(quota_res.status_code, 200)

        updated_bal = conn.execute(
            "SELECT total_allocated FROM leave_balances WHERE user_id=? AND leave_type_id=?",
            (new_emp['id'], first_type)
        ).fetchone()
        self.assertEqual(updated_bal['total_allocated'], 25.0)
        conn.close()

    def test_admin_reports_and_csv_export(self):
        # Log in as Admin
        self.client.post('/login', data={
            'email': 'admin@company.com',
            'password': 'admin123'
        })

        # Access reports page
        rep_res = self.client.get('/admin/reports')
        self.assertEqual(rep_res.status_code, 200)
        self.assertIn(b'Leave Records & Compliance Audit', rep_res.data)

        # Export CSV
        csv_res = self.client.get('/admin/reports/export')
        self.assertEqual(csv_res.status_code, 200)
        self.assertEqual(csv_res.mimetype, 'text/csv')
        self.assertIn(b'Application ID,Employee ID,Full Name', csv_res.data)

    def test_leave_application_half_day_fallback(self):
        # Test half-day application where end_date is omitted by browser form submission
        self.client.post('/login', data={
            'email': 'michael.c@company.com',
            'password': 'employee123'
        })

        conn = get_db_connection()
        user = conn.execute("SELECT id FROM users WHERE email='michael.c@company.com'").fetchone()
        sl_type = conn.execute("SELECT id FROM leave_types WHERE code='SL'").fetchone()
        conn.close()

        res = self.client.post('/leave/apply', data={
            'leave_type_id': sl_type['id'],
            'start_date': '2026-11-17', # Tuesday
            'end_date': '',            # Omitted or disabled in browser
            'is_half_day': 'on',
            'reason': 'Medical doctor appointment in morning'
        }, follow_redirects=True)

        self.assertEqual(res.status_code, 200)
        self.assertIn(b'0.5 day(s)', res.data)

        # Check DB recorded 0.5 days
        conn = get_db_connection()
        app_row = conn.execute(
            "SELECT * FROM leave_applications WHERE user_id=? AND start_date='2026-11-17'",
            (user['id'],)
        ).fetchone()
        self.assertIsNotNone(app_row)
        self.assertEqual(app_row['total_days'], 0.5)
        self.assertEqual(app_row['end_date'], '2026-11-17')
        self.assertEqual(app_row['is_half_day'], 1)
        conn.close()

    def test_leave_rejection_flow(self):
        # 1. Employee applies
        self.client.post('/login', data={
            'email': 'sarah.j@company.com',
            'password': 'employee123'
        })
        conn = get_db_connection()
        sarah = conn.execute("SELECT id FROM users WHERE email='sarah.j@company.com'").fetchone()
        lt = conn.execute("SELECT id FROM leave_types WHERE code='AL'").fetchone()

        bal_before = conn.execute(
            "SELECT * FROM leave_balances WHERE user_id=? AND leave_type_id=? AND year=2026",
            (sarah['id'], lt['id'])
        ).fetchone()
        conn.close()

        self.client.post('/leave/apply', data={
            'leave_type_id': lt['id'],
            'start_date': '2026-11-23',
            'end_date': '2026-11-24',
            'reason': 'Trip to visit family'
        }, follow_redirects=True)

        conn = get_db_connection()
        app_row = conn.execute(
            "SELECT id FROM leave_applications WHERE user_id=? AND start_date='2026-11-23' AND status='pending'",
            (sarah['id'],)
        ).fetchone()
        app_id = app_row['id']
        conn.close()

        # 2. Admin rejects
        self.client.get('/logout')
        self.client.post('/login', data={'email': 'admin@company.com', 'password': 'admin123'})
        reject_res = self.client.post(f'/admin/reject/{app_id}', data={
            'manager_remarks': 'Understaffed team during sprint delivery'
        }, follow_redirects=True)
        self.assertEqual(reject_res.status_code, 200)

        # Balance pending hold should be released back, used_days unchanged
        conn = get_db_connection()
        bal_after = conn.execute(
            "SELECT * FROM leave_balances WHERE user_id=? AND leave_type_id=? AND year=2026",
            (sarah['id'], lt['id'])
        ).fetchone()
        self.assertEqual(bal_after['pending_days'], bal_before['pending_days'])
        self.assertEqual(bal_after['used_days'], bal_before['used_days'])

        app_checked = conn.execute("SELECT status, manager_remarks FROM leave_applications WHERE id=?", (app_id,)).fetchone()
        self.assertEqual(app_checked['status'], 'rejected')
        self.assertIn('Understaffed', app_checked['manager_remarks'])
        conn.close()

    def test_insufficient_balance_rejection(self):
        self.client.post('/login', data={'email': 'john.doe@company.com', 'password': 'employee123'})
        conn = get_db_connection()
        lt = conn.execute("SELECT id FROM leave_types WHERE code='AL'").fetchone()
        conn.close()

        # Request 4 weeks of leave (20 business days), when balance is ~15
        res = self.client.post('/leave/apply', data={
            'leave_type_id': lt['id'],
            'start_date': '2026-11-02',
            'end_date': '2026-11-30',
            'reason': 'One full month holiday trip'
        }, follow_redirects=True)

        self.assertEqual(res.status_code, 200)
        self.assertIn(b'Insufficient leave balance', res.data)

    def test_overlapping_leave_rejection(self):
        self.client.post('/login', data={'email': 'john.doe@company.com', 'password': 'employee123'})
        conn = get_db_connection()
        lt = conn.execute("SELECT id FROM leave_types WHERE code='AL'").fetchone()
        conn.close()

        # Apply for first block: 2026-12-14 to 2026-12-16
        res1 = self.client.post('/leave/apply', data={
            'leave_type_id': lt['id'],
            'start_date': '2026-12-14',
            'end_date': '2026-12-16',
            'reason': 'Winter holiday block 1'
        }, follow_redirects=True)
        self.assertIn(b'submitted successfully', res1.data)

        # Apply for overlapping block: 2026-12-15 to 2026-12-17
        res2 = self.client.post('/leave/apply', data={
            'leave_type_id': lt['id'],
            'start_date': '2026-12-15',
            'end_date': '2026-12-17',
            'reason': 'Overlapping leave request'
        }, follow_redirects=True)
        self.assertIn(b'already have an existing leave application covering one or more of these dates', res2.data)

    def test_admin_validation_and_audit_logs(self):
        self.client.post('/login', data={'email': 'admin@company.com', 'password': 'admin123'})

        # Test duplicate email rejection
        res_dup = self.client.post('/admin/employees/add', data={
            'full_name': 'Duplicate Person',
            'email': 'john.doe@company.com',
            'department': 'Engineering',
            'designation': 'Software Engineer',
            'role': 'employee',
            'password': 'password123'
        }, follow_redirects=True)
        self.assertIn(b'already exists', res_dup.data)

        # Test invalid negative quota rejected
        res_neg = self.client.post('/admin/employees/quota', data={
            'user_id': 1,
            'leave_type_id': 1,
            'total_allocated': -5.0
        }, follow_redirects=True)
        self.assertIn(b'Invalid quota values provided', res_neg.data)

        # Verify audit logs captured records
        conn = get_db_connection()
        logs = conn.execute("SELECT COUNT(*) as count FROM audit_logs").fetchone()['count']
        self.assertGreater(logs, 0)
        conn.close()

if __name__ == '__main__':
    unittest.main()
