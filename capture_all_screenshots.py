import os
import subprocess
import time
from app import app

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'report_assets')
os.makedirs(OUTPUT_DIR, exist_ok=True)
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def capture_html(html_content, filename):
    static_abs = os.path.join(os.path.dirname(__file__), 'static').replace('\\', '/')
    html_content = html_content.replace('href="/static/', f'href="file:///{static_abs}/')
    html_content = html_content.replace('src="/static/', f'src="file:///{static_abs}/')

    html_file = os.path.join(OUTPUT_DIR, f"{filename}.html")
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(html_content)

    img_file = os.path.join(OUTPUT_DIR, f"{filename}.png")
    cmd = [
        EDGE_PATH,
        '--headless=new',
        '--disable-gpu',
        '--window-size=1280,820',
        f'--screenshot={img_file}',
        f'file:///{html_file.replace(os.sep, "/")}'
    ]
    subprocess.run(cmd, capture_output=True)
    print(f"Captured: {filename}.png")
    return img_file

def main():
    app.config['TESTING'] = True
    client = app.test_client()

    # 1. Login Screen
    res = client.get('/login')
    capture_html(res.get_data(as_text=True), '01_login_portal')

    # 2. Employee Dashboard
    client.post('/login', data={'email': 'john.doe@company.com', 'password': 'employee123'})
    res = client.get('/dashboard')
    capture_html(res.get_data(as_text=True), '02_employee_dashboard')

    # 3. Apply Leave Form (Empty)
    res = client.get('/leave/apply')
    capture_html(res.get_data(as_text=True), '03_apply_leave_empty')

    # 4. Apply Leave Form with Live Calculation Badge active
    html_calc = res.get_data(as_text=True)
    # Inject values into start_date, end_date and unhide badge
    html_calc = html_calc.replace('id="start_date" name="start_date"', 'id="start_date" name="start_date" value="2026-10-15"')
    html_calc = html_calc.replace('id="end_date" name="end_date"', 'id="end_date" name="end_date" value="2026-10-16"')
    html_calc = html_calc.replace('<option value="3"', '<option value="3" selected')
    html_calc = html_calc.replace('id="calculated_days_badge" class="hidden"', 'id="calculated_days_badge" class="mt-2 p-3 rounded-xl text-sm font-medium bg-emerald-50 text-emerald-700 border border-emerald-200 flex items-center gap-2"')
    html_calc = html_calc.replace('<span id="calculated_days_text"></span>', '<span id="calculated_days_text">2.0 working day(s) calculated (weekends and company holidays excluded)</span>')
    capture_html(html_calc, '04_apply_leave_calculated')

    # 5. Apply Leave with Overdraft Warning Banner
    html_warn = html_calc.replace('id="balance_warning" class="hidden', 'id="balance_warning" class="p-3 rounded-xl text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200')
    html_warn = html_warn.replace('<div id="balance_warning" class="p-3 rounded-xl text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200">', '<div id="balance_warning" class="p-3 rounded-xl text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200"><i class="fa-solid fa-triangle-exclamation mr-1.5"></i> Warning: Requested duration exceeds your remaining balance for this category. Please adjust dates.')
    capture_html(html_warn, '05_apply_leave_warning')

    # 6. Leave History (All)
    res = client.get('/leave/history?status=all')
    capture_html(res.get_data(as_text=True), '06_leave_history_all')

    # 7. Leave History (Pending with Cancel Button)
    res = client.get('/leave/history?status=pending')
    capture_html(res.get_data(as_text=True), '07_leave_history_pending')

    # 8. Company Holidays Schedule
    res = client.get('/holidays')
    capture_html(res.get_data(as_text=True), '08_holiday_schedule')

    # 9. Admin Dashboard
    client.get('/logout')
    client.post('/login', data={'email': 'admin@company.com', 'password': 'admin123'})
    res = client.get('/admin/dashboard')
    capture_html(res.get_data(as_text=True), '09_admin_dashboard')

    # 10. Admin Approvals Queue
    res = client.get('/admin/approvals')
    html_appr = res.get_data(as_text=True)
    capture_html(html_appr, '10_admin_approvals_queue')

    # 11. Admin Approvals with Approve Modal Open
    html_appr_mod = html_appr.replace('id="approveModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 hidden', 'id="approveModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex')
    html_appr_mod = html_appr_mod.replace('id="approveModalEmpName" class="text-slate-900"></span>', 'id="approveModalEmpName" class="text-slate-900">John Doe (EMP-1001)</span>')
    html_appr_mod = html_appr_mod.replace('id="approveModalDays" class="font-bold text-indigo-600"></span>', 'id="approveModalDays" class="font-bold text-indigo-600">2.0</span>')
    html_appr_mod = html_appr_mod.replace('placeholder="e.g., Approved. Handover coverage arranged with team."', 'value="Approved. Coverage arranged with engineering team."')
    capture_html(html_appr_mod, '11_admin_approve_modal')

    # 12. Admin Approvals with Reject Modal Open
    html_rej_mod = html_appr.replace('id="rejectModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 hidden', 'id="rejectModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex')
    html_rej_mod = html_rej_mod.replace('id="rejectModalEmpName" class="text-slate-900"></span>', 'id="rejectModalEmpName" class="text-slate-900">Michael Chang (EMP-1003)</span>')
    html_rej_mod = html_rej_mod.replace('placeholder="e.g., Critical sprint release scheduled on these dates. Please reschedule."', 'Critical product release deadline scheduled during these dates. Please reschedule for following week.')
    capture_html(html_rej_mod, '12_admin_reject_modal')

    # 13. Admin Personnel Directory
    res = client.get('/admin/employees')
    html_emp = res.get_data(as_text=True)
    capture_html(html_emp, '13_admin_employees_directory')

    # 14. Admin Add Employee Modal Open
    html_emp_add = html_emp.replace('id="addEmployeeModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 hidden', 'id="addEmployeeModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex')
    capture_html(html_emp_add, '14_admin_add_employee_modal')

    # 15. Admin Adjust Quota Modal Open
    html_emp_q = html_emp.replace('id="quotaModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 hidden', 'id="quotaModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex')
    html_emp_q = html_emp_q.replace('id="quotaModalEmpName" class="text-slate-900"></span>', 'id="quotaModalEmpName" class="text-slate-900">Sarah Jenkins (EMP-1002)</span>')
    html_emp_q = html_emp_q.replace('placeholder="e.g. 20"', 'value="22.0"')
    capture_html(html_emp_q, '15_admin_adjust_quota_modal')

    # 16. Admin Compliance Reports & CSV Export View
    res = client.get('/admin/reports')
    capture_html(res.get_data(as_text=True), '16_admin_reports_audit')

    print("All 16 comprehensive UI state screenshots captured successfully!")

if __name__ == '__main__':
    main()
