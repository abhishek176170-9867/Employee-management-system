import os
import subprocess
import time
from app import app
from database import get_db_connection

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'report_assets')
os.makedirs(OUTPUT_DIR, exist_ok=True)
EDGE_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

def render_page(client, route, filename, is_admin=False, is_logged_in=True):
    if is_logged_in:
        email = 'admin@company.com' if is_admin else 'john.doe@company.com'
        pwd = 'admin123' if is_admin else 'employee123'
        client.post('/login', data={'email': email, 'password': pwd})

    res = client.get(route)
    html_text = res.get_data(as_text=True)

    # Replace relative /static/ links with absolute file:/// links so Edge loads CSS/JS
    static_abs = os.path.join(os.path.dirname(__file__), 'static').replace('\\', '/')
    html_text = html_text.replace('href="/static/', f'href="file:///{static_abs}/')
    html_text = html_text.replace('src="/static/', f'src="file:///{static_abs}/')

    html_file = os.path.join(OUTPUT_DIR, f"{filename}.html")
    with open(html_file, 'w', encoding='utf-8') as f:
        f.write(html_text)

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
    print(f"Captured: {img_file}")
    return img_file

def main():
    app.config['TESTING'] = True
    client = app.test_client()

    # 1. Login Screen
    render_page(client, '/login', '01_login_portal', is_logged_in=False)

    # 2. Employee Dashboard
    render_page(client, '/dashboard', '02_employee_dashboard', is_admin=False)

    # 3. Employee Apply Leave Form
    render_page(client, '/leave/apply', '03_apply_leave', is_admin=False)

    # 4. Employee History
    render_page(client, '/leave/history', '04_leave_history', is_admin=False)

    # 5. Company Holidays
    render_page(client, '/holidays', '05_holiday_schedule', is_admin=False)

    # 6. Admin Dashboard
    render_page(client, '/admin/dashboard', '06_admin_dashboard', is_admin=True)

    # 7. Admin Approvals Queue
    render_page(client, '/admin/approvals', '07_admin_approvals', is_admin=True)

    # 8. Admin Personnel & Quotas
    render_page(client, '/admin/employees', '08_admin_employees', is_admin=True)

    # 9. Admin Reports & Export
    render_page(client, '/admin/reports', '09_admin_reports', is_admin=True)

    print("All screenshots generated successfully in", OUTPUT_DIR)

if __name__ == '__main__':
    main()
