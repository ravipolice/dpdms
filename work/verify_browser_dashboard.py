from pathlib import Path

base = Path(r"C:\Users\ravip\AndroidStudioProjects\DPDMS")
html = base / "outputs" / "Employee_Transfer_Dashboard_Browser.html"
server = base / "outputs" / "employee_transfer_browser_dashboard" / "server.py"
launcher = base / "outputs" / "employee_transfer_browser_dashboard" / "Start_Dashboard.bat"
stopper = base / "outputs" / "employee_transfer_browser_dashboard" / "Stop_Dashboard.bat"

for path in [html, server, launcher, stopper]:
    assert path.exists(), path

content = html.read_text(encoding="utf-8")
assert "window.EMBEDDED_DATA" in content
assert "District Police Data Management System" in content

print("browser dashboard verified")
