from pathlib import Path

base = Path(r"C:\Users\ravip\AndroidStudioProjects\DPDMS")
html = (base / "outputs" / "employee_transfer_browser_dashboard" / "index.html").read_text(encoding="utf-8")
assert "localDuplicateIssues" in html
assert "duplicateWarning" in html
assert "KGID already exists" in html
assert "Mobile already exists" in html
assert "Email already exists" in html
print("live duplicate check verified")
