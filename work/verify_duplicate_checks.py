import importlib.util
import shutil
from pathlib import Path

base = Path(r"C:\Users\ravip\AndroidStudioProjects\DPDMS")
server_path = base / "outputs" / "employee_transfer_browser_dashboard" / "server.py"
source = base / "outputs" / "District_Employee_Transfer_Database_IMPORTED.xlsx"
temp = base / "work" / "duplicate_test_copy.xlsx"
shutil.copy2(source, temp)

spec = importlib.util.spec_from_file_location("dashboard_server", server_path)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
module.WORKBOOK_PATH = temp
module.BACKUP_DIR = base / "work" / "duplicate_test_backups"

base_payload = {
    "kgid": "DUPTEST001",
    "name": "Duplicate Test Employee",
    "rank": "PC",
    "designation": "PC",
    "mobile": "9999912345",
    "currentDistrict": "Chikkaballapura",
    "currentUnit": "Duplicate Test PS",
    "sourceProfile": {"mobile2": "", "isApproved": "Yes", "isAdmin": "No"},
}
base_payload["sourceProfile"]["email"] = "duplicate-test@example.com"
assert module.save_employee(base_payload)["ok"]

try:
    module.save_employee({**base_payload, "name": "Duplicate KGID"})
    raise AssertionError("duplicate KGID was not blocked")
except ValueError as exc:
    assert "Duplicate KGID" in str(exc)

try:
    module.save_employee({**base_payload, "kgid": "DUPTEST002", "name": "Duplicate Mobile"})
    raise AssertionError("duplicate mobile was not blocked")
except ValueError as exc:
    assert "Duplicate mobile" in str(exc)

try:
    module.save_employee({
        **base_payload,
        "kgid": "DUPTEST003",
        "name": "Duplicate Email",
        "mobile": "9999912346",
        "sourceProfile": {**base_payload["sourceProfile"], "mobile2": "", "email": "duplicate-test@example.com"},
    })
    raise AssertionError("duplicate email was not blocked")
except ValueError as exc:
    assert "Duplicate email" in str(exc)

edit_payload = {**base_payload, "originalKgid": "DUPTEST001", "name": "Edited Same Employee"}
assert module.save_employee(edit_payload)["ok"]

print("duplicate checks verified")
