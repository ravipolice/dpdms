import importlib.util
import json
import base64
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("dashboard_server", APP_DIR / "server.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)

data = module.load_data()
html = (APP_DIR / "index.html").read_text(encoding="utf-8")

# Embed logo as base64 in standalone HTML
logo_path = APP_DIR / "logo.png"
if logo_path.exists():
    logo_base64 = base64.b64encode(logo_path.read_bytes()).decode("utf-8")
    html = html.replace('src="logo.png"', f'src="data:image/png;base64,{logo_base64}"')

payload = json.dumps(data, ensure_ascii=False)
standalone = html.replace(
    "<script>",
    f"<script>window.EMBEDDED_DATA = {payload};</script>\n  <script>",
    1,
)
(APP_DIR.parent / "Employee_Transfer_Dashboard_Browser.html").write_text(standalone, encoding="utf-8")
print(APP_DIR.parent / "Employee_Transfer_Dashboard_Browser.html")
