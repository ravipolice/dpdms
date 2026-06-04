# 🔒 DATA PROTECTION RULES — DPDMS
## Chikkaballapura District Police Data Management System

> **IMPORTANT:** This file defines strict rules to protect sensitive employee data.
> These rules apply to **all AI assistants, developers, and automation tools** working on this project.

---

## 🚫 NEVER Touch These Files (AI Must Not Edit or Delete)

| File / Pattern | Why Protected |
|---|---|
| `outputs/*.xlsx` | Live employee database — only the server may read/write |
| `outputs/backups/*.xlsx` | Backup snapshots — never delete manually |
| `outputs/employee_transfer_browser_dashboard/backups/` | Rolling server backups — server auto-prunes to 30 max |

> These files are **excluded from git** via `.gitignore` to prevent accidental version control overwrites.

---

## ✅ Safe Files (AI May Edit Code Only)

| File | Purpose |
|---|---|
| `outputs/employee_transfer_browser_dashboard/index.html` | Dashboard UI |
| `outputs/employee_transfer_browser_dashboard/server.py` | Backend server |
| `outputs/Code.gs` | Google Apps Script |
| `DPDMS/index.html` | Legacy dashboard |

---

## 🛡️ Built-In Protection Layers

### 1. Automatic Backups (Server-Side)
Every time the server performs a **write operation** it first calls `backup_workbook()` which:
- Creates a timestamped copy: `District_Employee_Transfer_Database_IMPORTED_<operation>_<YYYYMMDD_HHMMSS>.xlsx`
- Automatically **prunes backups older than the last 30** to save disk space

### 2. Startup Backup
Each time `server.py` starts, it creates a `_startup_` backup immediately.

### 3. Genuinity Cross-Validation
Before any employee data is written:
- **Frontend** validates format, logic (dates, ages, overlaps) and warns about mismatches with the profile
- **Server** independently re-validates the same rules
- A mismatch warning is shown to the user and requires confirmation

### 4. Git Version Control (Code Only)
- All **source code changes** are committed to git
- Data files (`.xlsx`) are **gitignored** — git never touches them

---

## 🔄 How to Restore from Backup

1. Stop the server
2. Navigate to `outputs/backups/`
3. Find the most recent `.xlsx` file before the problem occurred
4. Copy it to `outputs/` and rename it to `District_Employee_Transfer_Database_IMPORTED.xlsx`
5. Restart the server

---

## 📋 Backup Naming Convention

```
District_Employee_Transfer_Database_IMPORTED_<label>_<YYYYMMDD>_<HHMMSS>.xlsx

Labels:
  startup  — taken when server starts
  save     — before saving/editing an employee record
  delete   — before deleting an employee record
  transfer — before saving a transfer application
  withdraw — before withdrawing a transfer application
```

---

*Last updated: 2026-06-04 | Protected by: Chikkaballapura DPDMS Data Governance Rules*
