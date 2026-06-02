# 🛡️ District Police Data Management System (DPDMS) — Road Map
**Chikkaballapura Police District | DPDMS Project**

---

## 📌 Project Goal

Build a reliable, searchable, and reportable **Employee Transfer Management System** (District Police Data Management System) that allows district police administration to track service history, calculate transfer eligibility, and generate priority transfer lists — starting with Excel/browser and growing into a full mobile/web application.

---

## 🏗️ Architecture Overview

```
Service Registers / HRMS / Mobile App
              ↓
       District_Employee_Transfer_Database.xlsx
              ↓
         Python HTTP Server (local)
              ↓
     Browser Dashboard (index.html)
          ↙           ↘
    Reports          Edit / Add
    Transfer List     Employee Records
    Retirement Due    Posting History
    Seniority List    CSV Export
              ↓
      (Future) Apps Script API
              ↓
        Android App / Web Portal
```

---

## ✅ Phase 1 — Data Collection & Structure

**Status: COMPLETE**

### Excel Workbook Structure

| Sheet | Purpose |
|---|---|
| **Employee Master** | One row per employee — all basic details |
| **Posting History** | Multiple rows per KGID — complete service history |
| **Imported Emp Profiles** | Source import from HRMS / mobile app database |

### Employee Master Fields (20 columns)

| Col | Field | Auto-Calculated |
|---|---|---|
| A | KGID | — |
| B | Employee Name | — |
| C | Rank | — |
| D | Designation | — |
| E | DOB | — |
| F | Appointment Date | — |
| G | Mobile | — |
| H | Home District | — |
| I | Retirement Date | ✅ `EOMONTH(EDATE(DOB, 720), 0)` |
| J | Current District | — |
| K | Current Sub-Division | — |
| L | Current Unit | — |
| M | Present Posting Date | — |
| N | Total Service Years | ✅ `YEARFRAC(DOA, TODAY())` |
| O | Current Station Years | ✅ `YEARFRAC(PPD, TODAY())` |
| P | Current Sub-Division Years | ✅ SUMPRODUCT from History |
| Q | District Service Years | ✅ SUMPRODUCT from History |
| R | Transfer Eligible | ✅ `IF(O>=3, "Yes", "No")` |
| S | Priority | ✅ `IF(O>=5,"High", IF(O>=3,"Medium","Low"))` |
| T | Remarks | — |

### Posting History Fields (12 columns)

| Field | Description |
|---|---|
| KGID | Links to Employee Master |
| From Date | Start of this posting |
| To Date | End (blank = current posting) |
| Police Station / Unit | Station name |
| Sub-Division | Sub-division |
| District | District name |
| Rank Held | Rank during this posting |
| Order Number | Transfer/posting order |
| Duration Years | Auto: `YEARFRAC(From, IF(To="",TODAY(),To))` |
| Overlap Check | Flags overlapping dates |
| Completeness | Flags missing fields |

---

## ✅ Phase 2 — Browser Dashboard

**Status: COMPLETE**

### Dashboard Tabs

#### 📊 Dashboard Tab
- **6 KPI Cards**: Total · Transfer Due · High Priority · Retiring in 12M · Stations · Sub-Divisions
- **Station-wise Strength** horizontal bar chart
- **Rank Distribution** bar chart

#### 🔍 Search Employees Tab
- **5 Filters**: Quick Search · Rank · Station · Category · Transfer Status
- **Employee Profile** with 6 animated collapsible sections:
  - 👤 Basic Information
  - 📍 Current Posting (incl. Sub-Division years, District years)
  - 📜 Complete Posting History table
  - 🔄 Transfer Analysis
  - 🗃 Source Profile Data
  - 📊 District Snapshot
- **Edit / Delete / Print** actions per employee

#### 📋 Reports Tab

| Report | Description |
|---|---|
| 🔴 Transfer Priority | All eligible employees sorted by years served |
| 📅 Retirement Due | Filter by 3 / 6 / 12 / 24 months window |
| 📜 Seniority List | All employees sorted by Date of Appointment |
| 🏢 Station Strength | Working strength + transfer due per station |
| 🗺️ Sub-Division Report | Staff count + due for transfer per sub-division |

All reports have **⬇ Export CSV** button.

---

## 🔄 Phase 3 — Data Entry & Validation

**Status: IN PROGRESS**

### Priority Tasks

1. **Fill DOB / DOA** for all employees (retirement and service calculations depend on this)
2. **Enter Present Posting Date** for each employee (drives transfer eligibility)
3. **Complete Sub-Division** column (required for sub-division reports)
4. **Import Posting History** from service books for employees with 3+ years service

### Validation Checklist

| Check | Tool | Status |
|---|---|---|
| Missing dates | Dashboard highlights with `—` | Done |
| Duplicate KGID | Auto-blocked on Add/Edit | Done |
| Duplicate Mobile | Auto-blocked on Add/Edit | Done |
| Duplicate Email | Auto-blocked on Add/Edit | Done |
| Overlapping postings | Column J in Posting History | Done |
| Missing fields | Column K in Posting History | Done |

---

## 🎯 Phase 4 — Colour-Coded Priority System

**Status: COMPLETE**

| Priority | Condition | Badge Colour |
|---|---|---|
| 🟢 Low | < 3 years in present station | Green |
| 🟡 Medium | 3–4.9 years | Amber |
| 🔴 High | 5+ years | Red |

The dashboard automatically assigns Priority and Transfer Eligible based on the Excel formulas.

---

## 📋 Phase 5 — Reports & Export

**Status: COMPLETE**

### Available Reports

1. **Transfer Priority List** — Filter by High / Medium / Low priority; exportable
2. **Retirement Due** — 3 / 6 / 12 / 24 / 60 month window; exportable
3. **Seniority List** — Sorted by Date of Appointment; exportable
4. **Station Strength** — Working count + % due for transfer per station
5. **Sub-Division Report** — Staff, due, high priority per sub-division

### Future Reports (Phase 7+)
- Employees never posted in a particular sub-division
- Vacancy position (requires sanctioned strength data)
- Promotion eligibility list
- District transfer summary sheet

---

## ✅ Phase 6 — Google Sheets Backend & Secure Multi-Role Authentication

**Status: COMPLETE**

### Google Sheets Integration
- Transitioned backend database from local Excel files to **Google Sheets**.
- Built a **Google Apps Script API (`Code.gs`)** that acts as the backend server.
- Supports secure cross-origin (CORS) fetch requests directly from the browser dashboard.

### Secure Multi-Role Authentication
- **Role Enforcement**:
  - **User**: Locked to self-service profile access; can only edit personal contact fields (Mobile, Email, Blood Group).
  - **Admin**: Can view dashboards, reports, and search. Can edit and add employees but cannot delete them.
  - **Super Admin**: Master control to add, edit, delete, and configure system roles/passwords.
- **Client-Side Security**: Native SHA-256 password hashing via browser Web Cryptography APIs before transmission.
- **Session Security**: 6-hour temporary authentication session tokens stored in Google Apps Script `ScriptCache` and client `localStorage`.

### Implemented Auth Screens & Flows
1. **Login Screen**: KGID and password credentials gate.
2. **Registration Screen**: Allows new users to register by entering their KGID, Mobile, and Email. Identity is verified against the master sheet records before account creation.
3. **Registration OTP Screen**: Verifies email ownership via a **100% free email OTP code** sent using Google Apps Script's `MailApp`.
4. **Forgot Password Recovery Screen**: Users enter their KGID and registered email to trigger a password-reset verification OTP.
5. **Reset Password Screen**: Verifies the reset OTP and updates the user's password in the database.
6. **Super Admin Tab ("Users & Roles")**: Direct control panel to list users, suspend accounts, reset passwords, and toggle dynamic feature permissions for the Admin role.

---

## 🚀 Phase 7 — Future Enhancements Planned

### Data Improvements
- [ ] Add **Sanctioned Strength** per station (for vacancy tracking)
- [ ] Add **Category** column: Civil / AR / CAR / Traffic / KSRP / Other
- [ ] Complete posting history for employees with 3+ years service
- [ ] Add **Sub-Division** for all current postings

### Dashboard Improvements
- [ ] Posting history entry form (add individual postings via browser)
- [ ] Batch import from CSV
- [ ] Station-wise strength chart with vacancy overlay
- [ ] Search by Sub-Division

### Export Improvements
- [ ] PDF export of transfer priority list
- [ ] Print-ready district report (all stations)
- [ ] Seniority certificate generation

---

## 📱 Phase 8 — Mobile / Web Application (Future)

### Architecture
```
Google Sheets (Cloud)
    ↓
Google Apps Script (CORS API / Code.gs)
    ↓
Web App (HTML/JS Dashboard)   or     Android App (Kotlin)
    ↓                                    ↓
Admin Portal                         Field Access (Search only)
  Edit employees                     View posting history
  Generate reports                   Transfer eligibility
  Approve transfers                  Notifications
```

### Mobile App Features
- 🔍 Search by KGID / Name / Mobile
- 📋 View complete posting history
- 📊 Transfer eligibility status
- 📄 Export transfer lists as PDF
- 🔒 Admin-only: edit, add, approve transfers

### Technology Options
| Component | Option A | Option B |
|---|---|---|
| Backend | Google Apps Script (existing) | Firebase / Google Cloud |
| Database | Google Sheets (existing) | Cloud Firestore |
| Frontend | React + Vite | Google Sites |
| Mobile | Android (Kotlin) | Progressive Web App (PWA) |

---

## 📁 Current File Structure

Consolidated under the unified `dpdms/` root folder:
```
dpdms/
├── roadmap.md                                          ← Project documentation and roadmap
├── DPDMS/
│   ├── index.html                                      ← Standalone cloud-sync dashboard (production)
│   └── vercel.json                                     ← Cloud hosting configuration
├── outputs/
│   ├── Code.gs                                         ← Google Apps Script backend code
│   ├── District_Employee_Transfer_Database_IMPORTED.xlsx ← Local Excel template
│   ├── Employee_Transfer_Dashboard_Browser.html        ← Standalone cloud-sync dashboard (compiled build)
│   ├── logo.png                                        ← Logo image file
│   └── employee_transfer_browser_dashboard/            ← Web dashboard source folder
│       ├── index.html                                  ← Source HTML layout and scripts
│       ├── server.py                                   ← Local Python HTTP development server
│       ├── build_standalone.py                         ← Standalone HTML compiler
│       ├── Start_Dashboard.bat                         ← Script to start local development server
│       └── Stop_Dashboard.bat                          ← Script to stop local development server
└── work/                                               ← Python utilities, data migration scripts, and tests
```

---

## 🕒 Timeline Estimate

| Phase | Effort | Status |
|---|---|---|
| Phase 1 — Data Structure | Done | ✅ |
| Phase 2 — Browser Dashboard | Done | ✅ |
| Phase 3 — Data Entry | 2–4 weeks | 🔄 |
| Phase 4 — Priority System | Done | ✅ |
| Phase 5 — Reports | Done | ✅ |
| Phase 6 — Cloud Backend & Auth | Done | ✅ |
| Phase 7 — Future Enhancements | 1–2 weeks | 📋 |
| Phase 8 — Mobile App | 4–8 weeks | 🔮 |

---

*Document generated: 30-05-2026 | Chikkaballapura Police District*
