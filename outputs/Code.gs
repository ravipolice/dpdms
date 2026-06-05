/**
 * Chikkaballapura District Police Data Management System (DPDMS)
 * Google Apps Script Backend (Code.gs)
 * 
 * Instructions:
 * 1. Open your Google Sheet imported from the Excel database.
 * 2. Click Extensions -> Apps Script.
 * 3. Delete any default code and paste this script.
 * 4. Create a sheet named "System Users" if it does not exist (it will auto-create on first login with default Super Admin: KGID: 12345, Password: admin123).
 * 5. Deploy as Web App (Deploy -> New Deployment -> Select "Web App" -> Execute as "Me", Access: "Anyone").
 * 6. Copy the Web App URL and paste it into your index.html.
 */

var STATION_SUBDIVISION_MAP = {
  "Chikkaballapura Town PS": "Chikkaballapura",
  "Chikkaballapura Rural PS": "Chikkaballapura",
  "Chikkaballapura Women PS": "Chikkaballapura",
  "Chikkaballapura Traffic PS": "Chikkaballapura",
  "Gowribidanur Town PS": "Chikkaballapura",
  "Gowribidanur Rural PS": "Chikkaballapura",
  "Manchenahalli PS": "Chikkaballapura",
  "Gudibanda PS": "Chikkaballapura",
  "Nandigiridhama PS": "Chikkaballapura",
  "Peresandra PS": "Chikkaballapura",
  "Sdpo Cbpura": "Chikkaballapura",
  "CPI Office Cbpura": "Chikkaballapura",
  "CPI Off Gudibande": "Chikkaballapura",
  "CPI Off Gowribidanur": "Chikkaballapura",
  "Dysp Office Cbp": "Chikkaballapura",
  "DPO": "Chikkaballapura",
  "Control Room": "Chikkaballapura",
  "DSB": "Chikkaballapura",
  "Cen PS": "Chikkaballapura",
  "DAR Chikkaballapura": "Chikkaballapura",
  "Smmc": "Chikkaballapura",
  "DCRB": "Chikkaballapura",

  "Chintamani Town PS": "Chintamani",
  "Chintamani Rural PS": "Chintamani",
  "Shidlagatta Town PS": "Chintamani",
  "Shidlaghatta Town PS": "Chintamani",
  "Sidlaghatta Town PS": "Chintamani",
  "Sidlagatta Town PS": "Chintamani",
  "Shidlaghatta Rural PS": "Chintamani",
  "Shidlagatta Rural PS": "Chintamani",
  "Sidlaghatta Rural PS": "Chintamani",
  "Sidlagatta Rural PS": "Chintamani",
  "Bagepalli PS": "Chintamani",
  "Chelur PS": "Chintamani",
  "Pathapalya PS": "Chintamani",
  "Dibburahalli PS": "Chintamani",
  "Batlahalli PS": "Chintamani",
  "Kencharlahalli PS": "Chintamani",
  "Dysp Off Cmi": "Chintamani",
  "CPI Office Cheluru": "Chintamani",
  "CPI Off Kencharlahalli": "Chintamani",
  "CPI Off Shidlaghatta": "Chintamani",
  "CPI Off Sidlaghatta": "Chintamani"
};

// CORS and response helpers
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function checkSession(token) {
  if (!token) return null;
  var cached = CacheService.getScriptCache().get(token);
  if (cached) {
    return JSON.parse(cached);
  }
  return null;
}

// ── GET REQUESTS ──────────────────────────────────────────────
function doGet(e) {
  try {
    var action = e.parameter.action;
    var token = e.parameter.token;
    var session = checkSession(token);
    
    if (action === "check-session") {
      return jsonResponse({ ok: !!session, session: session });
    }
    
    if (!session) {
      return jsonResponse({ ok: false, error: "Unauthorized. Please log in again." });
    }
    
    if (action === "data") {
      var data = loadData(session.kgid, session.role);
      return jsonResponse(data);
    }
    
    if (action === "check-duplicate") {
      var query = {
        kgid: e.parameter.kgid || "",
        originalKgid: e.parameter.originalKgid || "",
        mobile: e.parameter.mobile || "",
        mobile2: e.parameter.mobile2 || "",
        email: e.parameter.email || ""
      };
      var dup = duplicateStatus(query);
      return jsonResponse(dup);
    }
    
    if (action === "list-users" && session.role === "Super Admin") {
      return jsonResponse({ ok: true, users: listUsers() });
    }
    
    return jsonResponse({ ok: false, error: "Invalid GET action." });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ── POST REQUESTS ─────────────────────────────────────────────
function doPost(e) {
  try {
    var payload;
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = e.parameter;
    }
    
    var action = payload.action;
    
    if (action === "login") {
      return handleLogin(payload.kgid, payload.passwordHash);
    }
    
    if (action === "google-login") {
      return handleGoogleLogin(payload.idToken);
    }
    
    if (action === "send-registration-otp") {
      return handleSendRegistrationOtp(payload.kgid, payload.mobile, payload.email, payload.passwordHash);
    }
    
    if (action === "verify-registration-otp") {
      return handleVerifyRegistrationOtp(payload.kgid, payload.otp);
    }
    
    if (action === "send-reset-otp") {
      return handleSendResetOtp(payload.kgid, payload.email);
    }
    
    if (action === "verify-reset-otp") {
      return handleVerifyResetOtp(payload.kgid, payload.otp, payload.passwordHash);
    }
    
    var token = payload.token;
    var session = checkSession(token);
    if (!session) {
      return jsonResponse({ ok: false, error: "Unauthorized. Please log in again." });
    }
    
    if (action === "save") {
      return handleSaveEmployee(payload.employee, session);
    }
    
    if (action === "save-transfer-request") {
      return handleSaveTransferRequest(payload.transferRequest, session);
    }
    
    if (action === "delete-transfer-request") {
      return handleDeleteTransferRequest(payload, session);
    }
    
    if (action === "delete") {
      if (session.role !== "Super Admin") {
        return jsonResponse({ ok: false, error: "Access Denied. Super Admin role required." });
      }
      return handleDeleteEmployee(payload.kgid);
    }
    
    if (action === "manage-user") {
      if (session.role !== "Super Admin") {
        return jsonResponse({ ok: false, error: "Access Denied. Super Admin role required." });
      }
      return handleManageUser(payload.userPayload);
    }
    
    if (action === "save-permissions") {
      if (session.role !== "Super Admin") {
        return jsonResponse({ ok: false, error: "Access Denied. Super Admin role required." });
      }
      return handleSavePermissions(payload.rolePermissions);
    }
    
    return jsonResponse({ ok: false, error: "Invalid POST action." });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ── LOGIN HANDLER ─────────────────────────────────────────────
function handleLogin(kgid, passwordHash) {
  kgid = String(kgid || "").trim();
  passwordHash = String(passwordHash || "").trim();
  if (!kgid || !passwordHash) {
    return jsonResponse({ ok: false, error: "KGID and Password are required." });
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName("System Users");
  
  // Create sheet with default super admin if missing
  if (!usersSheet) {
    usersSheet = ss.insertSheet("System Users");
    usersSheet.appendRow(["KGID", "PasswordHash", "Role", "Email", "Status"]);
    // Default pwd: admin123 -> SHA-256 hash below
    usersSheet.appendRow(["12345", "240713a1a08e96dd00cbe6e05342a98379c6fa076cf47e8e58319e64e52541c8", "Super Admin", "", "Active"]);
  }
  
  var data = usersSheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var kIdx = headers.indexOf("KGID");
  var pIdx = headers.indexOf("PasswordHash");
  var rIdx = headers.indexOf("Role");
  var sIdx = headers.indexOf("Status");
  
  var userRow = -1;
  var targetUserExists = false;
  for (var i = 1; i < data.length; i++) {
    var checkKgid = String(data[i][kIdx]).trim();
    if (checkKgid === kgid) {
      userRow = i;
    }
    if (checkKgid === "1953036") {
      targetUserExists = true;
    }
  }
  
  // Seed requested Super Admin user if not present
  if (!targetUserExists) {
    usersSheet.appendRow(["1953036", "db1bddcedf91b4fd0146205d4f238117a98d35f25bd8ce3840c46eead9f1b966", "Super Admin", "", "Active"]);
    data = usersSheet.getDataRange().getValues();
    if (kgid === "1953036") {
      userRow = data.length - 1;
    }
  }
  
  // If user doesn't exist in System Users sheet, they must register first.
  if (userRow === -1) {
    return jsonResponse({ ok: false, error: "This KGID has not been registered yet. Please click 'Register Here' to set up your account." });
  }
  
  var dbHash = String(data[userRow][pIdx]).trim();
  var dbRole = String(data[userRow][rIdx]).trim();
  var dbStatus = String(data[userRow][sIdx]).trim();
  
  if (dbStatus === "Inactive") {
    return jsonResponse({ ok: false, error: "Your account is deactivated." });
  }
  
  if (dbHash === passwordHash) {
    var token = Utilities.getUuid();
    
    // Get name if possible
    var name = kgid;
    var masterSheet = ss.getSheetByName("Employee Master");
    if (masterSheet) {
      var masterData = masterSheet.getDataRange().getValues();
      var mHeaders = masterData[0].map(function(h) { return String(h).trim(); });
      var mKIdx = mHeaders.indexOf("KGID");
      var mNIdx = mHeaders.indexOf("Employee Name");
      for (var j = 1; j < masterData.length; j++) {
        if (String(masterData[j][mKIdx]).trim() === kgid) {
          name = masterData[j][mNIdx];
          break;
        }
      }
    }
    
    var session = { kgid: kgid, role: dbRole, name: name };
    CacheService.getScriptCache().put(token, JSON.stringify(session), 21600); // 6 hours
    return jsonResponse({ ok: true, token: token, role: dbRole, kgid: kgid, name: name });
  }
  
  return jsonResponse({ ok: false, error: "Incorrect Password." });
}

// ── UTILITIES ─────────────────────────────────────────────────
function asText(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) {
    return formatDate(val);
  }
  return String(val).trim();
}

function normalizeStation(name) {
  var n = asText(name).trim();
  var nu = n.toUpperCase();
  if (nu === "DSB") return "DSB";
  if (nu === "DCRB") return "DCRB";
  if (nu === "DPO") return "DPO";
  return n;
}

function formatDate(dateObj) {
  if (!dateObj) return "";
  var d = new Date(dateObj);
  var day = ("0" + d.getDate()).slice(-2);
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  var year = d.getFullYear();
  return day + "-" + month + "-" + year;
}

function yearsBetween(start, end) {
  if (!start) return "";
  var startDate = new Date(start);
  if (isNaN(startDate.getTime())) return "";
  var endDate = end ? new Date(end) : new Date();
  var diffTime = endDate.getTime() - startDate.getTime();
  var diffDays = diffTime / (1000 * 60 * 60 * 24);
  return Math.round((diffDays / 365.25) * 10) / 10;
}

function parseDateStr(str) {
  if (!str) return null;
  var parts = str.split("-");
  if (parts.length === 3) {
    // dd-mm-yyyy to yyyy, mm-1, dd
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(str);
}

function getRetirementDate(dob) {
  if (!dob) return "";
  var dobDate = new Date(dob);
  if (isNaN(dobDate.getTime())) return "";
  var retYear = dobDate.getFullYear() + 60;
  var retMonth = dobDate.getMonth() + 1;
  // Last day of month
  var endOfMonth = new Date(retYear, retMonth, 0);
  return endOfMonth;
}

function getSheetRowsAsDicts(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var list = [];
  for (var r = 1; r < data.length; r++) {
    var dict = {};
    var hasVal = false;
    for (var c = 0; c < headers.length; c++) {
      var val = data[r][c];
      if (val !== null && val !== "") hasVal = true;
      dict[headers[c]] = val;
    }
    if (hasVal) list.push(dict);
  }
  return list;
}

// ── LOAD DATABASE ─────────────────────────────────────────────
function loadData(loggedKgid, role) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var masterSheet = ss.getSheetByName("Employee Master");
  var historySheet = ss.getSheetByName("Posting History");
  var importedSheet = ss.getSheetByName("Imported Emp Profiles");
  var importedDicts = [];
  if (importedSheet) {
    var data = importedSheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var modified = false;
    if (headers.indexOf("email") === -1) {
      importedSheet.getRange(1, headers.length + 1).setValue("email");
      headers.push("email");
      modified = true;
    }
    if (headers.indexOf("photoUrl") === -1) {
      importedSheet.getRange(1, headers.length + 1).setValue("photoUrl");
      headers.push("photoUrl");
      modified = true;
    }
    importedDicts = getSheetRowsAsDicts(importedSheet);
  }
  
  var transferRequestsSheet = ss.getSheetByName("Transfer Requests");
  var transferRequests = [];
  if (transferRequestsSheet) {
    var trDicts = getSheetRowsAsDicts(transferRequestsSheet);
    trDicts.forEach(function(tr) {
      var trKgid = asText(tr["KGID"]);
      if (trKgid) {
        transferRequests.push({
          "kgid": trKgid,
          "name": asText(tr["Employee Name"]),
          "rank": asText(tr["Rank"]),
          "currentStation": asText(tr["Current Station"]),
          "preference1": asText(tr["Preference 1"]),
          "preference2": asText(tr["Preference 2"]),
          "preference3": asText(tr["Preference 3"]),
          "preference4": asText(tr["Preference 4"]) || "",
          "preference5": asText(tr["Preference 5"]) || "",
          "transferCategory": asText(tr["Transfer Category"]) || "",
          "reason": asText(tr["Reason"]),
          "remarks": asText(tr["Remarks"]),
          "applicationDate": asText(tr["Application Date"]),
          "status": asText(tr["Status"]) || "Pending",
          "approvedStation": asText(tr["Approved Station"]) || ""
        });
      }
    });
  }
  
  var masterDicts = getSheetRowsAsDicts(masterSheet);
  var historyDicts = getSheetRowsAsDicts(historySheet);
  
  // Index imported by kgid
  var sourceByKgid = {};
  importedDicts.forEach(function(row) {
    var k = asText(row["kgid"]);
    if (k) sourceByKgid[k] = row;
  });
  
  // Index posting history by kgid
  var historyByKgid = {};
  historyDicts.forEach(function(item) {
    var kgid = asText(item["KGID"]);
    if (!kgid) return;
    
    var fromDate = item["From Date"];
    var toDate = item["To Date"];
    
    item["_fromDisplay"] = formatDate(fromDate);
    item["_toDisplay"] = toDate ? formatDate(toDate) : "Present";
    item["_durationYears"] = yearsBetween(fromDate, toDate);
    
    if (!historyByKgid[kgid]) historyByKgid[kgid] = [];
    historyByKgid[kgid].push(item);
  });
  
  var employees = [];
  masterDicts.forEach(function(emp) {
    var kgid = asText(emp["KGID"]);
    if (!kgid) return;
    
    // For standard users, only return their own record!
    if (role === "User" && kgid !== loggedKgid) return;
    
    var dob = emp["DOB"];
    var doa = emp["Appointment Date"];
    var currentSince = emp["Present Posting Date"];
    
    var totalService = yearsBetween(doa);
    var currentStationYears = yearsBetween(currentSince);
    
    var priority = (currentStationYears !== "" && currentStationYears >= 5) ? "High" : 
                   (currentStationYears !== "" && currentStationYears >= 3) ? "Medium" : "Low";
    var eligible = (currentStationYears !== "" && currentStationYears >= 3) ? "Yes" : "No";
    
    var postings = historyByKgid[kgid] || [];
    var source = sourceByKgid[kgid] || {};
    
    var longest = 0;
    postings.forEach(function(p) {
      var d = Number(p["_durationYears"] || 0);
      if (d > longest) longest = d;
    });
    
    var retirementValue = emp["Retirement Date"] || getRetirementDate(dob);
    var category = asText(emp["Category"]);
    var typeOfTransfer = asText(emp["Type of Transfer"]) || "Regular";
    
    var currentSubdivRaw = asText(emp["Current Sub-Division"]);
    if (!currentSubdivRaw) {
      currentSubdivRaw = STATION_SUBDIVISION_MAP[normalizeStation(emp["Current Unit"])] || "";
    }
    
    var currentDistrictRaw = asText(emp["Current District"]);
    
    // Calculate subdiv years & district years
    var subdivYears = 0;
    var districtYears = 0;
    postings.forEach(function(p) {
      var duration = Number(p["_durationYears"] || 0);
      var pSub = asText(p["Sub-Division"]) || STATION_SUBDIVISION_MAP[asText(p["Police Station / Unit"])] || "";
      var pDist = asText(p["District"]);
      if (currentSubdivRaw && pSub === currentSubdivRaw) {
        subdivYears += duration;
      }
      if (currentDistrictRaw && pDist === currentDistrictRaw) {
        districtYears += duration;
      }
    });
    subdivYears = Math.round(subdivYears * 10) / 10;
    districtYears = Math.round(districtYears * 10) / 10;
    
    var retDateStr = formatDate(retirementValue);
    var retirementMonthsLeft = "";
    if (retDateStr) {
      try {
        var parts = retDateStr.split("-");
        var rd = new Date(parts[2], parts[1] - 1, parts[0]);
        var today = new Date();
        retirementMonthsLeft = (rd.getFullYear() - today.getFullYear()) * 12 + (rd.getMonth() - today.getMonth());
      } catch(e) {}
    }
    
    employees.push({
      "kgid": kgid,
      "name": asText(emp["Employee Name"]),
      "rank": asText(emp["Rank"]),
      "designation": asText(emp["Designation"]),
      "dob": formatDate(dob),
      "doa": formatDate(doa),
      "mobile": asText(emp["Mobile"]),
      "homeDistrict": asText(emp["Home District"]),
      "retirementDate": formatDate(retirementValue),
      "currentDistrict": asText(emp["Current District"]),
      "currentSubDivision": currentSubdivRaw,
      "currentUnit": normalizeStation(emp["Current Unit"]),
      "currentSince": formatDate(currentSince),
      "totalServiceYears": totalService,
      "currentStationYears": currentStationYears,
      "transferEligible": eligible,
      "priority": priority,
      "longestPostingYears": longest,
      "postings": postings.map(function(p) {
        return {
          "from": p["_fromDisplay"],
          "to": p["_toDisplay"],
          "station": normalizeStation(p["Police Station / Unit"]),
          "subDivision": asText(p["Sub-Division"]) || STATION_SUBDIVISION_MAP[normalizeStation(p["Police Station / Unit"])] || "",
          "district": asText(p["District"]),
          "rank": asText(p["Rank Held"]),
          "orderNumber": asText(p["Order Number"]),
          "durationYears": p["_durationYears"]
        };
      }),
      "sourceProfile": {
        "metalNumber": asText(source["metalNumber"]),
        "bloodGroup": asText(source["bloodGroup"]),
        "email": asText(source["email"]),
        "mobile2": asText(source["mobile2"]),
        "isAdmin": asText(source["isAdmin"]),
        "isApproved": asText(source["isApproved"]),
        "createdAt": asText(source["createdAt"]),
        "updatedAt": asText(source["updatedAt"]),
        "height": asText(source["height"]),
        "weight": asText(source["weight"]),
        "caste": asText(source["caste"]),
        "subCaste": asText(source["subCaste"]),
        "familyDetails": asText(source["familyDetails"]),
        "photoUrl": asText(source["photoUrl"] || source["photo"] || source["photo_url"] || source["image"] || ""),
        "lockedFields": asText(source["lockedFields"] || "")
      },
      "category": category,
      "typeOfTransfer": typeOfTransfer,
      "subDivisionYears": subdivYears,
      "districtServiceYears": districtYears,
      "retirementMonthsLeft": retirementMonthsLeft
    });
  });
  
  // Sort employees by name
  employees.sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });
  
  var summary = {};
  if (role !== "User") {
    // Generate summaries only for Admins / Super Admins
    var stationsMap = {};
    var ranksMap = {};
    var subdivisionsMap = {};
    
    var transferDueCount = 0;
    var highPriorityCount = 0;
    var retirementDue12mCount = 0;
    
    employees.forEach(function(e) {
      if (e.transferEligible === "Yes") transferDueCount++;
      if (e.priority === "High") highPriorityCount++;
      if (typeof e.retirementMonthsLeft === "number" && e.retirementMonthsLeft >= 0 && e.retirementMonthsLeft <= 12) {
        retirementDue12mCount++;
      }
      
      if (e.currentUnit) {
        stationsMap[e.currentUnit] = (stationsMap[e.currentUnit] || 0) + 1;
      }
      if (e.rank) {
        ranksMap[e.rank] = (ranksMap[e.rank] || 0) + 1;
      }
      if (e.currentSubDivision) {
        subdivisionsMap[e.currentSubDivision] = (subdivisionsMap[e.currentSubDivision] || 0) + 1;
      }
    });
    
    summary = {
      "totalEmployees": employees.length,
      "transferDue": transferDueCount,
      "highPriority": highPriorityCount,
      "retirementDue": employees.filter(function(e) { return e.retirementDate; }).length,
      "retirementDue12m": retirementDue12mCount,
      "stations": Object.keys(stationsMap).sort().map(function(k) { return { name: k, working: stationsMap[k] }; }),
      "ranks": Object.keys(ranksMap).sort().map(function(k) { return { name: k, working: ranksMap[k] }; }),
      "subdivisions": Object.keys(subdivisionsMap).sort().map(function(k) { return { name: k, working: subdivisionsMap[k] }; })
    };
  }
  
  var permissions = getRolePermissions();
  
  return {
    "workbook": "Google Sheets Live",
    "editable": true,
    "lastLoaded": formatDate(new Date()) + " " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss"),
    "employees": employees,
    "summary": summary,
    "permissions": permissions,
    "transferRequests": transferRequests
  };
}

// ── SAVE EMPLOYEE HANDLER ─────────────────────────────────────
function handleSaveEmployee(payload, session) {
  var kgid = String(payload.kgid || "").trim();
  var origKgid = String(payload.originalKgid || "").trim();
  var isAdd = !origKgid;
  
  var perms = getRolePermissions();
  if (session.role === "User") {
    // Regular users can only edit their OWN profile, and cannot ADD new ones!
    if (isAdd || kgid !== session.kgid || origKgid !== session.kgid) {
      return jsonResponse({ ok: false, error: "Access Denied. You can only edit your own details." });
    }
  } else if (session.role === "Admin") {
    var adminPerms = perms["Admin"] || { canAddEmployee: true, canEditEmployee: true };
    if (isAdd && !adminPerms.canAddEmployee) {
      return jsonResponse({ ok: false, error: "Access Denied. Admins are currently restricted from adding employees." });
    }
    if (!isAdd && !adminPerms.canEditEmployee) {
      return jsonResponse({ ok: false, error: "Access Denied. Admins are currently restricted from editing employee details." });
    }
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Employee Master");
  var historySheet = ss.getSheetByName("Posting History");
  var importedSheet = ss.getSheetByName("Imported Emp Profiles");
  
  var masterData = masterSheet.getDataRange().getValues();
  var masterHeaders = masterData[0].map(function(h) { return String(h).trim(); });
  
  var targetRow = -1;
  if (!isAdd) {
    // Find existing row by KGID
    var kIdx = masterHeaders.indexOf("KGID");
    for (var r = 1; r < masterData.length; r++) {
      if (String(masterData[r][kIdx]).trim() === origKgid) {
        targetRow = r + 1; // 1-indexed for Sheet row
        break;
      }
    }
    if (targetRow === -1) {
      return jsonResponse({ ok: false, error: "Employee row not found in Employee Master sheet." });
    }
  } else {
    // Check duplicates
    var dup = duplicateStatus({ kgid: kgid, mobile: payload.mobile });
    if (dup.duplicate) {
      return jsonResponse({ ok: false, error: dup.issues.map(function(i) { return i.message; }).join(", ") });
    }
    targetRow = masterSheet.getLastRow() + 1;
    masterSheet.appendRow([kgid]); // seed row
  }
  
  // Set values based on role limits
  // (User can only change mobile, email, bloodGroup)
  var src = payload.sourceProfile || {};
  if (!isAdd && importedSheet) {
    var impData = importedSheet.getDataRange().getValues();
    var impHeaders = impData[0].map(function(h) { return String(h).trim(); });
    var impKIdx = impHeaders.indexOf("kgid");
    var impRow = -1;
    for (var r = 1; r < impData.length; r++) {
      if (String(impData[r][impKIdx]).trim() === origKgid) {
        impRow = r + 1;
        break;
      }
    }
    
    var impCols = {};
    impHeaders.forEach(function(h, idx) { impCols[h] = idx + 1; });
    
    var existingLockedFields = "";
    if (impRow !== -1 && impCols["lockedFields"]) {
      existingLockedFields = String(importedSheet.getRange(impRow, impCols["lockedFields"]).getValue() || "");
    }
    var lockedList = existingLockedFields.split(",").map(function(s) { return s.trim(); }).filter(Boolean);
    
    if (session.role === "User") {
      // Revert locked fields to existing database values
      if (lockedList.indexOf("mobile") !== -1) {
        if (impRow !== -1 && impCols["mobile1"]) {
          payload.mobile = String(importedSheet.getRange(impRow, impCols["mobile1"]).getValue() || "");
        }
      }
      var fieldsToCheck = ["mobile2", "bloodGroup", "email", "height", "weight", "caste", "subCaste", "familyDetails", "photoUrl"];
      fieldsToCheck.forEach(function(fKey) {
        if (lockedList.indexOf(fKey) !== -1) {
          if (impRow !== -1 && impCols[fKey]) {
            src[fKey] = String(importedSheet.getRange(impRow, impCols[fKey]).getValue() || "");
          } else {
            src[fKey] = "";
          }
        }
      });
      src.lockedFields = existingLockedFields;
    } else {
      src.lockedFields = src.lockedFields || "";
    }
  }

  var masterCols = {};
  masterHeaders.forEach(function(h, idx) { masterCols[h] = idx + 1; });
  
  if (session.role === "User") {
    // Restrict edits to contact info
    if (masterCols["Mobile"]) masterSheet.getRange(targetRow, masterCols["Mobile"]).setValue(payload.mobile);
  } else {
    // Admin & Super Admin have full master edit rights
    if (masterCols["KGID"]) masterSheet.getRange(targetRow, masterCols["KGID"]).setValue(kgid);
    if (masterCols["Employee Name"]) masterSheet.getRange(targetRow, masterCols["Employee Name"]).setValue(payload.name);
    if (masterCols["Rank"]) masterSheet.getRange(targetRow, masterCols["Rank"]).setValue(payload.rank);
    if (masterCols["Designation"]) masterSheet.getRange(targetRow, masterCols["Designation"]).setValue(payload.designation);
    if (masterCols["DOB"]) masterSheet.getRange(targetRow, masterCols["DOB"]).setValue(parseDateStr(payload.dob));
    if (masterCols["Appointment Date"]) masterSheet.getRange(targetRow, masterCols["Appointment Date"]).setValue(parseDateStr(payload.doa));
    if (masterCols["Mobile"]) masterSheet.getRange(targetRow, masterCols["Mobile"]).setValue(payload.mobile);
    if (masterCols["Home District"]) masterSheet.getRange(targetRow, masterCols["Home District"]).setValue(payload.homeDistrict);
    if (masterCols["Current District"]) masterSheet.getRange(targetRow, masterCols["Current District"]).setValue(payload.currentDistrict || "Chikkaballapura");
    if (masterCols["Current Sub-Division"]) masterSheet.getRange(targetRow, masterCols["Current Sub-Division"]).setValue(payload.currentSubDivision);
    if (masterCols["Current Unit"]) masterSheet.getRange(targetRow, masterCols["Current Unit"]).setValue(normalizeStation(payload.currentUnit));
    if (masterCols["Present Posting Date"]) masterSheet.getRange(targetRow, masterCols["Present Posting Date"]).setValue(parseDateStr(payload.currentSince));
    if (masterCols["Category"]) masterSheet.getRange(targetRow, masterCols["Category"]).setValue(payload.category);
    if (!masterCols["Type of Transfer"]) {
      var nextCol = masterHeaders.length + 1;
      masterSheet.getRange(1, nextCol).setValue("Type of Transfer");
      masterCols["Type of Transfer"] = nextCol;
      masterHeaders.push("Type of Transfer");
    }
    masterSheet.getRange(targetRow, masterCols["Type of Transfer"]).setValue(payload.typeOfTransfer || "Regular");
    if (masterCols["Remarks"]) masterSheet.getRange(targetRow, masterCols["Remarks"]).setValue(payload.remarks);
    
    // Set formulas
    if (masterCols["Retirement Date"]) masterSheet.getRange(targetRow, masterCols["Retirement Date"]).setFormula('=IF(E' + targetRow + '="","",EOMONTH(EDATE(E' + targetRow + ',60*12),0))');
    if (masterCols["Total Service Years"]) masterSheet.getRange(targetRow, masterCols["Total Service Years"]).setFormula('=IF(F' + targetRow + '="","",ROUND(YEARFRAC(F' + targetRow + ',TODAY()),1))');
    if (masterCols["Current Station Years"]) masterSheet.getRange(targetRow, masterCols["Current Station Years"]).setFormula('=IF(M' + targetRow + '="","",ROUND(YEARFRAC(M' + targetRow + ',TODAY()),1))');
    if (masterCols["Current Sub-Division Years"]) masterSheet.getRange(targetRow, masterCols["Current Sub-Division Years"]).setFormula('=IF(A' + targetRow + '="","",ROUND(SUMPRODUCT((\'Posting History\'!$A$2:$A$1000=A' + targetRow + ')*(\'Posting History\'!$E$2:$E$1000=K' + targetRow + ')*(\'Posting History\'!$B$2:$B$1000<>"")*(IF(\'Posting History\'!$C$2:$C$1000="",TODAY(),\'Posting History\'!$C$2:$C$1000)-\'Posting History\'!$B$2:$B$1000))/365.25,1))');
    if (masterCols["District Service Years"]) masterSheet.getRange(targetRow, masterCols["District Service Years"]).setFormula('=IF(A' + targetRow + '="","",ROUND(SUMPRODUCT((\'Posting History\'!$A$2:$A$1000=A' + targetRow + ')*(\'Posting History\'!$F$2:$F$1000=J' + targetRow + ')*(\'Posting History\'!$B$2:$B$1000<>"")*(IF(\'Posting History\'!$C$2:$C$1000="",TODAY(),\'Posting History\'!$C$2:$C$1000)-\'Posting History\'!$B$2:$B$1000))/365.25,1))');
    if (masterCols["Transfer Eligible"]) masterSheet.getRange(targetRow, masterCols["Transfer Eligible"]).setFormula('=IF(O' + targetRow + '="","",IF(O{row}>=3,"Yes","No"))'.replace('{row}', targetRow));
    if (masterCols["Priority"]) masterSheet.getRange(targetRow, masterCols["Priority"]).setFormula('=IF(O' + targetRow + '="","",IF(O' + targetRow + '>=5,"High",IF(O' + targetRow + '>=3,"Medium","Low")))');
  }
  
  // Save Posting History if Admin or Super Admin and postings payload is present
  if (session.role !== "User" && payload.postings) {
    var historyData = historySheet.getDataRange().getValues();
    var historyHeaders = historyData[0].map(function(h) { return String(h).trim(); });
    var hKIdx = historyHeaders.indexOf("KGID");
    
    // Delete all existing postings in Posting History for original_kgid / kgid
    var searchKgid = origKgid ? origKgid : kgid;
    for (var r = historyData.length - 1; r >= 1; r--) {
      if (String(historyData[r][hKIdx]).trim() === searchKgid) {
        historySheet.deleteRow(r + 1);
      }
    }
    
    var hCols = {};
    historyHeaders.forEach(function(h, idx) { hCols[h] = idx + 1; });
    
    payload.postings.forEach(function(p) {
      var hRow = historySheet.getLastRow() + 1;
      historySheet.appendRow([kgid]); // seed row
      
      var pTo = (p.to && p.to !== "Present") ? parseDateStr(p.to) : null;
      
      if (hCols["KGID"]) historySheet.getRange(hRow, hCols["KGID"]).setValue(kgid);
      if (hCols["From Date"]) historySheet.getRange(hRow, hCols["From Date"]).setValue(parseDateStr(p.from));
      if (hCols["To Date"]) {
        if (pTo) {
          historySheet.getRange(hRow, hCols["To Date"]).setValue(pTo);
        } else {
          historySheet.getRange(hRow, hCols["To Date"]).clearContent();
        }
      }
      if (hCols["Police Station / Unit"]) historySheet.getRange(hRow, hCols["Police Station / Unit"]).setValue(normalizeStation(p.station));
      if (hCols["Sub-Division"]) historySheet.getRange(hRow, hCols["Sub-Division"]).setValue(p.subDivision);
      if (hCols["District"]) historySheet.getRange(hRow, hCols["District"]).setValue(p.district || "Chikkaballapura");
      if (hCols["Rank Held"]) historySheet.getRange(hRow, hCols["Rank Held"]).setValue(p.rank);
      if (hCols["Order Number"]) historySheet.getRange(hRow, hCols["Order Number"]).setValue(p.orderNumber);
      if (hCols["Notes"]) historySheet.getRange(hRow, hCols["Notes"]).setValue("Maintained from Google Sheets DPDMS.");
      
      // Set formulas
      if (hCols["Duration Years"]) historySheet.getRange(hRow, hCols["Duration Years"]).setFormula('=IF(B' + hRow + '="","",ROUND((IF(C' + hRow + '="",TODAY(),C' + hRow + ')-B' + hRow + ')/365.25,1))');
      if (hCols["Overlap Check"]) historySheet.getRange(hRow, hCols["Overlap Check"]).setFormula('=IF(A' + hRow + '="","",IF(COUNTIFS($A:$A,A' + hRow + ',$B:$B,"<="&IF(C' + hRow + '="",TODAY(),C' + hRow + '),$C:$C,">="&B' + hRow + ')>1,"Review","OK"))');
      if (hCols["Missing Date Check"]) historySheet.getRange(hRow, hCols["Missing Date Check"]).setFormula('=IF(A' + hRow + '="","",IF(OR(B' + hRow + '="",D' + hRow + '="",E' + hRow + '="",F' + hRow + '="",G{row}=""),"Missing","OK"))'.replace('{row}', hRow));
    });
  }
  
  // Save Imported Emp Profiles (contact details email, mobile2, metal number, bloodGroup)
  var src = payload.sourceProfile || {};
  if (importedSheet) {
    var impData = importedSheet.getDataRange().getValues();
    var impHeaders = impData[0].map(function(h) { return String(h).trim(); });
    
    var impKIdx = impHeaders.indexOf("kgid");
    var impRow = -1;
    for (var r = 1; r < impData.length; r++) {
      if (String(impData[r][impKIdx]).trim() === origKgid) {
        impRow = r + 1;
        break;
      }
    }
    
    if (impRow === -1) {
      impRow = importedSheet.getLastRow() + 1;
      importedSheet.appendRow([kgid]);
    }
    
    var impCols = {};
    impHeaders.forEach(function(h, idx) { impCols[h] = idx + 1; });
    
    if (impCols["kgid"]) importedSheet.getRange(impRow, impCols["kgid"]).setValue(kgid);
    if (impCols["name"]) importedSheet.getRange(impRow, impCols["name"]).setValue(payload.name);
    if (impCols["mobile1"]) importedSheet.getRange(impRow, impCols["mobile1"]).setValue(payload.mobile);
    if (impCols["mobile2"]) importedSheet.getRange(impRow, impCols["mobile2"]).setValue(src.mobile2);
    if (impCols["rank"]) importedSheet.getRange(impRow, impCols["rank"]).setValue(payload.rank);
    if (impCols["station"]) importedSheet.getRange(impRow, impCols["station"]).setValue(normalizeStation(payload.currentUnit));
    if (impCols["district"]) importedSheet.getRange(impRow, impCols["district"]).setValue(payload.currentDistrict || "Chikkaballapura");
    if (impCols["metalNumber"]) importedSheet.getRange(impRow, impCols["metalNumber"]).setValue(src.metalNumber);
    if (impCols["bloodGroup"]) importedSheet.getRange(impRow, impCols["bloodGroup"]).setValue(src.bloodGroup);
    // Auto-create missing schema columns if they are not in impCols
    var schemaCols = ["email", "photoUrl", "lockedFields", "caste", "subCaste", "height", "weight", "familyDetails"];
    schemaCols.forEach(function(col) {
      if (!impCols[col]) {
        var nextCol = impHeaders.length + 1;
        importedSheet.getRange(1, nextCol).setValue(col);
        impCols[col] = nextCol;
        impHeaders.push(col);
      }
    });

    importedSheet.getRange(impRow, impCols["email"]).setValue(src.email || "");
    if (impCols["isAdmin"]) importedSheet.getRange(impRow, impCols["isAdmin"]).setValue(src.isAdmin || "No");
    if (impCols["isApproved"]) importedSheet.getRange(impRow, impCols["isApproved"]).setValue(src.isApproved || "Yes");
    importedSheet.getRange(impRow, impCols["height"]).setValue(src.height || "");
    importedSheet.getRange(impRow, impCols["weight"]).setValue(src.weight || "");
    importedSheet.getRange(impRow, impCols["caste"]).setValue(src.caste || "");
    importedSheet.getRange(impRow, impCols["subCaste"]).setValue(src.subCaste || "");
    importedSheet.getRange(impRow, impCols["familyDetails"]).setValue(src.familyDetails || "");
    importedSheet.getRange(impRow, impCols["photoUrl"]).setValue(src.photoUrl || "");
    importedSheet.getRange(impRow, impCols["lockedFields"]).setValue(src.lockedFields || "");
    
    if (impCols["updatedAt"]) importedSheet.getRange(impRow, impCols["updatedAt"]).setValue(formatDate(new Date()) + " " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm"));
  }
  
  return jsonResponse({ ok: true, message: "Saved employee " + payload.name + " (" + kgid + ")" });
}

// ── DELETE EMPLOYEE HANDLER ───────────────────────────────────
function handleDeleteEmployee(kgid) {
  kgid = String(kgid || "").trim();
  if (!kgid) return jsonResponse({ ok: false, error: "KGID is required." });
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Employee Master");
  var historySheet = ss.getSheetByName("Posting History");
  var importedSheet = ss.getSheetByName("Imported Emp Profiles");
  var usersSheet = ss.getSheetByName("System Users");
  
  // Remove from Master
  var masterData = masterSheet.getDataRange().getValues();
  var kIdx = masterData[0].indexOf("KGID");
  for (var r = masterData.length - 1; r >= 1; r--) {
    if (String(masterData[r][kIdx]).trim() === kgid) {
      masterSheet.deleteRow(r + 1);
    }
  }
  
  // Remove from History
  var historyData = historySheet.getDataRange().getValues();
  var hKIdx = historyData[0].indexOf("KGID");
  for (var r = historyData.length - 1; r >= 1; r--) {
    if (String(historyData[r][hKIdx]).trim() === kgid) {
      historySheet.deleteRow(r + 1);
    }
  }
  
  // Remove from Imported
  if (importedSheet) {
    var impData = importedSheet.getDataRange().getValues();
    var impKIdx = impData[0].indexOf("kgid");
    for (var r = impData.length - 1; r >= 1; r--) {
      if (String(impData[r][impKIdx]).trim() === kgid) {
        importedSheet.deleteRow(r + 1);
      }
    }
  }
  
  // Remove from System Users
  if (usersSheet) {
    var usersData = usersSheet.getDataRange().getValues();
    var uKIdx = usersData[0].indexOf("KGID");
    for (var r = usersData.length - 1; r >= 1; r--) {
      if (String(usersData[r][uKIdx]).trim() === kgid) {
        usersSheet.deleteRow(r + 1);
      }
    }
  }
  
  return jsonResponse({ ok: true, message: "Deleted employee " + kgid });
}

// ── DUPLICATE STATUS CHECKER ──────────────────────────────────
function duplicateStatus(query) {
  var kgid = String(query.kgid || "").trim();
  var originalKgid = String(query.originalKgid || "").trim();
  var mobile = String(query.mobile || "").trim();
  var isAdd = !originalKgid;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Employee Master");
  var masterData = masterSheet.getDataRange().getValues();
  var headers = masterData[0].map(function(h) { return String(h).trim(); });
  var kIdx = headers.indexOf("KGID");
  var nIdx = headers.indexOf("Employee Name");
  var mIdx = headers.indexOf("Mobile");
  
  var issues = [];
  function sameRecord(rowKgid) {
    return !isAdd && (rowKgid === originalKgid || rowKgid === kgid);
  }
  
  for (var r = 1; r < masterData.length; r++) {
    var rowKgid = String(masterData[r][kIdx]).trim();
    var rowName = String(masterData[r][nIdx]).trim();
    var rowMobile = String(masterData[r][mIdx]).trim();
    
    if (sameRecord(rowKgid)) continue;
    
    if (kgid && rowKgid === kgid) {
      issues.push({ field: "kgid", message: "KGID already exists: " + kgid + " - " + rowName });
    }
    if (mobile && rowMobile === mobile) {
      issues.push({ field: "mobile", message: "Mobile already exists: " + mobile + " - " + rowName });
    }
  }
  
  return { ok: true, duplicate: issues.length > 0, issues: issues };
}

// ── LIST USERS ────────────────────────────────────────────────
function listUsers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName("System Users");
  if (!usersSheet) return [];
  var rows = getSheetRowsAsDicts(usersSheet);
  return rows.map(function(r) {
    return {
      kgid: asText(r["KGID"]),
      role: asText(r["Role"]),
      email: asText(r["Email"]),
      status: asText(r["Status"])
    };
  });
}

// ── MANAGE USER (Super Admin Actions) ──────────────────────────
function handleManageUser(payload) {
  var targetKgid = String(payload.kgid || "").trim();
  var action = payload.action; // "save", "delete", "create"
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName("System Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("System Users");
    usersSheet.appendRow(["KGID", "PasswordHash", "Role", "Email", "Status"]);
  }
  
  var usersData = usersSheet.getDataRange().getValues();
  var headers = usersData[0].map(function(h) { return String(h).trim(); });
  var kIdx = headers.indexOf("KGID");
  
  var targetRow = -1;
  for (var r = 1; r < usersData.length; r++) {
    if (String(usersData[r][kIdx]).trim() === targetKgid) {
      targetRow = r + 1;
      break;
    }
  }
  
  var uCols = {};
  headers.forEach(function(h, idx) { uCols[h] = idx + 1; });
  
  if (action === "create" || action === "save") {
    if (targetRow === -1) {
      targetRow = usersSheet.getLastRow() + 1;
      usersSheet.appendRow([targetKgid]);
    }
    
    if (payload.passwordHash && uCols["PasswordHash"]) {
      usersSheet.getRange(targetRow, uCols["PasswordHash"]).setValue(payload.passwordHash);
    }
    if (payload.role && uCols["Role"]) {
      usersSheet.getRange(targetRow, uCols["Role"]).setValue(payload.role);
    }
    if (uCols["Email"]) {
      usersSheet.getRange(targetRow, uCols["Email"]).setValue(payload.email || "");
    }
    if (payload.status && uCols["Status"]) {
      usersSheet.getRange(targetRow, uCols["Status"]).setValue(payload.status);
    }
    return jsonResponse({ ok: true, message: "User account saved: " + targetKgid });
  }
  
  if (action === "delete") {
    if (targetRow !== -1) {
      usersSheet.deleteRow(targetRow);
      return jsonResponse({ ok: true, message: "User account deleted: " + targetKgid });
    }
    return jsonResponse({ ok: false, error: "User account not found." });
  }
  
  return jsonResponse({ ok: false, error: "Invalid manage-user action." });
}

// ── ROLE PERMISSIONS HELPERS ───────────────────────────────────
function getRolePermissions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Role Permissions");
  if (!sheet) {
    sheet = ss.insertSheet("Role Permissions");
    sheet.appendRow(["Role", "CanViewDashboard", "CanViewReports", "CanAddEmployee", "CanEditEmployee"]);
    sheet.appendRow(["Admin", "Yes", "Yes", "Yes", "Yes"]);
    sheet.appendRow(["User", "No", "No", "No", "No"]);
    sheet.appendRow(["Super Admin", "Yes", "Yes", "Yes", "Yes"]);
  }
  var data = sheet.getDataRange().getValues();
  var perms = {};
  for (var i = 1; i < data.length; i++) {
    var roleName = String(data[i][0]).trim();
    perms[roleName] = {
      canViewDashboard: String(data[i][1]).trim() === "Yes",
      canViewReports: String(data[i][2]).trim() === "Yes",
      canAddEmployee: String(data[i][3]).trim() === "Yes",
      canEditEmployee: String(data[i][4]).trim() === "Yes"
    };
  }
  return perms;
}

function handleSavePermissions(rolePermissions) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Role Permissions");
  if (!sheet) {
    sheet = ss.insertSheet("Role Permissions");
    sheet.appendRow(["Role", "CanViewDashboard", "CanViewReports", "CanAddEmployee", "CanEditEmployee"]);
  }
  var data = sheet.getDataRange().getValues();
  var targetRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === "Admin") {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
    sheet.appendRow(["Admin"]);
  }
  
  sheet.getRange(targetRow, 1).setValue("Admin");
  sheet.getRange(targetRow, 2).setValue(rolePermissions.canViewDashboard ? "Yes" : "No");
  sheet.getRange(targetRow, 3).setValue(rolePermissions.canViewReports ? "Yes" : "No");
  sheet.getRange(targetRow, 4).setValue(rolePermissions.canAddEmployee ? "Yes" : "No");
  sheet.getRange(targetRow, 5).setValue(rolePermissions.canEditEmployee ? "Yes" : "No");
  
  return jsonResponse({ ok: true, message: "Admin role permissions updated successfully." });
}

// ── REGISTRATION & OTP HANDLERS ───────────────────────────────
function handleSendRegistrationOtp(kgid, mobile, email, passwordHash) {
  kgid = String(kgid || "").trim();
  mobile = String(mobile || "").trim();
  email = String(email || "").trim();
  passwordHash = String(passwordHash || "").trim();
  
  if (!kgid || !mobile || !email || !passwordHash) {
    return jsonResponse({ ok: false, error: "KGID, Mobile, Email, and Password are all required." });
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Check if user is already registered in System Users
  var usersSheet = ss.getSheetByName("System Users");
  if (usersSheet) {
    var usersData = usersSheet.getDataRange().getValues();
    var uKIdx = usersData[0].indexOf("KGID");
    for (var r = 1; r < usersData.length; r++) {
      if (String(usersData[r][uKIdx]).trim() === kgid) {
        return jsonResponse({ ok: false, error: "This KGID is already registered. Please log in." });
      }
    }
  }
  
  // 2. Check if KGID exists in Employee Master
  var masterSheet = ss.getSheetByName("Employee Master");
  if (!masterSheet) {
    return jsonResponse({ ok: false, error: "Employee database sheet not found." });
  }
  
  var masterData = masterSheet.getDataRange().getValues();
  var mHeaders = masterData[0].map(function(h) { return String(h).trim(); });
  var mKIdx = mHeaders.indexOf("KGID");
  var mMIdx = mHeaders.indexOf("Mobile");
  
  var targetRow = -1;
  for (var r = 1; r < masterData.length; r++) {
    if (String(masterData[r][mKIdx]).trim() === kgid) {
      targetRow = r + 1;
      break;
    }
  }
  
  if (targetRow === -1) {
    return jsonResponse({ ok: false, error: "KGID not found in Employee Master database." });
  }
  
  // 3. Verify mobile number
  var rawMasterMobile = asText(masterData[targetRow - 1][mMIdx]);
  var normMasterMobile = rawMasterMobile.replace(/\D/g, "");
  var normInputMobile = mobile.replace(/\D/g, "");
  
  if (normMasterMobile !== "") {
    // Compare last 10 digits to handle country code or formatting differences
    var m10 = normMasterMobile.slice(-10);
    var i10 = normInputMobile.slice(-10);
    if (m10 !== i10) {
      return jsonResponse({ ok: false, error: "Entered mobile number does not match our records for this KGID." });
    }
  }
  
  // 4. Generate 6-digit OTP code
  var otp = String(Math.floor(100000 + Math.random() * 900000));
  
  // 5. Store in script cache (valid for 15 minutes)
  var cacheData = {
    otp: otp,
    kgid: kgid,
    mobile: mobile,
    email: email,
    passwordHash: passwordHash
  };
  CacheService.getScriptCache().put("reg_" + kgid, JSON.stringify(cacheData), 900);
  
  // 6. Send verification email via Google MailApp
  try {
    var subject = "DPDMS Registration Verification OTP";
    var body = "Hello,\n\n" +
               "You are setting up a sign-in account for the Chikkaballapura District Police Data Management System (DPDMS).\n\n" +
               "Your 6-digit verification code is:\n\n" +
               "   " + otp + "\n\n" +
               "This code is valid for 15 minutes. If you did not request this, please ignore this email.\n\n" +
               "Regards,\n" +
               "Chikkaballapura Police Admin";
    MailApp.sendEmail(email, subject, body);
  } catch (mailErr) {
    return jsonResponse({ ok: false, error: "Failed to send email verification: " + mailErr.toString() });
  }
  
  return jsonResponse({ ok: true, message: "Verification code sent to " + email });
}

function handleVerifyRegistrationOtp(kgid, otp) {
  kgid = String(kgid || "").trim();
  otp = String(otp || "").trim();
  
  if (!kgid || !otp) {
    return jsonResponse({ ok: false, error: "KGID and OTP code are required." });
  }
  
  // 1. Fetch cached registration details
  var cached = CacheService.getScriptCache().get("reg_" + kgid);
  if (!cached) {
    return jsonResponse({ ok: false, error: "Verification session expired or not found. Please register again." });
  }
  
  var data = JSON.parse(cached);
  if (data.otp !== otp) {
    return jsonResponse({ ok: false, error: "Incorrect verification code." });
  }
  
  // OTP is correct! Complete registration
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 2. Add to System Users
  var usersSheet = ss.getSheetByName("System Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("System Users");
    usersSheet.appendRow(["KGID", "PasswordHash", "Role", "Email", "Status"]);
  }
  usersSheet.appendRow([data.kgid, data.passwordHash, "User", data.email, "Active"]);
  
  // 3. Update Mobile in Employee Master
  var masterSheet = ss.getSheetByName("Employee Master");
  if (masterSheet) {
    var masterData = masterSheet.getDataRange().getValues();
    var mHeaders = masterData[0].map(function(h) { return String(h).trim(); });
    var mKIdx = mHeaders.indexOf("KGID");
    var mMIdx = mHeaders.indexOf("Mobile");
    
    var targetRow = -1;
    for (var r = 1; r < masterData.length; r++) {
      if (String(masterData[r][mKIdx]).trim() === data.kgid) {
        targetRow = r + 1;
        break;
      }
    }
    
    if (targetRow !== -1) {
      masterSheet.getRange(targetRow, mMIdx + 1).setValue(data.mobile);
    }
  }
  
  // 4. Update Imported Emp Profiles if exists
  var importedSheet = ss.getSheetByName("Imported Emp Profiles");
  if (importedSheet) {
    var impData = importedSheet.getDataRange().getValues();
    var impHeaders = impData[0].map(function(h) { return String(h).trim(); });
    var impKIdx = impHeaders.indexOf("kgid");
    var impRow = -1;
    for (var r = 1; r < impData.length; r++) {
      if (String(impData[r][impKIdx]).trim() === data.kgid) {
        impRow = r + 1;
        break;
      }
    }
    
    if (impRow === -1) {
      impRow = importedSheet.getLastRow() + 1;
      importedSheet.appendRow([data.kgid]);
    }
    
    var impCols = {};
    impHeaders.forEach(function(h, idx) { impCols[h] = idx + 1; });
    
    if (impCols["kgid"]) importedSheet.getRange(impRow, impCols["kgid"]).setValue(data.kgid);
    if (impCols["mobile1"]) importedSheet.getRange(impRow, impCols["mobile1"]).setValue(data.mobile);
    if (impCols["email"]) importedSheet.getRange(impRow, impCols["email"]).setValue(data.email);
    if (impCols["updatedAt"]) importedSheet.getRange(impRow, impCols["updatedAt"]).setValue(formatDate(new Date()) + " " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm"));
  }
  
  // 5. Clean up cache
  CacheService.getScriptCache().remove("reg_" + kgid);
  
  return jsonResponse({ ok: true, message: "Registration successful! You can now log in with your password." });
}

// ── PASSWORD RESET OTP HANDLERS ───────────────────────────────
function handleSendResetOtp(kgid, email) {
  kgid = String(kgid || "").trim();
  email = String(email || "").trim();
  
  if (!kgid || !email) {
    return jsonResponse({ ok: false, error: "KGID and Email are required." });
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName("System Users");
  if (!usersSheet) {
    return jsonResponse({ ok: false, error: "User accounts database not found." });
  }
  
  var usersData = usersSheet.getDataRange().getValues();
  var headers = usersData[0].map(function(h) { return String(h).trim(); });
  var kIdx = headers.indexOf("KGID");
  var eIdx = headers.indexOf("Email");
  var sIdx = headers.indexOf("Status");
  
  var targetRow = -1;
  for (var r = 1; r < usersData.length; r++) {
    if (String(usersData[r][kIdx]).trim() === kgid) {
      targetRow = r + 1;
      break;
    }
  }
  
  if (targetRow === -1) {
    return jsonResponse({ ok: false, error: "User account not found for this KGID." });
  }
  
  var dbEmail = String(usersData[targetRow - 1][eIdx]).trim().toLowerCase();
  var dbStatus = String(usersData[targetRow - 1][sIdx]).trim();
  
  if (dbStatus === "Inactive") {
    return jsonResponse({ ok: false, error: "This user account is suspended. Please contact support." });
  }
  
  if (dbEmail !== "" && dbEmail !== email.toLowerCase()) {
    return jsonResponse({ ok: false, error: "Entered email address does not match our records." });
  }
  
  // Generate OTP code
  var otp = String(Math.floor(100000 + Math.random() * 900000));
  
  // Cache reset code
  var cacheData = { otp: otp, kgid: kgid, email: email };
  CacheService.getScriptCache().put("reset_" + kgid, JSON.stringify(cacheData), 900);
  
  // Send reset email
  try {
    var subject = "DPDMS Password Reset OTP Verification";
    var body = "Hello,\n\n" +
               "A password reset request was made for your Chikkaballapura DPDMS account (KGID: " + kgid + ").\n\n" +
               "Your password reset verification code is:\n\n" +
               "   " + otp + "\n\n" +
               "This code is valid for 15 minutes. If you did not request this change, please ignore this email.\n\n" +
               "Regards,\n" +
               "Chikkaballapura Police Admin";
    MailApp.sendEmail(email, subject, body);
  } catch (mailErr) {
    return jsonResponse({ ok: false, error: "Failed to send reset email: " + mailErr.toString() });
  }
  
  return jsonResponse({ ok: true, message: "Reset code successfully sent to " + email });
}

function handleVerifyResetOtp(kgid, otp, newPasswordHash) {
  kgid = String(kgid || "").trim();
  otp = String(otp || "").trim();
  newPasswordHash = String(newPasswordHash || "").trim();
  
  if (!kgid || !otp || !newPasswordHash) {
    return jsonResponse({ ok: false, error: "KGID, OTP, and New Password are required." });
  }
  
  var cached = CacheService.getScriptCache().get("reset_" + kgid);
  if (!cached) {
    return jsonResponse({ ok: false, error: "Verification session expired or not found. Please try again." });
  }
  
  var data = JSON.parse(cached);
  if (data.otp !== otp) {
    return jsonResponse({ ok: false, error: "Incorrect verification code." });
  }
  
  // Verification successful! Update password
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersSheet = ss.getSheetByName("System Users");
  if (!usersSheet) {
    return jsonResponse({ ok: false, error: "User accounts database not found." });
  }
  
  var usersData = usersSheet.getDataRange().getValues();
  var headers = usersData[0].map(function(h) { return String(h).trim(); });
  var kIdx = headers.indexOf("KGID");
  var pIdx = headers.indexOf("PasswordHash");
  
  var targetRow = -1;
  for (var r = 1; r < usersData.length; r++) {
    if (String(usersData[r][kIdx]).trim() === kgid) {
      targetRow = r + 1;
      break;
    }
  }
  
  if (targetRow === -1) {
    return jsonResponse({ ok: false, error: "User account not found." });
  }
  
  usersSheet.getRange(targetRow, pIdx + 1).setValue(newPasswordHash);
  
  // Clean up cache
  CacheService.getScriptCache().remove("reset_" + kgid);
  
  return jsonResponse({ ok: true, message: "Password updated successfully." });
}

function handleGoogleLogin(idToken) {
  idToken = String(idToken || "").trim();
  if (!idToken) return jsonResponse({ ok: false, error: "ID Token is required." });
  
  try {
    var response = UrlFetchApp.fetch("https://oauth2.googleapis.com/tokeninfo?id_token=" + idToken, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return jsonResponse({ ok: false, error: "Failed to verify Google identity. Invalid token." });
    }
    
    var payload = JSON.parse(response.getContentText());
    var email = String(payload.email || "").trim().toLowerCase();
    var emailVerified = payload.email_verified;
    
    if (!email || (emailVerified !== "true" && emailVerified !== true)) {
      return jsonResponse({ ok: false, error: "Google email is not verified." });
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var usersSheet = ss.getSheetByName("System Users");
    if (!usersSheet) {
      usersSheet = ss.insertSheet("System Users");
      usersSheet.appendRow(["KGID", "PasswordHash", "Role", "Email", "Status"]);
      usersSheet.appendRow(["12345", "240713a1a08e96dd00cbe6e05342a98379c6fa076cf47e8e58319e64e52541c8", "Super Admin", "", "Active"]);
      usersSheet.appendRow(["1953036", "db1bddcedf91b4fd0146205d4f238117a98d35f25bd8ce3840c46eead9f1b966", "Super Admin", "", "Active"]);
    }
    
    var data = usersSheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var kIdx = headers.indexOf("KGID");
    var eIdx = headers.indexOf("Email");
    var rIdx = headers.indexOf("Role");
    var sIdx = headers.indexOf("Status");
    
    var userRow = -1;
    for (var i = 1; i < data.length; i++) {
      var dbEmail = String(data[i][eIdx]).trim().toLowerCase();
      if (dbEmail === email) {
        userRow = i;
        break;
      }
    }
    
    if (userRow === -1) {
      var importedSheet = ss.getSheetByName("Imported Emp Profiles");
      var foundKgid = "";
      var foundName = "";
      var foundIsAdmin = "No";
      
      if (importedSheet) {
        var impData = importedSheet.getDataRange().getValues();
        var impHeaders = impData[0].map(function(h) { return String(h).trim(); });
        var impKIdx = impHeaders.indexOf("kgid");
        var impEIdx = impHeaders.indexOf("email");
        var impNIdx = impHeaders.indexOf("name");
        var impAIdx = impHeaders.indexOf("isAdmin");
        for (var j = 1; j < impData.length; j++) {
          if (String(impData[j][impEIdx]).trim().toLowerCase() === email) {
            foundKgid = String(impData[j][impKIdx]).trim();
            foundName = String(impData[j][impNIdx]).trim();
            if (impAIdx !== -1) foundIsAdmin = String(impData[j][impAIdx]).trim();
            break;
          }
        }
      }
      
      if (!foundKgid) {
        return jsonResponse({ ok: false, error: "Your Google email (" + email + ") is not registered in our records. Please register your account first." });
      }
      
      var autoRole = foundIsAdmin === "Yes" ? "Admin" : "User";
      if (foundKgid === "1953036" || foundKgid === "12345" || email === "ravipolice@gmail.com") {
        autoRole = "Super Admin";
      }
      
      usersSheet.appendRow([foundKgid, "", autoRole, email, "Active"]);
      var token = Utilities.getUuid();
      var session = { kgid: foundKgid, role: autoRole, name: foundName };
      CacheService.getScriptCache().put(token, JSON.stringify(session), 21600); // 6 hours
      return jsonResponse({ ok: true, token: token, role: autoRole, kgid: foundKgid, name: foundName });
    }
    
    var kgid = String(data[userRow][kIdx]).trim();
    var dbRole = String(data[userRow][rIdx]).trim();
    if (email === "ravipolice@gmail.com") {
      dbRole = "Super Admin";
    }
    var dbStatus = String(data[userRow][sIdx]).trim();
    
    if (dbStatus === "Inactive") {
      return jsonResponse({ ok: false, error: "Your account is deactivated." });
    }
    
    var name = kgid;
    var masterSheet = ss.getSheetByName("Employee Master");
    if (masterSheet) {
      var masterData = masterSheet.getDataRange().getValues();
      var mHeaders = masterData[0].map(function(h) { return String(h).trim(); });
      var mKIdx = mHeaders.indexOf("KGID");
      var mNIdx = mHeaders.indexOf("Employee Name");
      for (var j = 1; j < masterData.length; j++) {
        if (String(masterData[j][mKIdx]).trim() === kgid) {
          name = masterData[j][mNIdx];
          break;
        }
      }
    }
    
    var token = Utilities.getUuid();
    var session = { kgid: kgid, role: dbRole, name: name };
    CacheService.getScriptCache().put(token, JSON.stringify(session), 21600); // 6 hours
    return jsonResponse({ ok: true, token: token, role: dbRole, kgid: kgid, name: name });
  } catch(err) {
    return jsonResponse({ ok: false, error: "Google verification failed: " + err.toString() });
  }
}


function handleSaveTransferRequest(payload, session) {
  var kgid = String(payload.kgid || "").trim();
  if (!kgid) {
    return jsonResponse({ ok: false, error: "KGID is required." });
  }
  
  if (session.role === "User" && kgid !== session.kgid) {
    return jsonResponse({ ok: false, error: "Access Denied. You can only apply for your own transfer." });
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transfer Requests");
  if (!sheet) {
    sheet = ss.insertSheet("Transfer Requests");
    sheet.appendRow(["KGID", "Employee Name", "Rank", "Current Station", "Preference 1", "Preference 2", "Preference 3", "Reason", "Remarks", "Application Date", "Status", "Approved Station"]);
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var kIdx = headers.indexOf("KGID");
  
  var targetRow = -1;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][kIdx]).trim() === kgid) {
      targetRow = r + 1;
      break;
    }
  }
  
  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
    sheet.appendRow([kgid]);
  }
  
  var cols = {};
  headers.forEach(function(h, idx) { cols[h] = idx + 1; });
  
  // Set values
  if (cols["KGID"]) sheet.getRange(targetRow, cols["KGID"]).setValue(kgid);
  if (cols["Employee Name"]) sheet.getRange(targetRow, cols["Employee Name"]).setValue(payload.name);
  if (cols["Rank"]) sheet.getRange(targetRow, cols["Rank"]).setValue(payload.rank);
  if (cols["Current Station"]) sheet.getRange(targetRow, cols["Current Station"]).setValue(normalizeStation(payload.currentStation));
  if (cols["Preference 1"]) sheet.getRange(targetRow, cols["Preference 1"]).setValue(normalizeStation(payload.preference1));
  if (cols["Preference 2"]) sheet.getRange(targetRow, cols["Preference 2"]).setValue(normalizeStation(payload.preference2));
  if (cols["Preference 3"]) sheet.getRange(targetRow, cols["Preference 3"]).setValue(normalizeStation(payload.preference3));
  if (cols["Preference 4"]) sheet.getRange(targetRow, cols["Preference 4"]).setValue(normalizeStation(payload.preference4 || ""));
  if (cols["Preference 5"]) sheet.getRange(targetRow, cols["Preference 5"]).setValue(normalizeStation(payload.preference5 || ""));
  if (cols["Transfer Category"]) sheet.getRange(targetRow, cols["Transfer Category"]).setValue(payload.transferCategory || "");
  if (cols["Reason"]) sheet.getRange(targetRow, cols["Reason"]).setValue(payload.reason);
  if (cols["Remarks"]) sheet.getRange(targetRow, cols["Remarks"]).setValue(payload.remarks);
  if (cols["Application Date"]) sheet.getRange(targetRow, cols["Application Date"]).setValue(payload.applicationDate || formatDate(new Date()));
  
  // Admin can change status; user sets it to Pending
  var status = payload.status || "Pending";
  if (session.role === "User") {
    status = "Pending";
  }
  if (cols["Status"]) sheet.getRange(targetRow, cols["Status"]).setValue(status);
  
  if (!cols["Approved Station"]) {
    var nextCol = headers.length + 1;
    sheet.getRange(1, nextCol).setValue("Approved Station");
    cols["Approved Station"] = nextCol;
    headers.push("Approved Station");
  }
  if (cols["Approved Station"]) {
    sheet.getRange(targetRow, cols["Approved Station"]).setValue(payload.approvedStation || "");
  }
  
  return jsonResponse({ ok: true, message: "Saved transfer request for " + payload.name + " (" + kgid + ")" });
}

function handleDeleteTransferRequest(payload, session) {
  var kgid = String(payload.kgid || "").trim();
  if (!kgid) {
    return jsonResponse({ ok: false, error: "KGID is required." });
  }
  
  if (session.role === "User" && kgid !== session.kgid) {
    return jsonResponse({ ok: false, error: "Access Denied." });
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transfer Requests");
  if (!sheet) {
    return jsonResponse({ ok: false, error: "Transfer Requests sheet not found." });
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var kIdx = headers.indexOf("KGID");
  
  var targetRow = -1;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][kIdx]).trim() === kgid) {
      targetRow = r + 1;
      break;
    }
  }
  
  if (targetRow !== -1) {
    sheet.deleteRow(targetRow);
    return jsonResponse({ ok: true, message: "Withdrew transfer request for KGID " + kgid });
  }
  
  return jsonResponse({ ok: false, error: "Transfer request not found." });
}

// ── CUSTOM MENU FOR IMPORTING DATA ────────────────────────────
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("DPDMS Admin")
    .addItem("Import Data from External Sheet", "importExternalEmployeeData")
    .addItem("Split Remarks to Separate Columns", "migrateRemarksToSeparateColumns")
    .addToUi();
}

function importExternalEmployeeData() {
  var ui = SpreadsheetApp.getUi();
  var defaultUrl = "https://docs.google.com/spreadsheets/d/199WoCqPzhSB-zS2qkM_TdmjYJ8oNFtHASLLyXf0VMBk/edit";
  
  var response = ui.prompt(
    "Import External Employee Data",
    "Please enter the URL of the external Google Sheet you want to import from.\n\n" +
    "Suggested sheets:\n" +
    "1. First Sheet: https://docs.google.com/spreadsheets/d/199WoCqPzhSB-zS2qkM_TdmjYJ8oNFtHASLLyXf0VMBk/edit\n" +
    "2. Second Sheet: https://docs.google.com/spreadsheets/d/1Js_V1wmQdETvG02rWqIT4y3Pac3TnEVzVsQ1tOly7U0/edit\n" +
    "3. Height/Weight Sheet: https://docs.google.com/spreadsheets/d/15-VhnB4wVhvA4dM78YTPYkyA_0Gu8J6186cHKz_G9IQ/edit",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var externalUrl = response.getResponseText().trim();
  if (!externalUrl) {
    externalUrl = defaultUrl;
  }
  
  try {
    var masterSs = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = masterSs.getSheetByName("Employee Master");
    var importedSheet = masterSs.getSheetByName("Imported Emp Profiles");
    
    if (!masterSheet) {
      ui.alert("Error: 'Employee Master' sheet not found in the active spreadsheet.");
      return;
    }
    
    // Open source sheet
    var srcSs = SpreadsheetApp.openByUrl(externalUrl);
    var srcSheet = srcSs.getSheets()[0]; // get first sheet
    var srcData = srcSheet.getDataRange().getValues();
    if (srcData.length <= 1) {
      ui.alert("Error: The source spreadsheet is empty or has no data rows.");
      return;
    }
    
    var srcHeaders = srcData[0].map(function(h) { return String(h).trim(); });
    var srcKgidIdx = -1;
    for (var i = 0; i < srcHeaders.length; i++) {
      var hLower = srcHeaders[i].toLowerCase();
      if (hLower === "kgid" || hLower === "k.g.i.d" || hLower === "employee kgid" || hLower === "employee k.g.i.d") {
        srcKgidIdx = i;
        break;
      }
    }
    
    if (srcKgidIdx === -1) {
      ui.alert("Error: Could not find a 'KGID' column in the source sheet. Headers found: " + srcHeaders.join(", "));
      return;
    }
    
    // Map headers for Employee Master
    var masterData = masterSheet.getDataRange().getValues();
    var masterHeaders = masterData[0].map(function(h) { return String(h).trim(); });
    var masterKgidIdx = masterHeaders.indexOf("KGID");
    if (masterKgidIdx === -1) {
      ui.alert("Error: 'KGID' column not found in 'Employee Master' sheet.");
      return;
    }
    
    // Map headers for Imported Emp Profiles
    var importedHeaders = [];
    var importedKgidIdx = -1;
    var importedData = [];
    if (importedSheet) {
      importedData = importedSheet.getDataRange().getValues();
      if (importedData.length > 0) {
        importedHeaders = importedData[0].map(function(h) { return String(h).trim(); });
        importedKgidIdx = importedHeaders.indexOf("kgid");
      }
    }
    
    var masterHeaderMap = {}; // srcColIndex -> masterColIndex (1-indexed)
    var importedHeaderMap = {}; // srcColIndex -> importedColIndex (1-indexed)
    
    srcHeaders.forEach(function(srcHeader, srcIdx) {
      if (srcIdx === srcKgidIdx) return; // skip KGID as it is target key
      
      var sLower = srcHeader.toLowerCase();
      
      // 1. Try to match Imported Emp Profiles first (bloodGroup, height, weight, caste, etc.)
      if (importedSheet && importedKgidIdx !== -1) {
        var matchedImportedIdx = -1;
        for (var iIdx = 0; iIdx < importedHeaders.length; iIdx++) {
          var iHeader = importedHeaders[iIdx];
          var iLower = iHeader.toLowerCase();
          if (iLower === sLower ||
              (sLower === "blood group" && iLower === "bloodgroup") ||
              (sLower === "bloodgroup" && iLower === "bloodgroup") ||
              (sLower === "mobile 2" && iLower === "mobile2") ||
              (sLower === "photo" && iLower === "photourl") ||
              (sLower === "photo url" && iLower === "photourl") ||
              (sLower === "image" && iLower === "photourl") ||
              (sLower === "height (in cms)" && iLower === "height") ||
              (sLower === "weight (in kgs)" && iLower === "weight") ||
              (sLower === "caste / category" && iLower === "caste") ||
              (sLower === "sub caste" && iLower === "subcaste")) {
            matchedImportedIdx = iIdx;
            break;
          }
        }
        if (matchedImportedIdx !== -1) {
          importedHeaderMap[srcIdx] = matchedImportedIdx + 1; // 1-indexed
          return;
        }
        
        // Dynamically add height or weight if they are in source but not in imported sheet
        if (sLower === "height" || sLower === "height (in cms)" || sLower === "height(in cms)" || sLower === "height in cms") {
          var newCol = importedHeaders.length + 1;
          importedSheet.getRange(1, newCol).setValue("height");
          importedHeaders.push("height");
          importedHeaderMap[srcIdx] = newCol;
          return;
        }
        if (sLower === "weight" || sLower === "weight (in kgs)" || sLower === "weight(in kgs)" || sLower === "weight in kgs") {
          var newCol = importedHeaders.length + 1;
          importedSheet.getRange(1, newCol).setValue("weight");
          importedHeaders.push("weight");
          importedHeaderMap[srcIdx] = newCol;
          return;
        }
      }
      
      // 2. Try to find matching master header (case-insensitive)
      var matchedMasterIdx = -1;
      for (var mIdx = 0; mIdx < masterHeaders.length; mIdx++) {
        var mHeader = masterHeaders[mIdx];
        var mLower = mHeader.toLowerCase();
        if (mLower === sLower ||
            (sLower === "employee name" && mLower === "employee name") ||
            (sLower === "name" && mLower === "employee name") ||
            (sLower === "appointment date" && mLower === "appointment date") ||
            (sLower === "date of appointment" && mLower === "appointment date") ||
            (sLower === "date of birth" && mLower === "dob") ||
            (sLower === "unit" && mLower === "current unit") ||
            (sLower === "subdivision" && mLower === "current sub-division") ||
            (sLower === "sub-division" && mLower === "current sub-division") ||
            (sLower === "sub division" && mLower === "current sub-division")) {
          matchedMasterIdx = mIdx;
          break;
        }
      }
      if (matchedMasterIdx !== -1) {
        masterHeaderMap[srcIdx] = matchedMasterIdx + 1; // 1-indexed for sheet column
      }
    });
    
    // Index existing master rows by KGID
    var masterKgidRowMap = {};
    for (var r = 1; r < masterData.length; r++) {
      var k = String(masterData[r][masterKgidIdx]).trim();
      if (k) {
        masterKgidRowMap[k] = r + 1;
      }
    }
    
    // Index existing imported rows by KGID
    var importedKgidRowMap = {};
    if (importedSheet && importedKgidIdx !== -1) {
      for (var r = 1; r < importedData.length; r++) {
        var k = String(importedData[r][importedKgidIdx]).trim();
        if (k) {
          importedKgidRowMap[k] = r + 1;
        }
      }
    }
    
    var addedCount = 0;
    var updatedCount = 0;
    var importedUpdatedCount = 0;
    
    // Columns mapping
    var mCols = {};
    masterHeaders.forEach(function(h, idx) { mCols[h] = idx + 1; });
    
    for (var r = 1; r < srcData.length; r++) {
      var row = srcData[r];
      var kgid = String(row[srcKgidIdx] || "").trim();
      if (!kgid) continue;
      
      var hasMasterUpdates = Object.keys(masterHeaderMap).length > 0;
      var hasImportedUpdates = Object.keys(importedHeaderMap).length > 0;
      
      // 1. Update Employee Master
      if (hasMasterUpdates) {
        var destMasterRow = masterKgidRowMap[kgid];
        var isNewMaster = !destMasterRow;
        
        if (isNewMaster) {
          destMasterRow = masterSheet.getLastRow() + 1;
          masterSheet.getRange(destMasterRow, masterKgidIdx + 1).setValue(kgid);
          masterKgidRowMap[kgid] = destMasterRow;
          addedCount++;
        } else {
          updatedCount++;
        }
        
        // Update mapped columns
        for (var srcIdxStr in masterHeaderMap) {
          var srcIdx = parseInt(srcIdxStr);
          var destCol = masterHeaderMap[srcIdx];
          var val = row[srcIdx];
          if (val !== null && val !== "") {
            var destHeader = masterHeaders[destCol - 1];
            if (destHeader === "DOB" || destHeader === "Appointment Date" || destHeader === "Present Posting Date") {
              if (typeof val === "string") {
                val = parseDateStr(val);
              }
            }
            masterSheet.getRange(destMasterRow, destCol).setValue(val);
          }
        }
        
        // Set default formulas for new rows
        if (isNewMaster) {
          if (mCols["Retirement Date"]) masterSheet.getRange(destMasterRow, mCols["Retirement Date"]).setFormula('=IF(E' + destMasterRow + '="","",EOMONTH(EDATE(E' + destMasterRow + ',60*12),0))');
          if (mCols["Total Service Years"]) masterSheet.getRange(destMasterRow, mCols["Total Service Years"]).setFormula('=IF(F' + destMasterRow + '="","",ROUND(YEARFRAC(F' + destMasterRow + ',TODAY()),1))');
          if (mCols["Current Station Years"]) masterSheet.getRange(destMasterRow, mCols["Current Station Years"]).setFormula('=IF(M' + destMasterRow + '="","",ROUND(YEARFRAC(M' + destMasterRow + ',TODAY()),1))');
          if (mCols["Current Sub-Division Years"]) masterSheet.getRange(destMasterRow, mCols["Current Sub-Division Years"]).setFormula('=IF(A' + destMasterRow + '="","",ROUND(SUMPRODUCT((\'Posting History\'!$A$2:$A$1000=A' + destMasterRow + ')*(\'Posting History\'!$E$2:$E$1000=K' + destMasterRow + ')*(\'Posting History\'!$B$2:$B$1000<>"")*(IF(\'Posting History\'!$C$2:$C$1000="",TODAY(),\'Posting History\'!$C$2:$C$1000)-\'Posting History\'!$B$2:$B$1000))/365.25,1))');
          if (mCols["District Service Years"]) masterSheet.getRange(destMasterRow, mCols["District Service Years"]).setFormula('=IF(A' + destMasterRow + '="","",ROUND(SUMPRODUCT((\'Posting History\'!$A$2:$A$1000=A' + destMasterRow + ')*(\'Posting History\'!$F$2:$F$1000=J' + destMasterRow + ')*(\'Posting History\'!$B$2:$B$1000<>"")*(IF(\'Posting History\'!$C$2:$C$1000="",TODAY(),\'Posting History\'!$C$2:$C$1000)-\'Posting History\'!$B$2:$B$1000))/365.25,1))');
          if (mCols["Transfer Eligible"]) masterSheet.getRange(destMasterRow, mCols["Transfer Eligible"]).setFormula('=IF(O' + destMasterRow + '="","",IF(O' + destMasterRow + '>=3,"Yes","No"))');
          if (mCols["Priority"]) masterSheet.getRange(destMasterRow, mCols["Priority"]).setFormula('=IF(O' + destMasterRow + '="","",IF(O' + destMasterRow + '>=5,"High",IF(O' + destMasterRow + '>=3,"Medium","Low")))');
        }
      }
      
      // 2. Update Imported Emp Profiles
      if (importedSheet && importedKgidIdx !== -1 && hasImportedUpdates) {
        var destImportedRow = importedKgidRowMap[kgid];
        var isNewImported = !destImportedRow;
        
        if (isNewImported) {
          destImportedRow = importedSheet.getLastRow() + 1;
          importedSheet.getRange(destImportedRow, importedKgidIdx + 1).setValue(kgid);
          importedKgidRowMap[kgid] = destImportedRow;
        }
        importedUpdatedCount++;
        
        for (var srcIdxStr in importedHeaderMap) {
          var srcIdx = parseInt(srcIdxStr);
          var destCol = importedHeaderMap[srcIdx];
          var val = row[srcIdx];
          if (val !== null && val !== "") {
            importedSheet.getRange(destImportedRow, destCol).setValue(val);
          }
        }
      }
    }
    
    ui.alert("Import Complete!\n\nAdded: " + addedCount + " new employees\nUpdated (Master): " + updatedCount + " records\nUpdated (Profiles): " + importedUpdatedCount + " records");
  } catch (e) {
    ui.alert("Error during import: " + e.toString());
  }
}

// ── SPLIT REMARKS INTO SEPARATE COLUMNS ───────────────────────
function migrateRemarksToSeparateColumns() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    "Migrate Profile Details",
    "This will parse the profile details (Blood, Metal, Email, Approved, Admin, Created, Updated) from the 'Remarks' column in 'Employee Master', move them to their respective separate columns in 'Imported Emp Profiles', and clean up the 'Remarks' column.\n\nDo you want to proceed?",
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) {
    return;
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName("Employee Master");
    var importedSheet = ss.getSheetByName("Imported Emp Profiles");
    
    if (!masterSheet) {
      ui.alert("Error: 'Employee Master' sheet not found.");
      return;
    }
    
    if (!importedSheet) {
      importedSheet = ss.insertSheet("Imported Emp Profiles");
      importedSheet.appendRow([
        "kgid", "name", "mobile1", "mobile2", "rank", "station", "district",
        "metalNumber", "bloodGroup", "email", "isAdmin", "isApproved",
        "createdAt", "updatedAt", "isDeleted", "photoUrl", "lockedFields",
        "caste", "subCaste", "height", "weight", "familyDetails"
      ]);
    }
    
    var masterData = masterSheet.getDataRange().getValues();
    var masterHeaders = masterData[0].map(function(h) { return String(h).trim(); });
    var kgidCol = masterHeaders.indexOf("KGID") + 1;
    var nameCol = masterHeaders.indexOf("Employee Name") + 1;
    var remarksCol = masterHeaders.indexOf("Remarks") + 1;
    
    if (kgidCol === 0 || remarksCol === 0) {
      ui.alert("Error: KGID or Remarks column not found in 'Employee Master'.");
      return;
    }
    
    var importedData = importedSheet.getDataRange().getValues();
    var importedHeaders = importedData[0].map(function(h) { return String(h).trim(); });
    var impKgidCol = importedHeaders.indexOf("kgid") + 1;
    var impNameCol = importedHeaders.indexOf("name") + 1;
    
    if (impKgidCol === 0) {
      ui.alert("Error: 'kgid' column not found in 'Imported Emp Profiles'.");
      return;
    }
    
    // Map imported headers to column index
    var impCols = {};
    importedHeaders.forEach(function(h, idx) { impCols[h] = idx + 1; });
    
    // Index existing imported rows by KGID
    var impKgidRowMap = {};
    for (var r = 1; r < importedData.length; r++) {
      var k = String(importedData[r][impKgidCol - 1]).trim();
      if (k) {
        impKgidRowMap[k] = r + 1;
      }
    }
    
    var migratedCount = 0;
    var cleanedCount = 0;
    
    for (var r = 1; r < masterData.length; r++) {
      var kgid = String(masterData[r][kgidCol - 1]).trim();
      var name = String(masterData[r][nameCol - 1]).trim();
      var remarksVal = String(masterData[r][remarksCol - 1]).trim();
      
      if (!kgid || !remarksVal || remarksVal === "-" || remarksVal === "") continue;
      
      var parseResult = parseRemarksString(remarksVal);
      var profileData = parseResult.profileData;
      var cleanRemarks = parseResult.cleanRemarks;
      
      if (Object.keys(profileData).length > 0) {
        var destRow = impKgidRowMap[kgid];
        if (!destRow) {
          destRow = importedSheet.getLastRow() + 1;
          importedSheet.cell = importedSheet.getRange(destRow, impKgidCol).setValue(kgid);
          importedSheet.getRange(destRow, impNameCol).setValue(name);
          impKgidRowMap[kgid] = destRow;
        }
        
        // Write profile details to Imported Emp Profiles
        for (var k in profileData) {
          if (impCols[k]) {
            var cell = importedSheet.getRange(impKgidRowMap[kgid], impCols[k]);
            var currentVal = String(cell.getValue()).trim();
            if (!currentVal || currentVal === "" || currentVal === "??" || currentVal === "None") {
              cell.setValue(profileData[k]);
            }
          }
        }
        
        migratedCount++;
        
        // Update Remarks column in Employee Master with cleaned string
        if (cleanRemarks !== remarksVal) {
          masterSheet.getRange(r + 1, remarksCol).setValue(cleanRemarks ? cleanRemarks : "");
          cleanedCount++;
        }
      }
    }
    
    ui.alert("Migration Completed Successfully!\n\nParsed and migrated profiles for: " + migratedCount + " employees\nCleaned up Remarks column for: " + cleanedCount + " rows");
  } catch (e) {
    ui.alert("Error during migration: " + e.toString());
  }
}

function parseRemarksString(remarksStr) {
  if (!remarksStr) return { profileData: {}, cleanRemarks: "" };
  
  var parts = remarksStr.split(";");
  var profileData = {};
  var remainingParts = [];
  
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim();
    if (!part) continue;
    
    if (part.indexOf(":") !== -1) {
      var splitIdx = part.indexOf(":");
      var k = part.substring(0, splitIdx).trim().toLowerCase();
      var v = part.substring(splitIdx + 1).trim();
      
      if (k === "metal" || k === "metal number" || k === "metalno") {
        profileData["metalNumber"] = v;
      } else if (k === "blood" || k === "blood group" || k === "bloodgroup") {
        profileData["bloodGroup"] = v;
      } else if (k === "email") {
        profileData["email"] = v;
      } else if (k === "approved") {
        profileData["isApproved"] = v;
      } else if (k === "admin") {
        profileData["isAdmin"] = v;
      } else if (k === "created") {
        profileData["createdAt"] = v;
      } else if (k === "updated") {
        profileData["updatedAt"] = v;
      } else {
        remainingParts.push(part);
      }
    } else {
      remainingParts.push(part);
    }
  }
  
  var cleanRemarks = remainingParts.join("; ").trim();
  return { profileData: profileData, cleanRemarks: cleanRemarks };
}




