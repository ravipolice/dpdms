const ExcelJS = require('c:/Users/ravip/AndroidStudioProjects/Excel/node_modules/exceljs');
const fs = require('fs');
const path = require('path');

// Domain Mappings
const STATION_SUBDIVISION_MAP = {
    "Chikkaballapura Town PS": "Chikkaballapura",
    "Chikkaballapura Rural PS": "Chikkaballapura",
    "Chikkaballapura Women PS": "Chikkaballapura",
    "Chikkaballapura Traffic PS": "Chikkaballapura",
    "Gowribidanur Town PS": "Chikkaballapura",
    "Gowribidanur Rural PS": "Chikkaballapura",
    "Manchenahalli PS": "Chikkaballapura",
    "Gudibanda PS": "Chikkaballapura",
    "Gudibande PS": "Chikkaballapura",
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
    "DAR": "Chikkaballapura",
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

const STATION_CIRCLE_MAP = {
    "Chikkaballapura Town PS": "Chikkaballapura",
    "Chikkaballapura Rural PS": "Chikkaballapura",
    "Chikkaballapura Traffic PS": "Chikkaballapura",
    "Nandigiridhama PS": "Chikkaballapura",
    "Nandi Hills PS": "Chikkaballapura",
    "Nandhi Hills PS": "Chikkaballapura",
    "Cen PS": "Chikkaballapura",
    "CPI Office Cbpura": "Chikkaballapura",
    "Gowribidanur Town PS": "Gouribidanur",
    "Gowribidanur Rural PS": "Gouribidanur",
    "Manchenahalli PS": "Gouribidanur",
    "CPI Off Gowribidanur": "Gouribidanur",
    "Gudibanda PS": "Bagepalli",
    "Gudibande PS": "Bagepalli",
    "Peresandra PS": "Bagepalli",
    "CPI Off Gudibande": "Bagepalli",
    "Kencharlahalli PS": "Kencharlahalli",
    "Batlahalli PS": "Kencharlahalli",
    "Bhatlahalli PS": "Kencharlahalli",
    "CPI Off Kencharlahalli": "Kencharlahalli",
    "Chintamani Town PS": "Chintamani",
    "Chintamani Rural PS": "Chintamani",
    "Sidlaghatta Town PS": "Sidlaghatta",
    "Sidlagatta Town PS": "Sidlaghatta",
    "Shidlaghatta Town PS": "Sidlaghatta",
    "Shidlagatta Town PS": "Sidlaghatta",
    "Sidlaghatta Rural PS": "Sidlaghatta",
    "Sidlagatta Rural PS": "Sidlaghatta",
    "Shidlaghatta Rural PS": "Sidlaghatta",
    "Shidlagatta Rural PS": "Sidlaghatta",
    "Dibburhalli PS": "Sidlaghatta",
    "Dibburahalli PS": "Sidlaghatta",
    "CPI Off Shidlaghatta": "Sidlaghatta",
    "CPI Off Sidlaghatta": "Sidlaghatta",
    "Chelur PS": "Bagepalli",
    "Pathapalya PS": "Bagepalli",
    "CPI Office Cheluru": "Bagepalli",
    "Bagepalli PS": "Bagepalli",
    "Chikkaballapura Women PS": "District HQ & Sub-Division HQ",
    "Cyber Crime PS": "District HQ & Sub-Division HQ",
    "Sdpo Cbpura": "District HQ & Sub-Division HQ",
    "Dysp Office Cbp": "District HQ & Sub-Division HQ",
    "DPO": "District HQ & Sub-Division HQ",
    "Control Room": "District HQ & Sub-Division HQ",
    "DSB": "District HQ & Sub-Division HQ",
    "DAR Chikkaballapura": "District HQ & Sub-Division HQ",
    "DAR": "District HQ & Sub-Division HQ",
    "Smmc": "District HQ & Sub-Division HQ",
    "DCRB": "District HQ & Sub-Division HQ",
    "Dysp Off Cmi": "District HQ & Sub-Division HQ"
};

const CIRCLE_SUBDIVISION_MAP = {
    "Chikkaballapura": "Chikkaballapura",
    "Gouribidanur": "Chikkaballapura",
    "Bagepalli": "Chintamani",
    "Chintamani": "Chintamani",
    "Sidlaghatta": "Chintamani",
    "Kencharlahalli": "Chintamani",
    "District HQ & Sub-Division HQ": "Chikkaballapura"
};

// Helper function to clean cell values
function clean(val) {
    if (val === null || val === undefined) return '';
    // Handle exceljs formula object
    if (typeof val === 'object' && val.result !== undefined) {
        val = val.result;
    }
    if (typeof val === 'object' && val.formula !== undefined) {
        return ''; // Don't use formulas as direct values
    }
    if (val instanceof Date) {
        return val;
    }
    return String(val).trim();
}

function cleanString(val) {
    const v = clean(val);
    if (v instanceof Date) return formatDate(v);
    return v;
}

// Helper to parse dates
function parseDate(val) {
    if (!val) return null;
    // exceljs cell object format
    if (typeof val === 'object' && val.result !== undefined) {
        val = val.result;
    }
    if (val instanceof Date) {
        if (isNaN(val.getTime())) return null;
        return val;
    }
    if (typeof val === 'number') {
        return new Date((val - 25569) * 86400 * 1000);
    }
    if (typeof val === 'string') {
        const s = val.trim();
        if (!s) return null;
        const fmts = [
            /^\d{4}-\d{2}-\d{2}$/,
            /^\d{2}-\d{2}-\d{4}$/,
            /^\d{2}\/\d{2}\/\d{4}$/,
            /^\d{4}\/\d{2}\/\d{2}$/
        ];
        let parts;
        if (s.includes('-')) {
            parts = s.split('-');
        } else if (s.includes('/')) {
            parts = s.split('/');
        }
        if (parts && parts.length === 3) {
            if (parts[0].length === 4) { // YYYY-MM-DD
                return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
            } else if (parts[2].length === 4) { // DD-MM-YYYY
                return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
            }
        }
        const d = new Date(s);
        if (!isNaN(d.getTime())) return d;
    }
    return null;
}

// Format Date as DD-MM-YYYY
function formatDate(d) {
    if (!d) return '';
    const dateObj = parseDate(d);
    if (!dateObj) return '';
    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const year = dateObj.getFullYear();
    return `${day}-${month}-${year}`;
}

// Calculate years between two dates
function yearsBetween(start, end) {
    const startDate = parseDate(start);
    if (!startDate) return '';
    const endDate = parseDate(end) || new Date();
    const diffTime = Math.abs(endDate - startDate);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return parseFloat((diffDays / 365.25).toFixed(1));
}

// Calculate retirement date (60 years, end of month)
function getRetirementDate(dob) {
    const dobDate = parseDate(dob);
    if (!dobDate) return null;
    const year = dobDate.getFullYear() + 60;
    const month = dobDate.getMonth(); // 0-indexed
    // Get last day of that month
    const retirement = new Date(year, month + 1, 0);
    return retirement;
}

// Normalize station name
function normalizeStation(val) {
    const s = cleanString(val);
    const su = s.toUpperCase();
    if (su === "DSB") return "DSB";
    if (su === "DCRB") return "DCRB";
    if (su === "DPO") return "DPO";
    if (su === "DAR" || su === "DAR CHIKKABALLAPURA") return "DAR";
    return s;
}

function getCircleForStation(station) {
    if (!station) return '';
    const s = station.trim();
    if (STATION_CIRCLE_MAP[s]) return STATION_CIRCLE_MAP[s];
    // Case insensitive search
    const sLow = s.toLowerCase();
    for (const [k, v] of Object.entries(STATION_CIRCLE_MAP)) {
        if (k.toLowerCase() === sLow) return v;
    }
    return '';
}

function getSubdivisionForStation(station) {
    if (!station) return '';
    const s = station.trim();
    if (STATION_SUBDIVISION_MAP[s]) return STATION_SUBDIVISION_MAP[s];
    // Case insensitive search
    const sLow = s.toLowerCase();
    for (const [k, v] of Object.entries(STATION_SUBDIVISION_MAP)) {
        if (k.toLowerCase() === sLow) return v;
    }
    return '';
}

// Core merge function
async function mergeExcelData(srcPath, masterPath) {
    console.log(`Loading source workbook: ${srcPath}`);
    const srcWb = new ExcelJS.Workbook();
    await srcWb.xlsx.readFile(srcPath);
    
    // Find active sheet (first sheet, or named 'Emp Profiles')
    let srcSheet = srcWb.worksheets[0];
    for (const sheet of srcWb.worksheets) {
        if (sheet.name.toLowerCase().includes('profiles') || sheet.name === 'Emp Profiles') {
            srcSheet = sheet;
            break;
        }
    }
    console.log(`Source sheet: ${srcSheet.name} (Rows: ${srcSheet.rowCount})`);

    // Extract headers and map indices
    const srcHeaders = [];
    srcSheet.getRow(1).eachCell((cell, colNum) => {
        srcHeaders[colNum] = String(cell.value || '').trim();
    });
    
    let kgidColIdx = -1;
    for (let i = 1; i < srcHeaders.length; i++) {
        const h = srcHeaders[i];
        if (!h) continue;
        const hl = h.toLowerCase();
        if (hl === 'kgid' || hl === 'k.g.i.d' || hl === 'employee kgid' || hl === 'k.g.i.d.') {
            kgidColIdx = i;
            break;
        }
    }

    if (kgidColIdx === -1) {
        throw new Error("Could not find 'KGID' column in the uploaded Excel sheet. Found headers: " + srcHeaders.filter(Boolean).join(', '));
    }

    console.log(`KGID column index: ${kgidColIdx}`);

    // Map source headers to target headers in Employee Master & Imported Emp Profiles
    // We create custom mapping lists
    const headerMapping = {};
    const importedHeaderMapping = {};

    const employeeMasterFields = {
        "Employee Name": ["name", "employee name", "emp name", "fullname"],
        "Rank": ["rank", "designation", "post"],
        "Designation": ["designation", "rank", "post"],
        "DOB": ["dob", "date of birth", "birth date"],
        "Appointment Date": ["doa", "appointment date", "date of appointment", "doj", "joining date"],
        "Mobile": ["mobile", "mobile1", "phone", "contact", "mobile number"],
        "Home District": ["home district", "native district", "native", "home_district"],
        "Current District": ["current district", "district", "current_district"],
        "Current Sub-Division": ["current sub-division", "subdivision", "sub-division", "current_subdivision"],
        "Current Unit": ["current unit", "unit", "station", "police station", "present station", "current_unit"],
        "Present Posting Date": ["present posting date", "posting date", "posting_date", "current unit since", "current_unit_since", "date from"],
        "Remarks": ["remarks", "notes", "remark"],
        "Category": ["category", "caste category"],
        "Type of Transfer": ["type of transfer", "transfer type"]
    };

    const importedEmpProfileFields = {
        "mobile2": ["mobile2", "alternate mobile", "alt phone"],
        "metalNumber": ["metalnumber", "metal number", "metal_number", "metal no"],
        "bloodGroup": ["bloodgroup", "blood group", "blood_group", "blood"],
        "email": ["email", "email address", "email_id"],
        "isAdmin": ["isadmin", "admin"],
        "isApproved": ["isapproved", "approved"],
        "createdAt": ["createdat", "created"],
        "updatedAt": ["updatedat", "updated"],
        "isDeleted": ["isdeleted", "deleted"],
        "height": ["height", "height_cm"],
        "weight": ["weight", "weight_kg"],
        "caste": ["caste"],
        "subCaste": ["subcaste", "sub caste", "sub_caste"],
        "familyDetails": ["familydetails", "family details", "family"],
        "photoUrl": ["photourl", "photo", "image", "photo url"],
        "lockedFields": ["lockedfields", "locked fields"],
        "darWing": ["darwing", "dar wing", "dar_wing", "wing"]
    };

    for (let i = 1; i < srcHeaders.length; i++) {
        const h = srcHeaders[i];
        if (!h) continue;
        const hl = h.toLowerCase();
        
        // Find Employee Master Mapping
        for (const [targetField, aliases] of Object.entries(employeeMasterFields)) {
            if (aliases.includes(hl)) {
                headerMapping[targetField] = i;
            }
        }

        // Find Imported Emp Profiles Mapping
        for (const [targetField, aliases] of Object.entries(importedEmpProfileFields)) {
            if (aliases.includes(hl)) {
                importedHeaderMapping[targetField] = i;
            }
        }
    }

    console.log("Mapped Employee Master fields:", Object.keys(headerMapping));
    console.log("Mapped Imported Emp Profiles fields:", Object.keys(importedHeaderMapping));

    // Load Master Workbook
    console.log(`Loading master workbook: ${masterPath}`);
    const masterWb = new ExcelJS.Workbook();
    await masterWb.xlsx.readFile(masterPath);

    const masterSheet = masterWb.getWorksheet("Employee Master");
    const historySheet = masterWb.getWorksheet("Posting History");
    let importedSheet = masterWb.getWorksheet("Imported Emp Profiles");
    if (!importedSheet) {
        importedSheet = masterWb.addWorksheet("Imported Emp Profiles");
    }

    // Build index of KGID to row number in Employee Master
    const masterKgidMap = {};
    const masterHeaders = [];
    masterSheet.getRow(1).eachCell((cell, colNum) => {
        masterHeaders[colNum] = String(cell.value || '').trim();
    });
    const mKgidColIdx = masterHeaders.indexOf("KGID");
    if (mKgidColIdx === -1) {
        throw new Error("Master Employee Master sheet is missing 'KGID' column!");
    }

    for (let r = 2; r <= masterSheet.rowCount; r++) {
        const cellVal = cleanString(masterSheet.getCell(r, mKgidColIdx).value);
        if (cellVal) {
            masterKgidMap[cellVal] = r;
        }
    }

    // Build index of KGID to row number in Imported Emp Profiles
    const importedKgidMap = {};
    const importedHeaders = [];
    importedSheet.getRow(1).eachCell((cell, colNum) => {
        importedHeaders[colNum] = String(cell.value || '').trim();
    });
    let impKgidColIdx = importedHeaders.indexOf("kgid");
    if (impKgidColIdx === -1) {
        // Initialize header row for Imported Emp Profiles if empty
        const defaultImpHeaders = [
            "kgid", "name", "mobile1", "mobile2", "rank", "station", "district",
            "metalNumber", "bloodGroup", "email", "isAdmin", "isApproved",
            "createdAt", "updatedAt", "isDeleted", "height", "weight", "caste",
            "subCaste", "familyDetails", "photoUrl", "lockedFields", "darWing"
        ];
        importedSheet.getRow(1).values = defaultImpHeaders;
        importedHeaders.push(...defaultImpHeaders);
        impKgidColIdx = 1;
    }

    for (let r = 2; r <= importedSheet.rowCount; r++) {
        const cellVal = cleanString(importedSheet.getCell(r, impKgidColIdx).value);
        if (cellVal) {
            importedKgidMap[cellVal] = r;
        }
    }

    // Get History Headers
    const historyHeaders = [];
    historySheet.getRow(1).eachCell((cell, colNum) => {
        historyHeaders[colNum] = String(cell.value || '').trim();
    });

    let addedCount = 0;
    let updatedCount = 0;
    let transferCount = 0;

    // Iterate through source rows starting from row 2
    for (let r = 2; r <= srcSheet.rowCount; r++) {
        const srcRow = srcSheet.getRow(r);
        const kgid = cleanString(srcRow.getCell(kgidColIdx).value);
        if (!kgid) continue;

        const masterRowIdx = masterKgidMap[kgid];
        const isNew = !masterRowIdx;
        let targetRowIdx = masterRowIdx;

        if (isNew) {
            targetRowIdx = masterSheet.rowCount + 1;
            masterSheet.getCell(targetRowIdx, mKgidColIdx).value = kgid;
            masterKgidMap[kgid] = targetRowIdx;
            addedCount++;
        } else {
            updatedCount++;
        }

        // Keep track of unit change (transfers)
        const oldUnitCell = masterSheet.getCell(targetRowIdx, masterHeaders.indexOf("Current Unit"));
        const oldUnit = oldUnitCell ? cleanString(oldUnitCell.value) : '';
        let newUnit = oldUnit;
        if (headerMapping["Current Unit"]) {
            newUnit = normalizeStation(srcRow.getCell(headerMapping["Current Unit"]).value);
        }

        const isTransfer = !isNew && oldUnit && newUnit && oldUnit.toLowerCase() !== newUnit.toLowerCase();
        if (isTransfer) {
            transferCount++;
        }

        // Copy Employee Master Fields
        for (const [targetField, srcColIdx] of Object.entries(headerMapping)) {
            const colIdx = masterHeaders.indexOf(targetField);
            if (colIdx === -1) continue;

            let val = srcRow.getCell(srcColIdx).value;
            // Clean / Format
            if (targetField === "DOB" || targetField === "Appointment Date" || targetField === "Present Posting Date") {
                val = parseDate(val);
            } else if (targetField === "Current Unit") {
                val = normalizeStation(val);
            } else {
                val = cleanString(val);
            }

            // Only overwrite if value is present in uploaded row (or if it's a new row)
            if (val !== '' && val !== null && val !== undefined) {
                masterSheet.getCell(targetRowIdx, colIdx).value = val;
            }
        }

        // Automatically fill Current Sub-Division and Circle if empty or changed
        const unitColIdx = masterHeaders.indexOf("Current Unit");
        const subdivColIdx = masterHeaders.indexOf("Current Sub-Division");
        const circleColIdx = masterHeaders.indexOf("Circle");
        const districtColIdx = masterHeaders.indexOf("Current District");

        const currentUnitVal = cleanString(masterSheet.getCell(targetRowIdx, unitColIdx).value);
        if (currentUnitVal) {
            const calculatedSubdiv = getSubdivisionForStation(currentUnitVal);
            const calculatedCircle = getCircleForStation(currentUnitVal);
            
            if (calculatedSubdiv && subdivColIdx !== -1) {
                const currentSubdivVal = cleanString(masterSheet.getCell(targetRowIdx, subdivColIdx).value);
                if (!currentSubdivVal || isTransfer || isNew) {
                    masterSheet.getCell(targetRowIdx, subdivColIdx).value = calculatedSubdiv;
                }
            }
            if (calculatedCircle && circleColIdx !== -1) {
                const currentCircleVal = cleanString(masterSheet.getCell(targetRowIdx, circleColIdx).value);
                if (!currentCircleVal || isTransfer || isNew) {
                    masterSheet.getCell(targetRowIdx, circleColIdx).value = calculatedCircle;
                }
            }
        }

        // Set Formulas for Employee Master row
        const rNum = targetRowIdx;
        const colLetterE = excelColLetter(masterHeaders.indexOf("DOB"));
        const colLetterF = excelColLetter(masterHeaders.indexOf("Appointment Date"));
        const colLetterM = excelColLetter(masterHeaders.indexOf("Present Posting Date"));
        const colLetterA = excelColLetter(masterHeaders.indexOf("KGID"));
        const colLetterK = excelColLetter(masterHeaders.indexOf("Current Sub-Division"));
        const colLetterJ = excelColLetter(masterHeaders.indexOf("Current District"));
        const colLetterO = excelColLetter(masterHeaders.indexOf("Current Station Years"));

        if (masterHeaders.indexOf("Retirement Date") !== -1) {
            masterSheet.getCell(rNum, masterHeaders.indexOf("Retirement Date")).value = { formula: `=IF(${colLetterE}${rNum}="","",EOMONTH(EDATE(${colLetterE}${rNum},60*12),0))` };
        }
        if (masterHeaders.indexOf("Total Service Years") !== -1) {
            masterSheet.getCell(rNum, masterHeaders.indexOf("Total Service Years")).value = { formula: `=IF(${colLetterF}${rNum}="","",ROUND(YEARFRAC(${colLetterF}${rNum},TODAY()),1))` };
        }
        if (masterHeaders.indexOf("Current Station Years") !== -1) {
            masterSheet.getCell(rNum, masterHeaders.indexOf("Current Station Years")).value = { formula: `=IF(${colLetterM}${rNum}="","",ROUND(YEARFRAC(${colLetterM}${rNum},TODAY()),1))` };
        }
        if (masterHeaders.indexOf("Current Sub-Division Years") !== -1) {
            masterSheet.getCell(rNum, masterHeaders.indexOf("Current Sub-Division Years")).value = { formula: `=IF(${colLetterA}${rNum}="","",ROUND(SUMPRODUCT(('Posting History'!$A$2:$A$3000=${colLetterA}${rNum})*('Posting History'!$E$2:$E$3000=${colLetterK}${rNum})*('Posting History'!$B$2:$B$3000<>"")*(IF('Posting History'!$C$2:$C$3000="",TODAY(),'Posting History'!$C$2:$C$3000)-'Posting History'!$B$2:$B$3000))/365.25,1))` };
        }
        if (masterHeaders.indexOf("District Service Years") !== -1) {
            masterSheet.getCell(rNum, masterHeaders.indexOf("District Service Years")).value = { formula: `=IF(${colLetterA}${rNum}="","",ROUND(SUMPRODUCT(('Posting History'!$A$2:$A$3000=${colLetterA}${rNum})*('Posting History'!$F$2:$F$3000=${colLetterJ}${rNum})*('Posting History'!$B$2:$B$3000<>"")*(IF('Posting History'!$C$2:$C$3000="",TODAY(),'Posting History'!$C$2:$C$3000)-'Posting History'!$B$2:$B$3000))/365.25,1))` };
        }
        if (masterHeaders.indexOf("Transfer Eligible") !== -1) {
            masterSheet.getCell(rNum, masterHeaders.indexOf("Transfer Eligible")).value = { formula: `=IF(${colLetterO}${rNum}="","",IF(${colLetterO}${rNum}>=3,"Yes","No"))` };
        }
        if (masterHeaders.indexOf("Priority") !== -1) {
            masterSheet.getCell(rNum, masterHeaders.indexOf("Priority")).value = { formula: `=IF(${colLetterO}${rNum}="","",IF(${colLetterO}${rNum}>=5,"High",IF(${colLetterO}${rNum}>=3,"Medium","Low")))` };
        }

        // Copy Profile Details to Imported Emp Profiles
        let impRowIdx = importedKgidMap[kgid];
        if (!impRowIdx) {
            impRowIdx = importedSheet.rowCount + 1;
            importedSheet.getCell(impRowIdx, impKgidColIdx).value = kgid;
            importedKgidMap[kgid] = impRowIdx;
        }

        // Default set general fields in profiles
        const nameIdx = importedHeaders.indexOf("name");
        if (nameIdx !== -1 && headerMapping["Employee Name"]) {
            importedSheet.getCell(impRowIdx, nameIdx).value = cleanString(srcRow.getCell(headerMapping["Employee Name"]).value);
        }
        const mobIdx = importedHeaders.indexOf("mobile1");
        if (mobIdx !== -1 && headerMapping["Mobile"]) {
            importedSheet.getCell(impRowIdx, mobIdx).value = cleanString(srcRow.getCell(headerMapping["Mobile"]).value);
        }
        const rankIdx = importedHeaders.indexOf("rank");
        if (rankIdx !== -1 && headerMapping["Rank"]) {
            importedSheet.getCell(impRowIdx, rankIdx).value = cleanString(srcRow.getCell(headerMapping["Rank"]).value);
        }
        const stationIdx = importedHeaders.indexOf("station");
        if (stationIdx !== -1 && headerMapping["Current Unit"]) {
            importedSheet.getCell(impRowIdx, stationIdx).value = normalizeStation(srcRow.getCell(headerMapping["Current Unit"]).value);
        }
        const distIdx = importedHeaders.indexOf("district");
        if (distIdx !== -1 && headerMapping["Current District"]) {
            importedSheet.getCell(impRowIdx, distIdx).value = cleanString(srcRow.getCell(headerMapping["Current District"]).value);
        }

        // Copy target fields mapped in imported
        for (const [targetField, srcColIdx] of Object.entries(importedHeaderMapping)) {
            const colIdx = importedHeaders.indexOf(targetField);
            if (colIdx === -1) continue;
            let val = srcRow.getCell(srcColIdx).value;
            if (targetField === 'isDeleted' || targetField === 'isAdmin' || targetField === 'isApproved') {
                if (val === true) val = "Yes";
                else if (val === false) val = "No";
                else val = cleanString(val);
            } else {
                val = cleanString(val);
            }
            if (val !== '' && val !== null && val !== undefined) {
                importedSheet.getCell(impRowIdx, colIdx).value = val;
            }
        }
        
        // Update profile updatedAt timestamp
        const updatedIdx = importedHeaders.indexOf("updatedAt");
        if (updatedIdx !== -1) {
            const now = new Date();
            const day = String(now.getDate()).padStart(2, '0');
            const month = String(now.getMonth() + 1).padStart(2, '0');
            const year = now.getFullYear();
            const hour = String(now.getHours()).padStart(2, '0');
            const min = String(now.getMinutes()).padStart(2, '0');
            importedSheet.getCell(impRowIdx, updatedIdx).value = `${day}-${month}-${year} ${hour}:${min}`;
        }

        // Update Posting History
        const employeeRank = cleanString(masterSheet.getCell(targetRowIdx, masterHeaders.indexOf("Rank")).value);
        const employeeDistrict = cleanString(masterSheet.getCell(targetRowIdx, masterHeaders.indexOf("Current District")).value) || "Chikkaballapura";
        const employeeSubdiv = cleanString(masterSheet.getCell(targetRowIdx, masterHeaders.indexOf("Current Sub-Division")).value);
        const employeeSince = masterSheet.getCell(targetRowIdx, masterHeaders.indexOf("Present Posting Date")).value;

        if (isNew) {
            // New employee posting history
            const histRowIdx = historySheet.rowCount + 1;
            const fromDateVal = employeeSince ? parseDate(employeeSince) : null;
            
            writeHistoryRow(historySheet, histRowIdx, {
                "KGID": kgid,
                "From Date": fromDateVal,
                "To Date": null,
                "Police Station / Unit": currentUnitVal,
                "Sub-Division": employeeSubdiv,
                "District": employeeDistrict,
                "Rank Held": employeeRank,
                "Order Number": "",
                "Notes": "Auto-created during Excel import merge."
            }, historyHeaders);
        } else if (isTransfer) {
            // Find existing posting that has no To Date, close it, and open a new one
            let openPostingRowIdx = -1;
            const hKgidIdx = historyHeaders.indexOf("KGID");
            const hToDateIdx = historyHeaders.indexOf("To Date");
            
            for (let i = 2; i <= historySheet.rowCount; i++) {
                const hKgid = cleanString(historySheet.getCell(i, hKgidIdx).value);
                if (hKgid === kgid) {
                    const toDate = historySheet.getCell(i, hToDateIdx).value;
                    if (!toDate || toDate === "Present" || toDate === "") {
                        openPostingRowIdx = i;
                        break;
                    }
                }
            }

            const transferDate = employeeSince ? parseDate(employeeSince) : new Date();

            if (openPostingRowIdx !== -1) {
                historySheet.getCell(openPostingRowIdx, hToDateIdx).value = transferDate;
                historySheet.getCell(openPostingRowIdx, historyHeaders.indexOf("Notes")).value = "Closed automatically on transfer.";
            }

            // Open new posting
            const histRowIdx = historySheet.rowCount + 1;
            writeHistoryRow(historySheet, histRowIdx, {
                "KGID": kgid,
                "From Date": transferDate,
                "To Date": null,
                "Police Station / Unit": newUnit,
                "Sub-Division": employeeSubdiv,
                "District": employeeDistrict,
                "Rank Held": employeeRank,
                "Order Number": "",
                "Notes": `Transfer detected from ${oldUnit} to ${newUnit}.`
            }, historyHeaders);
        }
    }

    // Update Table Ranges
    const tableSpecs = [
        ["Employee Master", "tblEmployeeMaster"],
        ["Posting History", "tblPostingHistory"],
        ["Imported Emp Profiles", "tblImportedEmpProfiles"],
    ];

    for (const [sheetName, tableName] of tableSpecs) {
        const sheet = masterWb.getWorksheet(sheetName);
        if (sheet && sheet.tables && sheet.tables[tableName]) {
            const lastRow = Math.max(2, sheet.rowCount);
            const lastColLetter = excelColLetter(sheet.columnCount);
            sheet.tables[tableName].ref = `A1:${lastColLetter}${lastRow}`;
        }
    }

    console.log("Saving master workbook...");
    await masterWb.xlsx.writeFile(masterPath);
    console.log("Workbook saved successfully!");

    return { added: addedCount, updated: updatedCount, transfers: transferCount };
}

function writeHistoryRow(sheet, rowIdx, data, headers) {
    for (const [header, val] of Object.entries(data)) {
        const colIdx = headers.indexOf(header);
        if (colIdx !== -1) {
            sheet.getCell(rowIdx, colIdx).value = val;
        }
    }
    // Write Posting History calculated formulas
    const colA = excelColLetter(headers.indexOf("KGID"));
    const colB = excelColLetter(headers.indexOf("From Date"));
    const colC = excelColLetter(headers.indexOf("To Date"));
    const colD = excelColLetter(headers.indexOf("Police Station / Unit"));
    const colF = excelColLetter(headers.indexOf("District"));
    const colG = excelColLetter(headers.indexOf("Rank Held"));

    sheet.getCell(rowIdx, headers.indexOf("Duration Years")).value = { formula: `=IF(${colB}${rowIdx}="","",ROUND((IF(${colC}${rowIdx}="",TODAY(),${colC}${rowIdx})-${colB}${rowIdx})/365.25,1))` };
    sheet.getCell(rowIdx, headers.indexOf("Overlap Check")).value = { formula: `=IF(${colA}${rowIdx}="","",IF(COUNTIFS($A:$A,${colA}${rowIdx},$B:$B,"<="&IF(${colC}${rowIdx}="",TODAY(),${colC}${rowIdx}),$C:$C,">="&${colB}${rowIdx})>1,"Review","OK"))` };
    sheet.getCell(rowIdx, headers.indexOf("Missing Date Check")).value = { formula: `=IF(${colA}${rowIdx}="","",IF(OR(${colB}${rowIdx}="",${colD}${rowIdx}="",${colF}${rowIdx}="",${colG}${rowIdx}=""),"Missing","OK"))` };
}

// Convert 1-based column index to Excel column letter (A, B, C... Z, AA, AB...)
function excelColLetter(col) {
    let temp;
    let letter = '';
    while (col > 0) {
        temp = (col - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        col = (col - temp - 1) / 26;
    }
    return letter;
}

// Rebuild JSON payload for embedded data
async function buildStandaloneJSON(masterPath) {
    console.log(`Building JSON payload from: ${masterPath}`);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(masterPath);

    const masterSheet = wb.getWorksheet("Employee Master");
    const historySheet = wb.getWorksheet("Posting History");
    const importedSheet = wb.getWorksheet("Imported Emp Profiles");
    const transferRequestsSheet = wb.getWorksheet("Transfer Requests");

    // Read headers
    const readHeaders = (sheet) => {
        const h = [];
        sheet.getRow(1).eachCell((cell, colNum) => {
            h[colNum] = String(cell.value || '').trim();
        });
        return h;
    };

    const mHeaders = readHeaders(masterSheet);
    const hHeaders = readHeaders(historySheet);
    const impHeaders = readHeaders(importedSheet);

    const sheetToRows = (sheet, headers) => {
        const rows = [];
        for (let r = 2; r <= sheet.rowCount; r++) {
            const row = sheet.getRow(r);
            // check if row is empty
            let hasVal = false;
            const obj = {};
            for (let c = 1; c < headers.length; c++) {
                const header = headers[c];
                if (!header) continue;
                const val = row.getCell(c).value;
                if (val !== null && val !== undefined && val !== '') {
                    hasVal = true;
                }
                obj[header] = val;
            }
            if (hasVal) {
                rows.push(obj);
            }
        }
        return rows;
    };

    const masterRows = sheetToRows(masterSheet, mHeaders);
    const historyRows = sheetToRows(historySheet, hHeaders);
    const importedRows = sheetToRows(importedSheet, impHeaders);

    const sourceByKgid = {};
    for (const r of importedRows) {
        const kgid = cleanString(r.kgid);
        if (kgid) {
            sourceByKgid[kgid] = r;
        }
    }

    const transferRequests = [];
    if (transferRequestsSheet) {
        const trHeaders = readHeaders(transferRequestsSheet);
        const trRows = sheetToRows(transferRequestsSheet, trHeaders);
        for (const tr of trRows) {
            const trKgid = cleanString(tr.KGID);
            if (trKgid) {
                transferRequests.push({
                    "kgid": trKgid,
                    "name": cleanString(tr["Employee Name"]),
                    "rank": cleanString(tr["Rank"]),
                    "currentStation": cleanString(tr["Current Station"]),
                    "preference1": cleanString(tr["Preference 1"]),
                    "preference2": cleanString(tr["Preference 2"]),
                    "preference3": cleanString(tr["Preference 3"]),
                    "preference4": cleanString(tr["Preference 4"]),
                    "preference5": cleanString(tr["Preference 5"]),
                    "transferCategory": cleanString(tr["Transfer Category"]),
                    "reason": cleanString(tr["Reason"]),
                    "remarks": cleanString(tr["Remarks"]),
                    "applicationDate": cleanString(tr["Application Date"]),
                    "status": cleanString(tr["Status"] || "Pending"),
                    "approvedStation": cleanString(tr["Approved Station"] || "")
                });
            }
        }
    }

    const historyByKgid = {};
    for (const item of historyRows) {
        const kgid = cleanString(item.KGID);
        if (!kgid) continue;
        const fromDate = item["From Date"];
        const toDate = item["To Date"];
        item._fromDisplay = formatDate(fromDate);
        item._toDisplay = formatDate(toDate) || "Present";
        item._durationYears = yearsBetween(fromDate, toDate);
        if (!historyByKgid[kgid]) {
            historyByKgid[kgid] = [];
        }
        historyByKgid[kgid].push(item);
    }

    const employees = [];
    for (const employee of masterRows) {
        const kgid = cleanString(employee.KGID);
        if (!kgid) continue;

        const dob = employee.DOB;
        const doa = employee["Appointment Date"] || employee.DOJ || employee.DOA;
        const currentSince = employee["Present Posting Date"];
        const totalService = yearsBetween(doa);
        const currentStationYears = yearsBetween(currentSince);
        
        const priority = currentStationYears !== "" && currentStationYears >= 5 ? "High" : 
                         currentStationYears !== "" && currentStationYears >= 3 ? "Medium" : "Low";
        const eligible = currentStationYears !== "" && currentStationYears >= 3 ? "Yes" : "No";

        const postings = historyByKgid[kgid] || [];
        const source = sourceByKgid[kgid] || {};
        
        let longest = 0;
        for (const p of postings) {
            const dur = parseFloat(p._durationYears) || 0;
            if (dur > longest) longest = dur;
        }

        const retirementValue = employee["Retirement Date"] || getRetirementDate(dob);
        const category = cleanString(employee.Category);
        const typeOfTransfer = cleanString(employee["Type of Transfer"]) || "Regular";
        const currentSubdivRaw = cleanString(employee["Current Sub-Division"]);
        const unit = normalizeStation(employee["Current Unit"]);

        // Calculate subdiv and district service years
        let subdivYears = 0;
        let districtYears = 0;
        const currentDistrict = cleanString(employee["Current District"]) || "Chikkaballapura";

        for (const p of postings) {
            const dur = parseFloat(p._durationYears) || 0;
            const pStation = normalizeStation(p["Police Station / Unit"]);
            const pSubdiv = cleanString(p["Sub-Division"]) || getSubdivisionForStation(pStation);
            const pDistrict = cleanString(p["District"]);

            if (currentSubdivRaw && pSubdiv.toLowerCase() === currentSubdivRaw.toLowerCase()) {
                subdivYears += dur;
            }
            if (currentDistrict && pDistrict.toLowerCase() === currentDistrict.toLowerCase()) {
                districtYears += dur;
            }
        }

        subdivYears = parseFloat(subdivYears.toFixed(1));
        districtYears = parseFloat(districtYears.toFixed(1));

        const retDateStr = formatDate(retirementValue);
        let retirementMonthsLeft = "";
        if (retDateStr) {
            const rd = parseDate(retDateStr);
            if (rd) {
                const today = new Date();
                retirementMonthsLeft = (rd.getFullYear() - today.getFullYear()) * 12 + (rd.getMonth() - today.getMonth());
            }
        }

        const employeeOut = {
            "kgid": kgid,
            "name": cleanString(employee["Employee Name"]),
            "rank": cleanString(employee.Rank),
            "designation": cleanString(employee.Designation),
            "dob": formatDate(dob),
            "doa": formatDate(doa),
            "mobile": cleanString(employee.Mobile),
            "homeDistrict": cleanString(employee["Home District"]),
            "retirementDate": retDateStr,
            "currentDistrict": currentDistrict,
            "currentSubDivision": currentSubdivRaw,
            "circle": cleanString(employee.Circle || employee.circle || getCircleForStation(unit)),
            "currentUnit": unit,
            "currentSince": formatDate(currentSince),
            "totalServiceYears": totalService,
            "currentStationYears": currentStationYears,
            "transferEligible": eligible,
            "priority": priority,
            "longestPostingYears": longest,
            "postings": postings.map(item => ({
                "from": item._fromDisplay,
                "to": item._toDisplay,
                "station": normalizeStation(item["Police Station / Unit"]),
                "subDivision": cleanString(item["Sub-Division"]) || getSubdivisionForStation(normalizeStation(item["Police Station / Unit"])),
                "district": cleanString(item.District),
                "rank": cleanString(item["Rank Held"]),
                "orderNumber": cleanString(item["Order Number"]),
                "durationYears": item._durationYears,
            })),
            "sourceProfile": {
                "metalNumber": cleanString(source.metalNumber),
                "bloodGroup": cleanString(source.bloodGroup),
                "email": cleanString(source.email),
                "mobile2": cleanString(source.mobile2),
                "isAdmin": cleanString(source.isAdmin),
                "isApproved": cleanString(source.isApproved),
                "createdAt": cleanString(source.createdAt),
                "updatedAt": cleanString(source.updatedAt),
                "height": cleanString(source.height),
                "weight": cleanString(source.weight),
                "caste": cleanString(source.caste),
                "subCaste": cleanString(source.subCaste),
                "familyDetails": cleanString(source.familyDetails),
                "photoUrl": cleanString(source.photoUrl),
                "lockedFields": cleanString(source.lockedFields),
                "darWing": cleanString(source.darWing),
            },
            "category": category,
            "typeOfTransfer": typeOfTransfer,
            "subDivisionYears": subdivYears,
            "districtServiceYears": districtYears,
            "retirementMonthsLeft": retirementMonthsLeft
        };
        employees.push(employeeOut);
    }

    const stations = Array.from(new Set(employees.map(e => e.currentUnit).filter(Boolean))).sort();
    const ranks = Array.from(new Set(employees.map(e => e.rank).filter(Boolean))).sort();
    const subdivisions = Array.from(new Set(employees.map(e => e.currentSubDivision).filter(Boolean))).sort();

    const summary = {
        "totalEmployees": employees.length,
        "transferDue": employees.filter(e => e.transferEligible === "Yes").length,
        "highPriority": employees.filter(e => e.priority === "High").length,
        "retirementDue": employees.filter(e => e.retirementDate).length,
        "retirementDue12m": employees.filter(e => typeof e.retirementMonthsLeft === 'number' && e.retirementMonthsLeft >= 0 && e.retirementMonthsLeft <= 12).length,
        "stations": stations.map(st => ({ "name": st, "working": employees.filter(e => e.currentUnit === st).length })),
        "ranks": ranks.map(rk => ({ "name": rk, "working": employees.filter(e => e.rank === rk).length })),
        "subdivisions": subdivisions.map(sub => ({ "name": sub, "working": employees.filter(e => e.currentSubDivision === sub).length })),
    };

    return {
        "workbook": masterPath,
        "editable": true,
        "lastLoaded": formatDate(new Date()) + " " + new Date().toTimeString().split(' ')[0],
        "employees": employees,
        "summary": summary,
        "transferRequests": transferRequests,
        "stationSubdivisionMap": STATION_SUBDIVISION_MAP
    };
}

// Update the index.html with new EMBEDDED_DATA
function updateIndexHtml(indexHtmlPath, jsonData) {
    console.log(`Updating index.html at ${indexHtmlPath}`);
    let content = fs.readFileSync(indexHtmlPath, 'utf8');

    const prefix = 'window.EMBEDDED_DATA = ';
    const startIdx = content.indexOf(prefix);
    if (startIdx === -1) {
        throw new Error("Could not find window.EMBEDDED_DATA inside index.html!");
    }

    const jsonStart = startIdx + prefix.length;
    const endIdx = content.indexOf('</script>', jsonStart);
    if (endIdx === -1) {
        throw new Error("Could not find </script> tag closing EMBEDDED_DATA inside index.html!");
    }

    const before = content.substring(0, startIdx);
    const after = content.substring(endIdx);
    
    const payloadStr = JSON.stringify(jsonData);
    const newContent = before + 'window.EMBEDDED_DATA = ' + payloadStr + ';' + after;
    
    fs.writeFileSync(indexHtmlPath, newContent, 'utf8');
    console.log('Successfully wrote updated JSON into index.html');
}

// Master function called by server
async function handleMergeAndSync(uploadedFilePath) {
    const outputsDir = path.join(__dirname, '..', 'outputs');
    const masterPath = path.join(outputsDir, 'District_Employee_Transfer_Database_IMPORTED.xlsx');
    const fallbackTemplatePath = path.join(outputsDir, 'District_Employee_Transfer_Database_Template_FIXED.xlsx');
    
    const activeMasterPath = fs.existsSync(masterPath) ? masterPath : fallbackTemplatePath;

    // 1. Merge uploaded Excel with the master database
    const stats = await mergeExcelData(uploadedFilePath, activeMasterPath);
    console.log(`Merge stats: Added: ${stats.added}, Updated: ${stats.updated}, Transfers: ${stats.transfers}`);

    // 2. Build new dashboard JSON structure
    const updatedJson = await buildStandaloneJSON(activeMasterPath);

    // 3. Update index.html
    const indexHtmlPath = path.join(__dirname, '..', 'DPDMS', 'index.html');
    updateIndexHtml(indexHtmlPath, updatedJson);

    // Also update standalone dashboard in outputs if it exists
    const standaloneHtmlPath = path.join(outputsDir, 'Employee_Transfer_Dashboard_Browser.html');
    if (fs.existsSync(standaloneHtmlPath)) {
        try {
            updateIndexHtml(standaloneHtmlPath, updatedJson);
        } catch (e) {
            console.warn("Could not update standalone dashboard html:", e.message);
        }
    }

    return {
        ok: true,
        added: stats.added,
        updated: stats.updated,
        transfers: stats.transfers,
        message: `Successfully merged data! Added: ${stats.added}, Updated: ${stats.updated}, Transfers detected: ${stats.transfers}`
    };
}



async function saveTransferRequest(payload) {
    const outputsDir = path.join(__dirname, '..', 'outputs');
    const masterPath = path.join(outputsDir, 'District_Employee_Transfer_Database_IMPORTED.xlsx');
    const fallbackTemplatePath = path.join(outputsDir, 'District_Employee_Transfer_Database_Template_FIXED.xlsx');
    const activeMasterPath = fs.existsSync(masterPath) ? masterPath : fallbackTemplatePath;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(activeMasterPath);

    let sheet = wb.getWorksheet("Transfer Requests");
    if (!sheet) {
        sheet = wb.addWorksheet("Transfer Requests");
    }

    const headers = [];
    sheet.getRow(1).eachCell((cell, colNum) => {
        headers[colNum] = String(cell.value || '').trim();
    });

    let kgidColIdx = headers.indexOf("KGID");
    if (kgidColIdx === -1) {
        const defaultHeaders = [
            "KGID", "Employee Name", "Rank", "Current Station", 
            "Preference 1", "Preference 2", "Preference 3", "Preference 4", "Preference 5",
            "Transfer Category", "Reason", "Remarks", "Application Date", "Status", "Approved Station"
        ];
        sheet.getRow(1).values = defaultHeaders;
        headers.push(...defaultHeaders);
        kgidColIdx = 1;
    }

    const kgid = cleanString(payload.kgid);
    if (!kgid) throw new Error("KGID is required.");

    // Find existing row by KGID
    let rowIdx = -1;
    for (let r = 2; r <= sheet.rowCount; r++) {
        if (cleanString(sheet.getCell(r, kgidColIdx).value) === kgid) {
            rowIdx = r;
            break;
        }
    }

    if (rowIdx === -1) {
        rowIdx = sheet.rowCount + 1;
    }

    const trValues = {
        "KGID": kgid,
        "Employee Name": cleanString(payload.name),
        "Rank": cleanString(payload.rank),
        "Current Station": cleanString(payload.currentStation),
        "Preference 1": cleanString(payload.preference1),
        "Preference 2": cleanString(payload.preference2),
        "Preference 3": cleanString(payload.preference3),
        "Preference 4": cleanString(payload.preference4),
        "Preference 5": cleanString(payload.preference5),
        "Transfer Category": cleanString(payload.transferCategory),
        "Reason": cleanString(payload.reason),
        "Remarks": cleanString(payload.remarks),
        "Application Date": cleanString(payload.applicationDate || formatDate(new Date())),
        "Status": cleanString(payload.status || "Pending"),
        "Approved Station": cleanString(payload.approvedStation || "")
    };

    for (const [header, val] of Object.entries(trValues)) {
        const col = headers.indexOf(header);
        if (col !== -1) {
            sheet.getCell(rowIdx, col).value = val;
        }
    }

    // Update table reference if exists
    if (sheet.tables && sheet.tables["tblTransferRequests"]) {
        const lastRow = Math.max(2, sheet.rowCount);
        const lastColLetter = excelColLetter(sheet.columnCount);
        sheet.tables["tblTransferRequests"].ref = "A1:" + lastColLetter + lastRow;
    }

    await wb.xlsx.writeFile(activeMasterPath);

    // Rebuild data in index.html
    const updatedJson = await buildStandaloneJSON(activeMasterPath);
    const indexHtmlPath = path.join(__dirname, '..', 'DPDMS', 'index.html');
    updateIndexHtml(indexHtmlPath, updatedJson);

    const standaloneHtmlPath = path.join(outputsDir, 'Employee_Transfer_Dashboard_Browser.html');
    if (fs.existsSync(standaloneHtmlPath)) {
        try {
            updateIndexHtml(standaloneHtmlPath, updatedJson);
        } catch (e) {}
    }

    return { ok: true, message: `Saved transfer request for ${trValues["Employee Name"]} (${kgid})` };
}

async function deleteTransferRequest(kgid) {
    const outputsDir = path.join(__dirname, '..', 'outputs');
    const masterPath = path.join(outputsDir, 'District_Employee_Transfer_Database_IMPORTED.xlsx');
    const fallbackTemplatePath = path.join(outputsDir, 'District_Employee_Transfer_Database_Template_FIXED.xlsx');
    const activeMasterPath = fs.existsSync(masterPath) ? masterPath : fallbackTemplatePath;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(activeMasterPath);

    const sheet = wb.getWorksheet("Transfer Requests");
    if (!sheet) {
        return { ok: false, error: "Transfer Requests worksheet not found." };
    }

    const headers = [];
    sheet.getRow(1).eachCell((cell, colNum) => {
        headers[colNum] = String(cell.value || '').trim();
    });

    const kgidColIdx = headers.indexOf("KGID");
    if (kgidColIdx === -1) {
        return { ok: false, error: "KGID column not found in Transfer Requests." };
    }

    let rowIdx = -1;
    for (let r = 2; r <= sheet.rowCount; r++) {
        if (cleanString(sheet.getCell(r, kgidColIdx).value) === String(kgid).trim()) {
            rowIdx = r;
            break;
        }
    }

    if (rowIdx === -1) {
        return { ok: false, error: `Transfer request for KGID ${kgid} not found.` };
    }

    sheet.spliceRows(rowIdx, 1);

    // Update table reference if exists
    if (sheet.tables && sheet.tables["tblTransferRequests"]) {
        const lastRow = Math.max(2, sheet.rowCount);
        const lastColLetter = excelColLetter(sheet.columnCount);
        sheet.tables["tblTransferRequests"].ref = "A1:" + lastColLetter + lastRow;
    }

    await wb.xlsx.writeFile(activeMasterPath);

    // Rebuild data in index.html
    const updatedJson = await buildStandaloneJSON(activeMasterPath);
    const indexHtmlPath = path.join(__dirname, '..', 'DPDMS', 'index.html');
    updateIndexHtml(indexHtmlPath, updatedJson);

    const standaloneHtmlPath = path.join(outputsDir, 'Employee_Transfer_Dashboard_Browser.html');
    if (fs.existsSync(standaloneHtmlPath)) {
        try {
            updateIndexHtml(standaloneHtmlPath, updatedJson);
        } catch (e) {}
    }

    return { ok: true, message: `Withdrew transfer request for KGID ${kgid}` };
}

async function saveEmployee(payload) {
    const outputsDir = path.join(__dirname, '..', 'outputs');
    const masterPath = path.join(outputsDir, 'District_Employee_Transfer_Database_IMPORTED.xlsx');
    const fallbackTemplatePath = path.join(outputsDir, 'District_Employee_Transfer_Database_Template_FIXED.xlsx');
    const activeMasterPath = fs.existsSync(masterPath) ? masterPath : fallbackTemplatePath;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(activeMasterPath);

    const masterSheet = wb.getWorksheet("Employee Master");
    const historySheet = wb.getWorksheet("Posting History");
    let importedSheet = wb.getWorksheet("Imported Emp Profiles");
    if (!importedSheet) {
        importedSheet = wb.addWorksheet("Imported Emp Profiles");
    }

    const masterHeaders = [];
    masterSheet.getRow(1).eachCell((cell, colNum) => {
        masterHeaders[colNum] = String(cell.value || '').trim();
    });

    const mKgidColIdx = masterHeaders.indexOf("KGID");
    const kgid = cleanString(payload.kgid);
    if (!kgid) throw new Error("KGID is required.");

    const name = cleanString(payload.name);
    if (!name) throw new Error("Employee Name is required.");

    const originalKgid = cleanString(payload.originalKgid);
    const isAdd = !originalKgid;

    // Check duplicate KGID in Employee Master
    let kgidRowIdx = -1;
    for (let r = 2; r <= masterSheet.rowCount; r++) {
        const val = cleanString(masterSheet.getCell(r, mKgidColIdx).value);
        if (val === kgid) {
            kgidRowIdx = r;
            break;
        }
    }

    if (isAdd && kgidRowIdx !== -1) {
        const existingName = cleanString(masterSheet.getCell(kgidRowIdx, masterHeaders.indexOf("Employee Name")).value);
        throw new Error(`Duplicate KGID found: ${kgid} already belongs to ${existingName || 'another employee'}.`);
    }

    if (!isAdd && originalKgid !== kgid && kgidRowIdx !== -1) {
        const existingName = cleanString(masterSheet.getCell(kgidRowIdx, masterHeaders.indexOf("Employee Name")).value);
        throw new Error(`Cannot change KGID to ${kgid}; it already belongs to ${existingName || 'another employee'}.`);
    }

    // Find or create row index
    let targetRowIdx = -1;
    if (!isAdd) {
        for (let r = 2; r <= masterSheet.rowCount; r++) {
            if (cleanString(masterSheet.getCell(r, mKgidColIdx).value) === originalKgid) {
                targetRowIdx = r;
                break;
            }
        }
    }
    if (targetRowIdx === -1) {
        targetRowIdx = masterSheet.rowCount + 1;
    }

    // Set employee master values
    const mValues = {
        "KGID": kgid,
        "Employee Name": name,
        "Rank": cleanString(payload.rank),
        "Designation": cleanString(payload.designation) || cleanString(payload.rank),
        "DOB": parseDate(payload.dob),
        "Appointment Date": parseDate(payload.doa),
        "Mobile": cleanString(payload.mobile),
        "Home District": cleanString(payload.homeDistrict),
        "Current District": cleanString(payload.currentDistrict) || "Chikkaballapura",
        "Current Sub-Division": cleanString(payload.currentSubDivision),
        "Current Unit": normalizeStation(payload.currentUnit),
        "Present Posting Date": parseDate(payload.currentSince),
        "Remarks": cleanString(payload.remarks),
        "Category": cleanString(payload.category),
        "Type of Transfer": cleanString(payload.typeOfTransfer) || "Regular"
    };

    // Ensure Type of Transfer header exists
    if (masterHeaders.indexOf("Type of Transfer") === -1) {
        const newCol = masterHeaders.length;
        masterSheet.getCell(1, newCol).value = "Type of Transfer";
        masterHeaders[newCol] = "Type of Transfer";
    }

    for (const [header, val] of Object.entries(mValues)) {
        const col = masterHeaders.indexOf(header);
        if (col !== -1) {
            masterSheet.getCell(targetRowIdx, col).value = val;
        }
    }

    // Write formulas in Employee Master
    const rNum = targetRowIdx;
    const colLetterE = excelColLetter(masterHeaders.indexOf("DOB"));
    const colLetterF = excelColLetter(masterHeaders.indexOf("Appointment Date"));
    const colLetterM = excelColLetter(masterHeaders.indexOf("Present Posting Date"));
    const colLetterA = excelColLetter(masterHeaders.indexOf("KGID"));
    const colLetterK = excelColLetter(masterHeaders.indexOf("Current Sub-Division"));
    const colLetterJ = excelColLetter(masterHeaders.indexOf("Current District"));
    const colLetterO = excelColLetter(masterHeaders.indexOf("Current Station Years"));

    masterSheet.getCell(rNum, masterHeaders.indexOf("Retirement Date")).value = { formula: "=IF(" + colLetterE + rNum + '="","",EOMONTH(EDATE(' + colLetterE + rNum + ",60*12),0))" };
    masterSheet.getCell(rNum, masterHeaders.indexOf("Total Service Years")).value = { formula: "=IF(" + colLetterF + rNum + '="","",ROUND(YEARFRAC(' + colLetterF + rNum + ",TODAY()),1))" };
    masterSheet.getCell(rNum, masterHeaders.indexOf("Current Station Years")).value = { formula: "=IF(" + colLetterM + rNum + '="","",ROUND(YEARFRAC(' + colLetterM + rNum + ",TODAY()),1))" };
    masterSheet.getCell(rNum, masterHeaders.indexOf("Current Sub-Division Years")).value = { formula: "=IF(" + colLetterA + rNum + '="","",ROUND(SUMPRODUCT((\'Posting History\'!$A$2:$A$3000=' + colLetterA + rNum + ")*('Posting History'!$E$2:$E$3000=" + colLetterK + rNum + ")*('Posting History'!$B$2:$B$3000<>\"\")*(IF('Posting History'!$C$2:$C$3000=\"\",TODAY(),'Posting History'!$C$2:$C$3000)-'Posting History'!$B$2:$B$3000))/365.25,1))" };
    masterSheet.getCell(rNum, masterHeaders.indexOf("District Service Years")).value = { formula: "=IF(" + colLetterA + rNum + '="","",ROUND(SUMPRODUCT((\'Posting History\'!$A$2:$A$3000=' + colLetterA + rNum + ")*('Posting History'!$F$2:$F$3000=" + colLetterJ + rNum + ")*('Posting History'!$B$2:$B$3000<>\"\")*(IF('Posting History'!$C$2:$C$3000=\"\",TODAY(),'Posting History'!$C$2:$C$3000)-'Posting History'!$B$2:$B$3000))/365.25,1))" };
    masterSheet.getCell(rNum, masterHeaders.indexOf("Transfer Eligible")).value = { formula: "=IF(" + colLetterO + rNum + '="","",IF(' + colLetterO + rNum + '>=3,"Yes","No"))' };
    masterSheet.getCell(rNum, masterHeaders.indexOf("Priority")).value = { formula: "=IF(" + colLetterO + rNum + '="","",IF(' + colLetterO + rNum + '>=5,"High",IF(' + colLetterO + rNum + '>=3,"Medium","Low")))' };

    // Update Posting History
    const historyHeaders = [];
    historySheet.getRow(1).eachCell((cell, colNum) => {
        historyHeaders[colNum] = String(cell.value || '').trim();
    });

    if (payload.postings) {
        // Delete all postings for originalKgid / kgid
        const searchKgid = originalKgid || kgid;
        const rowsToDelete = [];
        for (let r = 2; r <= historySheet.rowCount; r++) {
            if (cleanString(historySheet.getCell(r, 1).value) === searchKgid) {
                rowsToDelete.push(r);
            }
        }
        // Delete in reverse order
        for (let i = rowsToDelete.length - 1; i >= 0; i--) {
            historySheet.spliceRows(rowsToDelete[i], 1);
        }

        // Add postings from payload
        for (const p of payload.postings) {
            const hRowIdx = historySheet.rowCount + 1;
            writeHistoryRow(historySheet, hRowIdx, {
                "KGID": kgid,
                "From Date": parseDate(p.from),
                "To Date": (p.to && p.to !== "Present") ? parseDate(p.to) : null,
                "Police Station / Unit": normalizeStation(p.station),
                "Sub-Division": cleanString(p.subDivision) || getSubdivisionForStation(normalizeStation(p.station)),
                "District": cleanString(p.district) || "Chikkaballapura",
                "Rank Held": cleanString(p.rank),
                "Order Number": cleanString(p.orderNumber),
                "Notes": "Maintained from browser dashboard."
            }, historyHeaders);
        }
    } else {
        // Fallback single posting
        let hRowIdx = -1;
        for (let r = 2; r <= historySheet.rowCount; r++) {
            if (cleanString(historySheet.getCell(r, 1).value) === kgid) {
                hRowIdx = r;
                break;
            }
        }
        if (hRowIdx === -1) {
            hRowIdx = historySheet.rowCount + 1;
        }
        writeHistoryRow(historySheet, hRowIdx, {
            "KGID": kgid,
            "From Date": parseDate(payload.currentSince),
            "To Date": null,
            "Police Station / Unit": normalizeStation(payload.currentUnit),
            "Sub-Division": cleanString(payload.currentSubDivision),
            "District": cleanString(payload.currentDistrict) || "Chikkaballapura",
            "Rank Held": cleanString(payload.rank),
            "Order Number": "",
            "Notes": "Current profile maintained from browser dashboard."
        }, historyHeaders);
    }

    // Update Imported Emp Profiles
    const impHeaders = [];
    importedSheet.getRow(1).eachCell((cell, colNum) => {
        impHeaders[colNum] = String(cell.value || '').trim();
    });
    let impKgidCol = impHeaders.indexOf("kgid");
    if (impKgidCol === -1) {
        const defaultImpHeaders = [
            "kgid", "name", "mobile1", "mobile2", "rank", "station", "district",
            "metalNumber", "bloodGroup", "email", "isAdmin", "isApproved",
            "createdAt", "updatedAt", "isDeleted", "height", "weight", "caste",
            "subCaste", "familyDetails", "photoUrl", "lockedFields", "darWing"
        ];
        importedSheet.getRow(1).values = defaultImpHeaders;
        impHeaders.push(...defaultImpHeaders);
        impKgidCol = 1;
    }

    let impRowIdx = -1;
    const searchKgid2 = originalKgid || kgid;
    for (let r = 2; r <= importedSheet.rowCount; r++) {
        if (cleanString(importedSheet.getCell(r, impKgidCol).value) === searchKgid2) {
            impRowIdx = r;
            break;
        }
    }
    if (impRowIdx === -1) {
        impRowIdx = importedSheet.rowCount + 1;
    }

    const source = payload.sourceProfile || {};
    const impValues = {
        "kgid": kgid,
        "name": name,
        "mobile1": cleanString(payload.mobile),
        "mobile2": cleanString(source.mobile2),
        "rank": cleanString(payload.rank),
        "station": normalizeStation(payload.currentUnit),
        "district": cleanString(payload.currentDistrict) || "Chikkaballapura",
        "metalNumber": cleanString(source.metalNumber),
        "bloodGroup": cleanString(source.bloodGroup),
        "email": cleanString(source.email),
        "isAdmin": cleanString(source.isAdmin),
        "isApproved": cleanString(source.isApproved) || "Yes",
        "createdAt": cleanString(source.createdAt),
        "updatedAt": formatDate(new Date()) + " " + new Date().toTimeString().split(' ')[0],
        "isDeleted": "No",
        "height": cleanString(source.height),
        "weight": cleanString(source.weight),
        "caste": cleanString(source.caste),
        "subCaste": cleanString(source.subCaste),
        "familyDetails": cleanString(source.familyDetails),
        "photoUrl": cleanString(source.photoUrl),
        "lockedFields": cleanString(source.lockedFields),
        "darWing": cleanString(source.darWing)
    };

    // Ensure all headers exist in imported
    for (const h of Object.keys(impValues)) {
        if (impHeaders.indexOf(h) === -1) {
            const newCol = impHeaders.length;
            importedSheet.getCell(1, newCol).value = h;
            impHeaders[newCol] = h;
        }
    }

    for (const [header, val] of Object.entries(impValues)) {
        const col = impHeaders.indexOf(header);
        if (col !== -1) {
            importedSheet.getCell(impRowIdx, col).value = val;
        }
    }

    // Update table ranges
    const tableSpecs = [
        ["Employee Master", "tblEmployeeMaster"],
        ["Posting History", "tblPostingHistory"],
        ["Imported Emp Profiles", "tblImportedEmpProfiles"],
    ];
    for (const [sheetName, tableName] of tableSpecs) {
        const sheet = wb.getWorksheet(sheetName);
        if (sheet && sheet.tables && sheet.tables[tableName]) {
            const lastRow = Math.max(2, sheet.rowCount);
            const lastColLetter = excelColLetter(sheet.columnCount);
            sheet.tables[tableName].ref = "A1:" + lastColLetter + lastRow;
        }
    }

    await wb.xlsx.writeFile(activeMasterPath);

    // Rebuild data in index.html
    const updatedJson = await buildStandaloneJSON(activeMasterPath);
    const indexHtmlPath = path.join(__dirname, '..', 'DPDMS', 'index.html');
    updateIndexHtml(indexHtmlPath, updatedJson);

    const standaloneHtmlPath = path.join(outputsDir, 'Employee_Transfer_Dashboard_Browser.html');
    if (fs.existsSync(standaloneHtmlPath)) {
        try {
            updateIndexHtml(standaloneHtmlPath, updatedJson);
        } catch (e) {}
    }

    return { ok: true, message: `Saved ${name} (${kgid})` };
}

async function deleteEmployee(kgid) {
    const outputsDir = path.join(__dirname, '..', 'outputs');
    const masterPath = path.join(outputsDir, 'District_Employee_Transfer_Database_IMPORTED.xlsx');
    const fallbackTemplatePath = path.join(outputsDir, 'District_Employee_Transfer_Database_Template_FIXED.xlsx');
    const activeMasterPath = fs.existsSync(masterPath) ? masterPath : fallbackTemplatePath;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(activeMasterPath);

    const masterSheet = wb.getWorksheet("Employee Master");
    const historySheet = wb.getWorksheet("Posting History");
    const importedSheet = wb.getWorksheet("Imported Emp Profiles");
    const transferRequestsSheet = wb.getWorksheet("Transfer Requests");

    const masterHeaders = [];
    masterSheet.getRow(1).eachCell((cell, colNum) => {
        masterHeaders[colNum] = String(cell.value || '').trim();
    });
    const mKgidColIdx = masterHeaders.indexOf("KGID");

    let masterRowIdx = -1;
    for (let r = 2; r <= masterSheet.rowCount; r++) {
        if (cleanString(masterSheet.getCell(r, mKgidColIdx).value) === String(kgid).trim()) {
            masterRowIdx = r;
            break;
        }
    }

    if (masterRowIdx === -1) {
        throw new Error(`Employee ${kgid} not found in Employee Master.`);
    }

    const name = cleanString(masterSheet.getCell(masterRowIdx, masterHeaders.indexOf("Employee Name")).value);

    // Delete row from Employee Master
    masterSheet.spliceRows(masterRowIdx, 1);

    // Delete rows from Posting History
    const rowsToDelete = [];
    for (let r = 2; r <= historySheet.rowCount; r++) {
        if (cleanString(historySheet.getCell(r, 1).value) === String(kgid).trim()) {
            rowsToDelete.push(r);
        }
    }
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
        historySheet.spliceRows(rowsToDelete[i], 1);
    }

    // Delete row from Imported Emp Profiles
    if (importedSheet) {
        const impHeaders = [];
        importedSheet.getRow(1).eachCell((cell, colNum) => {
            impHeaders[colNum] = String(cell.value || '').trim();
        });
        const impKgidColIdx = impHeaders.indexOf("kgid");
        if (impKgidColIdx !== -1) {
            let impRowIdx = -1;
            for (let r = 2; r <= importedSheet.rowCount; r++) {
                if (cleanString(importedSheet.getCell(r, impKgidColIdx).value) === String(kgid).trim()) {
                    impRowIdx = r;
                    break;
                }
            }
            if (impRowIdx !== -1) {
                importedSheet.spliceRows(impRowIdx, 1);
            }
        }
    }

    // Delete row from Transfer Requests
    if (transferRequestsSheet) {
        const trHeaders = [];
        transferRequestsSheet.getRow(1).eachCell((cell, colNum) => {
            trHeaders[colNum] = String(cell.value || '').trim();
        });
        const trKgidColIdx = trHeaders.indexOf("KGID");
        if (trKgidColIdx !== -1) {
            let trRowIdx = -1;
            for (let r = 2; r <= transferRequestsSheet.rowCount; r++) {
                if (cleanString(transferRequestsSheet.getCell(r, trKgidColIdx).value) === String(kgid).trim()) {
                    trRowIdx = r;
                    break;
                }
            }
            if (trRowIdx !== -1) {
                transferRequestsSheet.spliceRows(trRowIdx, 1);
            }
        }
    }

    // Update table ranges
    const tableSpecs = [
        ["Employee Master", "tblEmployeeMaster"],
        ["Posting History", "tblPostingHistory"],
        ["Imported Emp Profiles", "tblImportedEmpProfiles"],
        ["Transfer Requests", "tblTransferRequests"],
    ];
    for (const [sheetName, tableName] of tableSpecs) {
        const sheet = wb.getWorksheet(sheetName);
        if (sheet && sheet.tables && sheet.tables[tableName]) {
            const lastRow = Math.max(2, sheet.rowCount);
            const lastColLetter = excelColLetter(sheet.columnCount);
            sheet.tables[tableName].ref = "A1:" + lastColLetter + lastRow;
        }
    }

    await wb.xlsx.writeFile(activeMasterPath);

    // Rebuild data in index.html
    const updatedJson = await buildStandaloneJSON(activeMasterPath);
    const indexHtmlPath = path.join(__dirname, '..', 'DPDMS', 'index.html');
    updateIndexHtml(indexHtmlPath, updatedJson);

    const standaloneHtmlPath = path.join(outputsDir, 'Employee_Transfer_Dashboard_Browser.html');
    if (fs.existsSync(standaloneHtmlPath)) {
        try {
            updateIndexHtml(standaloneHtmlPath, updatedJson);
        } catch (e) {}
    }

    return { ok: true, message: `Deleted ${name} (${kgid})` };
}

async function checkDuplicate(query) {
    const outputsDir = path.join(__dirname, '..', 'outputs');
    const masterPath = path.join(outputsDir, 'District_Employee_Transfer_Database_IMPORTED.xlsx');
    const fallbackTemplatePath = path.join(outputsDir, 'District_Employee_Transfer_Database_Template_FIXED.xlsx');
    const activeMasterPath = fs.existsSync(masterPath) ? masterPath : fallbackTemplatePath;

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(activeMasterPath);

    const masterSheet = wb.getWorksheet("Employee Master");
    const headers = [];
    masterSheet.getRow(1).eachCell((cell, colNum) => {
        headers[colNum] = String(cell.value || '').trim();
    });

    const mKgidColIdx = headers.indexOf("KGID");
    const mNameColIdx = headers.indexOf("Employee Name");
    const mMobileColIdx = headers.indexOf("Mobile");

    const kgid = cleanString(query.kgid);
    const originalKgid = cleanString(query.originalKgid);
    const mobile = cleanString(query.mobile);
    const mobile2 = cleanString(query.mobile2);
    const isAdd = !originalKgid;

    const sameRecord = (rowKgid) => {
        return (!isAdd) && (rowKgid === originalKgid || rowKgid === kgid);
    };

    const issues = [];
    for (let r = 2; r <= masterSheet.rowCount; r++) {
        const rowKgid = cleanString(masterSheet.getCell(r, mKgidColIdx).value);
        const rowName = cleanString(masterSheet.getCell(r, mNameColIdx).value);
        const rowMobile = cleanString(masterSheet.getCell(r, mMobileColIdx).value);

        if (kgid && rowKgid === kgid && !sameRecord(rowKgid)) {
            issues.push({ field: "kgid", message: `KGID already exists: ${kgid} - ${rowName}` });
        }
        for (const [field, mobileVal] of [["mobile", mobile], ["mobile2", mobile2]]) {
            if (mobileVal && rowMobile === mobileVal && !sameRecord(rowKgid) && !issues.some(i => i.field === field)) {
                issues.push({ field: field, message: `Mobile already exists: ${mobileVal} - ${rowName || rowKgid}` });
            }
        }
    }

    return issues;
}

module.exports = {
    handleMergeAndSync,
    saveTransferRequest,
    deleteTransferRequest,
    saveEmployee,
    deleteEmployee,
    checkDuplicate,
    buildStandaloneJSON
};
