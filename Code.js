/***** =========================================================================================
 *
 * ERP - Main Server Code
 *
 * ========================================================================================= *****/


/***** ========== CONFIG ========== *****/
const DOCS_FOLDER_ID = ''; // Optional: put a folder ID here. If blank, a folder "ERP_Attachments" is used/created.
const SPREADSHEET_ID = '1yTWsbGPyZK1j5da28EF66l_ZCeAdf_2Cab36IWkeWCU';


const SHEETS = {
  // System Sheets
  SYS_DF_VALIDATION: 'SYS_DF_Validation',
  SYS_SETTINGS: 'SYS_Settings',
  SYS_AUDIT_REPORT: 'SYS_Audit_Report',
  SYS_DOCUMENTS: 'SYS_Documents',
  SYS_PROFILE_VIEW: 'SYS_Profile_View',
  SYS_DROPDOWNS: 'SYS_Dropdowns',
  SYS_DYNAMIC_FORMS: 'SYS_Dynamic_Forms',
  SYS_USERS: 'SYS_Users',
  SYS_ROLES: 'SYS_Roles',
  SYS_PERMISSIONS: 'SYS_Permissions',
  SYS_AUDIT_LOG: 'SYS_Audit_Log',
  SYS_PUB_HOLIDAYS: 'SYS_PubHolidays',
  // HR Sheets
  HR_EMPLOYEES: 'HR_Employees',
  HR_ATTENDANCE: 'HR_Attendance',
  HR_LEAVE_REQUESTS: 'HR_Leave_Requests',
  HR_ABSENCE_DEDUCTIONS: 'HR_Absence_Deductions',
  HR_LEAVE: 'HR_Leave',
  HR_LEAVE_ANALYSIS: 'HR_Leave_Analysis',
  HR_ADVANCES: 'HR_Advances',
  HR_OVERTIME: 'HR_OverTime',
  HR_PENALTIES: 'HR_Penalties',
  HR_DEDUCTIONS: 'HR_Deductions',
  HR_PAYROLL: 'HR_Payroll',
  HR_PAYROLL_INPUTS: 'HR_Payroll_Inputs',
  HR_DASHBOARD: 'HR_Dashboard',
  HR_KPIS: 'HR_KPIs',
  // Project Sheets
  PRJ_MAIN: 'PRJ_Main',
  PRJ_TASKS: 'PRJ_Tasks',
  PRJ_COSTS: 'PRJ_Costs',
  PRJ_CLIENTS: 'PRJ_Clients',
  PRJ_MATERIALS: 'PRJ_Materials',
  PRJ_INDIR_EXP_ALLOCATIONS: 'PRJ_InDirExp_Allocations',
  PRJ_SCHEDULE_CALC: 'PRJ_Schedule_Calc',
  PRJ_DASHBOARD: 'PRJ_Dashboard',
  PRJ_KPIS: 'PRJ_KPIs',
  // Finance Sheets
  FIN_DIRECT_EXPENSES: 'FIN_DirectExpenses',
  FIN_INDIR_EXPENSE_REPEATED: 'FIN_InDirExpense_Repeated',
  FIN_INDIR_EXPENSE_ONCE: 'FIN_InDirExpense_Once',
  FIN_PROJECT_REVENUE: 'FIN_Project_Revenue',
  FIN_REVENUES: 'FIN_Revenues',
  FIN_JOURNAL: 'FIN_Journal',
  FIN_CUSTODY: 'FIN_Custody',
  FIN_CHART_OF_ACCOUNTS: 'FIN_ChartOfAccounts',
  FIN_GL_TOTALS: 'FIN_GL_Totals',
  FIN_DASHBOARD: 'FIN_Dashboard',
  FIN_KPIS: 'FIN_KPIs'
};

const HR_FORM_KEYS = {
  EMP_DATA: 'EMP001',
  ATTENDANCE: 'ATT001',
  LEAVE_REQUESTS: 'LEAVE001',
  ABSENCE_DEDUCTIONS: 'ABS001',
  OVERTIME: 'OT001',
  ADVANCES: 'ADV001',
  PAYROLL: 'PAY001'
};




/***** =========================================================================================
 *
 * SYSTEM CORE FUNCTIONS
 *
 * ========================================================================================= *****/


/***** ==================== ROUTING ==================== *****/


/**
 * ROUTING -- Handles Web App Requests
 * This is the main entry point for the web app. It serves the correct HTML page
 * based on the 'page' URL parameter (e.g., ?page=dashboard).
 */
function doGet(e) {
  const page = e?.parameter?.page || 'login';
  const fileMap = { dashboard: 'Dashboard', dynamicForm: 'DynamicForm', login: 'Login' };
  const fileName = fileMap[page] || fileMap.login;
  try {
    const template = HtmlService.createTemplateFromFile(fileName);
    template.cacheBuster = new Date().getTime();
    return template.evaluate()
      .setTitle(page === 'dashboard' ? 'ERP Dashboard' : page === 'dynamicForm' ? 'Dynamic Form' : 'ERP Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(`<pre>${String(err.stack)}</pre>`).setTitle('Error');
  }
}


/**
 * ROUTING -- Get Web App URL
 * A helper function that returns the base URL of the deployed web app.
 */
function getWebAppBaseUrl() { return ScriptApp.getService().getUrl(); }




/***** ==================== AUTH ==================== *****/


/**
 * AUTH -- User Login
 * Authenticates a user against the SYS_Users sheet using a salted hash.
 */
function login(username, password) {
  try {
    const sh = getSheet(SHEETS.SYS_USERS);
    if (!sh) return { success: false, message: 'Users sheet not found.' };
    const values = sh.getDataRange().getValues();
    if (!values.length) return { success: false, message: 'Users sheet empty.' };
    const headers = values.shift();
const ixMap = _getHeadersMap(headers);
const ix = { user: ixMap['User_Name'], salt: ixMap['Password_Salt'], hash: ixMap['Password_Hash'], status: ixMap['Status'], role: ixMap['Role_ID'], name: ixMap['Full_Name_EN'], id: ixMap['User_ID'], email:ixMap['Email'] };    if (Object.values(ix).some(i => i === -1)) return { success: false, message: 'Missing columns in SYS_Users.' };
    const row = values.find(r => String(r[ix.user]) === String(username));
    if (!row || String(row[ix.status]).trim().toUpperCase() !== 'ACTIVE' || hashPasswordWithSalt(password, row[ix.salt]) !== String(row[ix.hash] || '')) {
      return { success: false, message: 'Invalid username or password.' };
    }
    const user = { userId: row[ix.id], userName: row[ix.user], fullName: row[ix.name], email: row[ix.email], roleId: row[ix.role], status: row[ix.status] };
    logAction(user.userId, user.fullName, 'LOGIN_SUCCESS', 'User logged in', '');
    return { success: true, message: 'OK', user, redirectUrl: getWebAppBaseUrl() + '?page=dashboard' };
  } catch (err) {
    Logger.log('[login] ' + err);
    return { success: false, message: 'Login error.' };
  }
}


/**
 * AUTH -- User Logout
 * Clears any session-related data for the user.
 */
function logout(){
  try{
    return { success:true };
  }catch(e){
    return { success:false, message:String(e) };
  }
}


/**
 * AUTH -- Password Hashing Helper
 * Creates a salted SHA-256 hash for a given password.
 */
function hashPasswordWithSalt(pw, salt) {
  return sha256b64_(String(pw) + String(salt || ''));
}




/***** ==================== UTILS ==================== *****/


/**
 * UTILS -- Get Sheet by Name
 * A robust helper to safely retrieve a sheet object by its name.
 */
function getSheet(name) {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  } catch (err) {
    Logger.log(`[getSheet: ${name}] ${err}`);
    return null;
  }
}


/**
 * UTILS -- Log User Action
 * Appends a record to the SYS_Audit_Log sheet.
 */
function logAction(userId, fullName, actionKey, actionDesc, description) {
  try {
    getSheet(SHEETS.SYS_AUDIT_LOG).appendRow(['', userId, fullName, actionKey, actionDesc, new Date(), 'Web', description || '']);
  } catch (err) {
    Logger.log('[logAction] ' + err);
  }
}


/**
 * UTILS -- SHA-256 Hashing Utility
 * A generic helper function to compute a SHA-256 hash and return it as a Base64 string.
 */
function sha256b64_(s) {
  const bytes = Utilities.newBlob(String(s)).getBytes();
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes));
}


/**
 * UTILS -- Profile View ID Normalizer
 * A helper to robustly normalize IDs for matching (uppercase, no spaces, unified dashes).
 */
function _pv_norm_(s){
  return String(s||'').replace(/[\u200f\u200e\u202a-\u202e]/g,'').trim().toUpperCase().replace(/\s+/g,'').replace(/[-\u2013\u2014\u2212\u2010]/g,'-');
}


/**
 * UTILS -- Profile View Value Formatter
 * A helper to format values (especially dates and phone numbers) for display.
 */
function _pv_fmt_(val, fmt){
  if (val instanceof Date){
    const tz = Session.getScriptTimeZone();
    if (fmt === 'datetime') return Utilities.formatDate(val, tz, 'yyyy-MM-dd HH:mm');
    if (fmt === 'date') return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
    return Utilities.formatDate(val, tz, 'yyyy-MM-dd HH:mm');
  }
  if (!fmt) return val;
  const s = String(val == null ? '' : val);
  if (fmt === 'phone'){
    const digits = s.replace(/\D+/g,'');
    if (/^\d{9}$/.test(digits)) return '0' + digits;
    if (/^\d{10}$/.test(digits) && digits[0] !== '0') return '0' + digits;
  }
  return s;
}

/**
 * Header -> index map helper.
 * Returns -1 for missing keys (so existing callers using === -1 keep working).
 */
function _getHeadersMap(headers) {
  const map = {};
  if (!Array.isArray(headers)) return new Proxy(map, { get: () => -1 });
  headers.forEach((h, i) => {
    const key = (h === null || h === undefined) ? '' : String(h).trim();
    map[key] = i;
  });
  return new Proxy(map, {
    get(target, prop) {
      if (prop in target) return target[prop];
      return -1;
    }
  });
}

/* Server test:
   Run __test_headers() in the Apps Script editor.
*/
function __test_headers() {
  try {
    const m = _getHeadersMap(['A','B']);
    Logger.log('A=' + m['A'] + ' MISSING=' + m['X_MISSING']);
  } catch (e) { Logger.log('[__test_headers] ' + e); }
}



/**
 * UTILS -- Filter Rows by Key Helper
 * Filters a 2D array of rows where a specific column's value matches a key.
 */
function _filterRowsByKey_(headers, rows, keyColumn, keyValue) {
    const keyIndex = headers.indexOf(keyColumn);
    const normalizedKeyValue = _pv_norm_(keyValue);
    if (keyIndex === -1) return [];
    return rows.filter(r => _pv_norm_(r[keyIndex]) === normalizedKeyValue);
}

/**
 * UTILS -- Get Related Attachments Helper
 * Retrieves all document records from SYS_Documents related to a specific entity ID.
 */
function _getRelatedAttachments_(entityId, memoizedReadSheet) {
    const p = memoizedReadSheet(SHEETS.SYS_DOCUMENTS);
    const h = p.headers;
    if (!h.length) return [];
    const iId = h.indexOf('Entity_ID');
    const iLbl = h.indexOf('Label');
    const iFnm = h.indexOf('File_Name');
    const iUrl = h.indexOf('Drive_URL');
    if (iId === -1) return [];
    
    return p.rows
        .filter(r => _pv_norm_(r[iId]) === _pv_norm_(entityId))
        .map(r => ({ label: r[iLbl] || r[iFnm] || '', url: r[iUrl] || '' }));
}

/**
 * UTILS -- Dropdown Cache Helper
 * Caches the dropdown map for 10 minutes to improve performance.
 */
const __DD_CACHE = { mapByKey: null, ts: 0 };

function _getDropdownMap() {
    const now = Date.now();
    // Return cache if it's less than 10 minutes old
    if (__DD_CACHE.mapByKey && (now - __DD_CACHE.ts < 600 * 1000)) {
        return __DD_CACHE.mapByKey;
    }

    const sh = getSheet(SHEETS.SYS_DROPDOWNS);
    const map = new Map(); // key -> [values]
    if (sh) {
        const vals = sh.getDataRange().getValues();
        if (vals.length > 1) {
            const h = vals[0];
            const H = x => h.indexOf(x);
            const iK = H('Key'), iL2 = H('Level_2'), iA = H('Is_Active'), iS = H('Sort_Order');

            const rows = vals.slice(1)
                .filter(r => String(r[iA] || 'Yes').toUpperCase() !== 'NO')
                .map(r => ({
                    key: String(r[iK] || '').trim(),
                    txt: String(r[iL2] || '').trim(),
                    sort: Number(r[iS] || 999)
                }))
                .filter(x => x.key && x.txt);

            rows.sort((a, b) => a.key.localeCompare(b.key) || a.sort - b.sort || a.txt.localeCompare(b.txt, 'ar'));

            rows.forEach(x => {
                if (!map.has(x.key)) map.set(x.key, []);
                map.get(x.key).push(x.txt);
            });
        }
    }
    __DD_CACHE.mapByKey = map;
    __DD_CACHE.ts = now;
    return map;
}

/**
 * UTILS -- Get Dropdown Values by Key
 * Retrieves a sorted list of allowed values for a given dropdown key.
 */
function getAllowedDropdownValues(key) {
    const map = _getDropdownMap();
    return map.get(String(key || '').trim()) || [];
}

/**
 * UTILS -- Normalize Dropdown Value
 * Normalizes any value to an allowed dropdown value for the given key using synonyms.
 */
function normalizeDropdownValue(dropdownKey, value) {
    const v = String(value || '').trim();
    if (!v) return '';

    const allowed = getAllowedDropdownValues(dropdownKey);
    if (!allowed.length) return v; // No validation configured, return original
    if (allowed.includes(v)) return v; // Already perfect

    // Synonym map for common translations or aliases
    const __DD_SYNONYMS = {
        PRJ_STATUS: new Map([
            ['ACTIVE', 'نشط'], ['PAUSED', 'موقوف'], ['COMPLETED', 'مكتمل'], ['CANCELLED', 'ملغي']
        ]),
        PAY_STATUS: new Map([
            ['PAID', 'مدفوع'], ['UNPAID', 'غير مدفوع'], ['PENDING', 'معلق']
        ]),
        UNIT: new Map([
            ['SHEET', 'لوح'], ['PCS', 'قطعة'], ['M', 'متر'], ['KG', 'كجم']
        ])
    };

    const synMap = __DD_SYNONYMS[dropdownKey];
    if (synMap) {
        const normV = v.toUpperCase();
        // Check if the value is a key (e.g., "ACTIVE") or a value (e.g., "نشط") in the map
        for (const [key, val] of synMap.entries()) {
            if (key === normV || val.toUpperCase() === normV) {
                if (allowed.includes(val)) return val;
            }
        }
    }

    // Fallback: if no match, return the first allowed value as a safe default
    return allowed[0] || '';
}

/**
 * Get only the JS inside <script>...</script> blocks of an HTML file.
 */
function getRawJs(filename) {
  try {
    const html = HtmlService.createHtmlOutputFromFile(filename).getContent();
    const scriptRegex = /<script\b[^>]*>([\s\S]*?)<\/script>/gi;
    let match;
    let out = '';
    while ((match = scriptRegex.exec(html)) !== null) {
      out += match[1] + '\n';
    }
    return out.trim();
  } catch (e) {
    Logger.log('[getRawJs] ' + e);
    return '';
  }
}

/**
 * Debug helper: returns a small preview of extracted JS and leftover HTML.
 */
function getRawJsDebug(filename) {
  try {
    const html = HtmlService.createHtmlOutputFromFile(filename).getContent();
    const scriptRegex = /<script\b[^>]*>([\s\S]*?)<\/script>/gi;
    let match;
    let out = '';
    let htmlWithoutScripts = html.replace(scriptRegex, '');
    while ((match = scriptRegex.exec(html)) !== null) {
      out += match[1] + '\n';
    }
    return {
      jsPreview: out.trim().slice(0, 2000),
      leftoverHtmlPreview: htmlWithoutScripts.trim().slice(0, 2000),
      jsLength: out.length,
      leftoverLength: htmlWithoutScripts.trim().length
    };
  } catch (e) {
    Logger.log('[getRawJsDebug] ' + e);
    return { jsPreview: '', leftoverHtmlPreview: '', jsLength:0, leftoverLength:0 };
  }
}


/***** ==================== Auto-ID Generator ==================== *****/


/**
 * Auto-ID -- Get Next ID
 * Generates the next sequential ID for a given column (e.g., 'EMP-0004' -> 'EMP-0005').
 */
function getNextAutoValue(targetSheet, targetColumn, prefix, padLen) {
  try {
    const sh = getSheet(targetSheet);
    if (!sh) return '';
    const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const colIdx = header.indexOf(targetColumn);
    if (colIdx === -1) return '';
    const lastRow = sh.getLastRow();
    const data = (lastRow > 1) ? sh.getRange(2, colIdx + 1, lastRow - 1, 1).getDisplayValues() : [];
    let maxNum = 0;
    data.forEach(([val]) => {
      const m = String(val || '').match(/(\d+)\s*$/);
      if (m && m[1]) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    });
    const pfx = (prefix || '').toString().replace(/-+$/,'');
    return (pfx ? `${pfx}-` : '') + String(maxNum + 1).padStart(padLen || 4, '0');
  } catch (e) {
    Logger.log(`[getNextAutoValue: ${targetSheet}] ${e}`);
    return '';
  }
}
function __test_getRawJs_debug(){ Logger.log(JSON.stringify(getRawJsDebug('HR_Employees_JS'))); }


/**
 * Auto-ID -- Server-side Wrapper for Client
 * A client-callable wrapper for the getNextAutoValue function.
 */
function getNextAutoValueServer(targetSheet, targetColumn, prefix, padLen) {
  return getNextAutoValue(targetSheet, targetColumn, prefix, padLen);
}


/**
 * Auto-ID -- Get Next Document ID
 * A specialized helper to generate the next ID specifically for the SYS_Documents sheet.
 */
function nextDocId_() {
  return getNextAutoValue(SHEETS.SYS_DOCUMENTS, 'Doc_ID', 'DOC', 5);
}




/***** ==================== DROPDOWNS ==================== *****/


/**
 * DROPDOWNS -- Get Options by Key
 * Retrieves all dropdown options for a specific key from SYS_Dropdowns.
 */
function getDropdownOptions(dropdownKey) {
  const wantKey = String(dropdownKey || '').trim().toUpperCase();
  if (!wantKey) return [];
  const sh = getSheet(SHEETS.SYS_DROPDOWNS);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values.shift();
const ixMap = _getHeadersMap(header);
const ix = { Key: ixMap['Key'], L1: ixMap['Level_1'], L2: ixMap['Level_2'], Active: ixMap['Is_Active'], Sort: ixMap['Sort_Order'] };  if (ix.Key === -1 || ix.L2 === -1) return [];
  const rows = [];
  values.forEach(r => {
    if (String(r[ix.Key]||'').trim().toUpperCase() === wantKey && String(r[ix.Active]||'Yes').trim().toUpperCase() !== 'NO'){
      const ar = String(r[ix.L2]||'').trim();
      if(ar) rows.push({ value: ar, text_ar: ar, text_en: String(r[ix.L1]||''), sort: Number(r[ix.Sort]||999) });
    }
  });
  rows.sort((a,b)=>a.sort - b.sort || a.text_ar.localeCompare(b.text_ar,'ar'));
  return rows.map(({value, text_ar, text_en})=>({value, text_ar, text_en}));
}


/**
 * DROPDOWNS -- Get All Arabic Options
 * Retrieves all dropdowns from SYS_Dropdowns and maps them by key, using only Arabic values.
 */
function getSysDropdownsMap_Arabic() {
  const sh = getSheet(SHEETS.SYS_DROPDOWNS);
  if (!sh) return {};
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return {};
  const head = vals.shift();
  const H = h => head.indexOf(h);
  const cKey = H('Key'), cL2 = H('Level_2');
  const map = {};
  vals.forEach(r => {
    const k = String(r[cKey]||'').trim(), ar = String(r[cL2]||'').trim();
    if (k && ar) {
      if (!map[k]) map[k] = [];
      map[k].push({ value: ar, text_ar: ar });
    }
  });
  return map;
}




/***** ==================== DYNAMIC FORMS VALIDATOR ==================== *****/


/**
 * VALIDATOR -- Add Menu to Spreadsheet
 * Runs when the spreadsheet is opened to add a custom menu for validating the forms configuration.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('ERP Tools')
      .addItem('Validate Dynamic Forms', 'validateDynamicForms')
      .addItem('Validate Schema Constants', 'validateSchemaConstants') // If you added this before
      .addSeparator()
      .addItem('Generate Schema for Analyzer', 'generateSchemaForAnalyzer')
      .addToUi();
  } catch (e) {
    Logger.log('[onOpen] ' + e);
  }
}


/**
 * VALIDATOR -- Dynamic Forms Configuration Checker
 * Scans the SYS_Dynamic_Forms sheet for common errors (e.g., missing targets, bad ranges)
 * and generates a report in a new sheet for developers to review.
 */
function validateDynamicForms() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const shDF = ss.getSheetByName(SHEETS.SYS_DYNAMIC_FORMS);
    if (!shDF) throw new Error('SYS_Dynamic_Forms not found.');
    const values = shDF.getDataRange().getValues();
    if (!values.length) throw new Error('SYS_Dynamic_Forms is empty.');

    const headers = values.shift();
    const ix = _getHeadersMap(headers); // <-- REPLACED H HELPER

    const rptName = 'SYS_DF_Validation';
    let shRpt = ss.getSheetByName(rptName);
    if (shRpt) ss.deleteSheet(shRpt);
    shRpt = ss.insertSheet(rptName);

    const out = [['Status', 'Message', 'Form_ID', 'Field_ID', 'Row#']];
    const sheetCache = new Map(), headerCache = new Map(), seenTarget = new Set(), seenFieldId = new Set();
    const getSheetSafe = name => { if (!name) return null; if (sheetCache.has(name)) return sheetCache.get(name); const s = getSheet(name); sheetCache.set(name, s); return s; };
    const getHeadersSafe = name => { if (!name) return []; if (headerCache.has(name)) return headerCache.get(name); const s = getSheetSafe(name); const hdr = s ? s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0] : []; headerCache.set(name, hdr); return hdr; };

    values.forEach((row, i) => {
        // Use the ix map for clarity and performance
        const formId = String(row[ix['Form_ID']] || ''), fieldId = String(row[ix['Field_ID']] || ''), fType = String(row[ix['Field_Type']] || '').toLowerCase();
        const tgtSheet = String(row[ix['Target_Sheet']] || ''), tgtCol = String(row[ix['Target_Column']] || ''), ddKey = String(row[ix['Dropdown_Key']] || '');
        const srcSheet = String(row[ix['Source_Sheet']] || ''), srcRange = String(row[ix['Source_Range']] || '');
        const msgs = [];

        if (!formId) msgs.push({ lvl: 'ERROR', msg: 'Form_ID is blank.' });
        if (fType === 'dropdown' && !ddKey) msgs.push({ lvl: 'ERROR', msg: 'Dropdown field without Dropdown_Key.' });
        if ((fType === 'auto' || fType === 'autogen') && (!tgtSheet || !tgtCol)) msgs.push({ lvl: 'ERROR', msg: 'Auto/Autogen requires Target_Sheet and Target_Column.' });
        if (fType === 'autofill' && (!srcSheet || !srcRange)) msgs.push({ lvl: 'ERROR', msg: 'Autofill requires Source_Sheet and Source_Range.' });
        if (tgtSheet) { const t = getSheetSafe(tgtSheet); if (!t) msgs.push({ lvl: 'ERROR', msg: `Target_Sheet "${tgtSheet}" not found.` }); else if (tgtCol && getHeadersSafe(tgtSheet).indexOf(tgtCol) === -1) msgs.push({ lvl: 'ERROR', msg: `Target_Column "${tgtCol}" not found in "${tgtSheet}".` }); }
        if (formId && fieldId) { const k = `${formId}|${fieldId}`; if (seenFieldId.has(k)) msgs.push({ lvl: 'WARN', msg: 'Duplicate Field_ID.' }); else seenFieldId.add(k); }
        if (formId && tgtSheet && tgtCol) { const k2 = `${formId}|${tgtSheet}|${tgtCol}`; if (seenTarget.has(k2)) msgs.push({ lvl: 'WARN', msg: 'Duplicate Target mapping.' }); else seenTarget.add(k2); }
        if (!msgs.length) msgs.push({ lvl: 'OK', msg: 'Row looks good.' });

        msgs.forEach(m => out.push([m.lvl, m.msg, formId, fieldId, i + 2]));
    });

    shRpt.getRange(1, 1, out.length, out[0].length).setValues(out);
    shRpt.autoResizeColumns(1, out[0].length);
    const rng = shRpt.getRange(2, 1, shRpt.getLastRow() - 1, 1);
    const rules = [SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ERROR').setBackground('#ffd7d7').setBold(true).setRanges([rng]).build(), SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('WARN').setBackground('#fff7d6').setRanges([rng]).build()];
    shRpt.setConditionalFormatRules(rules);
    SpreadsheetApp.flush();
}




/***** ==================== DYNAMIC FORMS & PROFILES Engine ==================== *****/

/***** =========================================================
 * Projects – Core helpers & post-save processing (FULL)
 * ========================================================= *****/

/** حساب نهاية مخططة (أيام عمل: أحد→خميس؛ استبعاد جمعة/سبت والعطلات العامة) */
function calcPlannedEndDate(startDate, plannedDays){
  try{
    if (!(startDate instanceof Date) || !plannedDays || plannedDays <= 0) return '';
    const sh = getSheet(SHEETS.SYS_PUB_HOLIDAYS);
    const tz = Session.getScriptTimeZone();
    const hol = sh ? sh.getDataRange().getValues().slice(1)
      .map(r => r[0]).filter(d => d instanceof Date)
      .map(d => Utilities.formatDate(d, tz, 'yyyy-MM-dd')) : [];
    const isHoliday = (d) => hol.indexOf(Utilities.formatDate(d, tz, 'yyyy-MM-dd')) !== -1;

    const d = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    let remaining = Number(plannedDays);
    while (remaining > 0){
      d.setDate(d.getDate() + 1);
      const wd = d.getDay();            // 0=Sun, 5=Fri, 6=Sat
      const isWeekend = (wd === 5 || wd === 6);
      if (!isWeekend && !isHoliday(d)) remaining--;
    }
    return d;
  }catch(e){ Logger.log('[calcPlannedEndDate] '+e); return ''; }
}

/** علم الجدول الزمني (مبسّط) */
function computeScheduleFlag(startDate, plannedEndDate){
  try{
    if (!(startDate instanceof Date) || !(plannedEndDate instanceof Date)) return '';
    const today = new Date();
    if (today < startDate) return 'ON';
    if (today > plannedEndDate) return 'LATE';
    return 'ON';
  }catch(e){ Logger.log('[computeScheduleFlag] '+e); return ''; }
}

/** علم التكلفة (بسيط) */
function computeCostFlag(plannedMaterial, actualMaterial){
  try{
    const p = Number(plannedMaterial||0), a = Number(actualMaterial||0);
    if (!p && !a) return '';
    if (!p) return '';                  // لا نستطيع الحكم بدون مخطط
    return (a > p) ? 'OVERBUDGET' : 'OK';
  }catch(e){ Logger.log('[computeCostFlag] '+e); return ''; }
}

/** مجموع المواد/المصروفات المباشرة للمشروع (FIN_DirectExpenses.Amount) */
function rollupActualMaterials(projectId){
  try{
    const sh = getSheet(SHEETS.FIN_DIRECT_EXPENSES); if (!sh) return 0;
    const vals = sh.getDataRange().getValues(); if (!vals.length) return 0;
    const head = vals.shift();
    const H = h => head.indexOf(h);
    const iPrj = H('Project_ID'), iAmt = H('Amount');
    let sum = 0;
    vals.forEach(r => {
      if (String(r[iPrj]||'').trim().toUpperCase() === String(projectId||'').trim().toUpperCase()){
        const n = Number(String(r[iAmt]||'').toString().replace(/,/g,''));
        if (!isNaN(n)) sum += n;
      }
    });
    return sum;
  }catch(e){ Logger.log('[rollupActualMaterials] '+e); return 0; }
}

/** مجموع المُحصل من الإيرادات (FIN_Project_Revenue.Amount) */
function rollupTotalPayReceived(projectId){
  try{
    const sh = getSheet(SHEETS.FIN_PROJECT_REVENUE); if (!sh) return 0;
    const vals = sh.getDataRange().getValues(); if (!vals.length) return 0;
    const head = vals.shift();
    const H = h => head.indexOf(h);
    const iPrj = H('Project_ID'), iAmt = H('Amount');
    let sum = 0;
    vals.forEach(r => {
      if (String(r[iPrj]||'').trim().toUpperCase() === String(projectId||'').trim().toUpperCase()){
        const n = Number(String(r[iAmt]||'').toString().replace(/,/g,''));
        if (!isNaN(n)) sum += n;
      }
    });
    return sum;
  }catch(e){ Logger.log('[rollupTotalPayReceived] '+e); return 0; }
}

/** بعد الحفظ – حساب القيم المشتقة + مطابقة الـ Dropdowns ثم الكتابة */
function postProcessProject(projectId){
  try{
    if (!projectId) return;
    const sh = getSheet(SHEETS.PRJ_MAIN); if (!sh) return;

    const rng = sh.getDataRange(), vals = rng.getValues();
    if (!vals.length) return;
    const head = vals[0], rows = vals.slice(1);
    const H = h=> head.indexOf(h);

    const iID       = H('Project_ID');
    const iStart    = H('Start_Date');
    const iDays     = H('Planned_Days');
    const iPlnEnd   = H('Planned_End_Date');
    const iEnd      = H('End_Date');
    const iPlnMat   = H('Planned_Material_Expense');
    const iActMat   = H('Actual_Material_Expense');
    const iBudget   = H('Proj_Budget');
    const iPayRec   = H('Total_Pay_Received');
    const iPayPend  = H('Total_Pay_Pending');
    const iSch      = H('Schedule_Flag');
    const iCost     = H('Cost_Flag');
    const iStatus   = H('Status');
    const iStage    = H('Stage');
    const iUpdAt    = H('Updated_At');
    const iUpdBy    = H('Updated_By');

    const rIx = rows.findIndex(r => String(r[iID]||'').trim().toUpperCase() === String(projectId||'').trim().toUpperCase());
    if (rIx === -1) return;
    const row = rows[rIx].slice();                // نسخة قابلة للتعديل

    // 1) تاريخ نهاية مخطط
    const st = (iStart!==-1 && row[iStart] instanceof Date) ? row[iStart] : null;
    const pd = (iDays!==-1) ? Number(row[iDays]||0) : 0;
    const plnEnd = (st && pd>0) ? calcPlannedEndDate(st, pd) : '';

    // 2) Rollups & أعلام
    const actMat = rollupActualMaterials(projectId);
    const payRec = rollupTotalPayReceived(projectId);
    const budget = (iBudget!==-1) ? Number(String(row[iBudget]||'').toString().replace(/,/g,'')) : 0;
    const payPend = (budget ? (budget - payRec) : 0);

    const schFlag = (st && plnEnd) ? computeScheduleFlag(st, plnEnd) : '';
    const costFlag = computeCostFlag((iPlnMat!==-1 ? row[iPlnMat] : 0), actMat);
// 3) Normalization لقيم Status/Stage طبقًا لـ SYS_Dropdowns
if (iStatus !== -1) row[iStatus] = normalizeDropdownValue('PRJ_STATUS', row[iStatus]);
if (iStage  !== -1) row[iStage]  = normalizeDropdownValue('PROJECT_STAGE', row[iStage]);

    // 4) الكتابة إلى النسخة (row) قبل الإرسال
    if (iPlnEnd !== -1) row[iPlnEnd] = plnEnd || '';
    if (iActMat !== -1) row[iActMat] = actMat;
    if (iPayRec !== -1) row[iPayRec] = payRec;
    if (iPayPend!== -1) row[iPayPend]= payPend;
    if (iSch    !== -1) row[iSch]    = schFlag;
    if (iCost   !== -1) row[iCost]   = costFlag;

    // Audit
    if (iUpdAt !== -1) row[iUpdAt] = new Date();
    if (iUpdBy !== -1) row[iUpdBy] = 'SYSTEM';

    // 5) الكتابة للشيت
    sh.getRange(rIx + 2, 1, 1, head.length).setValues([row]);

  }catch(e){
    Logger.log('[postProcessProject] ' + e);
  }
}

/** فحص بصري: يعلّم أي قيم Status/Stage غير مسموح بها (لا يغيّر القيم) */
function validateProjectDropdowns(){
  const result = { checked: 0, invalid: 0 };
  try{
    const sh = getSheet(SHEETS.PRJ_MAIN); if (!sh) return result;
    const vals = sh.getDataRange().getValues(); if (!vals.length) return result;
    const head = vals[0], rows = vals.slice(1);
    const H = h => head.indexOf(h);
    const iStatus = H('Status'), iStage = H('Stage');
    const allowedStatus = new Set(getAllowedDropdownValues('PRJ_STATUS'));
    const allowedStage  = new Set(getAllowedDropdownValues('PROJECT_STAGE'));

    // إزالة تلوين سابق فقط في عمودي Status/Stage
    if (iStatus !== -1) sh.getRange(2, iStatus+1, sh.getLastRow()-1, 1).setBackground(null);
    if (iStage  !== -1) sh.getRange(2, iStage +1, sh.getLastRow()-1, 1).setBackground(null);

    rows.forEach((r, idx) => {
      if (iStatus !== -1){
        result.checked++;
        const ok = !r[iStatus] || allowedStatus.has(String(r[iStatus]).trim());
        if (!ok){ result.invalid++; sh.getRange(idx+2, iStatus+1).setBackground('#ffe6e6'); }
      }
      if (iStage !== -1){
        result.checked++;
        const ok = !r[iStage] || allowedStage.has(String(r[iStage]).trim());
        if (!ok){ result.invalid++; sh.getRange(idx+2, iStage+1).setBackground('#fff3cd'); }
      }
    });
  }catch(e){ Logger.log('[validateProjectDropdowns] ' + e); }
  return result;
}

/* ============================================================
   DROPDOWNS: helpers (generic normalization against SYS_Dropdowns)
   ============================================================ */

/** Arabic/Latin loose compare: trim + collapse spaces + uppercase */
function __normWord(v){
  return String(v||'')
    .replace(/[\u200f\u200e\u202a-\u202e]/g,'')
    .replace(/\s+/g,' ')
    .trim()
    .toUpperCase();
}

/** Read DF config once per call */
function __getDfConfigByForm(formId){
  const sh = getSheet(SHEETS.SYS_DYNAMIC_FORMS);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const head=vals[0]; const H=x=>head.indexOf(x);
  const iFID=H('Form_ID'), iFT=H('Field_Type'), iDK=H('Dropdown_Key'),
        iTS=H('Target_Sheet'), iTC=H('Target_Column'), iRID=H('Role_ID');
  const FID = String(formId||'').trim().toUpperCase();
  return vals.slice(1)
    .filter(r => String(r[iFID]||'').trim().toUpperCase() === FID)
    .map(r => ({
      type: String(r[iFT]||'').toLowerCase(),
      dropdownKey: String(r[iDK]||'').trim(),
      targetSheet: String(r[iTS]||'').trim(),
      targetColumn: String(r[iTC]||'').trim(),
      roles: String(r[iRID]||'ALL')
    }));
}

/** Normalize all dropdown fields in formData, based on DF (Field_Type=dropdown). */
function normalizeAllDropdownFields(formId, formData){
  const rows = __getDfConfigByForm(formId);
  rows.forEach(f=>{
    if (f.type === 'dropdown' && f.dropdownKey && f.targetColumn){
      if (formData.hasOwnProperty(f.targetColumn)){
        formData[f.targetColumn] = normalizeDropdownValue(f.dropdownKey, formData[f.targetColumn]);
      }
    }
  });
  return formData;
}

/* ============================================================
   FULL replacement: processFormSubmission (with normalization)
   ============================================================ */
function processFormSubmission(formId, formData, userId, roleId) {
    const diag = { formId, userId, roleId, plans: [] };
    try {
        formData = normalizeAllDropdownFields(formId, Object.assign({}, formData));

        const FID = String(formId || '').trim().toUpperCase();
        const ROLE = String(roleId || 'ALL').trim().toUpperCase();
        const WHO = String(userId || 'SYSTEM');

        const df = getSheet(SHEETS.SYS_DYNAMIC_FORMS);
        if (!df) return { success: false, message: 'SYS_Dynamic_Forms not found', diag };
        const values = df.getDataRange().getValues();
        if (!values.length) return { success: false, message: 'SYS_Dynamic_Forms empty', diag };

        const headers = values[0];
        const ixMap = _getHeadersMap(headers); // <-- REPLACED H HELPER
        const ix = {
            F_ID: ixMap['Form_ID'], FLD_ID: ixMap['Field_ID'], TYPE: ixMap['Field_Type'],
            TGT_S: ixMap['Target_Sheet'], TGT_C: ixMap['Target_Column'],
            DEF: ixMap['Default_Value'], R_ID: ixMap['Role_ID']
        };

        const targets = new Map();
        values.slice(1).forEach(row => {
            if (String(row[ix.F_ID] || '').trim().toUpperCase() !== FID) return;
            const rolesCell = String(row[ix.R_ID] || 'ALL').trim();
            const roles = rolesCell ? rolesCell.split(',').map(s => s.trim().toUpperCase()).filter(Boolean) : ['ALL'];
            const allowed = roles.includes('ALL') || ROLE === 'ALL' || roles.includes(ROLE);
            if (!allowed) return;
            const tgtSheet = String(row[ix.TGT_S] || '').trim();
            const tgtCol = String(row[ix.TGT_C] || '').trim();
            if (!tgtSheet || !tgtCol) return;
            if (!targets.has(tgtSheet)) targets.set(tgtSheet, { fields: [], pkCandidate: null });
            const bag = targets.get(tgtSheet);
            bag.fields.push({ col: tgtCol, type: String(row[ix.TYPE] || 'text').toLowerCase(), fieldId: String(row[ix.FLD_ID] || '').trim() });
            if (!bag.pkCandidate && /_ID$/i.test(tgtCol)) bag.pkCandidate = tgtCol;
        });

        if (!targets.size) {
            return { success: false, message: `No writable fields resolved for form ${FID} (role ${ROLE}).`, diag };
        }

        let totalWrites = 0, affectedRows = 0, newIds = [];
        for (const [sheetName, plan] of targets.entries()) {
            const sh = getSheet(sheetName);
            if (!sh) return { success: false, message: `Target sheet missing: ${sheetName}`, diag };
            const data = sh.getDataRange().getValues();
            const hdr = data.length ? data.shift() : [];
            const colMap = _getHeadersMap(hdr);
            let pkCol = plan.pkCandidate || hdr.find(h => /_ID$/i.test(h) && (formData[h] != null));
            const pkInHeaders = pkCol && (pkCol in colMap);
            let recordId = pkInHeaders ? String(formData[pkCol] || '').trim() : '';

            if (!recordId) {
                const autoField = plan.fields.find(f => (f.type === 'auto' || f.type === 'autogen') && f.col === pkCol);
                if (autoField) {
                    let prefix = '';
                    const dfRow = values.slice(1).find(r => String(r[ix.F_ID] || '').trim().toUpperCase() === FID && String(r[ix.TGT_S] || '').trim() === sheetName && String(r[ix.TGT_C] || '').trim() === pkCol);
                    if (dfRow) prefix = String(dfRow[ix.DEF] || '').trim().replace(/-+$/, '');
                    if (!prefix && pkCol === 'Employee_ID') prefix = 'EMP';
                    if (!prefix && pkCol === 'Project_ID') prefix = 'PRJ';
                    recordId = getNextAutoValue(sheetName, pkCol, prefix, 4);
                    if (pkInHeaders) formData[pkCol] = recordId;
                }
            }

            let rowIndex = -1;
            if (pkInHeaders && recordId) {
                rowIndex = data.findIndex(r => String(r[colMap[pkCol]] || '').trim().toUpperCase() === recordId.toUpperCase());
            }

            const row = (rowIndex === -1) ? new Array(hdr.length).fill('') : data[rowIndex].slice();
            let writesThisRow = 0;
            for (const f of plan.fields) {
                if (f.col in colMap) {
                    row[colMap[f.col]] = coerceValueForType(formData[f.col], f.type);
                    writesThisRow++;
                }
            }

            const now = new Date();
            if ('Created_At' in colMap && rowIndex === -1) row[colMap.Created_At] = now;
            if ('Created_By' in colMap && rowIndex === -1) row[colMap.Created_By] = WHO;
            if ('Updated_At' in colMap) row[colMap.Updated_At] = now;
            if ('Updated_By' in colMap) row[colMap.Updated_By] = WHO;

            if (rowIndex === -1 && row.every(v => v === '' || v === null)) {
                diag.plans.push({ sheet: sheetName, skipped: true, reason: 'all-empty row', pkCol, recordId });
                continue;
            }

            if (rowIndex === -1) {
                sh.appendRow(row);
            } else {
                sh.getRange(rowIndex + 2, 1, 1, hdr.length).setValues([row]);
            }
            affectedRows++;
            totalWrites += writesThisRow;
            if (recordId) newIds.push({ sheet: sheetName, pkCol, recordId });
            diag.plans.push({ sheet: sheetName, pkCol, recordId, rowIndex, writesThisRow });
        }

        if (FID === 'PRJ001' && newIds.length) {
            const prj = newIds.find(x => x.sheet === SHEETS.PRJ_MAIN) || newIds[0];
            if (prj && prj.recordId) postProcessProject(prj.recordId);
        }

        const msg = (affectedRows > 0) ? `Saved: ${affectedRows} row(s), ${totalWrites} field(s).` : 'Nothing to write (no mapped columns).';
        return { success: (affectedRows > 0), message: msg, ids: newIds, diag };

    } catch (e) {
        Logger.log(`[processFormSubmission ERROR] ${e}\n${e.stack}`);
        return { success: false, message: String(e), diag };
    }
}


/***** =========================================================
 * Direct Expenses "Cart" API (FULL)
 * ========================================================= *****/

/**
 * Save a "cart" of direct expenses as multiple rows in FIN_DirectExpenses.
 * items: Array of objects like:
 *  {
 *    materialId, name, category, sub1, sub2,
 *    unit, qty, unitPrice, totalPrice,        // choose unitPrice OR totalPrice (we'll compute the other)
 *    vendor, payStatus, payMethod, notes,     // optional
 *    date                                     // optional; default = today
 *  }
 * meta: { projectId, vatIncluded:boolean, vatRate:number (e.g. 0.14), createdBy:string }
 */
function saveDirectExpenseCart(items, meta){
  const res = { ok:false, wrote:0, totalExVAT:0, totalVAT:0, totalIncVAT:0, errors:[] };
  try{
    if (!meta || !meta.projectId) throw new Error('projectId is required in meta.');
    const projectId = String(meta.projectId).trim();
    const vatIncluded = !!meta.vatIncluded;
    const vatRate = (meta.vatRate != null ? Number(meta.vatRate) : 0);
    const who = meta.createdBy || 'SYSTEM';

    if (!Array.isArray(items) || !items.length) throw new Error('items list is empty.');

    const sh = getSheet(SHEETS.FIN_DIRECT_EXPENSES);
    if (!sh) throw new Error('FIN_DirectExpenses sheet missing.');

    const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const map = head.reduce((m,h,i)=>{ m[h]=i; return m; }, {});
// dropdown normalization targets (if present in SYS_Dropdowns)
const normUnit = v => normalizeDropdownValue('UNIT', v);
const normPayStatus = v => normalizeDropdownValue('PAY_STATUS', v);
    const normCategory = v => v; // if you keep categories free-text, leave as-is; or add a dropdown key.

    const rowsToAppend = [];
    items.forEach((it, idx) => {
      try{
        const qty = Number(it.qty||0);
        if (!qty || qty <= 0) throw new Error(`Item #${idx+1}: qty > 0 required`);

        const date = (it.date instanceof Date) ? it.date : new Date();

        // Decide prices
        let unitPrice = (it.unitPrice != null && it.unitPrice !== '') ? Number(it.unitPrice) : null;
        let totalPrice = (it.totalPrice != null && it.totalPrice !== '') ? Number(it.totalPrice) : null;

        if (unitPrice == null && totalPrice == null){
          // try default price from catalog
          const def = it.materialId ? getMaterialDefaultPrice(it.materialId) : null;
          if (def != null) unitPrice = Number(def);
        }
        if (unitPrice == null && totalPrice != null && qty){
          unitPrice = totalPrice / qty;
        }
        if (totalPrice == null && unitPrice != null){
          totalPrice = unitPrice * qty;
        }
        unitPrice = Number(unitPrice||0);
        totalPrice = Number(totalPrice||0);

        // VAT math
        let exVAT = 0, vatAmt = 0, incVAT = 0;
        if (vatRate > 0){
          if (vatIncluded){
            incVAT = totalPrice;
            exVAT = incVAT / (1 + vatRate);
            vatAmt = incVAT - exVAT;
          } else {
            exVAT = totalPrice;
            vatAmt = exVAT * vatRate;
            incVAT = exVAT + vatAmt;
          }
        } else {
          exVAT = totalPrice; vatAmt = 0; incVAT = totalPrice;
        }

        res.totalExVAT += exVAT;
        res.totalVAT   += vatAmt;
        res.totalIncVAT+= incVAT;

        const row = new Array(head.length).fill('');

        const set = (col, val) => { if (col in map) row[map[col]] = val; };

        set('Project_ID', projectId);
        set('Date', date);

        set('Category', normCategory(it.category||'مواد'));
        set('Level_1', it.sub1||'');
        set('Level_2', it.sub2||'');

        set('Material_ID', it.materialId||'');
        set('Material_Name', it.name || '');

        set('Unit', normUnit(it.unit||''));
        set('Qty', qty);
        set('Unit_Price', unitPrice);

        set('Amount', exVAT);                 // net amount
        set('VAT_Rate', vatRate);
        set('VAT_Amount', vatAmt);
        set('Total_With_VAT', incVAT);

        set('Vendor', it.vendor||'');
        set('Pay_Status', normPayStatus(it.payStatus||''));
        set('Pay_Method', it.payMethod||'');
        set('Notes', it.notes||'');

        set('Created_At', new Date());
        set('Created_By', who);
        set('Updated_At', new Date());
        set('Updated_By', who);

        rowsToAppend.push(row);
      }catch(e){
        res.errors.push(String(e));
      }
    });

    if (rowsToAppend.length){
      sh.getRange(sh.getLastRow()+1, 1, rowsToAppend.length, head.length).setValues(rowsToAppend);
      res.wrote = rowsToAppend.length;
      res.ok = true;
    }

    // Trigger project rollups after write
    try { postProcessProject(projectId); } catch(e){ Logger.log('[postProcessProject after cart] '+e); }

    return res;
  }catch(e){
    res.errors.push(String(e));
    return res;
  }
}

/**
 * DYNAMIC PROFILE VIEWS -- Get View Model
 * Core engine for all profile views. Reads SYS_Profile_View and returns a structured
 * "blueprint" of the tabs, sections, and fields to display.
 */
function getProfileViewModel(viewId, roleId){
  const sh = getSheet('SYS_Profile_View');
  if (!sh) return { viewId, tabs: [] };
  const values = sh.getDataRange().getValues();
  if (!values.length) return { viewId, tabs: [] };
  const headers = values.shift();
  const H = h => headers.indexOf(h);
  const ix = { V: H('View_ID'), T_ID: H('Tab_ID'), T_NM: H('Tab_Name'), S: H('Section_Header'), M: H('Mode'), SRC: H('Source_Sheet'), W: H('Where_Column'), F: H('Field_Column'), L: H('Field_Label'), FMT: H('Format'), SRT: H('Sort'), R: H('Role_ID') };
  const V = String(viewId||'').trim().toUpperCase();
  const ROLE = String(roleId||'ALL').trim().toUpperCase();
  const tabsMap = new Map();
  values.forEach(r => {
    if (String(r[ix.V]||'').trim().toUpperCase() !== V) return;
    const roles = String(r[ix.R]||'ALL').split(',').map(s=>s.trim().toUpperCase()).filter(Boolean);
    if (roles.length && !roles.includes('ALL') && ROLE !== 'ALL' && !roles.includes(ROLE)) return;
    const tabId = String(r[ix.T_ID]||'TAB').trim() || 'TAB';
    if (!tabsMap.has(tabId)) tabsMap.set(tabId, { id:tabId, name:String(r[ix.T_NM]||tabId), sections:new Map() });
    const tab = tabsMap.get(tabId);
    const secKey = String(r[ix.S]||'');
    if (!tab.sections.has(secKey)) tab.sections.set(secKey, []);
    tab.sections.get(secKey).push({ mode: String(r[ix.M]||'KV').toUpperCase(), sourceSheet: String(r[ix.SRC]||'').trim(), whereColumn: String(r[ix.W]||'').trim(), fieldColumn: String(r[ix.F]||'').trim(), fieldLabel: String(r[ix.L]||'').trim(), format: String(r[ix.FMT]||'').toLowerCase(), sort: Number(r[ix.SRT]||999) });
  });
  const tabs = Array.from(tabsMap.values()).map(t => {
    t.sections = Array.from(t.sections.entries()).map(([h, rows]) => ({ header: h, rows: rows.sort((a,b)=>a.sort-b.sort) }));
    return t;
  });
  return { viewId: V, tabs };
}


/**
 * DYNAMIC FORMS -- Get Form Model
 * Core function of the dynamic form engine. Reads SYS_Dynamic_Forms and builds
 * a JSON model of a form to be rendered by the client.
 */
function getDynamicFormModel(formId, userRoleId) {
  try {
    const sh = getSheet(SHEETS.SYS_DYNAMIC_FORMS);
    if (!sh) return { formId, tabs: [] };
    const values = sh.getDataRange().getValues();
    if (!values.length) return { formId, tabs: [] };
    const headers = values.shift();
    const H = h => headers.indexOf(h);
    const ix = { F_ID: H('Form_ID'), F_T: H('Form_Title'), T_ID: H('Tab_ID'), T_N: H('Tab_Name'), S_H: H('Section_Header'), FLD_ID: H('Field_ID'), LBL: H('Field_Label'), TYP: H('Field_Type'), SRC_S: H('Source_Sheet'), SRC_R: H('Source_Range'), M: H('Mandatory'), DDK: H('Dropdown_Key'), DEF: H('Default_Value'), TGT_S: H('Target_Sheet'), TGT_C: H('Target_Column'), R_ID: H('Role_ID') };
    const FID = String(formId||'').trim().toUpperCase();
    const ROLE = String(userRoleId||'ALL').trim().toUpperCase();
    let formTitle = '';
    const tabsMap = new Map();
    values.forEach(row => {
      if (String(row[ix.F_ID]||'').trim().toUpperCase() !== FID) return;
      const roles = String(row[ix.R_ID]||'ALL').split(',').map(s=>s.trim().toUpperCase()).filter(Boolean);
      if (roles.length && !roles.includes('ALL') && ROLE !== 'ALL' && !roles.includes(ROLE)) return;
      if (!formTitle) formTitle = String(row[ix.F_T]||'');
      const field = { fieldId: String(row[ix.FLD_ID]||''), label: String(row[ix.LBL]||''), type: String(row[ix.TYP]||'text').toLowerCase(), mandatory: ['Y','YES','TRUE','1'].includes(String(row[ix.M]||'').toUpperCase()), dropdownkey: String(row[ix.DDK]||''), defaultValue: row[ix.DEF], targetSheet: String(row[ix.TGT_S]||''), targetColumn: String(row[ix.TGT_C]||''), sourceSheet: String(row[ix.SRC_S]||''), sourceRange: String(row[ix.SRC_R]||'') };
      const ftype = field.type;
      if ((ftype === 'autogen' || ftype === 'auto') && field.targetSheet && field.targetColumn) {
        // [START] CORRECTED LOGIC
        // Use the prefix from the sheet, but provide a fallback based on the column name
        // to prevent issues with missing configuration.
        let prefix = (field.defaultValue || '').toString().trim().replace(/-+$/,'');
        if (!prefix) {
          if (field.targetColumn === 'Employee_ID') prefix = 'EMP';
          if (field.targetColumn === 'Project_ID') prefix = 'PRJ';
        }
        // [END] CORRECTED LOGIC
        field.defaultValue = getNextAutoValue(field.targetSheet, field.targetColumn, prefix, 4);
      } else if (ftype === 'autofill' && field.sourceSheet && field.sourceRange) {
        try { const src=getSheet(field.sourceSheet); if(src) field.defaultValue = src.getRange(field.sourceRange).getDisplayValue(); } catch(e){}
      }
      if (field.type === 'dropdown' && field.dropdownkey) field.options = getDropdownOptions(field.dropdownkey);
      const tabKey = String(row[ix.T_ID]||'TAB01');
      if (!tabsMap.has(tabKey)) tabsMap.set(tabKey, { id: tabKey, name: String(row[ix.T_N]||tabKey), sections: new Map() });
      const secKey = String(row[ix.S_H]||'');
      if (!tabsMap.get(tabKey).sections.has(secKey)) tabsMap.get(tabKey).sections.set(secKey, []);
      tabsMap.get(tabKey).sections.get(secKey).push(field);
    });
    const tabs = Array.from(tabsMap.values(), t => ({ id: t.id, name: t.name, sections: Array.from(t.sections.entries(), ([h, f]) => ({ header: h, fields: f })) }));
    return { formId, formTitle: formTitle||'Form', tabs };
  } catch (err) {
    Logger.log(`[getDynamicFormModel: ${formId}] ${err}`);
    return { formId, tabs: [] };
  }
}








/**
 * SUBMISSION -- Value Coercion Helper
 * Converts form string values into the correct data type (e.g., Date object).
 */
function coerceValueForType(v, type){
  const t = String(type||'text').toLowerCase();
  if (v === null || v === undefined) return '';
  const s = String(v).trim();

  if (t === 'date' || t === 'datetime' || t === 'datetime-local') {
    // dd/mm/yyyy [HH:mm]
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/);
    if (m) return new Date(m[3], m[2]-1, m[1], m[4]||0, m[5]||0);
    // ISO-ish pass-through
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return new Date(s);
    return s;
  }
  if (t === 'number' || t === 'currency') {
    const n = Number(s.replace(/,/g,''));
    return isNaN(n) ? s : n;
  }
  return s;
}



/**
 * SUBMISSION -- Get/Create Attachments Folder
 * A helper to get the designated Google Drive folder for attachments.
 */
function getDocsFolder_() {
  try {
    if (DOCS_FOLDER_ID) return DriveApp.getFolderById(DOCS_FOLDER_ID);
    const it = DriveApp.getFoldersByName('ERP_Attachments');
    return it.hasNext() ? it.next() : DriveApp.createFolder('ERP_Attachments');
  } catch (e) {
    Logger.log(`[getDocsFolder_] ${e}`);
    return DriveApp.getRootFolder();
  }
}


/**
 * SUBMISSION -- Attachment Uploader
 * Handles file uploads, saving to Drive and logging in SYS_Documents.
 */
function uploadFormAttachments(formId, entityId, files, userId) {
  try {
    if (!entityId || !files || !files.length) return {success:false, message:'Missing inputs.'};
    const folder = getDocsFolder_();
    const sh = getSheet(SHEETS.SYS_DOCUMENTS);
    if (!sh) return {success:false, message:'SYS_Documents not found.'};
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const results = files.map(f => {
      const blob = Utilities.newBlob(Utilities.base64Decode(f.data.split(',')[1]), f.mimeType, f.name);
      const gfile = folder.createFile(blob);
      const docId = nextDocId_();
      const row = { Doc_ID: docId, Entity: formId, Entity_ID: entityId, File_Name: f.name, Drive_URL: gfile.getUrl(), Uploaded_By: userId, Created_At: new Date() };
      sh.appendRow(headers.map(h => row[h]||''));
      return { docId, url: gfile.getUrl() };
    });
    return {success:true, message:'Attachments uploaded.', results};
  } catch (e) {
    Logger.log(`[uploadAttachments] ${e}`);
    return {success:false, message:'Upload failed: ' + e};
  }
}




/***** ==================== PUBLIC API Functions ==================== *****/


/**
 * PUBLIC API -- Get Dynamic Form Schema
 * Client-callable function to retrieve the structure of any dynamic form.
 */
function getHRFormSchema(formKey, userRoleId) { // Name kept for client compatibility
  return getDynamicFormModel(formKey, userRoleId);
}


/**
 * PUBLIC API -- Process Dynamic Form Submission
 * Client-callable function to submit data from any dynamic form.
 */
function processHRFormSubmission(formId, formData, userId, roleId){ // Name kept for client compatibility
  return processFormSubmission(formId, formData, userId, roleId);
}




/***** =========================================================================================
 *
 * HR MODULE FUNCTIONS
 *
 * ========================================================================================= *****/


/***** ==================== HR: Search ==================== *****/


/**
 * HR -- Search Employees
 * Searches employees by ID, name, email, mobile, department, or title.
 * Returns a list of matching employee records for the UI.
 */
function hrSearchEmployees(query, limit){
  const q = String(query || '').toLowerCase().trim();
  const LIM = Math.max(1, Math.min(+limit || 25, 100));
  const sh = getSheet(SHEETS.HR_EMPLOYEES);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (!values.length) return [];
  const headers = values.shift();
  const H = h => headers.indexOf(h);
  const ix = { ID: H('Employee_ID'), NameEN: H('Full_Name_EN'), NameAR: H('Full_Name_AR'), Email: H('Email'), Mob1: H('Mobile_Main'), Mob2: H('Mobile_Sub'), Dept: H('Department'), Title: H('Job_Title'), Status: H('Status') };
  const _norm_ = v => String(v || '').toLowerCase().trim();
  const mapRow = r => ({ Employee_ID: String(r[ix.ID]||''), Full_Name_EN: r[ix.NameEN], Full_Name_AR: r[ix.NameAR], Email: r[ix.Email], Mobile_Main: r[ix.Mob1], Department: r[ix.Dept], Job_Title: r[ix.Title], Status: r[ix.Status] });
  if (!q) return values.slice(-LIM).map(mapRow).reverse();
  const hits = [];
  for (const r of values) {
    const hay = [r[ix.ID], r[ix.NameEN], r[ix.NameAR], r[ix.Email], r[ix.Mob1], r[ix.Mob2], r[ix.Dept], r[ix.Title]].map(_norm_).join('|');
    if (hay.includes(q)) {
      hits.push(mapRow(r));
      if (hits.length >= LIM) break;
    }
  }
  return hits;
}




/***** ==================== HR: Profile View (Robust) ==================== *****/


/**
 * HR -- Get Employee Profile (Robust)
 * Fetches the main employee record and all related data (attendance, payroll, etc.)
 * from various sheets. Uses robust ID matching.
 */
function hrGetEmployeeProfile(employeeId) {
  const out = { ok: false, message: 'init', debug: {} };
  try {
    const rawId = String(employeeId || '').trim();
    if (!rawId) return { ok: false, message: 'Empty employeeId' };
    const empKey = _pv_norm_(rawId);
    const sh = getSheet(SHEETS.HR_EMPLOYEES);
    if (!sh) return { ok: false, message: 'HR_Employees sheet missing' };
    const vals = sh.getDataRange().getValues();
    if (!vals.length) return { ok: false, message: 'HR_Employees empty' };
    const headers = vals.shift();
    const idxEmp = headers.indexOf('Employee_ID');
    if (idxEmp === -1) return { ok: false, message: 'No Employee_ID column' };
    let personal = null;
    for (const r of vals) {
      if (_pv_norm_(r[idxEmp]) === empKey) {
        personal = {};
        headers.forEach((h, i) => { personal[h] = (r[i] instanceof Date) ? Utilities.formatDate(r[i], Session.getScriptTimeZone(), "yyyy-MM-dd") : r[i]; });
        break;
      }
    }
    if (!personal) return { ok: false, message: `No match for ${rawId}` };
    const getRelated = (s, k) => { try { const d = getSheet(s).getDataRange().getValues(), h = d.shift(), i=h.indexOf(k); return d.filter(r=>_pv_norm_(r[i])===empKey).map(r=>{const o={}; h.forEach((_h,_i)=>{o[_h]=(r[_i] instanceof Date)?Utilities.formatDate(r[_i],Session.getScriptTimeZone(),"yyyy-MM-dd HH:mm"):r[_i]});return o;});}catch(e){return[];}};
    return { ok: true, message: 'Profile loaded', personal, attendance: getRelated(SHEETS.HR_ATTENDANCE, 'Employee_ID'), payroll: getRelated(SHEETS.HR_PAYROLL, 'Employee_ID'), overtime: getRelated(SHEETS.HR_OVERTIME, 'Employee_ID'), advances: getRelated(SHEETS.HR_ADVANCES, 'Employee_ID'), deductions: getRelated(SHEETS.HR_DEDUCTIONS, 'Employee_ID'), attachments: getRelated(SHEETS.SYS_DOCUMENTS, 'Entity_ID'), debug: { rawId, empKey } };
  } catch (e) {
    return { ok: false, message: `Exception: ${e}`, debug: { stack: String(e.stack) } };
  }
}


/**
 * HR -- API Wrapper for Robust Profile View
 * This is the safe, client-callable function for the UI to get the employee profile.
 */
function apiHrGetEmployeeProfile(employeeId){
  try {
    const res = hrGetEmployeeProfile(employeeId);
    return res || { ok: false, message: 'Null response from hrGetEmployeeProfile' };
  } catch (e) {
    return { ok: false, message: `Exception: ${e}`, debug: { stack: String(e.stack) } };
  }
}




/***** ==================== HR: Profile View (Sheet-Driven) ==================== *****/


/**
 * HR -- Get Profile Data (Dynamic)
 * Core engine to build an employee's profile view, driven by the SYS_Profile_View sheet.
 */
function getEmployeeProfileDynamic(viewId, employeeId, roleId){
  const USE_CACHE = true, CACHE_TTL_SEC = 120, DEFAULT_TABLE_LIMIT = 200;
  const empKey = _pv_norm_(employeeId || '');
  if (!empKey) return { ok:false, message:'Empty employee ID.' };
  const cacheKey = `PV:${viewId}:${empKey}:${roleId||'ALL'}`;
  if (USE_CACHE) { try { const c = CacheService.getUserCache().get(cacheKey); if(c) return JSON.parse(c); } catch(e){} }
  const model = getProfileViewModel(viewId, roleId||'ALL');
  const memo = new Map();
  const readSheet = name => { if(memo.has(name)) return memo.get(name); const sh=getSheet(name), data=sh?sh.getDataRange().getValues():[], headers=data.length?data.shift():[]; const p={headers,rows:data}; memo.set(name,p); return p; };
  const dataTabs = model.tabs.map(tab => ({
    id: tab.id, name: tab.name, sections: (tab.sections||[]).map(sec => {
      const blocks = [];
      (sec.rows||[]).forEach(r => {
        const mode = r.mode.toUpperCase();
if (mode === 'ATTACH') { blocks.push({ type:'ATTACH', items: _getRelatedAttachments_(empKey, readSheet) }); return; }
        const pack = readSheet(r.sourceSheet);
        if (!pack.headers.length) return;
        if (mode === 'KV') {
const matches = _filterRowsByKey_(pack.headers, pack.rows, r.whereColumn, empKey);
          if (matches.length) { const idx=pack.headers.indexOf(r.fieldColumn); blocks.push({ type:'KV', label: r.fieldLabel||r.fieldColumn, value: _pv_fmt_(idx===-1?'':matches[0][idx], r.format) }); }
        } else if (mode === 'TABLE') {
          const matchesT = _filterRowsByKey_(pack.headers, pack.rows, r.whereColumn);
          if (!matchesT.length) return;
          const cols = r.fieldColumn.split(',').map(s=>s.trim()).filter(Boolean);
          const headers = cols.length ? cols : pack.headers;
          const rows = matchesT.slice(0, DEFAULT_TABLE_LIMIT).map(row => { const obj={}; headers.forEach(col=>{ const i=pack.headers.indexOf(col); obj[col]=_pv_fmt_(i===-1?'':row[i], r.format); }); return obj; });
          blocks.push({ type:'TABLE', headers, rows });
        }
      });
      return { header: sec.header, blocks };
    })
  }));
  const result = { ok:true, viewId: model.viewId, employeeId, tabs: dataTabs };
  if (USE_CACHE) { try { CacheService.getUserCache().put(cacheKey, JSON.stringify(result), CACHE_TTL_SEC); } catch(e){} }
  return result;
}


/**
 * HR -- API Wrapper for Dynamic Profile View
 * The safe, client-callable function for the UI to get the dynamic employee profile.
 */
function apiGetEmployeeProfileDynamic(viewId, employeeId, roleId){
  try{
    const res = getEmployeeProfileDynamic(viewId, employeeId, roleId||'ALL');
    return res || { ok:false, message:'Empty response' };
  }catch(e){
    return { ok:false, message:String(e) };
  }
}




/***** =========================================================================================
 *
 * PROJECTS MODULE FUNCTIONS
 *
 * ========================================================================================= *****/


/***** ==================== PROJECTS: Search ==================== *****/


/**
 * PROJECTS -- Search Projects
 * Searches projects by ID, name, client, or status from the PRJ_Main sheet.
 */
function prjSearchProjects(query, limit){
  const q = String(query || '').toLowerCase().trim();
  const LIM = Math.max(1, Math.min(+limit || 25, 100));
  const sh = getSheet(SHEETS.PRJ_MAIN);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (!values.length) return [];
  const headers = values.shift();
  const H = h => headers.indexOf(h);
  const ix = { ID: H('Project_ID'), Name: H('Project_Name'), Client: H('Client_Name'), Status: H('Status'), Start: H('Start_Date'), End: H('Planned_End_Date'), Budget: H('Proj_Budget') };
  const _norm_ = v => String(v || '').toLowerCase().trim();
  const fmtD = v => (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd") : v;
  const mapRow = r => ({ Project_ID: String(r[ix.ID]||''), Project_Name: r[ix.Name], Client_Name: r[ix.Client], Status: r[ix.Status], Start_Date: fmtD(r[ix.Start]), Planned_End_Date: fmtD(r[ix.End]), Proj_Budget: r[ix.Budget] });
  if (!q) return values.slice(-LIM).map(mapRow).reverse();
  const hits = [];
  for (const r of values) {
    const hay = [r[ix.ID], r[ix.Name], r[ix.Client], r[ix.Status]].map(_norm_).join('|');
    if (hay.includes(q)) {
      hits.push(mapRow(r));
      if (hits.length >= LIM) break;
    }
  }
  return hits;
}




/***** ==================== PROJECTS: Profile View (Sheet-Driven) ==================== *****/


/**
 * PROJECTS -- Get Profile Data (Dynamic)
 * Core engine to build a project's profile view, driven by the SYS_Profile_View sheet.
 */
function buildProjectProfileDynamic(viewId, projectId, roleId){
  const USE_CACHE = true, CACHE_TTL_SEC = 120, DEFAULT_TABLE_LIMIT = 200;


  const projKey = _pv_norm_(projectId || '');
  if (!projKey) return { ok:false, message:'Empty project ID.' };


  const cacheKey = `PV:${viewId}:${projKey}:${roleId||'ALL'}`;
  if (USE_CACHE) {
    try { const c = CacheService.getUserCache().get(cacheKey); if (c) return JSON.parse(c); } catch(e){}
  }


  const model = getProfileViewModel(viewId, roleId||'ALL');


  // memoized sheet reader
  const memo = new Map();
  const readSheet = (name) => {
    if (memo.has(name)) return memo.get(name);
    const sh = getSheet(name);
    const data = sh ? sh.getDataRange().getValues() : [];
    const headers = data.length ? data.shift() : [];
    const pack = { headers, rows: data };
    memo.set(name, pack);
    return pack;
  };


  const dataTabs = model.tabs.map(tab => {
    const sections = (tab.sections || []).map(sec => {
      const blocks = [];


      // 1) Collect all rows in this section by mode
      const rows = sec.rows || [];


      // ---- KV: group into a single block with items[] (this was the missing piece)
      const kvRows = rows.filter(r => String(r.mode||'').toUpperCase() === 'KV');
      if (kvRows.length) {
        // we only need the FIRST matching record from the source sheet, then map each KV row to an item
        // (All KV rows for this section should share the same source & where)
        const pack = readSheet(kvRows[0].sourceSheet || '');
        if (pack.headers.length) {
const matches = _filterRowsByKey_(pack.headers, pack.rows, kvRows[0].whereColumn, projKey);
          if (matches.length) {
            const rec = matches[0];
            const items = kvRows.map(r => {
              const idx = pack.headers.indexOf(r.fieldColumn);
              const value = (idx === -1) ? '' : _pv_fmt_(rec[idx], r.format);
              return { label: r.fieldLabel || r.fieldColumn, value, field: r.fieldColumn };
            });
            blocks.push({ type: 'KV', items });  // <-- now matches your viewer’s expectation
          }
        }
      }


      // ---- TABLE blocks
      rows.filter(r => String(r.mode||'').toUpperCase() === 'TABLE').forEach(r => {
        const pack = readSheet(r.sourceSheet);
        if (!pack.headers.length) return;
        const matches = _filterRowsByKey_(pack.headers, pack.rows, r.whereColumn, projKey);
        if (!matches.length) return;
        const cols = String(r.fieldColumn||'').split(',').map(s=>s.trim()).filter(Boolean);
        const headers = cols.length ? cols : pack.headers;
        const outRows = matches.slice(0, DEFAULT_TABLE_LIMIT).map(row => {
          const obj = {};
          headers.forEach(col => {
            const i = pack.headers.indexOf(col);
            obj[col] = _pv_fmt_(i===-1 ? '' : row[i], r.format);
          });
          return obj;
        });
        blocks.push({ type: 'TABLE', headers, rows: outRows });
      });


      // ---- ATTACH
     rows.filter(r => String(r.mode||'').toUpperCase() === 'ATTACH').forEach(() => blocks.push({ type:'ATTACH', items: _getRelatedAttachments_(projKey, readSheet) }));


      return { header: sec.header, blocks };
    });


    return { id: tab.id, name: tab.name, sections };
  });


  const result = { ok:true, viewId: model.viewId, projectId, tabs: dataTabs };
  if (USE_CACHE) {
    try { CacheService.getUserCache().put(cacheKey, JSON.stringify(result), CACHE_TTL_SEC); } catch(e){}
  }
  return result;
}

/**
 * PROJECTS -- API Wrapper for Profile View
 * The safe, client-callable function for the UI to get the project profile.
 */
function apiGetProjectProfileDynamic(viewId, projectId, roleId){
  try{
    const res = buildProjectProfileDynamic(viewId, projectId, roleId||'ALL');
    return res || { ok:false, message:'Empty response' };
  }catch(e){
    return { ok:false, message:String(e) };
  }
}

/***** ==================== PROJECTS: Core Calculations & Rollups ==================== *****/







/**
 * Sum revenues received = FIN_Project_Revenue by Project_ID (flexible columns).
 * Supported:
 *  - Amount_Received  (preferred)
 *  - Received_Amount
 *  - Amount where Status in ['Received','Paid','Collected'] (fallback)
 */
function rollupRevenuesReceived(projectId){
  try{
    const sh = getSheet(SHEETS.FIN_PROJECT_REVENUE); if (!sh) return 0;
    const vals = sh.getDataRange().getValues(); if (vals.length < 2) return 0;
    const head = vals.shift();
    const H = h => head.indexOf(h);

    const iPrj = H('Project_ID');
    if (iPrj === -1) return 0;

    const iAmtRecv = H('Amount_Received') !== -1 ? H('Amount_Received') : H('Received_Amount');
    const iAmt = H('Amount');
    const iStatus = H('Status');

    let sum = 0;
    vals.forEach(r=>{
      if (String(r[iPrj]||'').trim().toUpperCase() !== String(projectId).trim().toUpperCase()) return;

      if (iAmtRecv !== -1){
        const n = Number(String(r[iAmtRecv]||'').toString().replace(/,/g,''));
        if (!isNaN(n)) sum += n;
      } else if (iAmt !== -1){
        const status = String(r[iStatus]||'').trim().toUpperCase();
        const ok = ['RECEIVED','PAID','COLLECTED','مستلم','مدفوع'].includes(status);
        if (ok){
          const n = Number(String(r[iAmt]||'').toString().replace(/,/g,''));
          if (!isNaN(n)) sum += n;
        }
      }
    });
    return sum;
  }catch(e){
    Logger.log('[rollupRevenuesReceived] ' + e);
    return 0;
  }
}

/**
 * Post-process a saved project row in PRJ_Main:
 * - Planned_End_Date from Start_Date + Planned_Days (working days)
 * - Actual_Material_Expense rollup
 * - Total_Pay_Received rollup
 * - Total_Pay_Pending = Proj_Budget - Total_Pay_Received
 * - Schedule_Flag, Cost_Flag
 */
/** tiny helper: set a single cell if index exists */
function setCellIf_(sh, rowIndex0, headers, colName, value){
  const colIdx = headers.indexOf(colName);
  if (colIdx === -1) return false;
  // rowIndex0 is 0-based within "data rows" (i.e., excluding header)
  sh.getRange(rowIndex0 + 2, colIdx + 1, 1, 1).setValue(value);
  return true;
}

/** OPTIONAL: clear cached project profile so UI sees fresh values right away */
function invalidateProjectProfileCache_(viewId, projectId, roleId){
  try{
    const cacheKey = `PV:${String(viewId||'').toUpperCase()}:${_pv_norm_(projectId)}:${String(roleId||'ALL').toUpperCase()}`;
    CacheService.getUserCache().remove(cacheKey);
  }catch(e){}
}




/***** =========================================================================================
 *
 * DEVELOPER UTILITIES
 *
 * ========================================================================================= *****/


/**
 * UTILITY -- Generate Password Hash (for Testing)
 * A utility for developers to generate a password hash from the Apps Script editor.
 */
function generatePasswordHash() {
  const password = '210388'; // example only
  const salt = '';          // set if you use salts
  Logger.log(sha256b64_(password + salt));
}


/**
 * UTILITY -- Get Build Timestamp
 * A simple client-callable function to check if the server code is live.
 */
function pingBuild(){
  return { ok: true, buildAt: new Date().toISOString() };
}


/**
 * UTILITY -- Sanitize Project IDs
 * A one-time maintenance function to clean and standardize all 'Project_ID' values.
 */
function sanitizeProjectIDs() {
  try {
    const sh = getSheet(SHEETS.PRJ_MAIN);
    const dataRange = sh.getDataRange();
    const values = dataRange.getValues();
    const headers = values.shift();
    const colIndex = headers.indexOf('Project_ID');
    if (colIndex === -1) throw new Error('Project_ID column not found.');
    let count = 0;
    const cleaned = values.map(row => {
      const original = row[colIndex];
      const clean = _pv_norm_(original);
      if (original !== clean) count++;
      return [clean];
    });
    sh.getRange(2, colIndex + 1, cleaned.length, 1).setValues(cleaned);
    SpreadsheetApp.getUi().alert(`Sanitization complete. Cleaned ${count} Project IDs.`);
  } catch (e) {
    SpreadsheetApp.getUi().alert(`An error occurred: ${e}`);
  }
}


/**
 * UTILITY -- Log Sheet Schemas
 * A helper function to log the names and headers of multiple sheets.
 */
function logSchema(sheets, startNumber) {
  let counter = startNumber;
  let logOutput = '';
  const separator = '\n\n======================================================\n\n';
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const formattedHeaders = headers.join(' | ');
    logOutput += `-(` + counter + `)-{` + sheetName + `}--->[ ` + formattedHeaders + ` ]` + separator;
    counter++;
  });
  Logger.log(logOutput);
}


/**
 * UTILITY -- Log First Half of Schemas
 * Logs the schema for the first half of all sheets in the spreadsheet.
 */
function logFirstHalf() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const halfPoint = Math.ceil(sheets.length / 2);
  const firstHalf = sheets.slice(0, halfPoint);
  PropertiesService.getUserProperties().setProperty('logCounter', '1');
  logSchema(firstHalf, 1);
}


/**
 * UTILITY -- Log Second Half of Schemas
 * Logs the schema for the second half of all sheets in the spreadsheet.
 */
function logSecondHalf() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const startNumber = parseInt(PropertiesService.getUserProperties().getProperty('logCounter')) || 1;
  const halfPoint = Math.ceil(sheets.length / 2);
  const secondHalf = sheets.slice(halfPoint);
  logSchema(secondHalf, startNumber + halfPoint);
}

/**
 * UTILITY -- Generate Schema for Analyzer
 * Reads all sheets and their headers from the current spreadsheet and formats them
 * into a string that can be pasted directly into the Code Analyzer tool.
 */
function generateSchemaForAnalyzer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const schemaLines = [];

  sheets.forEach((sheet, i) => {
    const sheetName = sheet.getName();
    // Skip the validation report sheets themselves
    if (sheetName.startsWith('SYS_Schema_Validation_Report')) {
      return;
    }
    
    const lastColumn = sheet.getLastColumn();
    if (lastColumn > 0) {
      const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
      const headerString = headers.map(h => String(h || '').trim()).filter(Boolean).join(' | ');
      const line = `-(${i + 1})-{${sheetName}}--->[ ${headerString} ]`;
      schemaLines.push(line);
    } else {
      // Handle empty sheets
      const line = `-(${i + 1})-{${sheetName}}--->[ ]`;
      schemaLines.push(line);
    }
  });

  const fullSchemaString = schemaLines.join('\n');
  
  // Display the result in a dialog box with a textarea for easy copying
  const htmlOutput = HtmlService.createHtmlOutput(
      `<p>Copy the entire schema below and paste it into the analyzer tool:</p>
       <textarea style="width: 98%; height: 300px; font-family: monospace;">${fullSchemaString}</textarea>`
    )
    .setWidth(600)
    .setHeight(400);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generated ERP Schema');
}

function _smokeSaveProject(){
  const formId = 'PRJ001'; // your projects form key
  const payload = {
    Project_ID: '',                  // blank → autogen if configured
    Project_Name: 'Test From API',
    Client_Name: 'QA Client',
    Start_Date: '01/10/2025',
    Planned_End_Date: '15/11/2025',
    Status: 'نشط',
    Proj_Budget: '123456'
  };
  const res = processFormSubmission(formId, payload, 'DEV', 'ALL');
  Logger.log(JSON.stringify(res, null, 2));
}


/***** ==================== FINANCE: Direct Expenses (API) ==================== *****/

/** Utility: next DirectExp_ID like DEXP-00001 */
function nextDirectExpId_() {
  return getNextAutoValue(SHEETS.FIN_DIRECT_EXPENSES, 'DirectExp_ID', 'DEXP', 5);
}

/* ---------- Lite projects list (for pickers) ---------- */
function apiGetProjectsLite(limit){
  const sh = getSheet(SHEETS.PRJ_MAIN); if(!sh) return [];
  const vals = sh.getDataRange().getValues(); if (!vals.length) return [];
  const head = vals.shift(); const H = h => head.indexOf(h);
  const iID = H('Project_ID'), iName = H('Project_Name'), iClient = H('Client_Name'), iStat = H('Status');
  const out = vals.map(r => ({
    id: String(r[iID]||''),
    name: r[iName]||'',
    client: r[iClient]||'',
    status: r[iStat]||''
  })).filter(x => x.id);
  return (limit && +limit>0) ? out.slice(0, +limit) : out;
}

/** Catalog for Direct Expenses (matches your headers exactly)
 * Returns [{ id, name, unit, price, cat, sub1, sub2, vat:false }]
 * - Uses Name_AR for display name.
 * - Treats Active = "" or "Yes" as active. Only "No" hides the item.
 */
function apiGetMaterialsCatalog(query, limit){
  const sh = getSheet(SHEETS.PRJ_MATERIALS);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];

  const head = vals.shift();
  const H = h => head.indexOf(h);

  const ix = {
    id:   H('Material_ID'),
    nmAR: H('Name_AR'),
    // nmEN: H('Name_EN'), // keep available if you want later
    cat:  H('Category'),
    sub1: H('Subcategory'),
    sub2: H('Sub2'),
    unit: H('Default_Unit'),
    prc:  H('Default_Price'),
    act:  H('Active')
  };

  const norm = s => String(s||'').toLowerCase().trim();
  const q = norm(query);
  const LIM = Math.max(1, Math.min(+limit || 300, 1000));

  const rows = [];
  for (const r of vals){
    const activeRaw = (ix.act !== -1) ? String(r[ix.act]||'').trim() : '';
    const isInactive = activeRaw && activeRaw.toUpperCase() === 'NO';
    if (isInactive) continue; // blank or Yes -> include

    const o = {
      id:   String(ix.id   ===-1 ? '' : r[ix.id]  || ''),
      name: String(ix.nmAR ===-1 ? '' : r[ix.nmAR]|| ''),
      unit: String(ix.unit ===-1 ? '' : r[ix.unit]|| ''),
      price: Number(String(ix.prc===-1 ? '' : r[ix.prc]||'').toString().replace(/,/g,'')) || 0,
      cat:  String(ix.cat  ===-1 ? '' : r[ix.cat] || ''),
      sub1: String(ix.sub1 ===-1 ? '' : r[ix.sub1]|| ''),
      sub2: String(ix.sub2 ===-1 ? '' : r[ix.sub2]|| ''),
      vat: false  // you can extend later if you store per-item VAT
    };

    if (!q) { rows.push(o); if (rows.length>=LIM) break; continue; }

    const hay = (o.id+' '+o.name+' '+o.cat+' '+o.sub1+' '+o.sub2).toLowerCase();
    if (hay.includes(q)) { rows.push(o); if (rows.length>=LIM) break; }
  }

  return rows;
}

/**** =========================================================
 *  FINANCE: Direct Expenses — catalog, save-batch, search
 *  Sheets used:
 *    - SHEETS.PRJ_MATERIALS       (catalog)
 *    - SHEETS.FIN_DIRECT_EXPENSES (ledger)
 *    - SHEETS.PRJ_MAIN            (rollup/flags via postProcessProject)
 *  ========================================================= ****/

/** Back-compat alias so older front-end calls still work */
function getMaterialsCatalog(query, limit){
  return apiGetMaterialsCatalog(query, limit);
}

/** Catalog helper: price by Material_ID or Name */
function getMaterialDefaultPrice(materialKey){
  const key = String(materialKey||'').trim().toLowerCase();
  if (!key) return 0;

  const sh = getSheet(SHEETS.PRJ_MATERIALS);
  if (!sh) return 0;
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return 0;

  const head = vals.shift();
  const H = h => head.indexOf(h);

  const iId   = H('Material_ID');
  const iNmAR = H('Name_AR');
  const iPrc  = H('Default_Price');

  for (const r of vals){
    const id  = String(iId  ===-1 ? '' : r[iId]).trim().toLowerCase();
    const nm  = String(iNmAR===-1 ? '' : r[iNmAR]).trim().toLowerCase();
    if ((id && id === key) || (nm && nm === key)) {
      const raw = String(iPrc===-1 ? '' : r[iPrc]||'').toString().replace(/,/g,'');
      const n = Number(raw);
      return isNaN(n) ? 0 : n;
    }
  }
  return 0;
}

/**
 * Save cart lines
 * Supports BOTH signatures:
 *   1) apiFinSaveDirectExpensesBatch({lines:[...], meta:{...}})
 *   2) apiFinSaveDirectExpensesBatch(lines, meta)
 *
 * lines: [{materialId,name,unit,qty,unitPrice?,totalPrice?,category,sub1,sub2}]
 * meta:  { projectId, vendor, payStatus, payMethod,
 *          vatIncluded:boolean, vatRate:number, date?:iso, notes?:string, userId?:string }
 */
function apiFinSaveDirectExpensesBatch(payloadOrLines, maybeMeta){
  try{
    // ---- normalize arguments
    let lines = [];
    let meta  = {};
    if (Array.isArray(payloadOrLines)) {
      lines = payloadOrLines || [];
      meta  = maybeMeta || {};
    } else {
      const payload = payloadOrLines || {};
      lines = Array.isArray(payload.lines) ? payload.lines : [];
      meta  = payload.meta || {};
    }
    if (!lines.length) return { success:false, message:'No lines.' };

    // ---- open sheet & headers
    const sh = getSheet(SHEETS.FIN_DIRECT_EXPENSES);
    if (!sh) return { success:false, message:'FIN_DirectExpenses sheet missing.' };
    const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const H = h => head.indexOf(h);
    const ix = {
      id:   H('DirectExp_ID'),
      prj:  H('Project_ID'),
      date: H('Date'),
      cat:  H('Category'),
      s1:   H('Level_1'),
      s2:   H('Level_2'),
      matId:H('Material_ID'),
      name: H('Material_Name'),
      unit: H('Unit'),
      qty:  H('Qty'),
      uprice:H('Unit_Price'),
      amount:H('Amount'),
      vendor:H('Vendor'),
      paySt: H('Pay_Status'),
      payM:  H('Pay_Method'),
      vatInc:H('VAT_Included'),
      vatRt: H('VAT_Rate'),
      vatAmt:H('VAT_Amount'),
      totVat:H('Total_With_VAT'),
      notes: H('Notes'),
      cAt:   H('Created_At'),
      cBy:   H('Created_By'),
      uAt:   H('Updated_At'),
      uBy:   H('Updated_By'),
    };

    const today = meta.date ? new Date(meta.date) : new Date();
    const who   = String(meta.userId || 'SYSTEM');

    const rows = [];
    let count = 0;

    for (const ln of lines){
      const qty = Number(ln.qty||0);
      if (!qty) continue;

      // derive unit price
      let uPrice = (ln.unitPrice != null && ln.unitPrice !== '')
        ? Number(String(ln.unitPrice).toString().replace(/,/g,'')) : null;
      if (uPrice == null || isNaN(uPrice)) {
        uPrice = getMaterialDefaultPrice(ln.materialId || ln.name);
      }

      const amount = (ln.totalPrice != null && ln.totalPrice !== '')
        ? Number(String(ln.totalPrice).toString().replace(/,/g,''))
        : Number((uPrice||0) * qty);

      const vatIncluded = !!meta.vatIncluded;
      const vatRate     = Number(meta.vatRate||0);
      const vatAmount   = vatIncluded ? (amount - (amount/(1+vatRate))) : (amount * vatRate);
      const totalWithVat = vatIncluded ? amount : (amount + vatAmount);

      const row = new Array(head.length).fill('');
      if (ix.id    !== -1) row[ix.id]    = nextDirectExpId_();
      if (ix.prj   !== -1) row[ix.prj]   = String(meta.projectId||'').trim();
      if (ix.date  !== -1) row[ix.date]  = today;
      if (ix.cat   !== -1) row[ix.cat]   = ln.category || '';
      if (ix.s1    !== -1) row[ix.s1]    = ln.sub1 || '';
      if (ix.s2    !== -1) row[ix.s2]    = ln.sub2 || '';
      if (ix.matId !== -1) row[ix.matId] = ln.materialId || '';
      if (ix.name  !== -1) row[ix.name]  = ln.name || '';
      if (ix.unit  !== -1) row[ix.unit]  = ln.unit || '';
      if (ix.qty   !== -1) row[ix.qty]   = qty;
      if (ix.uprice!== -1) row[ix.uprice]= uPrice||0;
      if (ix.amount!== -1) row[ix.amount]= amount||0;
      if (ix.vendor!== -1) row[ix.vendor]= meta.vendor||'';
      if (ix.paySt !== -1) row[ix.paySt] = meta.payStatus||'';
      if (ix.payM  !== -1) row[ix.payM]  = meta.payMethod||'';
      if (ix.vatInc!== -1) row[ix.vatInc]= vatIncluded ? 'Yes' : 'No';
      if (ix.vatRt !== -1) row[ix.vatRt] = vatRate||0;
      if (ix.vatAmt!== -1) row[ix.vatAmt]= vatAmount||0;
      if (ix.totVat!== -1) row[ix.totVat]= totalWithVat||0;
      if (ix.notes !== -1) row[ix.notes] = meta.notes||'';
      if (ix.cAt   !== -1) row[ix.cAt]   = new Date();
      if (ix.cBy   !== -1) row[ix.cBy]   = who;
      if (ix.uAt   !== -1) row[ix.uAt]   = new Date();
      if (ix.uBy   !== -1) row[ix.uBy]   = who;

      rows.push(row);
      count++;
    }

    if (!rows.length) return { success:false, message:'Nothing to insert.' };

    sh.getRange(sh.getLastRow()+1, 1, rows.length, head.length).setValues(rows);

    // rollup → PRJ_Main and recompute flags
    try{
      const pid = String(meta.projectId||'').trim();
      if (pid) postProcessProject(pid);
    }catch(e){ Logger.log('[postProcessProject after DEXP] '+e); }

    return { success:true, message:`Saved ${count} expense(s).`, count };
  }catch(e){
    return { success:false, message:String(e) };
  }
}

/** Search Direct Expenses (by text + optional projectId) */
function finSearchDirectExpenses(query, limit, projectId){
  const sh = getSheet(SHEETS.FIN_DIRECT_EXPENSES); if (!sh) return [];
  const vals = sh.getDataRange().getValues(); if (!vals.length) return [];
  const head = vals.shift(); const H = h => head.indexOf(h);

  const ix = {
    id:  H('DirectExp_ID'),
    prj: H('Project_ID'),
    date:H('Date'),
    name:H('Material_Name'),
    amount:H('Amount'),
    vendor:H('Vendor'),
    cat:H('Category'),
    s1:H('Level_1'),
    s2:H('Level_2')
  };

  const q = String(query||'').trim().toLowerCase();
  const pid = String(projectId||'').trim().toLowerCase();
  const LIM = Math.max(1, Math.min(+limit||200, 1000));

  const out = [];
  for (const r of vals){
    const row = {
      DirectExp_ID: String(ix.id===-1?'':r[ix.id]||''),
      Project_ID:   String(ix.prj===-1?'':r[ix.prj]||''),
      Date:         (ix.date!==-1 && r[ix.date] instanceof Date)
                     ? Utilities.formatDate(r[ix.date],Session.getScriptTimeZone(),'yyyy-MM-dd')
                     : (r[ix.date]||''),
      Description:  String(ix.name===-1?'':r[ix.name]||''),
      Amount:       Number(ix.amount===-1?0:r[ix.amount]||0),
      Vendor:       String(ix.vendor===-1?'':r[ix.vendor]||''),
      Category:     String(ix.cat===-1?'':r[ix.cat]||''),
      Level_1:      String(ix.s1===-1?'':r[ix.s1]||''),
      Level_2:      String(ix.s2===-1?'':r[ix.s2]||'')
    };
    if (pid && row.Project_ID.toLowerCase() !== pid) continue;
    const hay = (row.DirectExp_ID + ' ' + row.Project_ID + ' ' + row.Description + ' ' + row.Vendor + ' ' + row.Category + ' ' + row.Level_1 + ' ' + row.Level_2).toLowerCase();
    if (!q || hay.includes(q)) {
      out.push(row);
      if (out.length >= LIM) break;
    }
  }
  return out;
}

/**
 * Return FIN_DirectExpenses rows filtered by Project_ID as objects keyed by header.
 * Add this to Code.gs (new function).
 */
function apiGetProjectDirectExpenses(projectId) {
  try {
    if (!projectId) return [];
    const sh = getSheet(SHEETS.FIN_DIRECT_EXPENSES);
    if (!sh) return [];
    const vals = sh.getDataRange().getValues();
    if (!vals || vals.length < 2) return [];
    const headers = vals[0].map(h => (h === null || h === undefined) ? '' : String(h).trim());
    const rows = vals.slice(1);
    const ix = _getHeadersMap(headers);
    const iProj = ix['Project_ID'];
    if (iProj === -1) return [];
    const out = rows
      .filter(r => String(r[iProj] || '').trim() === String(projectId).trim())
      .map(r => {
        const o = {};
        headers.forEach((h, j) => o[h] = r[j]);
        return o;
      });
    return out;
  } catch (e) {
    Logger.log('[apiGetProjectDirectExpenses] ' + e);
    return [];
  }
}

/* Server test:
   Run __test_apiGetProjectDirectExpenses() in Apps Script editor.
*/
function __test_apiGetProjectDirectExpenses() {
  try {
    const rows = apiGetProjectDirectExpenses('PRJ-0003');
    Logger.log(JSON.stringify(rows));
  } catch (e) { Logger.log('[__test_apiGetProjectDirectExpenses] ' + e); }
}

/**
 * Receive an array of files as data URLs, save to Drive (DOCS_FOLDER_ID or ERP_Attachments),
 * and append entries to SYS_Documents.
 * Expects payloadLines = [{ name, dataUrl }, ...], meta optional with Entity/Entity_ID/Uploaded_By.
 */
function apiUploadFinDexFilesBatch(payloadLines, meta) {
  try {
    if (!Array.isArray(payloadLines) || !payloadLines.length) return { success: false, message: 'No files' };
    meta = meta || {};
    // determine folder
    let folder;
    try {
      if (DOCS_FOLDER_ID && DOCS_FOLDER_ID.trim()) folder = DriveApp.getFolderById(DOCS_FOLDER_ID);
    } catch (e){}
    if (!folder) {
      const root = DriveApp.getRootFolder();
      const folders = DriveApp.getFoldersByName('ERP_Attachments');
      folder = folders.hasNext() ? folders.next() : root.createFolder('ERP_Attachments');
    }

    const sh = getSheet(SHEETS.SYS_DOCUMENTS);
    const results = [];
    const now = new Date();
    payloadLines.forEach(pl => {
      try {
        const fname = String(pl.name || 'file').replace(/[^\w.\- \u0600-\u06FF]/g,'').slice(0,200);
        const parts = String(pl.dataUrl || '').split(',');
        if (parts.length < 2) throw new Error('Invalid dataUrl');
        const base64 = parts.slice(1).join(',');
        const mimeMatch = parts[0].match(/data:([^;]+);/);
        const mime = mimeMatch ? mimeMatch[1] : 'application/octet-stream';
        const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime, fname);
        const file = folder.createFile(blob);
        const driveUrl = file.getUrl();
        const driveId = file.getId();
        const docId = (typeof nextDocId_ === 'function') ? nextDocId_() : getNextAutoValue(SHEETS.SYS_DOCUMENTS, 'Doc_ID', 'DOC', 5);
        if (sh) {
          const row = [docId, meta.Entity || '', meta.Entity_ID || '', fname, meta.Label || '', driveId, driveUrl, meta.Uploaded_By || '', now];
          sh.appendRow(row);
        }
        results.push({ Doc_ID: docId, Drive_File_ID: driveId, Drive_URL: driveUrl, File_Name: fname });
      } catch (inner) {
        Logger.log('[apiUploadFinDexFilesBatch:file] ' + inner);
      }
    });
    return { success: true, results: results };
  } catch (e) {
    Logger.log('[apiUploadFinDexFilesBatch] ' + e);
    return { success: false, message: String(e) };
  }
}

/* Server test:
   Run __test_apiUploadFinDexFilesBatch() in Apps Script editor.
*/
function __test_apiUploadFinDexFilesBatch() {
  try {
    const payload = [{
      name: 't.png',
      dataUrl: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGNgYAAAAAMAASsJTYQAAAAASUVORK5CYII='
    }];
    const meta = { Entity:'Project', Entity_ID:'PRJ-0003', Uploaded_By:'dev_test' };
    const res = apiUploadFinDexFilesBatch(payload, meta);
    Logger.log(JSON.stringify(res));
  } catch (e) { Logger.log('[__test_apiUploadFinDexFilesBatch] ' + e); }
}

/**
 * Server-side diagnostic: try to createHtmlOutputFromFile for many likely module names.
 * Run this from the Apps Script editor: select __scanProjectIncludes then Run -> View -> Logs
 */
function __scanProjectIncludes() {
  const candidates = [
    // Common module names from our conversation / blueprint
    'HR_Employees_JS','HR_Employees_UI',
    'Projects_Active_JS','Projects_Active_UI',
    'Finance_DirectExpenses_JS','Finance_DirectExpenses_UI','Finance_DirectExpenses_Modal_UI',
    'General_JS','General_UI','App_JS',
    'Dashboard_Home_JS','Dashboard_Home_UI',
    'HR_Employees_JS','HR_Employees',
    'Projects_Active','Projects_Active_JS',
    'HR_Employees_JS','HR_Employees_UI',
    // Some other likely names (based on earlier full blueprint)
    'HR_Employees_JS','HR_Employees_JS.html',
    'Projects_Active_JS','Projects_Active_UI',
    'Dashboard','Index','Login'
  ].filter((v,i,a)=>v && a.indexOf(v)===i);

  const results = { checked: [], found: [], missing: [], errors: [] };

  candidates.forEach(name => {
    try {
      // Attempt to create HTML output for the file. If it doesn't exist this will throw.
      const content = HtmlService.createHtmlOutputFromFile(name).getContent();
      results.checked.push({ name: name, ok: true, length: (content||'').length });
      results.found.push(name);
    } catch (e) {
      // Capture the exact error message for diagnosis
      results.checked.push({ name: name, ok: false, error: String(e) });
      // classify missing vs other errors
      if (String(e).indexOf('No HTML file named') !== -1) results.missing.push(name);
      else results.errors.push({ name: name, error: String(e) });
    }
  });

  // Also attempt to call getRawJs('Dashboard') but trap errors
  try {
    const raw = getRawJs('Dashboard');
    results.getRawJs = { success: true, length: (raw || '').length, preview: String(raw || '').slice(0,1200) };
  } catch (e) {
    results.getRawJs = { success: false, error: String(e) };
  }

  Logger.log(JSON.stringify(results, null, 2));
  return results;
}
/**
 * Server-side diagnostic: inspect Dashboard.js (getRawJs('Dashboard')) for presence
 * of expected function names. Run from Apps Script editor: select __scanModuleFunctionPresence then Run -> View -> Logs
 */
function __scanModuleFunctionPresence() {
  const fnNames = [
    // module loaders & critical UI functions
    'loadHREmployeesView','loadProjectsActiveView','loadFinanceDirectExpensesView',
    'selectSubTab','openFormForEdit','openForm','openNewForm','wireGlobalClosers',
    // dex/cart functions
    'renderLines','applyTotals','handleFileSelection','uploadAllAttachments','readFileAsDataURL',
    // dynamic form engine hooks (if present)
    'renderDynamicFormModel','getDynamicFormModel','getNextAutoValueServer',
    // misc
    'initDatePickers','loadModuleJsFromServer','injectModuleScript'
  ];

  const report = { checkedFile: 'Dashboard', len: 0, functions: {} };
  try {
    const raw = getRawJs('Dashboard') || '';
    report.len = raw.length;
    fnNames.forEach(fn => {
      // look for typical declarations: function name(, name = function, name:function, name: (for objects), or window.name
      const patterns = [
        new RegExp('\\bfunction\\s+' + fn + '\\s*\\(', 'm'),
        new RegExp('\\b' + fn + '\\s*=\\s*function\\s*\\(', 'm'),
        new RegExp('\\b' + fn + '\\s*:\\s*function\\s*\\(', 'm'),
        new RegExp('\\b' + fn + '\\s*=\\s*\\(', 'm'),
        new RegExp('window\\.' + fn + '\\s*=','m'),
        new RegExp('\\b' + fn + '\\s*\\=\\s*async','m'),
        new RegExp('\\b' + fn + '\\s*\\:\\s*\\(', 'm') // arrow shorthand
      ];
      const found = patterns.some(rx => rx.test(raw));
      report.functions[fn] = found;
    });
    Logger.log(JSON.stringify(report, null, 2));
    return report;
  } catch (e) {
    Logger.log('[__scanModuleFunctionPresence] ' + e);
    return { error: String(e) };
  }
}
/**
 * __findDupFunctionsInDashboard
 * Run from Apps Script editor. Returns counts of likely-duplicate function names in Dashboard.
 */
function __findDupFunctionsInDashboard(){
  const raw = (typeof getRawJs === 'function') ? (getRawJs('Dashboard') || '') : '';
  const names = [
    'setAddButton','clearAddButton','uniqSort','wireAdd','renderCards','renderSuggest',
    'runSearch','addBtn','renderBlocks','applyCols','addBtn','renderLines','renderDexGrid'
  ];
  const report = { fileLength: raw.length, counts: {}, snippets: {} };

  names.forEach(name=>{
    // patterns to match common function declaration/assignment styles
    const patterns = [
      new RegExp('\\bfunction\\s+' + name + '\\s*\\(', 'g'),
      new RegExp('\\b' + name + '\\s*=\\s*function\\s*\\(', 'g'),
      new RegExp('\\b' + name + '\\s*:\\s*function\\s*\\(', 'g'),
      new RegExp('\\b' + name + '\\s*=\\s*\\(', 'g'),         // arrow
      new RegExp('window\\.' + name + '\\s*=','g'),
      new RegExp('\\b' + name + '\\s*=\\s*async','g')
    ];
    let count = 0;
    patterns.forEach(rx => { try { const m = raw.match(rx); if (m) count += m.length; } catch(e){} });
    report.counts[name] = count;
    // also capture a small snippet of first occurrence (for context)
    try {
      const idx = raw.search(new RegExp('\\bfunction\\s+' + name + '\\s*\\(|' + name + '\\s*=\\s*function\\s*\\(|' + name + '\\s*=\\s*\\(', 'm'));
      if (idx >= 0) report.snippets[name] = raw.slice(Math.max(0, idx-120), Math.min(raw.length, idx+480));
      else report.snippets[name] = null;
    } catch(e){ report.snippets[name] = null; }
  });

  // count DOMContentLoaded handlers
  report.counts['DOMContentLoaded'] = (raw.match(/DOMContentLoaded/g) || []).length;

  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
