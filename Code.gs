/**
 * BUDGET SHEET — SHORTCUT BACKEND
 * ---------------------------------
 * Deploy this bound to your "my budget - Sean" spreadsheet [cite: 1]
 */

const CONFIG = {
  TRANSACTIONS_SHEET: 'Transactions',
  CATEGORIES_SHEET: 'Categories',
  HEADER_ROW: 9,       
  FIRST_DATA_ROW: 10,  
  LAST_FORMATTED_ROW: 2021, 
  // Pulls your secret from script properties (Project Settings > Script properties).
  SECRET: PropertiesService.getScriptProperties().getProperty('BUDGET_SECRET')
};

const DEFAULT_CATEGORIES = [
  'Groceries', 'Dining', 'Entertainment', 'Transportation',
  'Personal', 'Supplies', 'DOT Ticket', 'Clothes'
];

// ---------- ONE-TIME SETUP ----------

function setupCategories() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let catSheet = ss.getSheetByName(CONFIG.CATEGORIES_SHEET);

  if (!catSheet) {
    catSheet = ss.insertSheet(CONFIG.CATEGORIES_SHEET);
  }

  if (catSheet.getLastRow() === 0) {
    catSheet.getRange(1, 1).setValue('Category');
    catSheet.getRange(2, 1, DEFAULT_CATEGORIES.length, 1)
      .setValues(DEFAULT_CATEGORIES.map(c => [c]));
  }

  repointCategoryValidation_();
  Logger.log('Categories tab ready and Transactions!E validation repointed.');
}

function repointCategoryValidation_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);
  const catSheet = ss.getSheetByName(CONFIG.CATEGORIES_SHEET);

  const catRange = catSheet.getRange(2, 1, CONFIG.LAST_FORMATTED_ROW, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(catRange, true)
    .setAllowInvalid(false)
    .build();

  const targetRange = txSheet.getRange(
    CONFIG.FIRST_DATA_ROW, 5, 
    CONFIG.LAST_FORMATTED_ROW - CONFIG.FIRST_DATA_ROW + 1, 1
  );
  targetRange.setDataValidation(rule);
}

// ---------- WEB APP ENTRY POINTS ----------

function doGet(e) {
  const action = e.parameter.action;
  try {
    if (action === 'getCategories') {
      return jsonOut_({ ok: true, categories: getCategories_() });
    }
    return jsonOut_({ ok: false, error: 'Unknown or missing action for GET' });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
    
    if (!checkAuth_2(body)) {
      return jsonOut_({ ok: false, body: body, error: 'Unauthorized signature or expired request' });
    }

    switch (body.action) {
      case 'addTransaction':
        return jsonOut_(addTransaction_(body));
      case 'addCategory':
        return jsonOut_(addCategory_(body));
      case 'getCategories':
        return jsonOut_({ ok: true, categories: getCategories_() });
      default:
        return jsonOut_({ ok: false, error: 'Unknown action: ' + body.action });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message, body: body || null });
  }
}

/**
 * SHA-256 Auth Verification Strategy [cite: 30]
 */
function checkAuth_2(body) {
  const clientTimestamp = body.timestamp;
  const clientHash = body.key; // Maps correctly to body.key [cite: 31]

  if (!clientTimestamp || !clientHash) return false;
  
  // Enforce the 30-second freshness buffer (0.5 minutes) [cite: 31]
  if (!isTimestampFresh(timeStringClean_(clientTimestamp), 0.5)) return false;

  // Ordering requested: timestamp + secret 
  const rawString = clientTimestamp + CONFIG.SECRET; 
  
  const rawCipher = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, rawString, Utilities.Charset.UTF_8
  );

  let expectedHash = '';
  for (let i = 0; i < rawCipher.length; i++) {
    let byteStr = (rawCipher[i] & 0xFF).toString(16);
    if (byteStr.length === 1) byteStr = '0' + byteStr;
    expectedHash += byteStr;
  }

  // Prevent replay window exploits [cite: 34]
  if (isReplay_(expectedHash)) return false;

  return clientHash.toLowerCase() === expectedHash;
}

// ---------- CORE ACTIONS ----------

function addTransaction_(body) {
  const { date, name, amount, category } = body;
  if (!date || !name || amount === undefined || !category) {
    throw new Error('addTransaction requires date, name, amount, category');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TRANSACTIONS_SHEET);

  const nextRow = findNextEmptyRow_(sheet);

  const [y, m, d] = date.split('-').map(Number);
  const dateObj = new Date(y, m - 1, d);

  const amt = Number(amount);
  if (isNaN(amt)) throw new Error('amount must be numeric');

  sheet.getRange(nextRow, 1).setValue(dateObj);
  sheet.getRange(nextRow, 3).setValue(name);      
  sheet.getRange(nextRow, 4).setValue(amt);       
  sheet.getRange(nextRow, 5).setValue(category);

  extendFilter_(sheet, nextRow);

  return { ok: true, row: nextRow };
}

function findNextEmptyRow_(sheet) {
  const colA = sheet.getRange(
    CONFIG.FIRST_DATA_ROW, 1,
    sheet.getMaxRows() - CONFIG.FIRST_DATA_ROW + 1, 1
  ).getValues();

  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0] === '' || colA[i][0] === null) {
      return CONFIG.FIRST_DATA_ROW + i;
    }
  }
  throw new Error('No empty row found — sheet may need more rows added.');
}

function extendFilter_(sheet, newLastRow) {
  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    existingFilter.remove();
  }
  sheet.getRange(CONFIG.HEADER_ROW, 1, newLastRow - CONFIG.HEADER_ROW + 1, 5)
    .createFilter();
}

function getCategories_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catSheet = ss.getSheetByName(CONFIG.CATEGORIES_SHEET);
  if (!catSheet) throw new Error('Categories sheet not found — run setupCategories() first.');

  const values = catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 1).getValues();
  return values.map(r => r[0]).filter(v => v !== '' && v !== null);
}

function addCategory_(body) {
  const { category } = body;
  if (!category) throw new Error('addCategory requires a category name');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catSheet = ss.getSheetByName(CONFIG.CATEGORIES_SHEET);
  if (!catSheet) throw new Error('Categories sheet not found — run setupCategories() first.');

  const existing = getCategories_();
  if (existing.map(c => c.toLowerCase()).includes(category.toLowerCase())) {
    return { ok: true, message: 'Category already exists', categories: existing };
  }

  catSheet.getRange(catSheet.getLastRow() + 1, 1).setValue(category);
  repointCategoryValidation_(); 

  return { ok: true, categories: getCategories_() };
}

function isReplay_(hash) {
  const cache = CacheService.getScriptCache();
  if (cache.get(hash)) return true;
  cache.put(hash, '1', 50); // Keep cached for 50 seconds [cite: 55]
  return false;
}

// ---------- UTIL ----------

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function timeStringClean_(str) {
  return String(str).replace(/[^0-9]/g, '');
}

function isTimestampFresh(timeStr, maxMinutesDiff) {
  if (timeStr.length !== 12) return false;

  const year = parseInt(timeStr.substring(0, 4), 10);
  const month = parseInt(timeStr.substring(4, 6), 10) - 1;
  const day = parseInt(timeStr.substring(6, 8), 10);
  const hour = parseInt(timeStr.substring(8, 10), 10);
  const min = parseInt(timeStr.substring(10, 12), 10);
  
  const clientDate = new Date(year, month, day, hour, min);
  const serverDate = new Date();
  
  const timeDifferenceMs = Math.abs(serverDate - clientDate);
  const allowedDifferenceMs = maxMinutesDiff * 60 * 1000;
  
  return timeDifferenceMs <= allowedDifferenceMs;
}