/**
 * ============================================================================
 * NIRMAN FINANCE LTD — BACKEND SYSTEM (Google Apps Script)
 * ============================================================================
 * Professional Financial Management Platform
 * Database: Google Sheets
 * 
 * Sheets:
 * 1. Shareholders
 * 2. Deposits
 * 3. Withdrawals
 * 4. Borrowers
 * 5. Loans
 * 6. Loan_Installments
 * 7. Loan_Payments
 * 8. Transactions
 * 9. Audit_Logs
 * 10. Settings
 * ============================================================================
 */

// Configuration Constants
const CONFIG = {
  SPREADSHEET_ID: '', // Optional: Leave empty if script is created via Extensions > Apps Script inside your Google Sheet
  SHARE_UNIT_PRICE: 1000, // ৳ 1,000 per share
  DEFAULT_TIMEZONE: 'Asia/Dhaka',
  LOCK_TIMEOUT_MS: 20000,
  SHEET_NAMES: {
    SHAREHOLDERS: 'Shareholders',
    DEPOSITS: 'Deposits',
    WITHDRAWALS: 'Withdrawals',
    BORROWERS: 'Borrowers',
    LOANS: 'Loans',
    INSTALLMENTS: 'Loan_Installments',
    PAYMENTS: 'Loan_Payments',
    TRANSACTIONS: 'Transactions',
    AUDIT_LOGS: 'Audit_Logs',
    SETTINGS: 'Settings'
  }
};

/**
 * Web App entry point
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'ping') {
    return ContentService.createTextOutput(JSON.stringify({ success: true, message: 'Nirman Finance API online' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Nirman Finance LTD — Financial Management Platform')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * REST API Endpoint for external frontend hosting (e.g., Vercel, Netlify)
 */
function doPost(e) {
  let responseObj = { success: false, message: 'Invalid request' };
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }
    
    const action = payload.action;
    const data = payload.data;
    
    switch (action) {
      case 'ping':
        responseObj = { success: true, message: 'Nirman Finance API online', timestamp: new Date().toISOString() };
        break;
      case 'initializeDatabase':
        responseObj = initializeDatabase();
        break;
      case 'seedDemoData':
        responseObj = seedDemoData();
        break;
      case 'getDashboardData':
        responseObj = getDashboardData();
        break;
      case 'getShareholders':
        responseObj = getShareholders();
        break;
      case 'addShareholder':
        responseObj = addShareholder(data);
        break;
      case 'updateShareholder':
        responseObj = updateShareholder(data);
        break;
      case 'getShareholderDetails':
        responseObj = getShareholderDetails(data);
        break;
      case 'getDeposits':
        responseObj = getDeposits();
        break;
      case 'addDeposit':
        responseObj = addDeposit(data);
        break;
      case 'updateDeposit':
        responseObj = updateDeposit(data);
        break;
      case 'deleteDeposit':
        responseObj = deleteDeposit(data);
        break;
      case 'getWithdrawals':
        responseObj = getWithdrawals();
        break;
      case 'addWithdrawal':
        responseObj = addWithdrawal(data);
        break;
      case 'updateWithdrawal':
        responseObj = updateWithdrawal(data);
        break;
      case 'deleteWithdrawal':
        responseObj = deleteWithdrawal(data);
        break;
      case 'getBorrowers':
        responseObj = getBorrowers();
        break;
      case 'addBorrower':
        responseObj = addBorrower(data);
        break;
      case 'updateBorrower':
        responseObj = updateBorrower(data);
        break;
      case 'getLoans':
        responseObj = getLoans();
        break;
      case 'addLoan':
        responseObj = addLoan(data);
        break;
      case 'getLoanDetails':
        responseObj = getLoanDetails(data);
        break;
      case 'getInstallments':
        responseObj = getInstallments();
        break;
      case 'recordInstallmentPayment':
        responseObj = recordInstallmentPayment(data);
        break;
      case 'getTransactions':
        responseObj = getTransactions(data);
        break;
      case 'getReportData':
        responseObj = getReportData(data);
        break;
      case 'getSettings':
        responseObj = getSettings();
        break;
      case 'updateSettings':
        responseObj = updateSettings(data);
        break;
      default:
        responseObj = { success: false, message: 'Unknown action: ' + action };
        break;
    }
  } catch (err) {
    responseObj = { success: false, message: err.toString(), stack: err.stack };
  }
  
  return ContentService.createTextOutput(JSON.stringify(responseObj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * HTML include helper for modular files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Helper to get the active spreadsheet safely
 */
function getDB() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (err) {}
  if (CONFIG.SPREADSHEET_ID && String(CONFIG.SPREADSHEET_ID).trim() !== '') {
    return SpreadsheetApp.openById(String(CONFIG.SPREADSHEET_ID).trim());
  }
  throw new Error('Spreadsheet not bound. Please attach this script to a Google Sheet (Extensions > Apps Script) or specify CONFIG.SPREADSHEET_ID in Code.gs.');
}

/**
 * Get or create sheet by name with formatting
 */
function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#174879');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      headerRange.setFontFamily('Arial');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Initialize all database sheets with headers and formatting
 */
function initializeDatabase() {
  const ss = getDB();
  
  const schemas = [
    {
      name: CONFIG.SHEET_NAMES.SHAREHOLDERS,
      headers: [
        'Shareholder_ID', 'Full_Name', 'Phone', 'Email', 'Address', 
        'NID_Number', 'Joining_Date', 'Total_Deposited', 'Total_Withdrawn', 
        'Available_Balance', 'Number_Of_Shares', 'Total_Share_Value', 'Status', 'Notes', 'Created_At'
      ]
    },
    {
      name: CONFIG.SHEET_NAMES.DEPOSITS,
      headers: [
        'Deposit_ID', 'Shareholder_ID', 'Shareholder_Name', 'Amount', 
        'Deposit_Date', 'Reference_No', 'Notes', 'Created_At'
      ]
    },
    {
      name: CONFIG.SHEET_NAMES.WITHDRAWALS,
      headers: [
        'Withdrawal_ID', 'Shareholder_ID', 'Shareholder_Name', 'Amount', 
        'Withdrawal_Date', 'Reference_No', 'Notes', 'Created_At'
      ]
    },
    {
      name: CONFIG.SHEET_NAMES.BORROWERS,
      headers: [
        'Borrower_ID', 'Borrower_Type', 'Full_Name', 'Phone', 'Email', 
        'Address', 'NID_Or_Reg_No', 'Contact_Person', 'Registration_Date', 
        'Status', 'Notes', 'Created_At'
      ]
    },
    {
      name: CONFIG.SHEET_NAMES.LOANS,
      headers: [
        'Loan_ID', 'Borrower_ID', 'Borrower_Name', 'Borrower_Type', 
        'Loan_Amount', 'Loan_Date', 'Duration_Months', 'Monthly_Installment', 
        'First_Installment_Date', 'Total_Paid', 'Remaining_Amount', 
        'Loan_Status', 'Notes', 'Created_At'
      ]
    },
    {
      name: CONFIG.SHEET_NAMES.INSTALLMENTS,
      headers: [
        'Installment_ID', 'Loan_ID', 'Borrower_ID', 'Borrower_Name', 
        'Installment_Number', 'Due_Date', 'Due_Amount', 'Paid_Amount', 
        'Balance_Due', 'Status', 'Paid_Date', 'Payment_Reference'
      ]
    },
    {
      name: CONFIG.SHEET_NAMES.PAYMENTS,
      headers: [
        'Payment_ID', 'Loan_ID', 'Borrower_ID', 'Borrower_Name', 
        'Installment_ID', 'Amount_Paid', 'Payment_Date', 'Reference_No', 
        'Notes', 'Created_At'
      ]
    },
    {
      name: CONFIG.SHEET_NAMES.TRANSACTIONS,
      headers: [
        'Transaction_ID', 'Date', 'Type', 'Category', 'Entity_Type', 
        'Entity_ID', 'Entity_Name', 'Amount', 'Direction', 'Reference_No', 
        'Notes', 'Created_At'
      ]
    },
    {
      name: CONFIG.SHEET_NAMES.AUDIT_LOGS,
      headers: [
        'Log_ID', 'Timestamp', 'Action', 'Entity', 'Entity_ID', 
        'User', 'Details'
      ]
    },
    {
      name: CONFIG.SHEET_NAMES.SETTINGS,
      headers: ['Setting_Key', 'Setting_Value', 'Description']
    }
  ];

  schemas.forEach(s => {
    getOrCreateSheet(ss, s.name, s.headers);
  });

  // Initialize Default Settings if empty
  const settingsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  if (settingsSheet.getLastRow() <= 1) {
    const defaultSettings = [
      ['ORG_NAME', 'Nirman Finance LTD', 'Company official name'],
      ['ORG_PHONE', '+880 1712-345678', 'Official phone number'],
      ['ORG_EMAIL', 'info@nirmanfinance.com', 'Official email address'],
      ['ORG_ADDRESS', 'Level 7, Corporate Tower, Gulshan-2, Dhaka 1212', 'Registered office address'],
      ['CURRENCY_SYMBOL', '৳', 'Official currency symbol'],
      ['SHARE_PRICE', '1000', 'Fixed value per share in BDT'],
      ['LOAN_INTEREST_RATE', '0%', 'Fixed 0% interest loan policy'],
      ['SYSTEM_VERSION', '2.0.0', 'Platform build version']
    ];
    defaultSettings.forEach(row => settingsSheet.appendRow(row));
  }

  logAudit('INITIALIZE_DB', 'SYSTEM', 'SYS-001', 'Admin', 'Database schema initialized successfully');

  return {
    success: true,
    message: 'Database initialized successfully with all 10 standard tables and default configurations.'
  };
}

/**
 * Generate Next Unique Sequential ID (e.g. SH-1001, LN-1001)
 */
function getNextID(sheet, prefix, startNum = 1001) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return `${prefix}-${startNum}`;
  
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxId = startNum - 1;
  
  for (let i = 0; i < values.length; i++) {
    const val = String(values[i][0] || '').trim();
    if (val.startsWith(`${prefix}-`)) {
      const numPart = parseInt(val.replace(`${prefix}-`, ''), 10);
      if (!isNaN(numPart) && numPart > maxId) {
        maxId = numPart;
      }
    }
  }
  return `${prefix}-${maxId + 1}`;
}

/**
 * Audit Logger
 */
function logAudit(action, entity, entityId, user, details) {
  try {
    const ss = getDB();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.AUDIT_LOGS);
    if (!sheet) return;
    
    const logId = getNextID(sheet, 'LOG', 10001);
    const now = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([logId, now, action, entity, entityId, user || 'Admin', details || '']);
  } catch (e) {
    Logger.log('Audit log error: ' + e.message);
  }
}

/**
 * Record Central Transaction
 */
function recordTransaction(type, category, entityType, entityId, entityName, amount, direction, refNo, notes, dateStr) {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TRANSACTIONS);
  const txnId = getNextID(sheet, 'TXN', 5001);
  const now = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  const txnDate = dateStr || Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
  
  sheet.appendRow([
    txnId,
    txnDate,
    type,
    category,
    entityType,
    entityId,
    entityName,
    Number(amount) || 0,
    direction,
    refNo || '',
    notes || '',
    now
  ]);
  return txnId;
}

/**
 * Helper: Convert Sheet to Array of JSON Objects
 */
function sheetToJSON(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) return [];
  
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0].map(h => String(h).trim());
  const rows = [];
  
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      let val = row[c];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
      }
      obj[headers[c]] = val;
    }
    rows.push(obj);
  }
  return rows;
}

/**
 * ============================================================================
 * DASHBOARD API
 * ============================================================================
 */
function getDashboardData() {
  const ss = getDB();
  
  // Fetch required datasets
  const shareholders = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS));
  const deposits = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.DEPOSITS));
  const withdrawals = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.WITHDRAWALS));
  const loans = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.LOANS));
  const installments = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.INSTALLMENTS));
  const transactions = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.TRANSACTIONS));

  // Compute Metrics
  let totalShareholders = shareholders.length;
  let totalShares = 0;
  let totalDeposits = 0;
  let totalWithdrawals = 0;
  let totalLoanAmount = 0;
  let totalOutstandingLoanAmount = 0;
  let totalActiveLoans = 0;

  shareholders.forEach(s => {
    totalShares += Number(s.Number_Of_Shares) || 0;
    totalDeposits += Number(s.Total_Deposited) || 0;
    totalWithdrawals += Number(s.Total_Withdrawn) || 0;
  });

  const totalAvailableBalance = totalDeposits - totalWithdrawals;

  loans.forEach(l => {
    totalLoanAmount += Number(l.Loan_Amount) || 0;
    totalOutstandingLoanAmount += Number(l.Remaining_Amount) || 0;
    if (String(l.Loan_Status).toLowerCase() === 'active') {
      totalActiveLoans++;
    }
  });

  // Monthly trends (Last 6 Months)
  const monthlyData = getMonthlyTrends(deposits, withdrawals, loans);

  // Recent 6 Transactions
  const sortedTxns = transactions.slice(-10).reverse();

  // Upcoming & Overdue Installments
  const today = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
  const pendingInstallments = installments
    .filter(ins => ins.Status !== 'Paid')
    .sort((a, b) => new Date(a.Due_Date) - new Date(b.Due_Date))
    .slice(0, 8);

  return {
    success: true,
    data: {
      metrics: {
        totalShareholders,
        totalShares,
        totalDeposits,
        totalWithdrawals,
        totalAvailableBalance,
        totalActiveLoans,
        totalLoanAmount,
        totalOutstandingLoanAmount
      },
      charts: monthlyData,
      recentTransactions: sortedTxns,
      upcomingInstallments: pendingInstallments
    }
  };
}

/**
 * Monthly trend calculation for charts
 */
function getMonthlyTrends(deposits, withdrawals, loans) {
  const months = [];
  const d = new Date();
  for (let i = 5; i >= 0; i--) {
    const past = new Date(d.getFullYear(), d.getMonth() - i, 1);
    const key = Utilities.formatDate(past, CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM');
    const label = Utilities.formatDate(past, CONFIG.DEFAULT_TIMEZONE, 'MMM yyyy');
    months.push({ key, label, deposits: 0, withdrawals: 0, loanDisbursed: 0 });
  }

  deposits.forEach(dep => {
    const mKey = String(dep.Deposit_Date || '').substring(0, 7);
    const item = months.find(m => m.key === mKey);
    if (item) item.deposits += Number(dep.Amount) || 0;
  });

  withdrawals.forEach(wth => {
    const mKey = String(wth.Withdrawal_Date || '').substring(0, 7);
    const item = months.find(m => m.key === mKey);
    if (item) item.withdrawals += Number(wth.Amount) || 0;
  });

  loans.forEach(ln => {
    const mKey = String(ln.Loan_Date || '').substring(0, 7);
    const item = months.find(m => m.key === mKey);
    if (item) item.loanDisbursed += Number(ln.Loan_Amount) || 0;
  });

  return months;
}

/**
 * ============================================================================
 * SHAREHOLDER MANAGEMENT
 * ============================================================================
 */
function getShareholders() {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS);
  return { success: true, data: sheetToJSON(sheet) };
}

function addShareholder(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const ss = getDB();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS);
    
    // Check duplicate phone or NID only when provided
    const existing = sheetToJSON(sheet);
    if (data.NID_Number && String(data.NID_Number).trim() !== '') {
      const nidTrim = String(data.NID_Number).trim();
      if (existing.some(s => s.NID_Number && String(s.NID_Number).trim() === nidTrim)) {
        return { success: false, error: 'A shareholder with this NID already exists.' };
      }
    }
    if (data.Phone && String(data.Phone).trim() !== '') {
      const phoneTrim = String(data.Phone).trim();
      if (existing.some(s => s.Phone && String(s.Phone).trim() === phoneTrim)) {
        return { success: false, error: 'A shareholder with this Phone number already exists.' };
      }
    }

    const id = getNextID(sheet, 'SH', 1001);
    const initialShares = Number(data.Number_Of_Shares) || 0;
    const initialDeposited = initialShares * CONFIG.SHARE_UNIT_PRICE;
    const now = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const joinDate = data.Joining_Date || Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');

    sheet.appendRow([
      id,
      data.Full_Name || '',
      data.Phone || '',
      data.Email || '',
      data.Address || '',
      data.NID_Number || '',
      joinDate,
      initialDeposited, // Total_Deposited
      0,                // Total_Withdrawn
      initialDeposited, // Available_Balance
      initialShares,    // Number_Of_Shares
      initialDeposited, // Total_Share_Value
      data.Status || 'Active',
      data.Notes || '',
      now
    ]);

    // If initial shares were purchased with money, record initial deposit & transaction
    if (initialDeposited > 0) {
      const depSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEPOSITS);
      const depId = getNextID(depSheet, 'DEP', 2001);
      depSheet.appendRow([
        depId, id, data.Full_Name, initialDeposited, joinDate, 'INITIAL-SHARE', 'Initial share capital deposit', now
      ]);
      recordTransaction('DEPOSIT', 'Share Capital', 'Shareholder', id, data.Full_Name, initialDeposited, 'IN', 'INITIAL-SHARE', 'Initial share subscription', joinDate);
    }

    logAudit('CREATE_SHAREHOLDER', 'Shareholders', id, 'Admin', `Added shareholder ${data.Full_Name} with ${initialShares} shares`);

    return { success: true, message: 'Shareholder added successfully!', id: id };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function updateShareholder(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const ss = getDB();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: 'No shareholders found' };

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let rowIndex = -1;
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === data.Shareholder_ID) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'Shareholder not found' };
    }

    // Check duplicate phone or NID with other shareholders (excluding self) only when provided
    const existing = sheetToJSON(sheet);
    if (data.NID_Number && String(data.NID_Number).trim() !== '') {
      const nidTrim = String(data.NID_Number).trim();
      if (existing.some(s => s.Shareholder_ID !== data.Shareholder_ID && s.NID_Number && String(s.NID_Number).trim() === nidTrim)) {
        return { success: false, error: 'Another shareholder with this NID already exists.' };
      }
    }
    if (data.Phone && String(data.Phone).trim() !== '') {
      const phoneTrim = String(data.Phone).trim();
      if (existing.some(s => s.Shareholder_ID !== data.Shareholder_ID && s.Phone && String(s.Phone).trim() === phoneTrim)) {
        return { success: false, error: 'Another shareholder with this Phone number already exists.' };
      }
    }

    // Editable basic fields: Full_Name, Phone, Email, Address, NID_Number, Status, Notes
    sheet.getRange(rowIndex, 2).setValue(data.Full_Name || '');
    sheet.getRange(rowIndex, 3).setValue(data.Phone || '');
    sheet.getRange(rowIndex, 4).setValue(data.Email || '');
    sheet.getRange(rowIndex, 5).setValue(data.Address || '');
    sheet.getRange(rowIndex, 6).setValue(data.NID_Number || '');
    sheet.getRange(rowIndex, 13).setValue(data.Status || 'Active');
    sheet.getRange(rowIndex, 14).setValue(data.Notes || '');

    logAudit('UPDATE_SHAREHOLDER', 'Shareholders', data.Shareholder_ID, 'Admin', `Updated profile of ${data.Full_Name}`);

    return { success: true, message: 'Shareholder profile updated successfully!' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function getShareholderDetails(shareholderId) {
  const ss = getDB();
  const shareholders = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS));
  const sh = shareholders.find(s => s.Shareholder_ID === shareholderId);
  if (!sh) return { success: false, error: 'Shareholder not found' };

  const deposits = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.DEPOSITS))
    .filter(d => d.Shareholder_ID === shareholderId);
  const withdrawals = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.WITHDRAWALS))
    .filter(w => w.Shareholder_ID === shareholderId);

  return {
    success: true,
    data: {
      profile: sh,
      deposits: deposits.reverse(),
      withdrawals: withdrawals.reverse()
    }
  };
}

/**
 * ============================================================================
 * DEPOSIT MANAGEMENT
 * ============================================================================
 */
function getDeposits() {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEPOSITS);
  return { success: true, data: sheetToJSON(sheet).reverse() };
}

function addDeposit(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const amount = Number(data.Amount);
    if (!amount || amount <= 0) {
      return { success: false, error: 'Deposit amount must be a positive number.' };
    }

    const ss = getDB();
    const shSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS);
    const shList = sheetToJSON(shSheet);
    const shIndex = shList.findIndex(s => s.Shareholder_ID === data.Shareholder_ID);

    if (shIndex === -1) {
      return { success: false, error: 'Invalid Shareholder ID selected.' };
    }

    const sh = shList[shIndex];
    const rowNumber = shIndex + 2;

    const newTotalDeposited = (Number(sh.Total_Deposited) || 0) + amount;
    const currentWithdrawn = Number(sh.Total_Withdrawn) || 0;
    const newAvailableBalance = newTotalDeposited - currentWithdrawn;
    const newShares = Math.floor(newAvailableBalance / CONFIG.SHARE_UNIT_PRICE);
    const newShareValue = newShares * CONFIG.SHARE_UNIT_PRICE;

    // Update Shareholder balance & shares
    shSheet.getRange(rowNumber, 8).setValue(newTotalDeposited);
    shSheet.getRange(rowNumber, 10).setValue(newAvailableBalance);
    shSheet.getRange(rowNumber, 11).setValue(newShares);
    shSheet.getRange(rowNumber, 12).setValue(newShareValue);

    // Record Deposit
    const depSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEPOSITS);
    const depId = getNextID(depSheet, 'DEP', 2001);
    const now = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const depDate = data.Deposit_Date || Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');

    depSheet.appendRow([
      depId,
      sh.Shareholder_ID,
      sh.Full_Name,
      amount,
      depDate,
      data.Reference_No || '',
      data.Notes || '',
      now
    ]);

    // Central Transaction
    recordTransaction('DEPOSIT', 'Shareholder Deposit', 'Shareholder', sh.Shareholder_ID, sh.Full_Name, amount, 'IN', data.Reference_No, data.Notes, depDate);
    logAudit('CREATE_DEPOSIT', 'Deposits', depId, 'Admin', `Deposit of ৳ ${amount} recorded for ${sh.Full_Name}`);

    return {
      success: true,
      message: `Deposit of ৳ ${amount.toLocaleString()} recorded successfully! New shares: ${newShares}.`,
      depositId: depId
    };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function updateDeposit(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const depId = data.Deposit_ID;
    const newAmount = Number(data.Amount);
    if (!depId) {
      return { success: false, error: 'Deposit ID is required for editing.' };
    }
    if (!newAmount || newAmount <= 0) {
      return { success: false, error: 'Deposit amount must be a positive number.' };
    }

    const ss = getDB();
    const depSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEPOSITS);
    const depList = sheetToJSON(depSheet);
    const depIndex = depList.findIndex(d => d.Deposit_ID === depId);

    if (depIndex === -1) {
      return { success: false, error: 'Deposit record not found.' };
    }

    const oldDep = depList[depIndex];
    const oldAmount = Number(oldDep.Amount) || 0;
    const oldShId = oldDep.Shareholder_ID;
    const newShId = data.Shareholder_ID || oldShId;

    const shSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS);
    const shList = sheetToJSON(shSheet);
    const newShIndex = shList.findIndex(s => s.Shareholder_ID === newShId);
    if (newShIndex === -1) {
      return { success: false, error: 'Target Shareholder not found.' };
    }
    const newSh = shList[newShIndex];

    const depRowNumber = depIndex + 2;
    const newDate = data.Deposit_Date || oldDep.Deposit_Date;
    const newRef = (data.Reference_No !== undefined) ? data.Reference_No : oldDep.Reference_No;
    const newNotes = (data.Notes !== undefined) ? data.Notes : oldDep.Notes;

    depSheet.getRange(depRowNumber, 2).setValue(newSh.Shareholder_ID);
    depSheet.getRange(depRowNumber, 3).setValue(newSh.Full_Name);
    depSheet.getRange(depRowNumber, 4).setValue(newAmount);
    depSheet.getRange(depRowNumber, 5).setValue(newDate);
    depSheet.getRange(depRowNumber, 6).setValue(newRef);
    depSheet.getRange(depRowNumber, 7).setValue(newNotes);

    if (oldShId === newShId) {
      const diff = newAmount - oldAmount;
      const curDeposited = Number(newSh.Total_Deposited) || 0;
      const curWithdrawn = Number(newSh.Total_Withdrawn) || 0;
      const updatedDeposited = Math.max(0, curDeposited + diff);
      const updatedBalance = updatedDeposited - curWithdrawn;
      const updatedShares = Math.floor(updatedBalance / CONFIG.SHARE_UNIT_PRICE);
      const updatedShareVal = updatedShares * CONFIG.SHARE_UNIT_PRICE;

      const shRow = newShIndex + 2;
      shSheet.getRange(shRow, 8).setValue(updatedDeposited);
      shSheet.getRange(shRow, 10).setValue(updatedBalance);
      shSheet.getRange(shRow, 11).setValue(updatedShares);
      shSheet.getRange(shRow, 12).setValue(updatedShareVal);
    } else {
      const oldShIndex = shList.findIndex(s => s.Shareholder_ID === oldShId);
      if (oldShIndex !== -1) {
        const oldSh = shList[oldShIndex];
        const oldDepTotal = Math.max(0, (Number(oldSh.Total_Deposited) || 0) - oldAmount);
        const oldBal = oldDepTotal - (Number(oldSh.Total_Withdrawn) || 0);
        const oldShares = Math.floor(oldBal / CONFIG.SHARE_UNIT_PRICE);
        const oldRow = oldShIndex + 2;
        shSheet.getRange(oldRow, 8).setValue(oldDepTotal);
        shSheet.getRange(oldRow, 10).setValue(oldBal);
        shSheet.getRange(oldRow, 11).setValue(oldShares);
        shSheet.getRange(oldRow, 12).setValue(oldShares * CONFIG.SHARE_UNIT_PRICE);
      }
      const newDepTotal = (Number(newSh.Total_Deposited) || 0) + newAmount;
      const newBal = newDepTotal - (Number(newSh.Total_Withdrawn) || 0);
      const newShares = Math.floor(newBal / CONFIG.SHARE_UNIT_PRICE);
      const newRow = newShIndex + 2;
      shSheet.getRange(newRow, 8).setValue(newDepTotal);
      shSheet.getRange(newRow, 10).setValue(newBal);
      shSheet.getRange(newRow, 11).setValue(newShares);
      shSheet.getRange(newRow, 12).setValue(newShares * CONFIG.SHARE_UNIT_PRICE);
    }

    const txnSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TRANSACTIONS);
    const txnList = sheetToJSON(txnSheet);
    const txnIdx = txnList.findIndex(t => t.Type === 'DEPOSIT' && (t.Reference_No === oldDep.Reference_No || t.Reference_No === depId || (t.Entity_ID === oldShId && Math.abs(Number(t.Amount) - oldAmount) < 0.01)));
    if (txnIdx !== -1) {
      const tRow = txnIdx + 2;
      txnSheet.getRange(tRow, 2).setValue(newDate);
      txnSheet.getRange(tRow, 6).setValue(newSh.Shareholder_ID);
      txnSheet.getRange(tRow, 7).setValue(newSh.Full_Name);
      txnSheet.getRange(tRow, 8).setValue(newAmount);
      txnSheet.getRange(tRow, 10).setValue(newRef);
      txnSheet.getRange(tRow, 11).setValue(newNotes);
    }

    logAudit('UPDATE_DEPOSIT', 'Deposits', depId, 'Admin', `Deposit ${depId} updated: ৳ ${oldAmount} -> ৳ ${newAmount}`);
    return { success: true, message: `Deposit ${depId} updated successfully!` };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteDeposit(depositId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    if (!depositId) return { success: false, error: 'Deposit ID required.' };

    const ss = getDB();
    const depSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DEPOSITS);
    const depList = sheetToJSON(depSheet);
    const depIndex = depList.findIndex(d => d.Deposit_ID === depositId);

    if (depIndex === -1) return { success: false, error: 'Deposit record not found.' };

    const dep = depList[depIndex];
    const amount = Number(dep.Amount) || 0;
    const shId = dep.Shareholder_ID;

    // Delete deposit row
    depSheet.deleteRow(depIndex + 2);

    // Reverse shareholder balance & shares
    const shSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS);
    const shList = sheetToJSON(shSheet);
    const shIndex = shList.findIndex(s => s.Shareholder_ID === shId);
    if (shIndex !== -1) {
      const sh = shList[shIndex];
      const newTotalDeposited = Math.max(0, (Number(sh.Total_Deposited) || 0) - amount);
      const newBalance = newTotalDeposited - (Number(sh.Total_Withdrawn) || 0);
      const newShares = Math.floor(newBalance / CONFIG.SHARE_UNIT_PRICE);
      const shRow = shIndex + 2;
      shSheet.getRange(shRow, 8).setValue(newTotalDeposited);
      shSheet.getRange(shRow, 10).setValue(newBalance);
      shSheet.getRange(shRow, 11).setValue(newShares);
      shSheet.getRange(shRow, 12).setValue(newShares * CONFIG.SHARE_UNIT_PRICE);
    }

    // Delete matching transaction row
    const txnSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TRANSACTIONS);
    const txnList = sheetToJSON(txnSheet);
    const txnIdx = txnList.findIndex(t => t.Type === 'DEPOSIT' && (t.Reference_No === dep.Reference_No || t.Reference_No === depositId || (t.Entity_ID === shId && Math.abs(Number(t.Amount) - amount) < 0.01)));
    if (txnIdx !== -1) {
      txnSheet.deleteRow(txnIdx + 2);
    }

    logAudit('DELETE_DEPOSIT', 'Deposits', depositId, 'Admin', `Deposit ${depositId} of ৳ ${amount} deleted/reversed for ${dep.Shareholder_Name}`);
    return { success: true, message: `Deposit ${depositId} deleted. ৳ ${amount.toLocaleString()} reversed from balance.` };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * WITHDRAWAL MANAGEMENT
 * ============================================================================
 */
function getWithdrawals() {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.WITHDRAWALS);
  return { success: true, data: sheetToJSON(sheet).reverse() };
}

function addWithdrawal(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const amount = Number(data.Amount);
    if (!amount || amount <= 0) {
      return { success: false, error: 'Withdrawal amount must be a positive number.' };
    }

    const ss = getDB();
    const shSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS);
    const shList = sheetToJSON(shSheet);
    const shIndex = shList.findIndex(s => s.Shareholder_ID === data.Shareholder_ID);

    if (shIndex === -1) {
      return { success: false, error: 'Invalid Shareholder ID selected.' };
    }

    const sh = shList[shIndex];
    const rowNumber = shIndex + 2;

    const currentDeposited = Number(sh.Total_Deposited) || 0;
    const currentWithdrawn = Number(sh.Total_Withdrawn) || 0;
    const currentAvailable = currentDeposited - currentWithdrawn;

    // Strict validation: Never allow withdrawal exceeding available balance
    if (amount > currentAvailable) {
      return {
        success: false,
        error: `Insufficient balance! Shareholder ${sh.Full_Name} has only ৳ ${currentAvailable.toLocaleString()} available.`
      };
    }

    const newTotalWithdrawn = currentWithdrawn + amount;
    const newAvailableBalance = currentDeposited - newTotalWithdrawn;
    const newShares = Math.floor(newAvailableBalance / CONFIG.SHARE_UNIT_PRICE);
    const newShareValue = newShares * CONFIG.SHARE_UNIT_PRICE;

    // Update Shareholder
    shSheet.getRange(rowNumber, 9).setValue(newTotalWithdrawn);
    shSheet.getRange(rowNumber, 10).setValue(newAvailableBalance);
    shSheet.getRange(rowNumber, 11).setValue(newShares);
    shSheet.getRange(rowNumber, 12).setValue(newShareValue);

    // Record Withdrawal
    const wthSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.WITHDRAWALS);
    const wthId = getNextID(wthSheet, 'WTH', 3001);
    const now = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const wthDate = data.Withdrawal_Date || Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');

    wthSheet.appendRow([
      wthId,
      sh.Shareholder_ID,
      sh.Full_Name,
      amount,
      wthDate,
      data.Reference_No || '',
      data.Notes || '',
      now
    ]);

    // Central Transaction
    recordTransaction('WITHDRAWAL', 'Shareholder Withdrawal', 'Shareholder', sh.Shareholder_ID, sh.Full_Name, amount, 'OUT', data.Reference_No, data.Notes, wthDate);
    logAudit('CREATE_WITHDRAWAL', 'Withdrawals', wthId, 'Admin', `Withdrawal of ৳ ${amount} processed for ${sh.Full_Name}`);

    return {
      success: true,
      message: `Withdrawal of ৳ ${amount.toLocaleString()} processed successfully! Remaining shares: ${newShares}.`,
      withdrawalId: wthId
    };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function updateWithdrawal(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const wthId = data.Withdrawal_ID;
    const newAmount = Number(data.Amount);
    if (!wthId) {
      return { success: false, error: 'Withdrawal ID is required for editing.' };
    }
    if (!newAmount || newAmount <= 0) {
      return { success: false, error: 'Withdrawal amount must be a positive number.' };
    }

    const ss = getDB();
    const wthSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.WITHDRAWALS);
    const wthList = sheetToJSON(wthSheet);
    const wthIndex = wthList.findIndex(w => w.Withdrawal_ID === wthId);

    if (wthIndex === -1) {
      return { success: false, error: 'Withdrawal record not found.' };
    }

    const oldWth = wthList[wthIndex];
    const oldAmount = Number(oldWth.Amount) || 0;
    const oldShId = oldWth.Shareholder_ID;
    const newShId = data.Shareholder_ID || oldShId;

    const shSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS);
    const shList = sheetToJSON(shSheet);
    const newShIndex = shList.findIndex(s => s.Shareholder_ID === newShId);
    if (newShIndex === -1) {
      return { success: false, error: 'Target Shareholder not found.' };
    }
    const newSh = shList[newShIndex];

    const currentDeposited = Number(newSh.Total_Deposited) || 0;
    const currentWithdrawn = Number(newSh.Total_Withdrawn) || 0;
    const adjustedWithdrawn = (oldShId === newShId) ? (currentWithdrawn - oldAmount + newAmount) : (currentWithdrawn + newAmount);
    if (adjustedWithdrawn > currentDeposited) {
      return { success: false, error: `Withdrawal exceeds available balance (Available: ৳ ${(currentDeposited - (currentWithdrawn - (oldShId === newShId ? oldAmount : 0))).toLocaleString()})` };
    }

    const wthRowNumber = wthIndex + 2;
    const newDate = data.Withdrawal_Date || oldWth.Withdrawal_Date;
    const newRef = (data.Reference_No !== undefined) ? data.Reference_No : oldWth.Reference_No;
    const newNotes = (data.Notes !== undefined) ? data.Notes : oldWth.Notes;

    wthSheet.getRange(wthRowNumber, 2).setValue(newSh.Shareholder_ID);
    wthSheet.getRange(wthRowNumber, 3).setValue(newSh.Full_Name);
    wthSheet.getRange(wthRowNumber, 4).setValue(newAmount);
    wthSheet.getRange(wthRowNumber, 5).setValue(newDate);
    wthSheet.getRange(wthRowNumber, 6).setValue(newRef);
    wthSheet.getRange(wthRowNumber, 7).setValue(newNotes);

    if (oldShId === newShId) {
      const newBal = currentDeposited - adjustedWithdrawn;
      const newShares = Math.floor(newBal / CONFIG.SHARE_UNIT_PRICE);
      const shRow = newShIndex + 2;
      shSheet.getRange(shRow, 9).setValue(adjustedWithdrawn);
      shSheet.getRange(shRow, 10).setValue(newBal);
      shSheet.getRange(shRow, 11).setValue(newShares);
      shSheet.getRange(shRow, 12).setValue(newShares * CONFIG.SHARE_UNIT_PRICE);
    } else {
      const oldShIndex = shList.findIndex(s => s.Shareholder_ID === oldShId);
      if (oldShIndex !== -1) {
        const oldSh = shList[oldShIndex];
        const restoredWithdrawn = Math.max(0, (Number(oldSh.Total_Withdrawn) || 0) - oldAmount);
        const restoredBal = (Number(oldSh.Total_Deposited) || 0) - restoredWithdrawn;
        const restoredShares = Math.floor(restoredBal / CONFIG.SHARE_UNIT_PRICE);
        const oldRow = oldShIndex + 2;
        shSheet.getRange(oldRow, 9).setValue(restoredWithdrawn);
        shSheet.getRange(oldRow, 10).setValue(restoredBal);
        shSheet.getRange(oldRow, 11).setValue(restoredShares);
        shSheet.getRange(oldRow, 12).setValue(restoredShares * CONFIG.SHARE_UNIT_PRICE);
      }
      const newBal = currentDeposited - adjustedWithdrawn;
      const newShares = Math.floor(newBal / CONFIG.SHARE_UNIT_PRICE);
      const newRow = newShIndex + 2;
      shSheet.getRange(newRow, 9).setValue(adjustedWithdrawn);
      shSheet.getRange(newRow, 10).setValue(newBal);
      shSheet.getRange(newRow, 11).setValue(newShares);
      shSheet.getRange(newRow, 12).setValue(newShares * CONFIG.SHARE_UNIT_PRICE);
    }

    const txnSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TRANSACTIONS);
    const txnList = sheetToJSON(txnSheet);
    const txnIdx = txnList.findIndex(t => t.Type === 'WITHDRAWAL' && (t.Reference_No === oldWth.Reference_No || t.Reference_No === wthId || (t.Entity_ID === oldShId && Math.abs(Number(t.Amount) - oldAmount) < 0.01)));
    if (txnIdx !== -1) {
      const tRow = txnIdx + 2;
      txnSheet.getRange(tRow, 2).setValue(newDate);
      txnSheet.getRange(tRow, 6).setValue(newSh.Shareholder_ID);
      txnSheet.getRange(tRow, 7).setValue(newSh.Full_Name);
      txnSheet.getRange(tRow, 8).setValue(newAmount);
      txnSheet.getRange(tRow, 10).setValue(newRef);
      txnSheet.getRange(tRow, 11).setValue(newNotes);
    }

    logAudit('UPDATE_WITHDRAWAL', 'Withdrawals', wthId, 'Admin', `Withdrawal ${wthId} updated: ৳ ${oldAmount} -> ৳ ${newAmount}`);
    return { success: true, message: `Withdrawal ${wthId} updated successfully!` };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteWithdrawal(withdrawalId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    if (!withdrawalId) return { success: false, error: 'Withdrawal ID required.' };

    const ss = getDB();
    const wthSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.WITHDRAWALS);
    const wthList = sheetToJSON(wthSheet);
    const wthIndex = wthList.findIndex(w => w.Withdrawal_ID === withdrawalId);

    if (wthIndex === -1) return { success: false, error: 'Withdrawal record not found.' };

    const wth = wthList[wthIndex];
    const amount = Number(wth.Amount) || 0;
    const shId = wth.Shareholder_ID;

    // Delete withdrawal row
    wthSheet.deleteRow(wthIndex + 2);

    // Reverse shareholder balance & shares
    const shSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS);
    const shList = sheetToJSON(shSheet);
    const shIndex = shList.findIndex(s => s.Shareholder_ID === shId);
    if (shIndex !== -1) {
      const sh = shList[shIndex];
      const newTotalWithdrawn = Math.max(0, (Number(sh.Total_Withdrawn) || 0) - amount);
      const newBalance = (Number(sh.Total_Deposited) || 0) - newTotalWithdrawn;
      const newShares = Math.floor(newBalance / CONFIG.SHARE_UNIT_PRICE);
      const shRow = shIndex + 2;
      shSheet.getRange(shRow, 9).setValue(newTotalWithdrawn);
      shSheet.getRange(shRow, 10).setValue(newBalance);
      shSheet.getRange(shRow, 11).setValue(newShares);
      shSheet.getRange(shRow, 12).setValue(newShares * CONFIG.SHARE_UNIT_PRICE);
    }

    // Delete matching transaction row
    const txnSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TRANSACTIONS);
    const txnList = sheetToJSON(txnSheet);
    const txnIdx = txnList.findIndex(t => t.Type === 'WITHDRAWAL' && (t.Reference_No === wth.Reference_No || t.Reference_No === withdrawalId || (t.Entity_ID === shId && Math.abs(Number(t.Amount) - amount) < 0.01)));
    if (txnIdx !== -1) {
      txnSheet.deleteRow(txnIdx + 2);
    }

    logAudit('DELETE_WITHDRAWAL', 'Withdrawals', withdrawalId, 'Admin', `Withdrawal ${withdrawalId} of ৳ ${amount} deleted/reversed for ${wth.Shareholder_Name}`);
    return { success: true, message: `Withdrawal ${withdrawalId} deleted. ৳ ${amount.toLocaleString()} restored to balance.` };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * BORROWER MANAGEMENT
 * ============================================================================
 */
function getBorrowers() {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.BORROWERS);
  return { success: true, data: sheetToJSON(sheet) };
}

function addBorrower(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const ss = getDB();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.BORROWERS);

    if (!data.Full_Name || !data.Borrower_Type) {
      return { success: false, error: 'Borrower Name and Borrower Type are required.' };
    }

    const id = getNextID(sheet, 'BR', 1001);
    const now = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const regDate = data.Registration_Date || Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');

    sheet.appendRow([
      id,
      data.Borrower_Type,
      data.Full_Name,
      data.Phone || '',
      data.Email || '',
      data.Address || '',
      data.NID_Or_Reg_No || '',
      data.Contact_Person || '',
      regDate,
      data.Status || 'Active',
      data.Notes || '',
      now
    ]);

    logAudit('CREATE_BORROWER', 'Borrowers', id, 'Admin', `Added ${data.Borrower_Type}: ${data.Full_Name}`);

    return { success: true, message: 'Borrower registered successfully!', id: id };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function updateBorrower(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const ss = getDB();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.BORROWERS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: 'No borrowers found' };

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let rowIndex = -1;
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === data.Borrower_ID) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, error: 'Borrower not found' };
    }

    sheet.getRange(rowIndex, 2).setValue(data.Borrower_Type);
    sheet.getRange(rowIndex, 3).setValue(data.Full_Name || '');
    sheet.getRange(rowIndex, 4).setValue(data.Phone || '');
    sheet.getRange(rowIndex, 5).setValue(data.Email || '');
    sheet.getRange(rowIndex, 6).setValue(data.Address || '');
    sheet.getRange(rowIndex, 7).setValue(data.NID_Or_Reg_No || '');
    sheet.getRange(rowIndex, 8).setValue(data.Contact_Person || '');
    sheet.getRange(rowIndex, 10).setValue(data.Status || 'Active');
    sheet.getRange(rowIndex, 11).setValue(data.Notes || '');

    logAudit('UPDATE_BORROWER', 'Borrowers', data.Borrower_ID, 'Admin', `Updated details for ${data.Full_Name}`);

    return { success: true, message: 'Borrower details updated successfully!' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * LOAN MANAGEMENT (0% Interest, Exact Penny-Adjusted Installments)
 * ============================================================================
 */
function getLoans() {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOANS);
  return { success: true, data: sheetToJSON(sheet).reverse() };
}

function addLoan(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const amount = Number(data.Loan_Amount);
    const duration = parseInt(data.Duration_Months, 10);

    if (!amount || amount <= 0) {
      return { success: false, error: 'Loan amount must be greater than zero.' };
    }
    if (!duration || duration <= 0) {
      return { success: false, error: 'Loan duration must be at least 1 month.' };
    }
    if (!data.Borrower_ID || !data.Borrower_Name) {
      return { success: false, error: 'Please select a valid borrower.' };
    }

    const ss = getDB();
    const loanSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOANS);
    const instSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INSTALLMENTS);

    const loanId = getNextID(loanSheet, 'LN', 4001);
    const now = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const loanDate = data.Loan_Date || Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
    const firstInstDateStr = data.First_Installment_Date || loanDate;

    // Calculate Exact Installments
    // Formula: Regular = Math.floor(Total / Months), Final = Total - sum(Regular * (Months-1))
    const regularInstallment = Math.floor(amount / duration);
    const lastInstallment = amount - (regularInstallment * (duration - 1));

    loanSheet.appendRow([
      loanId,
      data.Borrower_ID,
      data.Borrower_Name,
      data.Borrower_Type || 'Outside Individual',
      amount,
      loanDate,
      duration,
      regularInstallment,
      firstInstDateStr,
      0,                  // Total_Paid
      amount,             // Remaining_Amount
      'Active',           // Loan_Status
      data.Notes || '',
      now
    ]);

    // Generate Installment Schedule
    const startDate = new Date(firstInstDateStr);
    for (let i = 1; i <= duration; i++) {
      const instId = `${loanId}-INS-${String(i).padStart(2, '0')}`;
      const dueDate = new Date(startDate.getFullYear(), startDate.getMonth() + (i - 1), startDate.getDate());
      const dueDateFormatted = Utilities.formatDate(dueDate, CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
      const dueAmount = (i === duration) ? lastInstallment : regularInstallment;

      instSheet.appendRow([
        instId,
        loanId,
        data.Borrower_ID,
        data.Borrower_Name,
        i,
        dueDateFormatted,
        dueAmount,
        0,                // Paid_Amount
        dueAmount,        // Balance_Due
        'Pending',        // Status
        '',               // Paid_Date
        ''                // Payment_Reference
      ]);
    }

    // Record Central Transaction for Loan Disbursement
    recordTransaction('LOAN_DISBURSEMENT', '0% Interest Loan', data.Borrower_Type || 'Borrower', data.Borrower_ID, data.Borrower_Name, amount, 'OUT', loanId, `Disbursement of 0% interest loan for ${duration} months`, loanDate);
    logAudit('CREATE_LOAN', 'Loans', loanId, 'Admin', `Disbursed ৳ ${amount} loan to ${data.Borrower_Name} for ${duration} months`);

    return {
      success: true,
      message: `Loan ${loanId} created successfully for ৳ ${amount.toLocaleString()} with ${duration} monthly installments.`,
      loanId: loanId
    };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

function getLoanDetails(loanId) {
  const ss = getDB();
  const loans = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.LOANS));
  const loan = loans.find(l => l.Loan_ID === loanId);
  if (!loan) return { success: false, error: 'Loan record not found' };

  const installments = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.INSTALLMENTS))
    .filter(ins => ins.Loan_ID === loanId)
    .sort((a, b) => Number(a.Installment_Number) - Number(b.Installment_Number));

  const payments = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.PAYMENTS))
    .filter(p => p.Loan_ID === loanId)
    .reverse();

  return {
    success: true,
    data: {
      loan,
      installments,
      payments
    }
  };
}

/**
 * ============================================================================
 * INSTALLMENT & PAYMENT MANAGEMENT
 * ============================================================================
 */
function getInstallments() {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INSTALLMENTS);
  const rows = sheetToJSON(sheet);

  // Check and dynamically tag overdue installments
  const today = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
  rows.forEach(row => {
    if (row.Status !== 'Paid' && row.Due_Date < today) {
      row.Is_Overdue = true;
      if (row.Status === 'Pending') {
        row.Status = 'Overdue';
      }
    } else {
      row.Is_Overdue = false;
    }
  });

  return { success: true, data: rows };
}

function recordInstallmentPayment(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const amountPaid = Number(data.Amount_Paid);
    if (!amountPaid || amountPaid <= 0) {
      return { success: false, error: 'Payment amount must be greater than zero.' };
    }

    const ss = getDB();
    const instSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INSTALLMENTS);
    const loanSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOANS);
    const paySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PAYMENTS);

    const instList = sheetToJSON(instSheet);
    const targetInstIndex = instList.findIndex(i => i.Installment_ID === data.Installment_ID);
    if (targetInstIndex === -1) {
      return { success: false, error: 'Installment not found' };
    }

    const targetInst = instList[targetInstIndex];
    const instRow = targetInstIndex + 2;
    const loanId = targetInst.Loan_ID;

    // Find loan row
    const loanList = sheetToJSON(loanSheet);
    const loanIndex = loanList.findIndex(l => l.Loan_ID === loanId);
    if (loanIndex === -1) {
      return { success: false, error: 'Associated loan record not found.' };
    }
    const loan = loanList[loanIndex];
    const loanRow = loanIndex + 2;

    const paymentDate = data.Payment_Date || Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd');
    const now = Utilities.formatDate(new Date(), CONFIG.DEFAULT_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    // Calculate installment balance
    const currentInstPaid = Number(targetInst.Paid_Amount) || 0;
    const dueAmount = Number(targetInst.Due_Amount) || 0;
    const newInstPaid = currentInstPaid + amountPaid;
    const newInstBalance = Math.max(0, dueAmount - newInstPaid);

    let newStatus = 'Partially Paid';
    if (newInstBalance === 0) {
      newStatus = 'Paid';
    }

    // Update Installment Row: Paid_Amount (col 8), Balance_Due (col 9), Status (col 10), Paid_Date (col 11), Payment_Reference (col 12)
    instSheet.getRange(instRow, 8).setValue(newInstPaid);
    instSheet.getRange(instRow, 9).setValue(newInstBalance);
    instSheet.getRange(instRow, 10).setValue(newStatus);
    instSheet.getRange(instRow, 11).setValue(paymentDate);
    instSheet.getRange(instRow, 12).setValue(data.Reference_No || '');

    // Update Loan Row: Total_Paid (col 10), Remaining_Amount (col 11), Loan_Status (col 12)
    const newTotalLoanPaid = (Number(loan.Total_Paid) || 0) + amountPaid;
    const newRemainingAmount = Math.max(0, (Number(loan.Loan_Amount) || 0) - newTotalLoanPaid);
    const newLoanStatus = (newRemainingAmount === 0) ? 'Completed' : 'Active';

    loanSheet.getRange(loanRow, 10).setValue(newTotalLoanPaid);
    loanSheet.getRange(loanRow, 11).setValue(newRemainingAmount);
    loanSheet.getRange(loanRow, 12).setValue(newLoanStatus);

    // Record Payment
    const paymentId = getNextID(paySheet, 'PAY', 6001);
    paySheet.appendRow([
      paymentId,
      loanId,
      loan.Borrower_ID,
      loan.Borrower_Name,
      targetInst.Installment_ID,
      amountPaid,
      paymentDate,
      data.Reference_No || '',
      data.Notes || '',
      now
    ]);

    // Record Central Transaction
    recordTransaction('LOAN_REPAYMENT', 'Installment Collection', loan.Borrower_Type, loan.Borrower_ID, loan.Borrower_Name, amountPaid, 'IN', data.Reference_No || paymentId, `Installment #${targetInst.Installment_Number} Payment for Loan ${loanId}`, paymentDate);
    logAudit('RECORD_PAYMENT', 'Loan_Payments', paymentId, 'Admin', `Received ৳ ${amountPaid} payment for Loan ${loanId} (${newStatus})`);

    return {
      success: true,
      message: `Payment of ৳ ${amountPaid.toLocaleString()} recorded successfully! Installment status: ${newStatus}.`,
      paymentId: paymentId
    };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * TRANSACTIONS LEDGER
 * ============================================================================
 */
function getTransactions(filters) {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.TRANSACTIONS);
  let rows = sheetToJSON(sheet).reverse();

  if (filters) {
    if (filters.type && filters.type !== 'ALL') {
      rows = rows.filter(r => r.Type === filters.type);
    }
    if (filters.direction && filters.direction !== 'ALL') {
      rows = rows.filter(r => r.Direction === filters.direction);
    }
    if (filters.dateFrom) {
      rows = rows.filter(r => r.Date >= filters.dateFrom);
    }
    if (filters.dateTo) {
      rows = rows.filter(r => r.Date <= filters.dateTo);
    }
    if (filters.search) {
      const q = filters.search.toLowerCase();
      rows = rows.filter(r => 
        String(r.Transaction_ID).toLowerCase().includes(q) ||
        String(r.Entity_Name).toLowerCase().includes(q) ||
        String(r.Reference_No).toLowerCase().includes(q)
      );
    }
  }

  return { success: true, data: rows };
}

/**
 * ============================================================================
 * REPORTS ENGINE
 * ============================================================================
 */
function getReportData(reportTypeOrParams, dateFrom, dateTo) {
  const ss = getDB();

  let reportType = reportTypeOrParams;
  let dFrom = dateFrom;
  let dTo = dateTo;

  if (typeof reportTypeOrParams === 'object' && reportTypeOrParams !== null) {
    reportType = reportTypeOrParams.type;
    dFrom = reportTypeOrParams.dateFrom;
    dTo = reportTypeOrParams.dateTo;
  }

  switch (reportType) {
    case 'shareholders': {
      const data = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS));
      return { success: true, data };
    }
    case 'deposits': {
      let data = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.DEPOSITS));
      if (dFrom) data = data.filter(d => d.Deposit_Date >= dFrom);
      if (dTo) data = data.filter(d => d.Deposit_Date <= dTo);
      return { success: true, data };
    }
    case 'withdrawals': {
      let data = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.WITHDRAWALS));
      if (dFrom) data = data.filter(w => w.Withdrawal_Date >= dFrom);
      if (dTo) data = data.filter(w => w.Withdrawal_Date <= dTo);
      return { success: true, data };
    }
    case 'loans': {
      let data = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.LOANS));
      if (dFrom) data = data.filter(l => l.Loan_Date >= dFrom);
      if (dTo) data = data.filter(l => l.Loan_Date <= dTo);
      return { success: true, data };
    }
    case 'borrowers': {
      const borrowers = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.BORROWERS));
      const loans = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.LOANS));
      
      const enriched = borrowers.map(b => {
        const bLoans = loans.filter(l => l.Borrower_ID === b.Borrower_ID);
        const totalBorrowed = bLoans.reduce((sum, l) => sum + (Number(l.Loan_Amount) || 0), 0);
        const totalRepaid = bLoans.reduce((sum, l) => sum + (Number(l.Total_Paid) || 0), 0);
        const currentBalance = totalBorrowed - totalRepaid;
        return {
          ...b,
          Total_Loans: bLoans.length,
          Total_Borrowed: totalBorrowed,
          Total_Repaid: totalRepaid,
          Current_Exposure: currentBalance
        };
      });
      return { success: true, data: enriched };
    }
    case 'installments': {
      let data = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.INSTALLMENTS));
      if (dFrom) data = data.filter(i => i.Due_Date >= dFrom);
      if (dTo) data = data.filter(i => i.Due_Date <= dTo);
      return { success: true, data };
    }
    case 'outstanding': {
      const loans = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.LOANS));
      const outstanding = loans.filter(l => Number(l.Remaining_Amount) > 0);
      return { success: true, data: outstanding };
    }
    case 'financial_summary': {
      let txns = sheetToJSON(ss.getSheetByName(CONFIG.SHEET_NAMES.TRANSACTIONS));
      if (dFrom) txns = txns.filter(t => t.Date >= dFrom);
      if (dTo) txns = txns.filter(t => t.Date <= dTo);

      let totalInflow = 0;
      let totalOutflow = 0;
      let breakdown = {
        deposits: 0,
        repayments: 0,
        withdrawals: 0,
        disbursements: 0
      };

      txns.forEach(t => {
        const amt = Number(t.Amount) || 0;
        if (t.Direction === 'IN') {
          totalInflow += amt;
          if (t.Type === 'DEPOSIT') breakdown.deposits += amt;
          if (t.Type === 'LOAN_REPAYMENT') breakdown.repayments += amt;
        } else if (t.Direction === 'OUT') {
          totalOutflow += amt;
          if (t.Type === 'WITHDRAWAL') breakdown.withdrawals += amt;
          if (t.Type === 'LOAN_DISBURSEMENT') breakdown.disbursements += amt;
        }
      });

      return {
        success: true,
        data: {
          totalInflow,
          totalOutflow,
          netLiquidity: totalInflow - totalOutflow,
          breakdown,
          transactionCount: txns.length
        }
      };
    }
    default:
      return { success: false, error: 'Unknown report type requested' };
  }
}

/**
 * ============================================================================
 * SETTINGS API
 * ============================================================================
 */
function getSettings() {
  const ss = getDB();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  const rows = sheetToJSON(sheet);
  const map = {};
  rows.forEach(r => {
    map[r.Setting_Key] = r.Setting_Value;
  });
  return { success: true, data: map };
}

function updateSettings(settingsMap) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
    const ss = getDB();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, error: 'Settings sheet is empty' };

    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i][0];
      if (settingsMap[key] !== undefined) {
        sheet.getRange(i + 2, 2).setValue(settingsMap[key]);
      }
    }

    logAudit('UPDATE_SETTINGS', 'Settings', 'CONFIG', 'Admin', 'Updated platform configuration values');
    return { success: true, message: 'Settings saved successfully!' };
  } catch (err) {
    return { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * SEED DEMO DATA (Optional Administrator Helper)
 * ============================================================================
 */
function seedDemoData() {
  initializeDatabase();
  const ss = getDB();

  // Check if shareholders already exist
  const shSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SHAREHOLDERS);
  if (shSheet.getLastRow() > 1) {
    return { success: false, message: 'Database already contains data. Seeding skipped to preserve existing records.' };
  }

  // 1. Add Realistic Shareholders
  addShareholder({
    Full_Name: 'Mohammad Rafiqul Islam',
    Phone: '+880 1711-234567',
    Email: 'rafiqul.islam@gmail.com',
    Address: 'House 42, Road 11, Banani, Dhaka',
    NID_Number: '19822694580001234',
    Joining_Date: '2026-01-10',
    Number_Of_Shares: 150,
    Status: 'Active',
    Notes: 'Founding member and director'
  });

  addShareholder({
    Full_Name: 'Begum Shamsun Nahar',
    Phone: '+880 1819-876543',
    Email: 'nahar.shamsun@yahoo.com',
    Address: 'Flat 4B, Concord Tower, Dhanmondi, Dhaka',
    NID_Number: '19792694580005678',
    Joining_Date: '2026-02-01',
    Number_Of_Shares: 200,
    Status: 'Active',
    Notes: 'Senior investor'
  });

  addShareholder({
    Full_Name: 'Engr. Tanvir Ahmed',
    Phone: '+880 1912-334455',
    Email: 'tanvir.ahmed@primeconsult.com',
    Address: 'Plot 15, Sector 4, Uttara, Dhaka',
    NID_Number: '19882694580009988',
    Joining_Date: '2026-03-15',
    Number_Of_Shares: 100,
    Status: 'Active',
    Notes: 'Advisory committee'
  });

  // 2. Add Borrowers
  addBorrower({
    Borrower_Type: 'Company / Business',
    Full_Name: 'Apex Logistics & Transport LTD',
    Phone: '+880 1713-998877',
    Email: 'finance@apexlogistics-bd.com',
    Address: 'Tejgaon I/A, Dhaka',
    NID_Or_Reg_No: 'C-184920/2021',
    Contact_Person: 'Kazi Mahmudul Hasan (CFO)',
    Registration_Date: '2026-04-01',
    Status: 'Active',
    Notes: 'Fleet expansion loan application'
  });

  addBorrower({
    Borrower_Type: 'Outside Individual',
    Full_Name: 'Farhan Kabir Chowdhury',
    Phone: '+880 1612-445566',
    Email: 'farhan.kabir@outlook.com',
    Address: 'Mirpur DOHS, Road 6, Dhaka',
    NID_Or_Reg_No: '19902694580007711',
    Contact_Person: '',
    Registration_Date: '2026-04-10',
    Status: 'Active',
    Notes: 'Small enterprise trader'
  });

  // 3. Issue Loans
  const loan1 = addLoan({
    Borrower_ID: 'BR-1001',
    Borrower_Name: 'Apex Logistics & Transport LTD',
    Borrower_Type: 'Company / Business',
    Loan_Amount: 600000,
    Loan_Date: '2026-04-15',
    Duration_Months: 12,
    First_Installment_Date: '2026-05-15',
    Notes: '0% Interest commercial transport loan'
  });

  const loan2 = addLoan({
    Borrower_ID: 'BR-1002',
    Borrower_Name: 'Farhan Kabir Chowdhury',
    Borrower_Type: 'Outside Individual',
    Loan_Amount: 150000,
    Loan_Date: '2026-05-01',
    Duration_Months: 6,
    First_Installment_Date: '2026-06-01',
    Notes: '0% Interest trade supply finance'
  });

  // 4. Record sample installment payment
  const instSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INSTALLMENTS);
  const instData = sheetToJSON(instSheet);
  if (instData.length > 0) {
    const firstInst = instData[0];
    recordInstallmentPayment({
      Installment_ID: firstInst.Installment_ID,
      Amount_Paid: Number(firstInst.Due_Amount),
      Payment_Date: '2026-05-14',
      Reference_No: 'CHQ-778901',
      Notes: 'First monthly installment paid via Account Payee Cheque'
    });
  }

  return { success: true, message: 'Demo data seeded successfully with realistic Bangladeshi corporate financial records!' };
}
