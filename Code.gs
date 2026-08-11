// ==========================================
// SECTION 1: CONFIGURATION CONSTANTS
// ==========================================
var SHEET_SETTING = 'Setting';
var SHEET_TRANSAKSI = 'Transaksi';
var SHEET_PINJAMAN = 'Pinjaman';
var SESSION_DURATION = 21600; // 6 hours

// ==========================================
// SECTION 2: UTILITY FUNCTIONS
// ==========================================

function hashSHA256(input) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  var txtHash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) {
      hashVal += 256;
    }
    if (hashVal.toString(16).length == 1) {
      txtHash += '0';
    }
    txtHash += hashVal.toString(16);
  }
  return txtHash.toLowerCase();
}

function generateToken() {
  return Utilities.getUuid();
}

function generateId(prefix) {
  return prefix + '_' + Utilities.getUuid().substring(0, 8).toUpperCase();
}

function storeSession(token, data) {
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(data), SESSION_DURATION);
}

function getSession(token) {
  if (!token) return null;
  var data = CacheService.getScriptCache().get('sess_' + token);
  if (data) {
    try {
      return JSON.parse(data);
    } catch (e) {
      return null;
    }
  }
  return null;
}

function destroySession(token) {
  if (token) {
    CacheService.getScriptCache().remove('sess_' + token);
  }
}

function getSheetByName(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getSheetData(sheetName) {
  var sheet = getSheetByName(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Only headers or empty
  
  var headers = data[0];
  var result = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    obj._rowIndex = i + 1; // 1-indexed sheet row
    result.push(obj);
  }
  return result;
}

function formatDate(date) {
  return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function jsonOutput(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// SECTION 3: WEB APP ENTRY POINTS
// ==========================================

function doGet(e) {
  if (!e.parameter || !e.parameter.action) {
    return jsonOutput({ success: true, message: 'API Backend Pinjaman Mingguan Google Apps Script Berjalan' });
  }
  return handleApiRequest(e.parameter, 'GET');
}

function doPost(e) {
  var params = {};
  if (e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      for (var key in body) {
        params[key] = body[key];
      }
    } catch (error) {
      // Ignored
    }
  }
  for (var k in e.parameter) {
    params[k] = e.parameter[k];
  }
  return handleApiRequest(params, 'POST');
}

function handleApiRequest(params, method) {
  try {
    if (params.action === 'login') {
      return jsonOutput(apiLogin(params.username, params.password));
    }

    var token = params.token;
    var session = getSession(token);
    if (!session) {
      return jsonOutput({ success: false, error: 'Unauthorized. Invalid or expired token.', code: 401 });
    }

    if (params.action.indexOf('admin/') === 0) {
      if (session.role !== 'admin') {
        return jsonOutput({ success: false, error: 'Forbidden. Admin access required.', code: 403 });
      }
      
      switch (params.action) {
        case 'admin/resot-list':
          return jsonOutput(apiGetResotList(token));
        case 'admin/nasabah-by-resot':
          return jsonOutput(apiGetNasabahByResot(token, params.resot));
        case 'admin/batch-payments':
          return jsonOutput(apiBatchPayments(token, params.resot, params.payments));
        case 'admin/renew':
          return jsonOutput(apiRenewLoan(token, params));
        case 'admin/users':
          return jsonOutput(apiGetUsers(token));
        case 'admin/add-user':
          return jsonOutput(apiAddUser(token, params));
        case 'admin/toggle-setting':
          return jsonOutput(apiToggleSetting(token, params));
        default:
          return jsonOutput({ success: false, error: 'Unknown admin action', code: 400 });
      }
    } else if (params.action.indexOf('nasabah/') === 0) {
      switch (params.action) {
        case 'nasabah/summary':
          return jsonOutput(apiGetNasabahSummary(token));
        case 'nasabah/change-password':
          return jsonOutput(apiChangePassword(token, params.oldPassword, params.newPassword));
        default:
          return jsonOutput({ success: false, error: 'Unknown nasabah action', code: 400 });
      }
    } else {
       return jsonOutput({ success: false, error: 'Unknown action', code: 400 });
    }
  } catch (err) {
    return jsonOutput({ success: false, error: err.toString(), code: 500 });
  }
}

// ==========================================
// SECTION 4: API - AUTHENTICATION
// ==========================================

/**
 * Handles user login.
 * @param {string} username - User username
 * @param {string} password - User password
 * @returns {Object} JSON response with token on success
 */
function apiLogin(username, password) {
  if (!username || !password) {
    return { success: false, error: 'Username and password are required', code: 400 };
  }
  
  var cleanUsername = String(username).trim().toLowerCase();
  var cleanPassword = String(password).trim();
  var users = getSheetData(SHEET_SETTING);
  var hashedPass = hashSHA256(cleanPassword);
  
  for (var i = 0; i < users.length; i++) {
    var user = users[i];
    if (String(user.Username).trim().toLowerCase() === cleanUsername && String(user.Password).trim() === hashedPass) {
      var token = generateToken();
      var sessionData = {
        userId: user.ID,
        username: user.Username,
        displayName: user.Display_Name,
        role: user.Role,
        sheetParam: user.Sheet_Param,
        noAnggota: user.No_Anggota,
        currentLoanId: user.Current_Loan_ID
      };
      storeSession(token, sessionData);
      
      return {
        success: true,
        token: token,
        user: {
          id: user.ID,
          displayName: user.Display_Name,
          role: user.Role,
          sheetParam: user.Sheet_Param,
          noAnggota: user.No_Anggota
        }
      };
    }
  }
  return { success: false, error: 'Invalid username or password', code: 401 };
}

// ==========================================
// SECTION 5: API - ADMIN FUNCTIONS
// ==========================================

/**
 * Gets unique active resot list.
 * @param {string} token - Session token
 * @returns {Object} JSON response with resot array
 */
function apiGetResotList(token) {
  var loans = getSheetData(SHEET_PINJAMAN);
  var resotMap = {};
  for (var i = 0; i < loans.length; i++) {
    if (loans[i].Status === 'Berjalan' && loans[i].Resot) {
      resotMap[loans[i].Resot] = true;
    }
  }
  var resots = Object.keys(resotMap).sort();
  return { success: true, data: resots };
}

/**
 * Gets active nasabah by resot.
 * @param {string} token - Session token
 * @param {string} resot - Resot identifier
 * @returns {Object} JSON response with nasabah array
 */
function apiGetNasabahByResot(token, resot) {
  if (!resot) {
    return { success: false, error: 'Resot parameter is required', code: 400 };
  }
  var loans = getSheetData(SHEET_PINJAMAN);
  var result = [];
  for (var i = 0; i < loans.length; i++) {
    if (loans[i].Status === 'Berjalan' && loans[i].Resot === resot) {
      result.push({
        noAnggota: loans[i].No_Anggota,
        nama: loans[i].Nama,
        pinjaman: loans[i].Pinjaman,
        saldo: loans[i].Saldo,
        tabungan: loans[i].Tabungan,
        idPinjaman: loans[i].ID_Pinjaman
      });
    }
  }
  return { success: true, data: result };
}

/**
 * Processes batch payments.
 * @param {string} token - Session token
 * @param {string} resot - Resot identifier
 * @param {string|Array} payments - Payment data array
 * @returns {Object} JSON response with process summary
 */
function apiBatchPayments(token, resot, payments) {
  if (!resot || !payments) {
    return { success: false, error: 'Resot and payments are required', code: 400 };
  }
  
  var paymentArr = typeof payments === 'string' ? JSON.parse(payments) : payments;
  if (!Array.isArray(paymentArr) || paymentArr.length === 0) {
    return { success: false, error: 'Payments must be a non-empty array', code: 400 };
  }

  var transSheet = getSheetByName(SHEET_TRANSAKSI);
  var pinjamanSheet = getSheetByName(SHEET_PINJAMAN);
  var loans = getSheetData(SHEET_PINJAMAN);
  
  var today = formatDate(new Date());
  var processed = 0;
  var totalNominal = 0;

  for (var i = 0; i < paymentArr.length; i++) {
    var p = paymentArr[i];
    var nominal = Number(p.nominal);
    if (isNaN(nominal) || nominal <= 0) continue;

    transSheet.appendRow([
      generateId('T'),
      today,
      resot,
      p.noAnggota,
      p.nama,
      nominal,
      p.idPinjaman,
      'Angsuran Mingguan'
    ]);

    for (var j = 0; j < loans.length; j++) {
      if (loans[j].ID_Pinjaman === p.idPinjaman) {
        var newSaldo = Math.max(0, loans[j].Saldo - nominal);
        var status = newSaldo === 0 ? 'Lunas' : loans[j].Status;
        
        var rowIndex = loans[j]._rowIndex;
        pinjamanSheet.getRange(rowIndex, 7).setValue(newSaldo); // Col 7: Saldo
        pinjamanSheet.getRange(rowIndex, 9).setValue(status);   // Col 9: Status
        
        loans[j].Saldo = newSaldo;
        loans[j].Status = status;
        
        processed++;
        totalNominal += nominal;
        break;
      }
    }
  }

  return { success: true, processed: processed, totalNominal: totalNominal, message: 'Pembayaran berhasil diproses' };
}

/**
 * Renews a loan.
 * @param {string} token - Session token
 * @param {Object} data - Renewal data
 * @returns {Object} JSON response
 */
function apiRenewLoan(token, data) {
  var idPinjaman = data.idPinjaman;
  var pinjamanBaru = Number(data.pinjamanBaru);
  
  if (!idPinjaman || isNaN(pinjamanBaru) || pinjamanBaru <= 0) {
    return { success: false, error: 'Invalid parameters', code: 400 };
  }

  var loans = getSheetData(SHEET_PINJAMAN);
  var oldLoan = null;
  for (var i = 0; i < loans.length; i++) {
    if (loans[i].ID_Pinjaman === idPinjaman) {
      oldLoan = loans[i];
      break;
    }
  }

  if (!oldLoan) {
    return { success: false, error: 'Loan not found', code: 404 };
  }

  var today = formatDate(new Date());
  var transSheet = getSheetByName(SHEET_TRANSAKSI);
  var pinjamanSheet = getSheetByName(SHEET_PINJAMAN);
  var sisaSaldoLama = Number(oldLoan.Saldo) || 0;

  // Insert trans Pelunasan
  if (sisaSaldoLama > 0) {
    transSheet.appendRow([
      generateId('T'),
      today,
      oldLoan.Resot,
      oldLoan.No_Anggota,
      oldLoan.Nama,
      sisaSaldoLama,
      idPinjaman,
      'Pelunasan Pembaruan'
    ]);
  }

  // Update old loan
  pinjamanSheet.getRange(oldLoan._rowIndex, 7).setValue(0); // Saldo
  pinjamanSheet.getRange(oldLoan._rowIndex, 9).setValue('Diperbarui'); // Status

  // Create new loan
  var newLoanId = generateId('P');
  var saldoBaru = pinjamanBaru * 1.2;
  var tabunganBaru = (Number(oldLoan.Tabungan) || 0) + (pinjamanBaru * 0.05);

  pinjamanSheet.appendRow([
    newLoanId,
    oldLoan.Resot,
    today,
    oldLoan.No_Anggota,
    oldLoan.Nama,
    pinjamanBaru,
    saldoBaru,
    tabunganBaru,
    'Berjalan'
  ]);

  // Update Setting Current_Loan_ID
  var settingsSheet = getSheetByName(SHEET_SETTING);
  var users = getSheetData(SHEET_SETTING);
  for (var j = 0; j < users.length; j++) {
    if (users[j].No_Anggota === oldLoan.No_Anggota) {
      settingsSheet.getRange(users[j]._rowIndex, 9).setValue(newLoanId); // Col 9: Current_Loan_ID
      break;
    }
  }

  return { 
    success: true, 
    newLoanId: newLoanId, 
    pinjamanBaru: pinjamanBaru, 
    saldoBaru: saldoBaru, 
    tabunganBaru: tabunganBaru, 
    message: 'Pinjaman berhasil diperbarui' 
  };
}

/**
 * Gets all users for admin.
 * @param {string} token - Session token
 * @returns {Object} JSON response
 */
function apiGetUsers(token) {
  var users = getSheetData(SHEET_SETTING);
  var result = [];
  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    result.push({
      id: u.ID,
      username: u.Username,
      displayName: u.Display_Name,
      role: u.Role,
      sheetParam: u.Sheet_Param,
      noAnggota: u.No_Anggota,
      isShowDetail: u.Is_Show_Detail,
      currentLoanId: u.Current_Loan_ID,
      isShowTabungan: u.Is_Show_Tabungan
    });
  }
  return { success: true, data: result };
}

/**
 * Adds a new user.
 * @param {string} token - Session token
 * @param {Object} data - User data
 * @returns {Object} JSON response
 */
function apiAddUser(token, data) {
  if (!data.username || !data.password || !data.displayName) {
    return { success: false, error: 'Username, password, and display name are required', code: 400 };
  }

  var users = getSheetData(SHEET_SETTING);
  for (var i = 0; i < users.length; i++) {
    if (users[i].Username === data.username) {
      return { success: false, error: 'Username already exists', code: 400 };
    }
  }

  var newId = generateId('U');
  var hashedPassword = hashSHA256(data.password);
  var settingsSheet = getSheetByName(SHEET_SETTING);

  settingsSheet.appendRow([
    newId,
    data.username,
    hashedPassword,
    data.displayName,
    data.role || 'nasabah',
    data.sheetParam || '',
    data.noAnggota || '',
    true, // Is_Show_Detail
    data.currentLoanId || '',
    true  // Is_Show_Tabungan
  ]);

  return { success: true, id: newId, message: 'User added successfully' };
}

/**
 * Toggles a user setting.
 * @param {string} token - Session token
 * @param {Object} data - Setting data
 * @returns {Object} JSON response
 */
function apiToggleSetting(token, data) {
  if (!data.userId || !data.field || data.value === undefined) {
    return { success: false, error: 'Missing required parameters', code: 400 };
  }

  var colIndex = 0;
  if (data.field === 'Is_Show_Detail') {
    colIndex = 8;
  } else if (data.field === 'Is_Show_Tabungan') {
    colIndex = 10;
  } else {
    return { success: false, error: 'Invalid field', code: 400 };
  }

  var users = getSheetData(SHEET_SETTING);
  var targetRow = 0;
  for (var i = 0; i < users.length; i++) {
    if (users[i].ID === data.userId) {
      targetRow = users[i]._rowIndex;
      break;
    }
  }

  if (targetRow === 0) {
    return { success: false, error: 'User not found', code: 404 };
  }

  var settingsSheet = getSheetByName(SHEET_SETTING);
  settingsSheet.getRange(targetRow, colIndex).setValue(data.value);

  return { success: true, message: 'Setting updated successfully' };
}

// ==========================================
// SECTION 6: API - NASABAH FUNCTIONS
// ==========================================

/**
 * Gets nasabah dashboard summary.
 * @param {string} token - Session token
 * @returns {Object} JSON response
 */
function apiGetNasabahSummary(token) {
  var session = getSession(token);
  var users = getSheetData(SHEET_SETTING);
  var user = null;

  for (var i = 0; i < users.length; i++) {
    if (users[i].ID === session.userId) {
      user = users[i];
      break;
    }
  }

  if (!user) {
    return { success: false, error: 'User not found', code: 404 };
  }

  var isShowDetail = (user.Is_Show_Detail === true || String(user.Is_Show_Detail).toUpperCase() === 'TRUE');
  var isShowTabungan = (user.Is_Show_Tabungan === true || String(user.Is_Show_Tabungan).toUpperCase() === 'TRUE');

  var currentLoan = null;
  if (user.Current_Loan_ID) {
    var loans = getSheetData(SHEET_PINJAMAN);
    for (var j = 0; j < loans.length; j++) {
      if (loans[j].ID_Pinjaman === user.Current_Loan_ID) {
        currentLoan = {
          idPinjaman: loans[j].ID_Pinjaman,
          resot: loans[j].Resot,
          tanggal: formatDate(loans[j].Tanggal),
          pinjaman: loans[j].Pinjaman,
          saldo: loans[j].Saldo,
          tabungan: isShowTabungan ? loans[j].Tabungan : null,
          status: loans[j].Status
        };
        break;
      }
    }
  }

  var transactions = [];
  if (isShowDetail && currentLoan) {
    var allTrans = getSheetData(SHEET_TRANSAKSI);
    for (var k = 0; k < allTrans.length; k++) {
      if (allTrans[k].ID_Pinjaman === currentLoan.idPinjaman) {
        transactions.push({
          id: allTrans[k].ID,
          tanggal: formatDate(allTrans[k].Tanggal),
          pembayaran: allTrans[k].Pembayaran,
          keterangan: allTrans[k].Keterangan
        });
      }
    }
  }

  return {
    success: true,
    data: {
      user: {
        displayName: user.Display_Name,
        noAnggota: user.No_Anggota,
        sheetParam: user.Sheet_Param
      },
      loan: currentLoan,
      transactions: transactions,
      isShowDetail: isShowDetail,
      isShowTabungan: isShowTabungan
    }
  };
}

/**
 * Changes nasabah password.
 * @param {string} token - Session token
 * @param {string} oldPassword - Old password
 * @param {string} newPassword - New password
 * @returns {Object} JSON response
 */
function apiChangePassword(token, oldPassword, newPassword) {
  if (!oldPassword || !newPassword || newPassword.length < 6) {
    return { success: false, error: 'Invalid input or new password too short', code: 400 };
  }

  var session = getSession(token);
  var users = getSheetData(SHEET_SETTING);
  var targetRow = 0;
  var storedHashedPass = '';

  for (var i = 0; i < users.length; i++) {
    if (users[i].ID === session.userId) {
      targetRow = users[i]._rowIndex;
      storedHashedPass = users[i].Password;
      break;
    }
  }

  if (targetRow === 0) {
    return { success: false, error: 'User not found', code: 404 };
  }

  if (hashSHA256(oldPassword) !== storedHashedPass) {
    return { success: false, error: 'Old password incorrect', code: 401 };
  }

  var settingsSheet = getSheetByName(SHEET_SETTING);
  settingsSheet.getRange(targetRow, 3).setValue(hashSHA256(newPassword)); // Col 3: Password

  return { success: true, message: 'Password changed successfully' };
}

// ==========================================
// SECTION 7: SETUP DATABASE
// ==========================================

function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function initSheet(name, headers, data) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      sheet.clear();
    }
    
    // Headers
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4a86c8');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    
    // Data
    if (data && data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
    
    // Auto Resize
    for (var col = 1; col <= headers.length; col++) {
      sheet.autoResizeColumn(col);
    }
  }

  // Setting
  var setHeaders = ['ID', 'Username', 'Password', 'Display_Name', 'Role', 'Sheet_Param', 'No_Anggota', 'Is_Show_Detail', 'Current_Loan_ID', 'Is_Show_Tabungan'];
  var setRows = [
    ['U001', 'admin', hashSHA256('admin123'), 'Administrator', 'admin', '', '', true, '', true],
    ['U002', 'budi', hashSHA256('pass123'), 'Budi Santoso', 'nasabah', 'Resot22', 'A001', true, 'P001', true],
    ['U003', 'siti', hashSHA256('pass123'), 'Siti Aminah', 'nasabah', 'Resot22', 'A002', true, 'P002', true],
    ['U004', 'andi', hashSHA256('pass123'), 'Andi Wijaya', 'nasabah', 'Resot23', 'A003', false, 'P003', false]
  ];
  initSheet(SHEET_SETTING, setHeaders, setRows);

  // Pinjaman
  var pinHeaders = ['ID_Pinjaman', 'Resot', 'Tanggal', 'No_Anggota', 'Nama', 'Pinjaman', 'Saldo', 'Tabungan', 'Status'];
  var pinRows = [
    ['P001', 'Resot22', '2026-07-01', 'A001', 'Budi Santoso', 1000000, 800000, 50000, 'Berjalan'],
    ['P002', 'Resot22', '2026-07-01', 'A002', 'Siti Aminah', 2000000, 1500000, 100000, 'Berjalan'],
    ['P003', 'Resot23', '2026-07-15', 'A003', 'Andi Wijaya', 1500000, 1200000, 75000, 'Berjalan']
  ];
  initSheet(SHEET_PINJAMAN, pinHeaders, pinRows);

  // Transaksi
  var transHeaders = ['ID', 'Tanggal', 'Resot', 'No_Anggota', 'Nama', 'Pembayaran', 'ID_Pinjaman', 'Keterangan'];
  var transRows = [
    ['T001', '2026-08-04', 'Resot22', 'A001', 'Budi Santoso', 100000, 'P001', 'Angsuran Mingguan'],
    ['T002', '2026-08-04', 'Resot22', 'A002', 'Siti Aminah', 250000, 'P002', 'Angsuran Mingguan'],
    ['T003', '2026-08-04', 'Resot23', 'A003', 'Andi Wijaya', 150000, 'P003', 'Angsuran Mingguan'],
    ['T004', '2026-08-11', 'Resot22', 'A001', 'Budi Santoso', 100000, 'P001', 'Angsuran Mingguan'],
    ['T005', '2026-08-11', 'Resot22', 'A002', 'Siti Aminah', 250000, 'P002', 'Angsuran Mingguan']
  ];
  initSheet(SHEET_TRANSAKSI, transHeaders, transRows);

  return 'Database berhasil diinisialisasi dengan data dummy!';
}
