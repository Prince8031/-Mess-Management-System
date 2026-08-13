/**
 * 🍽️ Mess Management System - Server Backend
 * Developed by Prince Kumar
 */

// ==========================================
// 1. CONFIGURATION & CONSTANTS
// ==========================================

const SHEETS = {
  USERS: 'Users',
  STUDENTS: 'Students',
  MANAGERS: 'Managers',
  MENU: 'Menu',
  ATTENDANCE: 'Attendance',
  BILLS: 'Bills',
  PAYMENTS: 'Payments',
  INVENTORY: 'Inventory',
  INVENTORY_TRANSACTIONS: 'InventoryTransactions',
  EXPENSES: 'Expenses',
  COMPLAINTS: 'Complaints',
  NOTICES: 'Notices',
  SETTINGS: 'Settings',
  ACTIVITY_LOG: 'ActivityLog'
};

const SCHEMAS = {
  [SHEETS.USERS]: ['UserID', 'Name', 'Email', 'PasswordHash', 'Role', 'Phone', 'Status', 'ProfileImage', 'CreatedAt', 'UpdatedAt', 'LastLogin'],
  [SHEETS.STUDENTS]: ['StudentID', 'UserID', 'RollNo', 'Name', 'Email', 'Phone', 'Course', 'Branch', 'Semester', 'Hostel', 'RoomNo', 'JoiningDate', 'Status', 'CreatedAt', 'UpdatedAt'],
  [SHEETS.MANAGERS]: ['ManagerID', 'UserID', 'Name', 'Email', 'Phone', 'MessName', 'JoiningDate', 'Status', 'CreatedAt', 'UpdatedAt'],
  [SHEETS.MENU]: ['MenuID', 'Date', 'MealType', 'Items', 'Description', 'Special', 'Status', 'CreatedBy', 'CreatedAt', 'UpdatedAt'],
  [SHEETS.ATTENDANCE]: ['AttendanceID', 'StudentID', 'Date', 'MealType', 'Status', 'MarkedBy', 'MarkedAt'],
  [SHEETS.BILLS]: ['BillID', 'StudentID', 'Month', 'Year', 'BaseFee', 'MealCharges', 'ExtraCharges', 'Fine', 'Discount', 'TotalAmount', 'PaidAmount', 'DueAmount', 'Status', 'DueDate', 'CreatedAt', 'UpdatedAt'],
  [SHEETS.PAYMENTS]: ['PaymentID', 'BillID', 'StudentID', 'Amount', 'PaymentMethod', 'TransactionID', 'PaymentDate', 'Status', 'VerifiedBy', 'Notes', 'CreatedAt'],
  [SHEETS.INVENTORY]: ['ItemID', 'ItemName', 'Category', 'Quantity', 'Unit', 'MinimumStock', 'PurchasePrice', 'Supplier', 'ExpiryDate', 'Status', 'LastUpdated', 'CreatedAt'],
  [SHEETS.INVENTORY_TRANSACTIONS]: ['TransactionID', 'ItemID', 'Type', 'Quantity', 'UnitPrice', 'Supplier', 'Date', 'Description', 'CreatedBy', 'CreatedAt'],
  [SHEETS.EXPENSES]: ['ExpenseID', 'Category', 'Description', 'Amount', 'Date', 'PaymentMethod', 'ReceiptURL', 'CreatedBy', 'CreatedAt'],
  [SHEETS.COMPLAINTS]: ['ComplaintID', 'StudentID', 'Category', 'Title', 'Description', 'Priority', 'Status', 'Response', 'CreatedAt', 'ResolvedAt', 'ResolvedBy'],
  [SHEETS.NOTICES]: ['NoticeID', 'Title', 'Description', 'TargetRole', 'Priority', 'PublishDate', 'ExpiryDate', 'CreatedBy', 'Status', 'CreatedAt'],
  [SHEETS.SETTINGS]: ['SettingKey', 'SettingValue', 'Description', 'UpdatedAt'],
  [SHEETS.ACTIVITY_LOG]: ['LogID', 'UserID', 'UserName', 'Action', 'Module', 'Details', 'IPAddress', 'Timestamp']
};

const DEFAULT_SETTINGS = [
  ['APP_NAME', 'Mess Management System', 'Application Name', new Date().toISOString()],
  ['MESS_NAME', 'Central University Campus Mess', 'Name of the Mess facility', new Date().toISOString()],
  ['BASE_MONTHLY_FEE', '2500', 'Standard base monthly charge in INR', new Date().toISOString()],
  ['CURRENCY', '₹', 'System currency symbol', new Date().toISOString()],
  ['DEFAULT_DUE_DAYS', '10', 'Days before monthly bill becomes overdue', new Date().toISOString()],
  ['SYSTEM_STATUS', 'Active', 'Overall system operational status', new Date().toISOString()],
  ['DEVELOPER_CREDIT', 'developed by Prince Kumar', 'System Developer Credit', new Date().toISOString()]
];

// ==========================================
// 2. WEB APP ENTRY POINT & INITIALIZATION
// ==========================================

function doGet() {
  initializeDatabase();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Mess Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initializeDatabase() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create missing sheets and write headers
    Object.keys(SCHEMAS).forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(SCHEMAS[sheetName]);
        sheet.getRange(1, 1, 1, SCHEMAS[sheetName].length).setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
      }
    });

    // Populate default settings if missing
    const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
    if (settingsSheet.getLastRow() <= 1) {
      DEFAULT_SETTINGS.forEach(setting => settingsSheet.appendRow(setting));
    }

    // Ensure Default Admin Exists
    const usersSheet = ss.getSheetByName(SHEETS.USERS);
    const users = sheetToObjects(usersSheet);
    const adminExists = users.some(u => u.Email.toLowerCase() === 'admin@mess.com');

    if (!adminExists) {
      const adminId = 'USR-00001';
      const now = formatDate(new Date(), true);
      const adminPassHash = hashPassword('Admin@123');
      usersSheet.appendRow([
        adminId,
        'System Administrator',
        'admin@mess.com',
        adminPassHash,
        'Admin',
        '9999999999',
        'Active',
        'https://ui-avatars.com/api/?name=Admin&background=2563eb&color=fff',
        now,
        now,
        ''
      ]);
      logActivity('SYSTEM', 'System Administrator', 'DATABASE_INIT', 'System', 'Initialized database and created default admin account');
    }

    return apiResponse(true, 'Database initialized successfully.');
  } catch (err) {
    return apiResponse(false, 'Database initialization error: ' + err.toString());
  } finally {
    lock.releaseLock();
  }
}

// Setup full demo dataset if required
function setupDemoData() {
  initializeDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = formatDate(new Date(), false);
  const nowTs = formatDate(new Date(), true);

  // Check if demo students exist
  const studentsSheet = ss.getSheetByName(SHEETS.STUDENTS);
  if (studentsSheet.getLastRow() <= 1) {
    const demoUsers = [
      ['USR-00002', 'Manager Ramesh', 'manager@mess.com', hashPassword('Manager@123'), 'Manager', '9876543210', 'Active', '', nowTs, nowTs, ''],
      ['USR-00003', 'Rahul Sharma', 'rahul@student.com', hashPassword('Student@123'), 'Student', '9811111111', 'Active', '', nowTs, nowTs, ''],
      ['USR-00004', 'Priya Patel', 'priya@student.com', hashPassword('Student@123'), 'Student', '9822222222', 'Active', '', nowTs, nowTs, ''],
      ['USR-00005', 'Amit Kumar', 'amit@student.com', hashPassword('Student@123'), 'Student', '9833333333', 'Active', '', nowTs, nowTs, '']
    ];
    const usersSheet = ss.getSheetByName(SHEETS.USERS);
    demoUsers.forEach(u => usersSheet.appendRow(u));

    const managersSheet = ss.getSheetByName(SHEETS.MANAGERS);
    managersSheet.appendRow(['MGR-00001', 'USR-00002', 'Manager Ramesh', 'manager@mess.com', '9876543210', 'Central Mess', now, 'Active', nowTs, nowTs]);

    const demoStudents = [
      ['STU-00001', 'USR-00003', 'CS2026-01', 'Rahul Sharma', 'rahul@student.com', '9811111111', 'B.Tech', 'CSE', '6th', 'Hostel A', '101', now, 'Active', nowTs, nowTs],
      ['STU-00002', 'USR-00004', 'CS2026-02', 'Priya Patel', 'priya@student.com', '9822222222', 'B.Tech', 'ECE', '6th', 'Hostel B', '204', now, 'Active', nowTs, nowTs],
      ['STU-00003', 'USR-00005', 'CS2026-03', 'Amit Kumar', 'amit@student.com', '9833333333', 'B.Tech', 'ME', '4th', 'Hostel A', '112', now, 'Active', nowTs, nowTs]
    ];
    demoStudents.forEach(s => studentsSheet.appendRow(s));

    // Demo Menu
    const menuSheet = ss.getSheetByName(SHEETS.MENU);
    const meals = [
      ['MEN-00001', now, 'Breakfast', 'Poha, Jalebi, Tea, Sprouts', 'Fresh breakfast served hot', 'Yes', 'Active', 'Manager Ramesh', nowTs, nowTs],
      ['MEN-00002', now, 'Lunch', 'Rice, Paneer Butter Masala, Dal Tadka, Roti, Salad', 'Special Paneer day', 'Yes', 'Active', 'Manager Ramesh', nowTs, nowTs],
      ['MEN-00003', now, 'Snacks', 'Samosa, Green Chutney, Tea', 'Evening refreshments', 'No', 'Active', 'Manager Ramesh', nowTs, nowTs],
      ['MEN-00004', now, 'Dinner', 'Chapati, Mix Veg, Chana Dal, Rice, Kheer', 'Sweet dessert included', 'Yes', 'Active', 'Manager Ramesh', nowTs, nowTs]
    ];
    meals.forEach(m => menuSheet.appendRow(m));

    // Demo Inventory
    const invSheet = ss.getSheetByName(SHEETS.INVENTORY);
    const inv = [
      ['INV-00001', 'Basmati Rice', 'Grocery', '250', 'Kg', '50', '60', 'Agro Suppliers', '2027-12-31', 'In Stock', nowTs, nowTs],
      ['INV-00002', 'Toor Dal', 'Grocery', '100', 'Kg', '30', '120', 'Agro Suppliers', '2027-06-30', 'In Stock', nowTs, nowTs],
      ['INV-00003', 'Cooking Oil', 'Grocery', '15', 'Ltr', '20', '140', 'Oil Traders', '2027-08-15', 'Low Stock', nowTs, nowTs],
      ['INV-00004', 'Commercial Gas Cylinder', 'Gas', '4', 'Units', '5', '1750', 'Indane Gas Agency', '2026-12-31', 'Low Stock', nowTs, nowTs]
    ];
    inv.forEach(i => invSheet.appendRow(i));

    // Demo Notice
    const noticesSheet = ss.getSheetByName(SHEETS.NOTICES);
    noticesSheet.appendRow(['NOT-00001', 'Special Sunday Feast', 'Sunday Lunch will include Special Paneer, Ice Cream and Gulab Jamun.', 'All', 'High', now, '2026-12-31', 'Manager Ramesh', 'Active', nowTs]);

    // Demo Bills & Payments
    const billsSheet = ss.getSheetByName(SHEETS.BILLS);
    billsSheet.appendRow(['BIL-00001', 'STU-00001', 'August', '2026', '2500', '1200', '0', '0', '0', '3700', '3700', '0', 'Paid', now, nowTs, nowTs]);
    billsSheet.appendRow(['BIL-00002', 'STU-00002', 'August', '2026', '2500', '1100', '0', '0', '100', '3500', '1500', '2000', 'Partial', now, nowTs, nowTs]);

    const paymentsSheet = ss.getSheetByName(SHEETS.PAYMENTS);
    paymentsSheet.appendRow(['PAY-00001', 'BIL-00001', 'STU-00001', '3700', 'UPI', 'TXN9988776655', now, 'Paid', 'Admin', 'Monthly Fee August', nowTs]);
    paymentsSheet.appendRow(['PAY-00002', 'BIL-00002', 'STU-00002', '1500', 'Cash', 'CASH-102', now, 'Paid', 'Manager Ramesh', 'Advance payment', nowTs]);

    // Demo Complaints
    const compSheet = ss.getSheetByName(SHEETS.COMPLAINTS);
    compSheet.appendRow(['CMP-00001', 'STU-00002', 'Food Quality', 'Chapati was cold', 'Dinner chapati served yesterday was quite cold.', 'Medium', 'Resolved', 'We have strictly advised staff to keep warm containers.', nowTs, nowTs, 'Manager Ramesh']);
  }
  return apiResponse(true, 'Demo data populated successfully.');
}

// ==========================================
// 3. UTILITIES & HELPER FUNCTIONS
// ==========================================

function getDb() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const ss = getDb();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet "' + name + '" not found. Please re-initialize database.');
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        val = formatDate(val, h.includes('At') || h.includes('Timestamp') || h === 'LastLogin');
      }
      obj[h] = val !== undefined && val !== null ? String(val) : '';
    });
    return obj;
  });
}

function hashPassword(password) {
  if (!password) return '';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + '_MESS_SALT_2026', Utilities.Charset.UTF_8);
  return digest.map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
}

function generateId(prefix, sheetName, colName) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const sheet = getSheet(sheetName);
    const data = sheetToObjects(sheet);
    if (data.length === 0) return prefix + '-00001';
    
    let maxNum = 0;
    data.forEach(item => {
      const val = item[colName] || '';
      if (val.startsWith(prefix + '-')) {
        const numPart = parseInt(val.split('-')[1], 10);
        if (!isNaN(numPart) && numPart > maxNum) maxNum = numPart;
      }
    });
    const nextNum = maxNum + 1;
    return prefix + '-' + ('00000' + nextNum).slice(-5);
  } finally {
    lock.releaseLock();
  }
}

function formatDate(dateObj, includeTime) {
  if (!dateObj) return '';
  const d = new Date(dateObj);
  if (isNaN(d.getTime())) return String(dateObj);
  
  const tz = Session.getScriptTimeZone() || 'GMT+5:30';
  const format = includeTime ? 'dd-MM-yyyy HH:mm:ss' : 'yyyy-MM-dd';
  return Utilities.formatDate(d, tz, format);
}

function apiResponse(success, message, data) {
  return {
    success: success,
    message: message,
    data: data || null,
    timestamp: formatDate(new Date(), true),
    developer: 'developed by Prince Kumar'
  };
}

function logActivity(userId, userName, action, module, details) {
  try {
    const sheet = getSheet(SHEETS.ACTIVITY_LOG);
    const logId = generateId('LOG', SHEETS.ACTIVITY_LOG, 'LogID');
    sheet.appendRow([
      logId,
      userId || 'GUEST',
      userName || 'Unknown',
      action,
      module,
      details || '',
      'WebApp-Internal',
      formatDate(new Date(), true)
    ]);
  } catch (e) {
    Logger.log('Logging failed: ' + e.toString());
  }
}

// ==========================================
// 4. AUTHENTICATION & SESSION MANAGEMENT
// ==========================================

function login(email, password) {
  try {
    if (!email || !password) return apiResponse(false, 'Email and password are required.');
    
    const usersSheet = getSheet(SHEETS.USERS);
    const users = sheetToObjects(usersSheet);
    const user = users.find(u => u.Email.toLowerCase().trim() === email.toLowerCase().trim());

    if (!user) {
      return apiResponse(false, 'Invalid credentials.');
    }

    if (user.Status !== 'Active') {
      return apiResponse(false, 'Your account is ' + user.Status + '. Please contact administrator.');
    }

    const hashedInput = hashPassword(password);
    if (user.PasswordHash !== hashedInput) {
      return apiResponse(false, 'Invalid credentials.');
    }

    // Generate Session Token
    const token = Utilities.getUuid();
    const sessionObj = {
      token: token,
      userId: user.UserID,
      email: user.Email,
      name: user.Name,
      role: user.Role,
      status: user.Status,
      loginTime: new Date().getTime()
    };

    const cache = CacheService.getUserCache();
    cache.put(token, JSON.stringify(sessionObj), 21600); // 6 hours

    // Update Last Login in Sheet
    const dataRange = usersSheet.getDataRange().getValues();
    const headers = dataRange[0];
    const userIdColIndex = headers.indexOf('UserID');
    const lastLoginColIndex = headers.indexOf('LastLogin');

    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][userIdColIndex] === user.UserID) {
        usersSheet.getRange(i + 1, lastLoginColIndex + 1).setValue(formatDate(new Date(), true));
        break;
      }
    }

    logActivity(user.UserID, user.Name, 'LOGIN', 'Authentication', 'User logged in successfully');

    // Fetch related Student or Manager record if applicable
    let profileData = {};
    if (user.Role === 'Student') {
      const students = sheetToObjects(getSheet(SHEETS.STUDENTS));
      profileData = students.find(s => s.UserID === user.UserID) || {};
    } else if (user.Role === 'Manager') {
      const managers = sheetToObjects(getSheet(SHEETS.MANAGERS));
      profileData = managers.find(m => m.UserID === user.UserID) || {};
    }

    return apiResponse(true, 'Login successful', {
      token: token,
      user: {
        userId: user.UserID,
        name: user.Name,
        email: user.Email,
        role: user.Role,
        phone: user.Phone,
        profileImage: user.ProfileImage || 'https://ui-avatars.com/api/?name=' + encodeURIComponent(user.Name) + '&background=2563eb&color=fff',
        studentId: profileData.StudentID || '',
        managerId: profileData.ManagerID || '',
        rollNo: profileData.RollNo || '',
        hostel: profileData.Hostel || '',
        roomNo: profileData.RoomNo || ''
      }
    });
  } catch (err) {
    return apiResponse(false, 'Login failed: ' + err.toString());
  }
}

function validateSession(token) {
  if (!token) return null;
  const cache = CacheService.getUserCache();
  const cached = cache.get(token);
  if (!cached) return null;
  try {
    return JSON.parse(cached);
  } catch (e) {
    return null;
  }
}

function authGuard(token, allowedRoles) {
  const session = validateSession(token);
  if (!session) {
    throw new Error('Session expired or invalid. Please login again.');
  }
  if (allowedRoles && allowedRoles.length > 0) {
    if (!allowedRoles.includes(session.role)) {
      throw new Error('Unauthorized action for role: ' + session.role);
    }
  }
  return session;
}

function logout(token) {
  try {
    const session = validateSession(token);
    if (session) {
      CacheService.getUserCache().remove(token);
      logActivity(session.userId, session.name, 'LOGOUT', 'Authentication', 'User logged out');
    }
    return apiResponse(true, 'Logged out successfully');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function changePassword(token, oldPassword, newPassword) {
  try {
    const session = authGuard(token, ['Admin', 'Manager', 'Student']);
    if (!newPassword || newPassword.length < 6) {
      return apiResponse(false, 'New password must be at least 6 characters long.');
    }

    const usersSheet = getSheet(SHEETS.USERS);
    const dataRange = usersSheet.getDataRange().getValues();
    const headers = dataRange[0];
    const userIdIdx = headers.indexOf('UserID');
    const passIdx = headers.indexOf('PasswordHash');
    const updatedIdx = headers.indexOf('UpdatedAt');

    const oldHash = hashPassword(oldPassword);
    const newHash = hashPassword(newPassword);

    for (let i = 1; i < dataRange.length; i++) {
      if (dataRange[i][userIdIdx] === session.userId) {
        if (dataRange[i][passIdx] !== oldHash) {
          return apiResponse(false, 'Current password does not match.');
        }
        usersSheet.getRange(i + 1, passIdx + 1).setValue(newHash);
        usersSheet.getRange(i + 1, updatedIdx + 1).setValue(formatDate(new Date(), true));
        logActivity(session.userId, session.name, 'CHANGE_PASSWORD', 'User', 'Password changed successfully');
        return apiResponse(true, 'Password updated successfully.');
      }
    }
    return apiResponse(false, 'User record not found.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 5. DASHBOARD AGGREGATOR DATA
// ==========================================

function getDashboardData(token) {
  try {
    const session = authGuard(token, ['Admin', 'Manager', 'Student']);
    const todayStr = formatDate(new Date(), false);

    const students = sheetToObjects(getSheet(SHEETS.STUDENTS));
    const managers = sheetToObjects(getSheet(SHEETS.MANAGERS));
    const menu = sheetToObjects(getSheet(SHEETS.MENU));
    const attendance = sheetToObjects(getSheet(SHEETS.ATTENDANCE));
    const bills = sheetToObjects(getSheet(SHEETS.BILLS));
    const payments = sheetToObjects(getSheet(SHEETS.PAYMENTS));
    const inventory = sheetToObjects(getSheet(SHEETS.INVENTORY));
    const expenses = sheetToObjects(getSheet(SHEETS.EXPENSES));
    const complaints = sheetToObjects(getSheet(SHEETS.COMPLAINTS));
    const notices = sheetToObjects(getSheet(SHEETS.NOTICES));

    let dashboard = {
      role: session.role,
      user: session
    };

    if (session.role === 'Admin' || session.role === 'Manager') {
      const activeStudents = students.filter(s => s.Status === 'Active').length;
      const totalManagers = managers.filter(m => m.Status === 'Active').length;
      
      const todayAttendance = attendance.filter(a => a.Date === todayStr && a.Status === 'Present');
      const bCount = todayAttendance.filter(a => a.MealType === 'Breakfast').length;
      const lCount = todayAttendance.filter(a => a.MealType === 'Lunch').length;
      const dCount = todayAttendance.filter(a => a.MealType === 'Dinner').length;

      const totalRev = payments.filter(p => p.Status === 'Paid').reduce((acc, p) => acc + (parseFloat(p.Amount) || 0), 0);
      const totalExp = expenses.reduce((acc, e) => acc + (parseFloat(e.Amount) || 0), 0);
      const pendingBillsCount = bills.filter(b => b.Status === 'Pending' || b.Status === 'Partial' || b.Status === 'Overdue').length;
      const pendingComplaintsCount = complaints.filter(c => c.Status === 'Pending' || c.Status === 'In Progress').length;
      const lowStockCount = inventory.filter(i => i.Status === 'Low Stock' || i.Status === 'Out of Stock').length;

      dashboard.kpis = {
        totalStudents: activeStudents,
        totalManagers: totalManagers,
        todayMeals: { breakfast: bCount, lunch: lCount, dinner: dCount, total: todayAttendance.length },
        totalRevenue: totalRev,
        totalExpenses: totalExp,
        pendingBills: pendingBillsCount,
        pendingComplaints: pendingComplaintsCount,
        lowStockItems: lowStockCount
      };

      dashboard.todayMenu = menu.filter(m => m.Date === todayStr && m.Status === 'Active');
      dashboard.recentComplaints = complaints.slice(-5).reverse();
      dashboard.activeNotices = notices.filter(n => n.Status === 'Active').slice(-5).reverse();

    } else if (session.role === 'Student') {
      const studentObj = students.find(s => s.UserID === session.userId);
      const studentId = studentObj ? studentObj.StudentID : '';

      const studentBills = bills.filter(b => b.StudentID === studentId);
      const studentPayments = payments.filter(p => p.StudentID === studentId);
      const studentAttendance = attendance.filter(a => a.StudentID === studentId);
      const studentComplaints = complaints.filter(c => c.StudentID === studentId);

      const totalAttMarked = studentAttendance.length;
      const totalPresent = studentAttendance.filter(a => a.Status === 'Present').length;
      const attPercentage = totalAttMarked > 0 ? ((totalPresent / totalAttMarked) * 100).toFixed(1) : '100.0';

      const latestBill = studentBills.length > 0 ? studentBills[studentBills.length - 1] : null;

      dashboard.studentKpis = {
        attendancePercentage: attPercentage + '%',
        totalBills: studentBills.length,
        currentDue: latestBill ? parseFloat(latestBill.DueAmount || 0) : 0,
        myComplaintsCount: studentComplaints.length
      };

      dashboard.todayMenu = menu.filter(m => m.Date === todayStr && m.Status === 'Active');
      dashboard.myLatestBill = latestBill;
      dashboard.recentPayments = studentPayments.slice(-5).reverse();
      dashboard.activeNotices = notices.filter(n => n.Status === 'Active' && (n.TargetRole === 'All' || n.TargetRole === 'Student')).slice(-5).reverse();
      dashboard.myComplaints = studentComplaints.slice(-5).reverse();
    }

    return apiResponse(true, 'Dashboard metrics retrieved', dashboard);
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 6. STUDENT MANAGEMENT
// ==========================================

function getStudents(token) {
  try {
    authGuard(token, ['Admin', 'Manager']);
    const students = sheetToObjects(getSheet(SHEETS.STUDENTS));
    return apiResponse(true, 'Students retrieved', students.reverse());
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function createStudent(token, data) {
  try {
    const session = authGuard(token, ['Admin']);
    if (!data.name || !data.email || !data.rollNo) {
      return apiResponse(false, 'Name, Email and Roll Number are required.');
    }

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      const ss = getDb();
      const usersSheet = ss.getSheetByName(SHEETS.USERS);
      const studentsSheet = ss.getSheetByName(SHEETS.STUDENTS);

      const users = sheetToObjects(usersSheet);
      if (users.some(u => u.Email.toLowerCase() === data.email.toLowerCase())) {
        return apiResponse(false, 'User with this email already exists.');
      }

      const userId = generateId('USR', SHEETS.USERS, 'UserID');
      const studentId = generateId('STU', SHEETS.STUDENTS, 'StudentID');
      const nowTs = formatDate(new Date(), true);
      const nowStr = formatDate(new Date(), false);
      const passHash = hashPassword(data.password || 'Student@123');

      // Add to Users Sheet
      usersSheet.appendRow([
        userId,
        data.name,
        data.email,
        passHash,
        'Student',
        data.phone || '',
        'Active',
        'https://ui-avatars.com/api/?name=' + encodeURIComponent(data.name) + '&background=2563eb&color=fff',
        nowTs,
        nowTs,
        ''
      ]);

      // Add to Students Sheet
      studentsSheet.appendRow([
        studentId,
        userId,
        data.rollNo,
        data.name,
        data.email,
        data.phone || '',
        data.course || 'B.Tech',
        data.branch || 'CSE',
        data.semester || '1st',
        data.hostel || 'Hostel A',
        data.roomNo || '',
        data.joiningDate || nowStr,
        'Active',
        nowTs,
        nowTs
      ]);

      logActivity(session.userId, session.name, 'CREATE_STUDENT', 'Students', 'Created student: ' + data.name + ' (' + studentId + ')');
      return apiResponse(true, 'Student registered successfully', { studentId: studentId });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function updateStudent(token, data) {
  try {
    const session = authGuard(token, ['Admin']);
    if (!data.studentId) return apiResponse(false, 'StudentID is required.');

    const ss = getDb();
    const studentsSheet = ss.getSheetByName(SHEETS.STUDENTS);
    const usersSheet = ss.getSheetByName(SHEETS.USERS);

    const studentsData = studentsSheet.getDataRange().getValues();
    const sHeaders = studentsData[0];
    const sIdIdx = sHeaders.indexOf('StudentID');
    const uIdIdx = sHeaders.indexOf('UserID');

    let userId = '';
    let rowIndex = -1;

    for (let i = 1; i < studentsData.length; i++) {
      if (studentsData[i][sIdIdx] === data.studentId) {
        rowIndex = i + 1;
        userId = studentsData[i][uIdIdx];
        break;
      }
    }

    if (rowIndex === -1) return apiResponse(false, 'Student not found.');

    const nowTs = formatDate(new Date(), true);

    // Update Students Sheet fields
    studentsSheet.getRange(rowIndex, sHeaders.indexOf('RollNo') + 1).setValue(data.rollNo);
    studentsSheet.getRange(rowIndex, sHeaders.indexOf('Name') + 1).setValue(data.name);
    studentsSheet.getRange(rowIndex, sHeaders.indexOf('Phone') + 1).setValue(data.phone);
    studentsSheet.getRange(rowIndex, sHeaders.indexOf('Course') + 1).setValue(data.course);
    studentsSheet.getRange(rowIndex, sHeaders.indexOf('Branch') + 1).setValue(data.branch);
    studentsSheet.getRange(rowIndex, sHeaders.indexOf('Semester') + 1).setValue(data.semester);
    studentsSheet.getRange(rowIndex, sHeaders.indexOf('Hostel') + 1).setValue(data.hostel);
    studentsSheet.getRange(rowIndex, sHeaders.indexOf('RoomNo') + 1).setValue(data.roomNo);
    studentsSheet.getRange(rowIndex, sHeaders.indexOf('Status') + 1).setValue(data.status);
    studentsSheet.getRange(rowIndex, sHeaders.indexOf('UpdatedAt') + 1).setValue(nowTs);

    // Update Users Sheet
    const usersData = usersSheet.getDataRange().getValues();
    const uHeaders = usersData[0];
    const uUserIdIdx = uHeaders.indexOf('UserID');

    for (let j = 1; j < usersData.length; j++) {
      if (usersData[j][uUserIdIdx] === userId) {
        usersSheet.getRange(j + 1, uHeaders.indexOf('Name') + 1).setValue(data.name);
        usersSheet.getRange(j + 1, uHeaders.indexOf('Phone') + 1).setValue(data.phone);
        usersSheet.getRange(j + 1, uHeaders.indexOf('Status') + 1).setValue(data.status);
        usersSheet.getRange(j + 1, uHeaders.indexOf('UpdatedAt') + 1).setValue(nowTs);
        break;
      }
    }

    logActivity(session.userId, session.name, 'UPDATE_STUDENT', 'Students', 'Updated student: ' + data.name);
    return apiResponse(true, 'Student updated successfully.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 7. MANAGER MANAGEMENT
// ==========================================

function getManagers(token) {
  try {
    authGuard(token, ['Admin']);
    const managers = sheetToObjects(getSheet(SHEETS.MANAGERS));
    return apiResponse(true, 'Managers list', managers.reverse());
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function createManager(token, data) {
  try {
    const session = authGuard(token, ['Admin']);
    if (!data.name || !data.email) return apiResponse(false, 'Name and Email are required.');

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(5000);
      const ss = getDb();
      const usersSheet = ss.getSheetByName(SHEETS.USERS);
      const managersSheet = ss.getSheetByName(SHEETS.MANAGERS);

      const users = sheetToObjects(usersSheet);
      if (users.some(u => u.Email.toLowerCase() === data.email.toLowerCase())) {
        return apiResponse(false, 'Email already exists.');
      }

      const userId = generateId('USR', SHEETS.USERS, 'UserID');
      const managerId = generateId('MGR', SHEETS.MANAGERS, 'ManagerID');
      const nowTs = formatDate(new Date(), true);
      const nowStr = formatDate(new Date(), false);

      usersSheet.appendRow([
        userId,
        data.name,
        data.email,
        hashPassword(data.password || 'Manager@123'),
        'Manager',
        data.phone || '',
        'Active',
        'https://ui-avatars.com/api/?name=' + encodeURIComponent(data.name) + '&background=0f172a&color=fff',
        nowTs,
        nowTs,
        ''
      ]);

      managersSheet.appendRow([
        managerId,
        userId,
        data.name,
        data.email,
        data.phone || '',
        data.messName || 'Central Mess',
        nowStr,
        'Active',
        nowTs,
        nowTs
      ]);

      logActivity(session.userId, session.name, 'CREATE_MANAGER', 'Managers', 'Created manager: ' + data.name);
      return apiResponse(true, 'Manager account created.');
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 8. MENU MANAGEMENT
// ==========================================

function getMenu(token, filters) {
  try {
    authGuard(token, ['Admin', 'Manager', 'Student']);
    let menu = sheetToObjects(getSheet(SHEETS.MENU));
    if (filters) {
      if (filters.date) menu = menu.filter(m => m.Date === filters.date);
      if (filters.mealType) menu = menu.filter(m => m.MealType === filters.mealType);
    }
    return apiResponse(true, 'Menu list', menu.reverse());
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function createMenu(token, data) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    if (!data.date || !data.mealType || !data.items) {
      return apiResponse(false, 'Date, Meal Type and Items are required.');
    }

    const sheet = getSheet(SHEETS.MENU);
    const menuId = generateId('MEN', SHEETS.MENU, 'MenuID');
    const nowTs = formatDate(new Date(), true);

    sheet.appendRow([
      menuId,
      data.date,
      data.mealType,
      data.items,
      data.description || '',
      data.special || 'No',
      'Active',
      session.name,
      nowTs,
      nowTs
    ]);

    logActivity(session.userId, session.name, 'CREATE_MENU', 'Menu', 'Added menu for ' + data.date + ' (' + data.mealType + ')');
    return apiResponse(true, 'Menu added successfully.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function updateMenu(token, data) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    const sheet = getSheet(SHEETS.MENU);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const idIdx = headers.indexOf('MenuID');

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idIdx] === data.menuId) {
        sheet.getRange(i + 1, headers.indexOf('Date') + 1).setValue(data.date);
        sheet.getRange(i + 1, headers.indexOf('MealType') + 1).setValue(data.mealType);
        sheet.getRange(i + 1, headers.indexOf('Items') + 1).setValue(data.items);
        sheet.getRange(i + 1, headers.indexOf('Description') + 1).setValue(data.description);
        sheet.getRange(i + 1, headers.indexOf('Special') + 1).setValue(data.special);
        sheet.getRange(i + 1, headers.indexOf('Status') + 1).setValue(data.status || 'Active');
        sheet.getRange(i + 1, headers.indexOf('UpdatedAt') + 1).setValue(formatDate(new Date(), true));
        logActivity(session.userId, session.name, 'UPDATE_MENU', 'Menu', 'Updated menu ID: ' + data.menuId);
        return apiResponse(true, 'Menu updated successfully.');
      }
    }
    return apiResponse(false, 'Menu record not found.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function deleteMenu(token, menuId) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    const sheet = getSheet(SHEETS.MENU);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf('MenuID');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === menuId) {
        sheet.deleteRow(i + 1);
        logActivity(session.userId, session.name, 'DELETE_MENU', 'Menu', 'Deleted menu ID: ' + menuId);
        return apiResponse(true, 'Menu item deleted.');
      }
    }
    return apiResponse(false, 'Menu not found.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 9. ATTENDANCE MANAGEMENT
// ==========================================

function getAttendance(token, filters) {
  try {
    const session = authGuard(token, ['Admin', 'Manager', 'Student']);
    let list = sheetToObjects(getSheet(SHEETS.ATTENDANCE));

    if (session.role === 'Student') {
      const students = sheetToObjects(getSheet(SHEETS.STUDENTS));
      const sObj = students.find(s => s.UserID === session.userId);
      if (sObj) {
        list = list.filter(a => a.StudentID === sObj.StudentID);
      } else {
        list = [];
      }
    }

    if (filters) {
      if (filters.date) list = list.filter(a => a.Date === filters.date);
      if (filters.mealType) list = list.filter(a => a.MealType === filters.mealType);
      if (filters.studentId) list = list.filter(a => a.StudentID === filters.studentId);
    }

    return apiResponse(true, 'Attendance records', list.reverse());
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function markAttendance(token, records) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    if (!Array.isArray(records) || records.length === 0) {
      return apiResponse(false, 'No attendance records provided.');
    }

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      const sheet = getSheet(SHEETS.ATTENDANCE);
      const existing = sheetToObjects(sheet);
      const nowTs = formatDate(new Date(), true);

      let markedCount = 0;
      records.forEach(rec => {
        if (!rec.studentId || !rec.date || !rec.mealType) return;

        // Check for duplicate Student + Date + MealType
        const dupIndex = existing.findIndex(e => e.StudentID === rec.studentId && e.Date === rec.date && e.MealType === rec.mealType);
        
        if (dupIndex !== -1) {
          // Update existing row
          const rowNum = dupIndex + 2;
          const headers = SCHEMAS[SHEETS.ATTENDANCE];
          sheet.getRange(rowNum, headers.indexOf('Status') + 1).setValue(rec.status || 'Present');
          sheet.getRange(rowNum, headers.indexOf('MarkedBy') + 1).setValue(session.name);
          sheet.getRange(rowNum, headers.indexOf('MarkedAt') + 1).setValue(nowTs);
        } else {
          // Append new
          const attId = generateId('ATT', SHEETS.ATTENDANCE, 'AttendanceID');
          sheet.appendRow([
            attId,
            rec.studentId,
            rec.date,
            rec.mealType,
            rec.status || 'Present',
            session.name,
            nowTs
          ]);
        }
        markedCount++;
      });

      logActivity(session.userId, session.name, 'MARK_ATTENDANCE', 'Attendance', 'Marked attendance for ' + markedCount + ' records');
      return apiResponse(true, 'Attendance saved successfully for ' + markedCount + ' students.');
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 10. BILLING MANAGEMENT
// ==========================================

function getBills(token, filters) {
  try {
    const session = authGuard(token, ['Admin', 'Manager', 'Student']);
    let list = sheetToObjects(getSheet(SHEETS.BILLS));

    if (session.role === 'Student') {
      const students = sheetToObjects(getSheet(SHEETS.STUDENTS));
      const sObj = students.find(s => s.UserID === session.userId);
      if (sObj) {
        list = list.filter(b => b.StudentID === sObj.StudentID);
      } else {
        list = [];
      }
    }

    if (filters) {
      if (filters.month) list = list.filter(b => b.Month === filters.month);
      if (filters.year) list = list.filter(b => b.Year === String(filters.year));
      if (filters.status) list = list.filter(b => b.Status === filters.status);
    }

    return apiResponse(true, 'Bills list', list.reverse());
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function generateMonthlyBills(token, params) {
  try {
    const session = authGuard(token, ['Admin']);
    const { month, year, baseFee, mealRate, extraCharges, fine, discount } = params;

    if (!month || !year) return apiResponse(false, 'Month and Year are required.');

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(15000);
      const ss = getDb();
      const students = sheetToObjects(ss.getSheetByName(SHEETS.STUDENTS)).filter(s => s.Status === 'Active');
      const attendance = sheetToObjects(ss.getSheetByName(SHEETS.ATTENDANCE));
      const billsSheet = ss.getSheetByName(SHEETS.BILLS);
      const existingBills = sheetToObjects(billsSheet);

      let createdCount = 0;
      const nowTs = formatDate(new Date(), true);
      const dueDate = formatDate(new Date(Date.now() + 10 * 24 * 60 * 60 * 1000), false);

      students.forEach(student => {
        // Prevent duplicate bill generation for same Student + Month + Year
        const alreadyBilled = existingBills.some(b => b.StudentID === student.StudentID && b.Month === month && String(b.Year) === String(year));
        if (alreadyBilled) return;

        // Calculate meal count for that month/year
        const studentMeals = attendance.filter(a => {
          if (a.StudentID !== student.StudentID || a.Status !== 'Present') return false;
          const aDate = new Date(a.Date);
          const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
          return monthNames[aDate.getMonth()] === month && String(aDate.getFullYear()) === String(year);
        }).length;

        const base = parseFloat(baseFee || 2500);
        const rate = parseFloat(mealRate || 40);
        const mealCharges = studentMeals * rate;
        const extra = parseFloat(extraCharges || 0);
        const fn = parseFloat(fine || 0);
        const disc = parseFloat(discount || 0);

        const totalAmount = base + mealCharges + extra + fn - disc;
        const billId = generateId('BIL', SHEETS.BILLS, 'BillID');

        billsSheet.appendRow([
          billId,
          student.StudentID,
          month,
          String(year),
          base.toFixed(2),
          mealCharges.toFixed(2),
          extra.toFixed(2),
          fn.toFixed(2),
          disc.toFixed(2),
          totalAmount.toFixed(2),
          '0.00',
          totalAmount.toFixed(2),
          'Pending',
          dueDate,
          nowTs,
          nowTs
        ]);
        createdCount++;
      });

      logActivity(session.userId, session.name, 'GENERATE_BILLS', 'Billing', 'Generated ' + createdCount + ' bills for ' + month + ' ' + year);
      return apiResponse(true, 'Generated ' + createdCount + ' monthly bills successfully.');
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 11. PAYMENT MANAGEMENT
// ==========================================

function getPayments(token, filters) {
  try {
    const session = authGuard(token, ['Admin', 'Manager', 'Student']);
    let list = sheetToObjects(getSheet(SHEETS.PAYMENTS));

    if (session.role === 'Student') {
      const students = sheetToObjects(getSheet(SHEETS.STUDENTS));
      const sObj = students.find(s => s.UserID === session.userId);
      if (sObj) {
        list = list.filter(p => p.StudentID === sObj.StudentID);
      } else {
        list = [];
      }
    }

    if (filters && filters.billId) list = list.filter(p => p.BillID === filters.billId);

    return apiResponse(true, 'Payments list', list.reverse());
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function recordPayment(token, data) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    const { billId, amount, paymentMethod, transactionId, notes } = data;

    if (!billId || !amount || parseFloat(amount) <= 0) {
      return apiResponse(false, 'Valid Bill ID and Amount are required.');
    }

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      const ss = getDb();
      const billsSheet = ss.getSheetByName(SHEETS.BILLS);
      const paymentsSheet = ss.getSheetByName(SHEETS.PAYMENTS);

      const billsData = billsSheet.getDataRange().getValues();
      const headers = billsData[0];
      const billIdIdx = headers.indexOf('BillID');
      const studentIdIdx = headers.indexOf('StudentID');
      const totalIdx = headers.indexOf('TotalAmount');
      const paidIdx = headers.indexOf('PaidAmount');
      const dueIdx = headers.indexOf('DueAmount');
      const statusIdx = headers.indexOf('Status');
      const updatedIdx = headers.indexOf('UpdatedAt');

      let billRow = -1;
      let studentId = '';
      let currentTotal = 0;
      let currentPaid = 0;

      for (let i = 1; i < billsData.length; i++) {
        if (billsData[i][billIdIdx] === billId) {
          billRow = i + 1;
          studentId = billsData[i][studentIdIdx];
          currentTotal = parseFloat(billsData[i][totalIdx]) || 0;
          currentPaid = parseFloat(billsData[i][paidIdx]) || 0;
          break;
        }
      }

      if (billRow === -1) return apiResponse(false, 'Bill not found.');

      const payAmt = parseFloat(amount);
      const newPaid = currentPaid + payAmt;
      const newDue = Math.max(0, currentTotal - newPaid);
      let newStatus = 'Pending';

      if (newDue <= 0) {
        newStatus = 'Paid';
      } else if (newPaid > 0) {
        newStatus = 'Partial';
      }

      // Update Bill Sheet
      billsSheet.getRange(billRow, paidIdx + 1).setValue(newPaid.toFixed(2));
      billsSheet.getRange(billRow, dueIdx + 1).setValue(newDue.toFixed(2));
      billsSheet.getRange(billRow, statusIdx + 1).setValue(newStatus);
      billsSheet.getRange(billRow, updatedIdx + 1).setValue(formatDate(new Date(), true));

      // Append Payment
      const payId = generateId('PAY', SHEETS.PAYMENTS, 'PaymentID');
      const nowStr = formatDate(new Date(), false);
      const nowTs = formatDate(new Date(), true);

      paymentsSheet.appendRow([
        payId,
        billId,
        studentId,
        payAmt.toFixed(2),
        paymentMethod || 'Cash',
        transactionId || 'TXN-' + Date.now(),
        nowStr,
        'Paid',
        session.name,
        notes || '',
        nowTs
      ]);

      logActivity(session.userId, session.name, 'RECORD_PAYMENT', 'Payments', 'Recorded payment of ₹' + payAmt + ' for Bill ID: ' + billId);
      return apiResponse(true, 'Payment recorded successfully.', { paymentId: payId, newDue: newDue, status: newStatus });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 12. INVENTORY MANAGEMENT
// ==========================================

function getInventory(token) {
  try {
    authGuard(token, ['Admin', 'Manager']);
    const list = sheetToObjects(getSheet(SHEETS.INVENTORY));
    return apiResponse(true, 'Inventory list', list.reverse());
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function createInventoryItem(token, data) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    if (!data.itemName || !data.category) return apiResponse(false, 'Item name and category are required.');

    const sheet = getSheet(SHEETS.INVENTORY);
    const itemId = generateId('INV', SHEETS.INVENTORY, 'ItemID');
    const nowTs = formatDate(new Date(), true);
    const qty = parseFloat(data.quantity) || 0;
    const minQty = parseFloat(data.minimumStock) || 10;
    let status = 'In Stock';
    if (qty <= 0) status = 'Out of Stock';
    else if (qty <= minQty) status = 'Low Stock';

    sheet.appendRow([
      itemId,
      data.itemName,
      data.category,
      qty,
      data.unit || 'Kg',
      minQty,
      parseFloat(data.purchasePrice || 0).toFixed(2),
      data.supplier || '',
      data.expiryDate || '',
      status,
      nowTs,
      nowTs
    ]);

    logActivity(session.userId, session.name, 'CREATE_INVENTORY', 'Inventory', 'Added item: ' + data.itemName);
    return apiResponse(true, 'Inventory item added successfully.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function addInventoryTransaction(token, data) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    const { itemId, type, quantity, unitPrice, supplier, description } = data;

    if (!itemId || !type || !quantity || parseFloat(quantity) <= 0) {
      return apiResponse(false, 'Item ID, Type (IN/OUT) and valid quantity are required.');
    }

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      const ss = getDb();
      const invSheet = ss.getSheetByName(SHEETS.INVENTORY);
      const trxSheet = ss.getSheetByName(SHEETS.INVENTORY_TRANSACTIONS);

      const invData = invSheet.getDataRange().getValues();
      const headers = invData[0];
      const idIdx = headers.indexOf('ItemID');
      const qtyIdx = headers.indexOf('Quantity');
      const minIdx = headers.indexOf('MinimumStock');
      const statusIdx = headers.indexOf('Status');
      const updatedIdx = headers.indexOf('LastUpdated');

      let itemRow = -1;
      let currQty = 0;
      let minStock = 10;

      for (let i = 1; i < invData.length; i++) {
        if (invData[i][idIdx] === itemId) {
          itemRow = i + 1;
          currQty = parseFloat(invData[i][qtyIdx]) || 0;
          minStock = parseFloat(invData[i][minIdx]) || 10;
          break;
        }
      }

      if (itemRow === -1) return apiResponse(false, 'Inventory item not found.');

      const trxQty = parseFloat(quantity);
      let newQty = currQty;

      if (type === 'IN') {
        newQty += trxQty;
      } else if (type === 'OUT') {
        if (trxQty > currQty) {
          return apiResponse(false, 'Insufficient stock. Available: ' + currQty);
        }
        newQty -= trxQty;
      } else {
        return apiResponse(false, 'Invalid transaction type.');
      }

      let newStatus = 'In Stock';
      if (newQty <= 0) newStatus = 'Out of Stock';
      else if (newQty <= minStock) newStatus = 'Low Stock';

      const nowTs = formatDate(new Date(), true);
      const nowStr = formatDate(new Date(), false);

      // Update Inventory item
      invSheet.getRange(itemRow, qtyIdx + 1).setValue(newQty);
      invSheet.getRange(itemRow, statusIdx + 1).setValue(newStatus);
      invSheet.getRange(itemRow, updatedIdx + 1).setValue(nowTs);

      // Log Transaction
      const trxId = generateId('TRX', SHEETS.INVENTORY_TRANSACTIONS, 'TransactionID');
      trxSheet.appendRow([
        trxId,
        itemId,
        type,
        trxQty,
        parseFloat(unitPrice || 0).toFixed(2),
        supplier || '',
        nowStr,
        description || '',
        session.name,
        nowTs
      ]);

      logActivity(session.userId, session.name, 'INVENTORY_TRANSACTION', 'Inventory', type + ' transaction for Item ID: ' + itemId + ' (Qty: ' + trxQty + ')');
      return apiResponse(true, 'Stock transaction recorded successfully.', { newQuantity: newQty, status: newStatus });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 13. EXPENSES MANAGEMENT
// ==========================================

function getExpenses(token) {
  try {
    authGuard(token, ['Admin', 'Manager']);
    const list = sheetToObjects(getSheet(SHEETS.EXPENSES));
    return apiResponse(true, 'Expenses list', list.reverse());
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function createExpense(token, data) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    if (!data.category || !data.amount) return apiResponse(false, 'Category and Amount are required.');

    const sheet = getSheet(SHEETS.EXPENSES);
    const expId = generateId('EXP', SHEETS.EXPENSES, 'ExpenseID');
    const nowTs = formatDate(new Date(), true);
    const nowStr = formatDate(new Date(), false);

    sheet.appendRow([
      expId,
      data.category,
      data.description || '',
      parseFloat(data.amount).toFixed(2),
      data.date || nowStr,
      data.paymentMethod || 'Cash',
      data.receiptURL || '',
      session.name,
      nowTs
    ]);

    logActivity(session.userId, session.name, 'CREATE_EXPENSE', 'Expenses', 'Added expense: ₹' + data.amount + ' (' + data.category + ')');
    return apiResponse(true, 'Expense recorded successfully.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 14. COMPLAINTS MANAGEMENT
// ==========================================

function getComplaints(token, filters) {
  try {
    const session = authGuard(token, ['Admin', 'Manager', 'Student']);
    let list = sheetToObjects(getSheet(SHEETS.COMPLAINTS));

    if (session.role === 'Student') {
      const students = sheetToObjects(getSheet(SHEETS.STUDENTS));
      const sObj = students.find(s => s.UserID === session.userId);
      if (sObj) list = list.filter(c => c.StudentID === sObj.StudentID);
      else list = [];
    }

    if (filters && filters.status) list = list.filter(c => c.Status === filters.status);

    return apiResponse(true, 'Complaints list', list.reverse());
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function createComplaint(token, data) {
  try {
    const session = authGuard(token, ['Student']);
    const students = sheetToObjects(getSheet(SHEETS.STUDENTS));
    const sObj = students.find(s => s.UserID === session.userId);

    if (!sObj) return apiResponse(false, 'Student profile not found.');
    if (!data.title || !data.category) return apiResponse(false, 'Title and Category are required.');

    const sheet = getSheet(SHEETS.COMPLAINTS);
    const cmpId = generateId('CMP', SHEETS.COMPLAINTS, 'ComplaintID');
    const nowTs = formatDate(new Date(), true);

    sheet.appendRow([
      cmpId,
      sObj.StudentID,
      data.category,
      data.title,
      data.description || '',
      data.priority || 'Medium',
      'Pending',
      '',
      nowTs,
      '',
      ''
    ]);

    logActivity(session.userId, session.name, 'CREATE_COMPLAINT', 'Complaints', 'Logged complaint: ' + data.title);
    return apiResponse(true, 'Complaint submitted successfully.', { complaintId: cmpId });
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function resolveComplaint(token, data) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    const { complaintId, status, response } = data;

    if (!complaintId || !status) return apiResponse(false, 'Complaint ID and Status are required.');

    const sheet = getSheet(SHEETS.COMPLAINTS);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const idIdx = headers.indexOf('ComplaintID');

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idIdx] === complaintId) {
        const nowTs = formatDate(new Date(), true);
        sheet.getRange(i + 1, headers.indexOf('Status') + 1).setValue(status);
        sheet.getRange(i + 1, headers.indexOf('Response') + 1).setValue(response || '');
        sheet.getRange(i + 1, headers.indexOf('ResolvedAt') + 1).setValue(nowTs);
        sheet.getRange(i + 1, headers.indexOf('ResolvedBy') + 1).setValue(session.name);

        logActivity(session.userId, session.name, 'RESOLVE_COMPLAINT', 'Complaints', 'Updated complaint ID ' + complaintId + ' to ' + status);
        return apiResponse(true, 'Complaint updated successfully.');
      }
    }
    return apiResponse(false, 'Complaint not found.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 15. NOTICES MANAGEMENT
// ==========================================

function getNotices(token) {
  try {
    const session = authGuard(token, ['Admin', 'Manager', 'Student']);
    let list = sheetToObjects(getSheet(SHEETS.NOTICES));

    if (session.role === 'Student') {
      const todayStr = formatDate(new Date(), false);
      list = list.filter(n => n.Status === 'Active' && (n.TargetRole === 'All' || n.TargetRole === 'Student') && (!n.ExpiryDate || n.ExpiryDate >= todayStr));
    }

    return apiResponse(true, 'Notices list', list.reverse());
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function createNotice(token, data) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    if (!data.title || !data.description) return apiResponse(false, 'Title and Description are required.');

    const sheet = getSheet(SHEETS.NOTICES);
    const noticeId = generateId('NOT', SHEETS.NOTICES, 'NoticeID');
    const nowTs = formatDate(new Date(), true);
    const nowStr = formatDate(new Date(), false);

    sheet.appendRow([
      noticeId,
      data.title,
      data.description,
      data.targetRole || 'All',
      data.priority || 'Medium',
      data.publishDate || nowStr,
      data.expiryDate || '',
      session.name,
      'Active',
      nowTs
    ]);

    logActivity(session.userId, session.name, 'CREATE_NOTICE', 'Notices', 'Published notice: ' + data.title);
    return apiResponse(true, 'Notice published successfully.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function deleteNotice(token, noticeId) {
  try {
    const session = authGuard(token, ['Admin', 'Manager']);
    const sheet = getSheet(SHEETS.NOTICES);
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const idIdx = headers.indexOf('NoticeID');

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][idIdx] === noticeId) {
        sheet.deleteRow(i + 1);
        logActivity(session.userId, session.name, 'DELETE_NOTICE', 'Notices', 'Deleted notice ID: ' + noticeId);
        return apiResponse(true, 'Notice deleted.');
      }
    }
    return apiResponse(false, 'Notice not found.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 16. REPORTS & ANALYTICS
// ==========================================

function getReports(token, reportType, filters) {
  try {
    authGuard(token, ['Admin', 'Manager']);
    const ss = getDb();

    if (reportType === 'RevenueVsExpenses') {
      const payments = sheetToObjects(ss.getSheetByName(SHEETS.PAYMENTS)).filter(p => p.Status === 'Paid');
      const expenses = sheetToObjects(ss.getSheetByName(SHEETS.EXPENSES));

      const monthMap = {};
      payments.forEach(p => {
        const m = p.PaymentDate ? p.PaymentDate.substring(0, 7) : 'Unknown';
        if (!monthMap[m]) monthMap[m] = { revenue: 0, expense: 0 };
        monthMap[m].revenue += (parseFloat(p.Amount) || 0);
      });

      expenses.forEach(e => {
        const m = e.Date ? e.Date.substring(0, 7) : 'Unknown';
        if (!monthMap[m]) monthMap[m] = { revenue: 0, expense: 0 };
        monthMap[m].expense += (parseFloat(e.Amount) || 0);
      });

      const labels = Object.keys(monthMap).sort();
      const revenueData = labels.map(l => monthMap[l].revenue);
      const expenseData = labels.map(l => monthMap[l].expense);

      return apiResponse(true, 'Report generated', { labels: labels, revenue: revenueData, expenses: expenseData });

    } else if (reportType === 'MealDistribution') {
      const attendance = sheetToObjects(ss.getSheetByName(SHEETS.ATTENDANCE)).filter(a => a.Status === 'Present');
      const bCount = attendance.filter(a => a.MealType === 'Breakfast').length;
      const lCount = attendance.filter(a => a.MealType === 'Lunch').length;
      const sCount = attendance.filter(a => a.MealType === 'Snacks').length;
      const dCount = attendance.filter(a => a.MealType === 'Dinner').length;

      return apiResponse(true, 'Meal report generated', {
        labels: ['Breakfast', 'Lunch', 'Snacks', 'Dinner'],
        counts: [bCount, lCount, sCount, dCount]
      });
    }

    return apiResponse(false, 'Unknown report type.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

// ==========================================
// 17. SYSTEM SETTINGS & ACTIVITY LOGS
// ==========================================

function getSettings(token) {
  try {
    authGuard(token, ['Admin', 'Manager', 'Student']);
    const list = sheetToObjects(getSheet(SHEETS.SETTINGS));
    const settingsObj = {};
    list.forEach(s => settingsObj[s.SettingKey] = s.SettingValue);
    return apiResponse(true, 'Settings retrieved', settingsObj);
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function updateSettings(token, settingsMap) {
  try {
    const session = authGuard(token, ['Admin']);
    const sheet = getSheet(SHEETS.SETTINGS);
    const rows = sheet.getDataRange().getValues();
    const nowTs = formatDate(new Date(), true);

    Object.keys(settingsMap).forEach(key => {
      let found = false;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(settingsMap[key]);
          sheet.getRange(i + 1, 4).setValue(nowTs);
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([key, settingsMap[key], 'System Setting', nowTs]);
      }
    });

    logActivity(session.userId, session.name, 'UPDATE_SETTINGS', 'Settings', 'Updated application settings');
    return apiResponse(true, 'Settings updated successfully.');
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}

function getActivityLogs(token) {
  try {
    authGuard(token, ['Admin']);
    const logs = sheetToObjects(getSheet(SHEETS.ACTIVITY_LOG));
    return apiResponse(true, 'Activity logs', logs.reverse().slice(0, 100));
  } catch (err) {
    return apiResponse(false, err.toString());
  }
}
