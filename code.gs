/**
 * Namana Retreat Hotel Management System
 * Google Apps Script Backend
 * 
 * Deployment Instructions:
 * 1. Open Google Apps Script (script.google.com)
 * 2. Create a new project and paste this code
 * 3. Deploy > New Deployment > Web App
 * 4. Set "Execute as" to "Me" and "Who has access" to "Anyone"
 * 5. Copy the deployment URL and paste it in the Settings page of the PWA
 * 
 * Endpoints:
 * - /bookings (GET, POST, PUT, DELETE)
 * - /checkins (GET, POST, PUT)
 * - /expenses (GET, POST, PUT, DELETE)
 * - /users (GET)
 * - /rooms (GET)
 */

// ============ CONFIGURATION ============
const SPREADSHEET_ID = ''; // Set your Google Spreadsheet ID here
const SHEETS = {
  BOOKINGS: 'Bookings',
  EXPENSES: 'Expenses',
  USERS: 'Users',
  ROOMS: 'Rooms',
  NOTIFICATIONS: 'Notifications',
  SETTINGS: 'Settings',
};

// ============ MAIN ENTRY POINT ============
function doGet(e) {
  return handleRequest(e || {}, 'GET');
}

function doPost(e) {
  return handleRequest(e || {}, 'POST');
}

function doPut(e) {
  return handleRequest(e || {}, 'PUT');
}

function doDelete(e) {
  return handleRequest(e || {}, 'DELETE');
}

function handleRequest(e, method) {
  // Safely access parameters with defaults
  const params = (e && e.parameter) ? e.parameter : {};
  const path = params.path || '';
  let data = {};
  
  try {
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return jsonResponse({ error: 'Invalid JSON data: ' + err.toString() }, 400);
  }

  try {
    switch (path) {
      case 'bookings':
        return handleBookings(method, data, params);
      case 'checkins':
        return handleCheckins(method, data, params);
      case 'expenses':
        return handleExpenses(method, data, params);
      case 'users':
        return handleUsers(method, data, params);
      case 'rooms':
        return handleRooms(method, data, params);
      case 'notifications':
        return handleNotifications(method, data, params);
      case 'settings':
        return handleSettings(method, data, params);
      case 'pageAccess':
        return handlePageAccess(method, data, params);
      case 'sync':
        return handleSync();
      default:
        return jsonResponse({ 
          status: 'ok', 
          message: 'Namana Retreat API', 
          version: '1.1.0',
          endpoints: ['bookings', 'checkins', 'expenses', 'users', 'rooms', 'notifications', 'settings', 'pageAccess', 'sync'],
          usage: 'Add ?path=endpoint to URL (e.g., ?path=bookings)'
        });
    }
  } catch (err) {
    return jsonResponse({ error: err.toString(), stack: err.stack }, 500);
  }
}

// ============ MONTHLY SHEET HELPERS ============
function getMonthlySheetName(baseName, dateStr) {
  // Format: "Bookings_January_2026" or "Expenses_January_2026"
  if (!dateStr) {
    const now = new Date();
    dateStr = now.toISOString().split('T')[0];
  }
  const date = new Date(dateStr);
  const monthName = Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM_yyyy');
  return baseName + '_' + monthName;
}

function getOrCreateMonthlySheet(baseName, dateStr) {
  const sheetName = getMonthlySheetName(baseName, dateStr);
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // Initialize headers based on base sheet type
    if (baseName === SHEETS.BOOKINGS) {
      sheet.appendRow(['id', 'fullName', 'phone', 'email', 'adults', 'kids', 'checkInDate', 'checkOutDate', 'rooms', 'plan', 'bookingType', 'totalAmount', 'advancePaid', 'pendingAmount', 'idType', 'idNumber', 'status', 'createdAt', 'createdBy', 'guestNote']);
    } else if (baseName === SHEETS.EXPENSES) {
      sheet.appendRow(['id', 'description', 'category', 'amount', 'date', 'createdBy']);
    }
    // Style the header row
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground('#4F46E5');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
  }
  
  return sheet;
}

// ============ BOOKINGS ============
function handleBookings(method, data, params) {
  // Determine which sheet to use based on date
  const dateStr = data.checkInDate || data.createdAt || params.month || null;
  const sheet = dateStr ? getOrCreateMonthlySheet(SHEETS.BOOKINGS, dateStr) : getSheet(SHEETS.BOOKINGS);
  
  switch (method) {
    case 'GET':
      // If month parameter provided, get from monthly sheet
      let bookings;
      if (params.month) {
        const monthSheet = getOrCreateMonthlySheet(SHEETS.BOOKINGS, params.month + '-01');
        bookings = getAllRows(monthSheet);
      } else {
        // Get from all monthly sheets + main sheet
        bookings = getAllBookingsFromAllSheets();
      }
      if (params.id) {
        const found = bookings.find(b => b.id === params.id);
        return jsonResponse(found || { error: 'Not found' }, found ? 200 : 404);
      }
      return jsonResponse(bookings);
      
    case 'POST':
      const newBooking = {
        id: 'g_' + new Date().getTime(),
        fullName: data.fullName || '',
        phone: data.phone || '',
        email: data.email || '',
        adults: data.adults || 2,
        kids: data.kids || 0,
        checkInDate: data.checkInDate || '',
        checkOutDate: data.checkOutDate || '',
        rooms: Array.isArray(data.rooms) ? data.rooms.join(',') : (data.rooms || ''),
        plan: data.plan || 'Room Only',
        bookingType: data.bookingType || 'Direct',
        totalAmount: data.totalAmount || 0,
        advancePaid: data.advancePaid || 0,
        pendingAmount: data.pendingAmount || 0,
        idType: data.idType || '',
        idNumber: data.idNumber || '',
        status: data.status || 'pending',
        createdAt: data.createdAt || new Date().toISOString().split('T')[0],
        createdBy: data.createdBy || '',
        guestNote: data.guestNote || '',
      };
      // Save to monthly sheet based on createdAt date
      const targetSheet = getOrCreateMonthlySheet(SHEETS.BOOKINGS, newBooking.createdAt);
      appendRow(targetSheet, newBooking);
      return jsonResponse({ success: true, booking: newBooking }, 201);
      
    case 'PUT':
      if (!data.id) return jsonResponse({ error: 'ID required' }, 400);
      // Try to find and update in any monthly sheet
      const updated = updateRowInAnySheet(SHEETS.BOOKINGS, 'id', data.id, data);
      return jsonResponse({ success: updated });
      
    case 'DELETE':
      if (!params.id && !data.id) return jsonResponse({ error: 'ID required' }, 400);
      const deleted = deleteRowInAnySheet(SHEETS.BOOKINGS, 'id', params.id || data.id);
      return jsonResponse({ success: deleted });
  }
}

// Helper to get all bookings from all monthly sheets
function getAllBookingsFromAllSheets() {
  const ss = getSpreadsheet();
  const allSheets = ss.getSheets();
  let allBookings = [];
  
  allSheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith(SHEETS.BOOKINGS + '_') || name === SHEETS.BOOKINGS) {
      const rows = getAllRows(sheet);
      allBookings = allBookings.concat(rows);
    }
  });
  
  return allBookings;
}

// Helper to update a row in any monthly sheet
function updateRowInAnySheet(baseName, keyField, keyValue, updates) {
  const ss = getSpreadsheet();
  const allSheets = ss.getSheets();
  
  for (const sheet of allSheets) {
    const name = sheet.getName();
    if (name.startsWith(baseName + '_') || name === baseName) {
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) continue;
      const headers = data[0];
      const keyIndex = headers.indexOf(keyField);
      if (keyIndex === -1) continue;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][keyIndex]) === String(keyValue)) {
          headers.forEach((header, colIndex) => {
            if (updates[header] !== undefined) {
              sheet.getRange(i + 1, colIndex + 1).setValue(updates[header]);
            }
          });
          return true;
        }
      }
    }
  }
  return false;
}

// Helper to delete a row in any monthly sheet
function deleteRowInAnySheet(baseName, keyField, keyValue) {
  const ss = getSpreadsheet();
  const allSheets = ss.getSheets();
  
  for (const sheet of allSheets) {
    const name = sheet.getName();
    if (name.startsWith(baseName + '_') || name === baseName) {
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) continue;
      const headers = data[0];
      const keyIndex = headers.indexOf(keyField);
      if (keyIndex === -1) continue;
      
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][keyIndex]) === String(keyValue)) {
          sheet.deleteRow(i + 1);
          return true;
        }
      }
    }
  }
  return false;
}

// ============ CHECK-INS ============
function handleCheckins(method, data, params) {
  const sheet = getSheet(SHEETS.BOOKINGS);
  
  switch (method) {
    case 'GET':
      const allBookings = getAllRows(sheet);
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const pendingCheckins = allBookings.filter(b => 
        b.status === 'pending' && b.checkInDate <= today
      );
      return jsonResponse(pendingCheckins);
      
    case 'POST':
      if (!data.id) return jsonResponse({ error: 'Booking ID required' }, 400);
      updateRow(sheet, 'id', data.id, {
        status: 'checked-in',
        idType: data.idType || '',
        idNumber: data.idNumber || '',
        guestNote: data.guestNote || '',
      });
      return jsonResponse({ success: true });
      
    case 'PUT':
      if (!data.id) return jsonResponse({ error: 'ID required' }, 400);
      updateRow(sheet, 'id', data.id, data);
      return jsonResponse({ success: true });
  }
}

// ============ EXPENSES ============
function handleExpenses(method, data, params) {
  // Determine which sheet to use based on date
  const dateStr = data.date || params.month || null;
  const sheet = dateStr ? getOrCreateMonthlySheet(SHEETS.EXPENSES, dateStr) : getSheet(SHEETS.EXPENSES);
  
  switch (method) {
    case 'GET':
      let expenses;
      if (params.month) {
        const monthSheet = getOrCreateMonthlySheet(SHEETS.EXPENSES, params.month + '-01');
        expenses = getAllRows(monthSheet);
      } else {
        // Get from all monthly sheets + main sheet
        expenses = getAllExpensesFromAllSheets();
      }
      if (params.id) {
        const found = expenses.find(e => e.id === params.id);
        return jsonResponse(found || { error: 'Not found' }, found ? 200 : 404);
      }
      return jsonResponse(expenses);
      
    case 'POST':
      const newExpense = {
        id: 'e_' + new Date().getTime(),
        description: data.description || '',
        category: data.category || 'Miscellaneous',
        amount: data.amount || 0,
        date: data.date || new Date().toISOString().split('T')[0],
        createdBy: data.createdBy || '',
      };
      // Save to monthly sheet based on date
      const targetSheet = getOrCreateMonthlySheet(SHEETS.EXPENSES, newExpense.date);
      appendRow(targetSheet, newExpense);
      return jsonResponse({ success: true, expense: newExpense }, 201);
      
    case 'PUT':
      if (!data.id) return jsonResponse({ error: 'ID required' }, 400);
      const updated = updateRowInAnySheet(SHEETS.EXPENSES, 'id', data.id, data);
      return jsonResponse({ success: updated });
      
    case 'DELETE':
      if (!params.id && !data.id) return jsonResponse({ error: 'ID required' }, 400);
      const deleted = deleteRowInAnySheet(SHEETS.EXPENSES, 'id', params.id || data.id);
      return jsonResponse({ success: deleted });
  }
}

// Helper to get all expenses from all monthly sheets
function getAllExpensesFromAllSheets() {
  const ss = getSpreadsheet();
  const allSheets = ss.getSheets();
  let allExpenses = [];
  
  allSheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith(SHEETS.EXPENSES + '_') || name === SHEETS.EXPENSES) {
      const rows = getAllRows(sheet);
      allExpenses = allExpenses.concat(rows);
    }
  });
  
  return allExpenses;
}

// ============ USERS ============
function handleUsers(method, data, params) {
  const sheet = getSheet(SHEETS.USERS);
  
  switch (method) {
    case 'GET':
      const users = getAllRows(sheet);
      // Don't return passwords in GET
      const safeUsers = users.map(u => {
        const { password, ...rest } = u;
        return rest;
      });
      return jsonResponse(safeUsers);
      
    case 'POST':
      // Login verification
      if (data.username && data.password) {
        const users = getAllRows(sheet);
        const user = users.find(u => u.username === data.username && u.password === data.password);
        if (user) {
          const { password, ...safeUser } = user;
          return jsonResponse({ success: true, user: safeUser });
        }
        return jsonResponse({ error: 'Invalid credentials' }, 401);
      }
      return jsonResponse({ error: 'Username and password required' }, 400);
  }
}

// ============ ROOMS ============
function handleRooms(method, data, params) {
  const rooms = [
    { id: '101', floor: 'Ground Floor', status: 'available' },
    { id: '102', floor: 'Ground Floor', status: 'available' },
    { id: '103', floor: 'First Floor', status: 'available' },
    { id: '104', floor: 'First Floor', status: 'available' },
  ];
  
  // Check occupancy from bookings
  const bookingSheet = getSheet(SHEETS.BOOKINGS);
  const bookings = getAllRows(bookingSheet);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  const occupiedRooms = bookings
    .filter(b => b.status === 'checked-in' && b.checkInDate <= today && b.checkOutDate > today)
    .flatMap(b => (b.rooms || '').split(','));
  
  const roomsWithStatus = rooms.map(r => ({
    ...r,
    status: occupiedRooms.includes(r.id) ? 'occupied' : 'available',
    occupant: bookings.find(b => b.status === 'checked-in' && (b.rooms || '').split(',').includes(r.id))?.fullName || null,
  }));
  
  return jsonResponse(roomsWithStatus);
}

// ============ NOTIFICATIONS ============
function handleNotifications(method, data, params) {
  const sheet = getSheet(SHEETS.NOTIFICATIONS);
  
  switch (method) {
    case 'GET':
      const notifications = getAllRows(sheet);
      // Return last 50 notifications, newest first
      return jsonResponse(notifications.slice(0, 50));
      
    case 'POST':
      const newNotif = {
        id: data.id || 'n_' + new Date().getTime(),
        title: data.title || '',
        message: data.message || '',
        type: data.type || 'system',
        read: false,
        timestamp: data.timestamp || new Date().toISOString(),
      };
      appendRow(sheet, newNotif);
      return jsonResponse({ success: true, notification: newNotif }, 201);
      
    case 'PUT':
      if (!data.id) return jsonResponse({ error: 'ID required' }, 400);
      updateRow(sheet, 'id', data.id, data);
      return jsonResponse({ success: true });
      
    case 'DELETE':
      if (!params.id && !data.id) return jsonResponse({ error: 'ID required' }, 400);
      deleteRow(sheet, 'id', params.id || data.id);
      return jsonResponse({ success: true });
  }
}

// ============ SETTINGS ============
function handleSettings(method, data, params) {
  const sheet = getSheet(SHEETS.SETTINGS);
  
  switch (method) {
    case 'GET':
      const settings = getAllRows(sheet);
      const settingsObj = {};
      settings.forEach(s => {
        settingsObj[s.key] = s.value;
      });
      return jsonResponse(settingsObj);
      
    case 'POST':
    case 'PUT':
      if (!data.key) return jsonResponse({ error: 'Key required' }, 400);
      // Check if setting exists
      const existing = getAllRows(sheet).find(s => s.key === data.key);
      if (existing) {
        updateRow(sheet, 'key', data.key, { value: data.value });
      } else {
        appendRow(sheet, { key: data.key, value: data.value });
      }
      return jsonResponse({ success: true });
      
    case 'DELETE':
      if (!params.key && !data.key) return jsonResponse({ error: 'Key required' }, 400);
      deleteRow(sheet, 'key', params.key || data.key);
      return jsonResponse({ success: true });
  }
}

// ============ PAGE ACCESS ============
function handlePageAccess(method, data, params) {
  const sheet = getSheet(SHEETS.SETTINGS);
  
  switch (method) {
    case 'GET':
      // Get page access from settings
      const settings = getAllRows(sheet);
      const pageAccess = {};
      settings.forEach(s => {
        if (s.key.startsWith('page_access_')) {
          const page = s.key.replace('page_access_', '');
          pageAccess[page] = s.value ? s.value.split(',') : [];
        }
      });
      // Return default if not set
      if (Object.keys(pageAccess).length === 0) {
        return jsonResponse({
          dashboard: ['admin', 'manager', 'staff'],
          checkin: ['admin', 'manager', 'staff'],
          bookings: ['admin', 'manager', 'staff'],
          expenses: ['admin', 'manager', 'staff'],
          reports: ['admin', 'manager'],
          settings: ['admin'],
        });
      }
      return jsonResponse(pageAccess);
      
    case 'POST':
    case 'PUT':
      if (!data.page || !data.roles) return jsonResponse({ error: 'Page and roles required' }, 400);
      const key = 'page_access_' + data.page;
      const value = Array.isArray(data.roles) ? data.roles.join(',') : data.roles;
      // Check if setting exists
      const existing = getAllRows(sheet).find(s => s.key === key);
      if (existing) {
        updateRow(sheet, 'key', key, { value: value });
      } else {
        appendRow(sheet, { key: key, value: value });
      }
      return jsonResponse({ success: true });
  }
}

// ============ SYNC ============
function handleSync() {
  // Get all data from all monthly sheets
  const bookings = getAllBookingsFromAllSheets();
  const expenses = getAllExpensesFromAllSheets();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  const rooms = [
    { id: '101', floor: 'Ground Floor', status: 'available' },
    { id: '102', floor: 'Ground Floor', status: 'available' },
    { id: '103', floor: 'First Floor', status: 'available' },
    { id: '104', floor: 'First Floor', status: 'available' },
  ];
  
  const occupiedRooms = bookings
    .filter(b => b.status === 'checked-in' && b.checkInDate <= today && b.checkOutDate > today)
    .flatMap(b => (b.rooms || '').split(','));
  
  const roomsWithStatus = rooms.map(r => ({
    ...r,
    status: occupiedRooms.includes(r.id) ? 'occupied' : 'available',
    occupant: bookings.find(b => b.status === 'checked-in' && (b.rooms || '').split(',').includes(r.id))?.fullName || null,
  }));
  
  // Get notifications
  const notifSheet = getSheet(SHEETS.NOTIFICATIONS);
  const notifications = getAllRows(notifSheet).slice(0, 50);
  
  // Get page access
  const settingsSheet = getSheet(SHEETS.SETTINGS);
  const settings = getAllRows(settingsSheet);
  const pageAccess = {};
  settings.forEach(s => {
    if (s.key && s.key.startsWith('page_access_')) {
      const page = s.key.replace('page_access_', '');
      pageAccess[page] = s.value ? s.value.split(',') : [];
    }
  });
  // Default page access if not set
  if (Object.keys(pageAccess).length === 0) {
    pageAccess.dashboard = ['admin', 'manager', 'staff'];
    pageAccess.checkin = ['admin', 'manager', 'staff'];
    pageAccess.bookings = ['admin', 'manager', 'staff'];
    pageAccess.expenses = ['admin', 'manager', 'staff'];
    pageAccess.reports = ['admin', 'manager'];
    pageAccess.settings = ['admin'];
  }
  
  // Get shared URL from settings
  const sharedUrlSetting = settings.find(s => s.key === 'apps_script_url');
  const sharedUrl = sharedUrlSetting ? sharedUrlSetting.value : '';
  
  // List all monthly sheets for reference
  const ss = getSpreadsheet();
  const monthlySheets = ss.getSheets()
    .map(s => s.getName())
    .filter(n => n.includes('_') && (n.startsWith('Bookings_') || n.startsWith('Expenses_')));
  
  const data = {
    bookings: bookings,
    expenses: expenses,
    rooms: roomsWithStatus,
    notifications: notifications,
    pageAccess: pageAccess,
    sharedUrl: sharedUrl,
    monthlySheets: monthlySheets,
    timestamp: new Date().toISOString(),
  };
  return jsonResponse(data);
}

// ============ HELPER FUNCTIONS ============
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  // Try to use active spreadsheet if bound to a spreadsheet
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    // If no active spreadsheet, create a new one
    const ss = SpreadsheetApp.create('Namana Retreat Data');
    Logger.log('Created new spreadsheet: ' + ss.getUrl());
    Logger.log('Update SPREADSHEET_ID in the script with: ' + ss.getId());
    return ss;
  }
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Initialize headers based on sheet type
    switch (name) {
      case SHEETS.BOOKINGS:
        sheet.appendRow(['id', 'fullName', 'phone', 'email', 'adults', 'kids', 'checkInDate', 'checkOutDate', 'rooms', 'plan', 'bookingType', 'totalAmount', 'advancePaid', 'pendingAmount', 'idType', 'idNumber', 'status', 'createdAt', 'createdBy', 'guestNote']);
        break;
      case SHEETS.EXPENSES:
        sheet.appendRow(['id', 'description', 'category', 'amount', 'date', 'createdBy']);
        break;
      case SHEETS.USERS:
        sheet.appendRow(['id', 'name', 'email', 'role', 'username', 'password']);
        // Add default users
        sheet.appendRow(['u1', 'Ashwin', 'ashwin@namana.com', 'admin', 'Ashwin', 'ASH@888']);
        sheet.appendRow(['u2', 'Thousif', 'thousif@namana.com', 'manager', 'Thousif', 'THO@2026']);
        sheet.appendRow(['u3', 'Manager', 'manager@namana.com', 'staff', 'Manager', 'Manage2026']);
        break;
      case SHEETS.NOTIFICATIONS:
        sheet.appendRow(['id', 'title', 'message', 'type', 'read', 'timestamp']);
        break;
      case SHEETS.SETTINGS:
        sheet.appendRow(['key', 'value']);
        break;
    }
    // Style the header row
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground('#4F46E5');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
  }
  return sheet;
}

function getAllRows(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

function appendRow(sheet, data) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
}

function updateRow(sheet, keyField, keyValue, updates) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIndex = headers.indexOf(keyField);
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyIndex]) === String(keyValue)) {
      headers.forEach((header, colIndex) => {
        if (updates[header] !== undefined) {
          sheet.getRange(i + 1, colIndex + 1).setValue(updates[header]);
        }
      });
      break;
    }
  }
}

function deleteRow(sheet, keyField, keyValue) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIndex = headers.indexOf(keyField);
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][keyIndex]) === String(keyValue)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

function jsonResponse(data, status = 200) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============ QUICK TEST ============
// Run this function from the Apps Script editor to verify everything works
function quickTest() {
  Logger.log('Testing API...');
  
  // Test root endpoint
  const root = handleRequest({}, 'GET');
  Logger.log('Root: ' + root.getContent());
  
  // Test bookings endpoint
  const bookings = handleRequest({ parameter: { path: 'bookings' } }, 'GET');
  Logger.log('Bookings: ' + bookings.getContent());
  
  // Test rooms endpoint
  const rooms = handleRequest({ parameter: { path: 'rooms' } }, 'GET');
  Logger.log('Rooms: ' + rooms.getContent());
  
  Logger.log('All tests passed!');
}

// ============ SETUP FUNCTION ============
// Run this once to initialize the spreadsheet
function setup() {
  const ss = getSpreadsheet();
  
  // Create main sheets
  [SHEETS.BOOKINGS, SHEETS.EXPENSES, SHEETS.USERS, SHEETS.ROOMS, SHEETS.NOTIFICATIONS, SHEETS.SETTINGS].forEach(name => {
    getSheet(name);
  });
  
  // Create current month's booking and expense sheets
  const now = new Date();
  const currentMonth = now.toISOString().split('T')[0];
  getOrCreateMonthlySheet(SHEETS.BOOKINGS, currentMonth);
  getOrCreateMonthlySheet(SHEETS.EXPENSES, currentMonth);
  
  // Set default page access
  const settingsSheet = getSheet(SHEETS.SETTINGS);
  const defaultAccess = {
    'page_access_dashboard': 'admin,manager,staff',
    'page_access_checkin': 'admin,manager,staff',
    'page_access_bookings': 'admin,manager,staff',
    'page_access_expenses': 'admin,manager,staff',
    'page_access_reports': 'admin,manager',
    'page_access_settings': 'admin',
  };
  
  Object.entries(defaultAccess).forEach(([key, value]) => {
    const existing = getAllRows(settingsSheet).find(s => s.key === key);
    if (!existing) {
      appendRow(settingsSheet, { key: key, value: value });
    }
  });
  
  // Log success
  Logger.log('Setup complete! Spreadsheet ID: ' + ss.getId());
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('Monthly sheets created for bookings and expenses.');
  Logger.log('Default page access configured.');
  Logger.log('Deploy this script as a Web App to get the API URL.');
}

// ============ TEST FUNCTION ============
function testAPI() {
  // Test creating a booking
  const testBooking = {
    fullName: 'Test Guest',
    phone: '+91 99999 99999',
    email: 'test@example.com',
    adults: 2,
    kids: 1,
    checkInDate: '2026-01-15',
    checkOutDate: '2026-01-18',
    rooms: ['101'],
    plan: 'Including Breakfast',
    totalAmount: 17250,
    advancePaid: 10000,
    pendingAmount: 7250,
    status: 'pending',
    createdBy: 'admin',
    guestNote: 'Test booking - can be deleted',
  };
  
  const result = handleBookings('POST', testBooking, {});
  Logger.log(result.getContent());
}
