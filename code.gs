/**
 * Namana Retreat Hotel Management System
 * Google Apps Script Backend - v1.2.0
 * 
 * Deployment Instructions:
 * 1. Open Google Apps Script (script.google.com)
 * 2. Create a new project and paste this code
 * 3. Run the setup() function once to initialize sheets
 * 4. Deploy > New Deployment > Web App
 * 5. Set "Execute as" to "Me" and "Who has access" to "Anyone"
 * 6. Copy the deployment URL and paste it in Admin Settings of the PWA
 * 
 * Endpoints:
 * - /bookings (GET, POST, PUT, DELETE)
 * - /checkins (GET, POST, PUT)
 * - /expenses (GET, POST, PUT, DELETE)
 * - /users (GET, POST - create, POST - login)
 * - /rooms (GET)
 * - /notifications (GET, POST, PUT, DELETE)
 * - /settings (GET, POST, PUT, DELETE)
 * - /pageAccess (GET, POST)
 * - /sync (GET - full data dump)
 */

// ============ CONFIGURATION ============
const SPREADSHEET_ID = ''; // Set your Google Spreadsheet ID here (leave empty to auto-create)
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
  const params = (e && e.parameter) ? e.parameter : {};
  const path = params.path || '';
  let data = {};
  
  // Try to parse JSON from postData
  try {
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.data) {
      // Handle FormData with 'data' field
      data = JSON.parse(e.parameter.data);
    }
  } catch (err) {
    // If JSON parse fails, try to use parameter data directly
    if (e && e.parameter && e.parameter.data) {
      try {
        data = JSON.parse(e.parameter.data);
      } catch (e2) {
        data = {};
      }
    }
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
          version: '1.2.0',
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
    if (baseName === SHEETS.BOOKINGS) {
      sheet.appendRow(['id', 'fullName', 'phone', 'email', 'adults', 'kids', 'checkInDate', 'checkOutDate', 'rooms', 'plan', 'bookingType', 'totalAmount', 'advancePaid', 'pendingAmount', 'idType', 'idNumber', 'status', 'createdAt', 'createdBy', 'guestNote']);
    } else if (baseName === SHEETS.EXPENSES) {
      sheet.appendRow(['id', 'description', 'category', 'amount', 'date', 'createdBy']);
    }
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground('#4F46E5');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
  }
  
  return sheet;
}

// ============ BOOKINGS ============
function handleBookings(method, data, params) {
  const dateStr = data.checkInDate || data.createdAt || params.month || null;
  
  switch (method) {
    case 'GET':
      let bookings;
      if (params.month) {
        const monthSheet = getOrCreateMonthlySheet(SHEETS.BOOKINGS, params.month + '-01');
        bookings = getAllRows(monthSheet);
      } else {
        bookings = getAllBookingsFromAllSheets();
      }
      if (params.id) {
        const found = bookings.find(b => String(b.id) === String(params.id));
        return jsonResponse(found || { error: 'Not found' }, found ? 200 : 404);
      }
      return jsonResponse(bookings);
      
    case 'POST':
      const newBooking = {
        id: data.id || ('g_' + new Date().getTime()),
        fullName: data.fullName || '',
        phone: String(data.phone || ''),
        email: data.email || '',
        adults: parseInt(data.adults) || 2,
        kids: parseInt(data.kids) || 0,
        checkInDate: data.checkInDate || '',
        checkOutDate: data.checkOutDate || '',
        rooms: Array.isArray(data.rooms) ? data.rooms.join(',') : (data.rooms || ''),
        plan: data.plan || 'Room Only',
        bookingType: data.bookingType || 'Direct',
        totalAmount: parseFloat(data.totalAmount) || 0,
        advancePaid: parseFloat(data.advancePaid) || 0,
        pendingAmount: parseFloat(data.pendingAmount) || 0,
        idType: data.idType || '',
        idNumber: data.idNumber || '',
        status: data.status || 'pending',
        createdAt: data.createdAt || new Date().toISOString(),
        createdBy: data.createdBy || '',
        guestNote: data.guestNote || '',
      };
      const targetSheet = getOrCreateMonthlySheet(SHEETS.BOOKINGS, newBooking.createdAt);
      appendRow(targetSheet, newBooking);
      
      // Create notification
      createNotification('New Booking Created', newBooking.fullName + ' booked Room ' + newBooking.rooms, 'booking');
      
      return jsonResponse({ success: true, booking: newBooking }, 201);
      
    case 'PUT':
      if (!data.id) return jsonResponse({ error: 'ID required' }, 400);
      const updated = updateRowInAnySheet(SHEETS.BOOKINGS, 'id', data.id, data);
      return jsonResponse({ success: updated });
      
    case 'DELETE':
      if (!params.id && !data.id) return jsonResponse({ error: 'ID required' }, 400);
      const deleted = deleteRowInAnySheet(SHEETS.BOOKINGS, 'id', params.id || data.id);
      return jsonResponse({ success: deleted });
  }
}

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
  switch (method) {
    case 'GET':
      const allBookings = getAllBookingsFromAllSheets();
      const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const activeCheckins = allBookings.filter(b => 
        b.status === 'checked-in'
      );
      return jsonResponse(activeCheckins);
      
    case 'POST':
      if (!data.id && !data.bookingId) return jsonResponse({ error: 'Booking ID required' }, 400);
      const bookingId = data.bookingId || data.id;
      
      // Update the booking status to checked-in
      const updated = updateRowInAnySheet(SHEETS.BOOKINGS, 'id', bookingId, {
        status: 'checked-in',
        idType: data.idType || '',
        idNumber: data.idNumber || '',
      });
      
      // Create notification
      const booking = getAllBookingsFromAllSheets().find(b => String(b.id) === String(bookingId));
      if (booking) {
        createNotification('Guest Checked In', booking.fullName + ' checked into Room ' + (data.room || booking.rooms), 'checkin');
      }
      
      return jsonResponse({ success: updated });
      
    case 'PUT':
      if (!data.id) return jsonResponse({ error: 'ID required' }, 400);
      updateRowInAnySheet(SHEETS.BOOKINGS, 'id', data.id, data);
      return jsonResponse({ success: true });
  }
}

// ============ EXPENSES ============
function handleExpenses(method, data, params) {
  const dateStr = data.date || params.month || null;
  
  switch (method) {
    case 'GET':
      let expenses;
      if (params.month) {
        const monthSheet = getOrCreateMonthlySheet(SHEETS.EXPENSES, params.month + '-01');
        expenses = getAllRows(monthSheet);
      } else {
        expenses = getAllExpensesFromAllSheets();
      }
      if (params.id) {
        const found = expenses.find(e => String(e.id) === String(params.id));
        return jsonResponse(found || { error: 'Not found' }, found ? 200 : 404);
      }
      return jsonResponse(expenses);
      
    case 'POST':
      const newExpense = {
        id: data.id || ('e_' + new Date().getTime()),
        description: data.description || '',
        category: data.category || 'Miscellaneous',
        amount: parseFloat(data.amount) || 0,
        date: data.date || new Date().toISOString().split('T')[0],
        createdBy: data.createdBy || '',
      };
      const targetSheet = getOrCreateMonthlySheet(SHEETS.EXPENSES, newExpense.date);
      appendRow(targetSheet, newExpense);
      
      // Create notification
      createNotification('Expense Recorded', newExpense.description + ' - ₹' + newExpense.amount, 'expense');
      
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
      const safeUsers = users.map(u => {
        const safe = {};
        for (const key in u) {
          if (key !== 'password') safe[key] = u[key];
        }
        return safe;
      });
      return jsonResponse(safeUsers);
      
    case 'POST':
      // Check if this is a login request or user creation
      if (data.username && data.password && !data.role && !data.id) {
        // Login verification
        const users = getAllRows(sheet);
        const user = users.find(u => u.username === data.username && u.password === data.password);
        if (user) {
          const safe = {};
          for (const key in user) {
            if (key !== 'password') safe[key] = user[key];
          }
          return jsonResponse({ success: true, user: safe });
        }
        return jsonResponse({ error: 'Invalid credentials' }, 401);
      }
      
      // User creation
      if (data.username && data.password && (data.role || data.id)) {
        const users = getAllRows(sheet);
        const existing = users.find(u => u.username === data.username);
        if (existing) {
          return jsonResponse({ error: 'Username already exists' }, 409);
        }
        
        const newUser = {
          id: data.id || ('u_' + new Date().getTime()),
          name: data.name || data.username,
          email: data.email || '',
          role: data.role || 'staff',
          username: data.username,
          password: data.password,
        };
        appendRow(sheet, newUser);
        
        const safe = {};
        for (const key in newUser) {
          if (key !== 'password') safe[key] = newUser[key];
        }
        return jsonResponse({ success: true, user: safe }, 201);
      }
      
      return jsonResponse({ error: 'Username and password required' }, 400);
      
    case 'DELETE':
      if (!params.id && !data.id) return jsonResponse({ error: 'ID required' }, 400);
      deleteRow(sheet, 'id', params.id || data.id);
      return jsonResponse({ success: true });
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
  
  const bookings = getAllBookingsFromAllSheets();
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
      return jsonResponse(notifications.slice(-50).reverse());
      
    case 'POST':
      const newNotif = {
        id: data.id || ('n_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 9)),
        title: data.title || '',
        message: data.message || '',
        type: data.type || 'system',
        read: data.read || false,
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

function createNotification(title, message, type) {
  try {
    const sheet = getSheet(SHEETS.NOTIFICATIONS);
    const notif = {
      id: 'n_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 9),
      title: title,
      message: message,
      type: type,
      read: false,
      timestamp: new Date().toISOString(),
    };
    appendRow(sheet, notif);
  } catch (e) {
    Logger.log('Failed to create notification: ' + e.toString());
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
        if (s.key) settingsObj[s.key] = s.value;
      });
      return jsonResponse(settingsObj);
      
    case 'POST':
    case 'PUT':
      if (!data.key) return jsonResponse({ error: 'Key required' }, 400);
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
      const settings = getAllRows(sheet);
      const pageAccess = {};
      settings.forEach(s => {
        if (s.key && s.key.startsWith('page_access_')) {
          const page = s.key.replace('page_access_', '');
          pageAccess[page] = s.value ? s.value.split(',') : [];
        }
      });
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
  let notifications = [];
  try {
    const notifSheet = getSheet(SHEETS.NOTIFICATIONS);
    notifications = getAllRows(notifSheet).slice(-50).reverse();
  } catch (e) {}
  
  // Get users
  let users = [];
  try {
    const usersSheet = getSheet(SHEETS.USERS);
    users = getAllRows(usersSheet).map(u => {
      const safe = {};
      for (const key in u) {
        if (key !== 'password') safe[key] = u[key];
      }
      return safe;
    });
  } catch (e) {}
  
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
  if (Object.keys(pageAccess).length === 0) {
    pageAccess.dashboard = ['admin', 'manager', 'staff'];
    pageAccess.checkin = ['admin', 'manager', 'staff'];
    pageAccess.bookings = ['admin', 'manager', 'staff'];
    pageAccess.expenses = ['admin', 'manager', 'staff'];
    pageAccess.reports = ['admin', 'manager'];
    pageAccess.settings = ['admin'];
  }
  
  const sharedUrlSetting = settings.find(s => s.key === 'apps_script_url');
  const sharedUrl = sharedUrlSetting ? sharedUrlSetting.value : '';
  
  const ss = getSpreadsheet();
  const monthlySheets = ss.getSheets()
    .map(s => s.getName())
    .filter(n => n.includes('_') && (n.startsWith('Bookings_') || n.startsWith('Expenses_')));
  
  const data = {
    bookings: bookings,
    expenses: expenses,
    rooms: roomsWithStatus,
    notifications: notifications,
    users: users,
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
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    const ss = SpreadsheetApp.create('Namana Retreat Data');
    Logger.log('Created new spreadsheet: ' + ss.getUrl());
    Logger.log('Update SPREADSHEET_ID with: ' + ss.getId());
    return ss;
  }
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    switch (name) {
      case SHEETS.BOOKINGS:
        sheet.appendRow(['id', 'fullName', 'phone', 'email', 'adults', 'kids', 'checkInDate', 'checkOutDate', 'rooms', 'plan', 'bookingType', 'totalAmount', 'advancePaid', 'pendingAmount', 'idType', 'idNumber', 'status', 'createdAt', 'createdBy', 'guestNote']);
        break;
      case SHEETS.EXPENSES:
        sheet.appendRow(['id', 'description', 'category', 'amount', 'date', 'createdBy']);
        break;
      case SHEETS.USERS:
        sheet.appendRow(['id', 'name', 'email', 'role', 'username', 'password']);
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

function jsonResponse(data, status) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============ SETUP FUNCTION ============
function setup() {
  const ss = getSpreadsheet();
  
  [SHEETS.BOOKINGS, SHEETS.EXPENSES, SHEETS.USERS, SHEETS.NOTIFICATIONS, SHEETS.SETTINGS].forEach(name => {
    getSheet(name);
  });
  
  const now = new Date();
  const currentMonth = now.toISOString().split('T')[0];
  getOrCreateMonthlySheet(SHEETS.BOOKINGS, currentMonth);
  getOrCreateMonthlySheet(SHEETS.EXPENSES, currentMonth);
  
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
  
  Logger.log('Setup complete! Spreadsheet ID: ' + ss.getId());
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('Deploy this script as a Web App to get the API URL.');
}