// Change this text if your expense heading in the sheet is different
var EXPENSE_TITLE = 'Monthly Operation';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Property & Tenant Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Dates can't be sent to the browser, so convert them to text.
function clean_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

// Turn "$1,200.50" or 1200.5 into a number (or null if not a number)
function toNum_(v) {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    var t = v.replace(/[\$,\s]/g, '');
    if (t !== '' && /^-?\d*\.?\d+$/.test(t)) return parseFloat(t);
  }
  return null;
}

// Finds the "Monthly Operation" table anywhere in the sheet
function findExpenses_(data) {
  var key = EXPENSE_TITLE.toLowerCase();
  var hr = -1, hc = -1;

  for (var r = 0; r < data.length && hr < 0; r++) {
    for (var c = 0; c < data[r].length; c++) {
      var v = data[r][c];
      if (typeof v === 'string' && v.toLowerCase().indexOf(key) > -1) {
        hr = r; hc = c; break;
      }
    }
  }

  if (hr < 0) return { found: false, title: EXPENSE_TITLE, items: [], total: 0 };

  var items = [];
  var total = null;
  var blank = 0;

  for (var i = hr + 1; i < data.length && i < hr + 100; i++) {
    var row = data[i];
    var label = String(clean_(row[hc]) === undefined ? '' : clean_(row[hc])).trim();
    var amount = null;
    var note = '';

    for (var c2 = hc + 1; c2 <= hc + 6 && c2 < row.length; c2++) {
      var cell = row[c2];
      var n = toNum_(cell);
      if (amount === null && n !== null) {
        amount = n;
      } else if (amount !== null && typeof cell === 'string' && cell.trim() !== '') {
        note = cell.trim();
        break;
      }
    }

    if (label === '' && amount === null) {
      blank++;
      if (blank >= 2) break;
      continue;
    }
    blank = 0;

    if (label === '' || amount === null) continue; // sub-headings or column titles

    if (/^total/i.test(label)) {
      total = amount;
      break;
    }
    items.push({ label: label, amount: amount, note: note });
  }

  var sum = 0;
  items.forEach(function (it) { sum += it.amount; });

  return {
    found: true,
    title: EXPENSE_TITLE,
    items: items,
    total: total !== null ? total : sum
  };
}

function getDashboardData(selectedSheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNames = ss.getSheets()
    .map(function (s) { return s.getName(); })
    .filter(function (n) { return n !== 'Lists'; });

  var activeSheetName = selectedSheetName || sheetNames[0];
  var sheet = ss.getSheetByName(activeSheetName);
  if (!sheet) throw new Error('Sheet not found: ' + activeSheetName);

  var data = sheet.getDataRange().getValues();
  var rooms = [];
  var overdueRooms = [];

  for (var r = 4; r < data.length; r++) {
    var roomNo = data[r][1];
    if (roomNo && roomNo.toString().trim() !== '') {
      var roomObj = {
        rowIndex: r + 1,
        floor: String(clean_(data[r][0]) || ''),
        room: roomNo.toString(),
        deposit: clean_(data[r][2]) || 0,
        monthlyRent: clean_(data[r][3]) || 0,
        rentStatus: String(clean_(data[r][4]) || ''),
        utilitiesBill: clean_(data[r][5]) || 0,
        utilitiesStatus: String(clean_(data[r][6]) || ''),
        totalCurrentDue: clean_(data[r][7]) || 0,
        overdueMonths: Number(data[r][8]) || 0,
        outstandingDebt: Number(data[r][9]) || 0,
        notes: String(clean_(data[r][10]) || '')
      };
      rooms.push(roomObj);

      if (roomObj.overdueMonths > 0 ||
          roomObj.outstandingDebt > 0 ||
          roomObj.rentStatus.includes('Unpaid')) {
        overdueRooms.push(roomObj);
      }
    }
  }

  var summary = {};
  for (var k = 3; k < data.length; k++) {
    var metricName = data[k][12];
    var metricVal = data[k][13];
    if (metricName && metricVal !== undefined && metricVal !== '') {
      summary[metricName.toString().trim()] = clean_(metricVal);
    }
  }

  return {
    sheetNames: sheetNames,
    activeSheetName: activeSheetName,
    summary: summary,
    rooms: rooms,
    overdueRooms: overdueRooms,
    expenses: findExpenses_(data)
  };
}

function updateRoomData(sheetName, rowIndex, rentStatus, utilitiesStatus, utilitiesBill, notes) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet not found.' };

    rowIndex = Number(rowIndex);
    sheet.getRange(rowIndex, 5).setValue(rentStatus);

    if (utilitiesBill !== '' && utilitiesBill !== null && utilitiesBill !== undefined) {
      sheet.getRange(rowIndex, 6).setValue(parseFloat(utilitiesBill) || 0);
    }

    sheet.getRange(rowIndex, 7).setValue(utilitiesStatus);
    sheet.getRange(rowIndex, 11).setValue(notes);

    return { success: true, message: 'Room updated successfully!' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
