// Google Apps Script backend for the Nakshatra Observatory Billing app.
// Paste this whole file into the Apps Script editor (Extensions > Apps Script)
// attached to your "Bills" Google Sheet, then deploy it as a Web App.
// See the setup guide for full step-by-step instructions.

const SHEET_NAME = 'Bills';
const API_KEY = 'bcadbc33-bd22-4052-b335-a8c28f343273'; // must match API_KEY in the HTML file exactly

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e.parameter.key !== API_KEY) return jsonOut({ error: 'unauthorized' });
  const action = e.parameter.action || 'list';
  if (action === 'list') return jsonOut({ bills: readBills(), _codeVersion: 'fmtCell-fix-v2' });
  return jsonOut({ error: 'unknown action' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ error: 'invalid request body' });
  }
  if (body.key !== API_KEY) return jsonOut({ error: 'unauthorized' });
  if (body.action === 'add') return jsonOut(addBill(body.bill));
  if (body.action === 'delete') return jsonOut(deleteBill(body.billNo));
  return jsonOut({ error: 'unknown action' });
}

// Sheets auto-detects strings that look like dates/times and silently converts the cell to
// a real Date value (anchored to 1899-12-30 for time-only values). This undoes that if it
// ever happens, so the app always gets back the plain string it originally sent.
function fmtCell(val, isTime) {
  // Duck-type instead of `instanceof Date` — values from getValues() in a web-app request
  // don't always pass a strict instanceof check even though they behave like dates.
  if (val === null || val === undefined || typeof val.getTime !== 'function') return val;
  const tz = Session.getScriptTimeZone();
  return isTime ? Utilities.formatDate(val, tz, 'hh:mm a') : Utilities.formatDate(val, tz, 'yyyy-MM-dd');
}

function readBills() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  rows.shift(); // drop header row
  return rows
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      billNo: r[0],
      date: fmtCell(r[1], false),
      time: fmtCell(r[2], true),
      customer: r[3],
      mobile: r[4],
      items: JSON.parse(r[5] || '[]'),
      gstRate: r[6],
      discountRate: r[7],
      discount: r[8],
      subtotal: r[9],
      gst: r[10],
      total: r[11],
      payment: r[12]
    }));
}

function addBill(bill) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // avoids two devices getting the same bill number if they save at the same instant
  try {
    const props = PropertiesService.getScriptProperties();
    const billNo = Number(props.getProperty('nextBillNo') || '1');
    const sheet = getSheet();
    const nextRow = sheet.getLastRow() + 1;
    const rowValues = [
      billNo, bill.date, bill.time, bill.customer, bill.mobile,
      JSON.stringify(bill.items || []), bill.gstRate || 0, bill.discountRate || 0,
      bill.discount || 0, bill.subtotal || 0, bill.gst || 0, bill.total || 0, bill.payment || ''
    ];
    const range = sheet.getRange(nextRow, 1, 1, rowValues.length);
    // Force the Date (B) and Time (C) columns to plain text *before* writing, so Sheets
    // never auto-converts them into a Date value (see fmtCell for why that's a problem).
    // Writing via setValues() on a pre-formatted range (rather than appendRow) makes the
    // text format actually stick.
    sheet.getRange(nextRow, 2, 1, 2).setNumberFormat('@');
    range.setValues([rowValues]);
    props.setProperty('nextBillNo', String(billNo + 1));
    return { billNo };
  } finally {
    lock.releaseLock();
  }
}

function deleteBill(billNo) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(billNo)) {
      sheet.deleteRow(i + 1);
      return { deleted: true };
    }
  }
  return { deleted: false };
}
