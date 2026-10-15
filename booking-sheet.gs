/**
 * Booking requests for Laura Husseiny Tutoring.
 * Paste this into a Google Sheet: Extensions → Apps Script. Setup steps are in README.md.
 *
 * Each requested session becomes one row with Status "Requested".
 * Change Status to "Booked" to confirm it, or "Declined" to free the slot again.
 * The website reads Date, Time and Status only; names and emails are never shown to visitors.
 */

const NOTIFY_EMAIL = "husseinylaura@gmail.com";
const SHEET_NAME = "Requests";
const HEADERS = ["Requested on", "Date", "Time", "Status", "Name", "Email", "Session", "Subject", "Where", "Booking", "Message", "Request ID"];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange("D2:D").setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Requested", "Booked", "Declined"], true).build());
  }
  return sh;
}

function asDate_(v) { return v instanceof Date ? Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd") : String(v).trim(); }
function asTime_(v) { return v instanceof Date ? Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "HH:mm") : String(v).trim().padStart(5, "0"); }

// Active (Requested or Booked) slots from today on, keyed "YYYY-MM-DD HH:MM".
function activeSlots_() {
  const today = Utilities.formatDate(new Date(), "Asia/Dubai", "yyyy-MM-dd");
  const rows = sheet_().getDataRange().getValues().slice(1);
  const out = {};
  rows.forEach(r => {
    const status = String(r[3]).trim().toLowerCase();
    if (status !== "requested" && status !== "booked") return;
    const date = asDate_(r[1]), time = asTime_(r[2]);
    if (date >= today) out[date + " " + time] = status;
  });
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// The website calls this to colour the calendar.
function doGet() {
  const slots = activeSlots_();
  return json_(Object.keys(slots).map(k => ({ date: k.slice(0, 10), time: k.slice(11), status: slots[k] })));
}

// The website calls this when a student sends a request.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const slots = (d.slots || []).filter(s => /^\d{4}-\d{2}-\d{2}$/.test(s.date) && /^\d{2}:\d{2}$/.test(s.time)).slice(0, 20);
    if (!slots.length || !d.name || !d.email) return json_({ ok: false });

    const active = activeSlots_();
    const taken = slots.map(s => s.date + " " + s.time).filter(k => active[k]);
    if (taken.length) return json_({ ok: false, taken });

    const id = Utilities.getUuid().slice(0, 8);
    const now = new Date();
    const clip = v => String(v || "").slice(0, 500);
    const sh = sheet_();
    slots.forEach(s => sh.appendRow([now, "'" + s.date, "'" + s.time, "Requested", clip(d.name), clip(d.email),
      clip(d.session), clip(d.subject), clip(d.mode), clip(d.when), clip(d.message), id]));

    const dates = slots.map(s => s.date + " " + s.time).join("\n");
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      replyTo: clip(d.email),
      subject: "New booking request: " + clip(d.name) + " (" + slots.length + (slots.length > 1 ? " sessions)" : " session)"),
      body: [
        "New request from your website.", "",
        "Name: " + clip(d.name), "Email: " + clip(d.email),
        "Booking: " + clip(d.when),
        "Session: " + clip(d.session), "Subject: " + clip(d.subject), "Where: " + clip(d.mode),
        d.message ? "Focus: " + clip(d.message) : "", "",
        "Dates (Dubai time):", dates, "",
        "Open the sheet and change Status to Booked or Declined:",
        SpreadsheetApp.getActiveSpreadsheet().getUrl()
      ].join("\n")
    });
    return json_({ ok: true, id });
  } catch (err) {
    return json_({ ok: false });
  } finally {
    lock.releaseLock();
  }
}
