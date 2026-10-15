/**
 * Booking requests for Laura Husseiny Tutoring.
 * Paste this into a Google Sheet (in husseinylaura@gmail.com): Extensions → Apps Script.
 *
 * How it works:
 *  - A student requests a slot on the website → a row appears here with Status "Requested",
 *    the slot shows as Requested on the website, and you get an email.
 *  - Change Status to "Approved" → the student gets a confirmation email from you,
 *    and the slot shows as Booked on the website.
 *  - Change Status to "Declined" → the student gets a polite email, and the slot is free again.
 *
 * ONE-TIME SETUP: after pasting, choose "setup" in the function menu at the top and click Run.
 * If an email ever doesn't go out, use the Bookings menu in the sheet → "Send confirmation emails now".
 * The website only ever reads Date, Time and Status; names and emails are never shown to visitors.
 */

const MY_EMAIL = "husseinylaura@gmail.com";
const MY_NAME = "Laura Husseiny";
const MY_WHATSAPP = "+971 55 292 8764";
const SHEET_NAME = "Requests";
const HEADERS = ["Requested on", "Date", "Time", "Status", "Name", "Email", "Session", "Subject", "Where",
                 "Booking", "Message", "Request ID", "Student emailed"];
const COL = { date: 2, time: 3, status: 4, name: 5, email: 6, session: 7, subject: 8, mode: 9, booking: 10, message: 11, id: 12, emailed: 13 };

/* ---------- one-time setup: run this once from the Apps Script editor ---------- */
function setup() {
  sheet_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "onStatusChange") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("onStatusChange").forSpreadsheet(ss).onEdit().create();
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  sh.getRange("D2:D").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Requested", "Approved", "Declined"], true).build());
  return sh;
}

const tz_ = () => SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
function asDate_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), "yyyy-MM-dd") : String(v).trim(); }
function asTime_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), "HH:mm") : String(v).trim().padStart(5, "0"); }
function niceDate_(ymd, time) {
  const [y, m, d] = ymd.split("-").map(Number);
  return Utilities.formatDate(new Date(y, m - 1, d, 12), tz_(), "EEEE d MMMM") + " at " + time;
}

// Requested or Approved slots from today on, keyed "YYYY-MM-DD HH:MM" → "requested" | "booked".
function activeSlots_() {
  const today = Utilities.formatDate(new Date(), "Asia/Dubai", "yyyy-MM-dd");
  const out = {};
  sheet_().getDataRange().getValues().slice(1).forEach(r => {
    const raw = String(r[COL.status - 1]).trim().toLowerCase();
    const kind = raw === "requested" ? "requested" : normStatus_(raw) === "Approved" ? "booked" : null;
    if (!kind) return;
    const date = asDate_(r[COL.date - 1]), time = asTime_(r[COL.time - 1]);
    if (date >= today) out[date + " " + time] = kind;
  });
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- the website reads the calendar from here ---------- */
function doGet() {
  const slots = activeSlots_();
  return json_(Object.keys(slots).map(k => ({ date: k.slice(0, 10), time: k.slice(11), status: slots[k] })));
}

/* ---------- the website sends requests here ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    d.email = String(d.email || "").trim(); d.name = String(d.name || "").trim();
    const slots = (d.slots || []).filter(s => /^\d{4}-\d{2}-\d{2}$/.test(s.date) && /^\d{2}:\d{2}$/.test(s.time)).slice(0, 20);
    if (!slots.length || !d.name || !/^\S+@\S+\.\S+$/.test(d.email || "")) return json_({ ok: false });

    const active = activeSlots_();
    const taken = slots.map(s => s.date + " " + s.time).filter(k => active[k]);
    if (taken.length) return json_({ ok: false, taken });

    const id = Utilities.getUuid().slice(0, 8);
    const now = new Date();
    const clip = v => String(v || "").slice(0, 500);
    const sh = sheet_();
    slots.forEach(s => sh.appendRow([now, "'" + s.date, "'" + s.time, "Requested", clip(d.name), clip(d.email),
      clip(d.session), clip(d.subject), clip(d.mode), clip(d.when), clip(d.message), id, ""]));

    MailApp.sendEmail({
      to: MY_EMAIL,
      replyTo: clip(d.email),
      subject: "New booking request: " + clip(d.name) + " (" + slots.length + (slots.length > 1 ? " sessions)" : " session)"),
      body: [
        "New request from your website.", "",
        "Name: " + clip(d.name), "Email: " + clip(d.email),
        "Booking: " + clip(d.when),
        "Session: " + clip(d.session), "Subject: " + clip(d.subject), "Where: " + clip(d.mode),
        d.message ? "Focus: " + clip(d.message) : "", "",
        "Dates (Dubai time):", slots.map(s => niceDate_(s.date, s.time)).join("\n"), "",
        "To approve or decline, change Status in the sheet:",
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

/* ---------- when you change Status, email the student ---------- */
function onStatusChange(e) {
  const range = e && e.range;
  if (!range || range.getSheet().getName() !== SHEET_NAME) return;
  const first = range.getColumn(), last = first + range.getNumColumns() - 1;
  if (COL.status < first || COL.status > last) return;

  const sh = range.getSheet();
  const top = Math.max(range.getRow(), 2), bottom = range.getRow() + range.getNumRows() - 1;
  if (bottom < top) return;
  emailStudents_(sh, top, sh.getRange(top, 1, bottom - top + 1, HEADERS.length).getValues());
}

/* ---------- Bookings menu in the sheet ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Bookings")
    .addItem("Send confirmation emails now", "sendPendingEmails")
    .addItem("Switch on automatic emails", "setup")
    .addToUi();
}

// Emails every Approved / Declined student who hasn't been emailed yet.
function sendPendingEmails() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const sent = emailStudents_(sh, 2, sh.getRange(2, 1, n, HEADERS.length).getValues());
  try { SpreadsheetApp.getActive().toast(sent ? sent + " email(s) sent" : "Nothing waiting to be sent", "Bookings"); } catch (err) {}
}

function normStatus_(v) {
  const s = String(v).trim().toLowerCase();
  if (s === "approved" || s === "approve" || s === "booked" || s === "confirmed") return "Approved";
  if (s === "declined" || s === "decline" || s === "rejected") return "Declined";
  return "";
}

function emailStudents_(sh, top, rows) {
  // group newly approved / declined rows by request, so a repeating booking gets one email
  const groups = {};
  rows.forEach((r, i) => {
    const status = normStatus_(r[COL.status - 1]);
    if (!status) return;
    if (!/^\S+@\S+\.\S+$/.test(String(r[COL.email - 1]).trim())) return;
    if (String(r[COL.emailed - 1]).indexOf(status) === 0) return; // already emailed for this status
    const key = r[COL.id - 1] + "|" + status;
    (groups[key] = groups[key] || { status, row: r, items: [] }).items.push({ rowNum: top + i, when: niceDate_(asDate_(r[COL.date - 1]), asTime_(r[COL.time - 1])) });
  });

  Object.keys(groups).forEach(k => {
    const g = groups[k], r = g.row;
    const firstName = String(r[COL.name - 1]).split(" ")[0] || "there";
    const list = g.items.map(x => "  • " + x.when).join("\n");
    const online = String(r[COL.mode - 1]).toLowerCase().indexOf("online") >= 0;
    const body = g.status === "Approved" ? [
        "Hi " + firstName + ",", "",
        "Your tutoring " + (g.items.length > 1 ? "sessions are" : "session is") + " confirmed:", "",
        list, "",
        "Session: " + r[COL.session - 1], "Subject: " + r[COL.subject - 1], "Where: " + r[COL.mode - 1], "",
        online ? "I'll send you the meeting link before we start." : "I'll message you the address details before we start.",
        "If you need to change anything, just reply to this email or message me on WhatsApp (" + MY_WHATSAPP + ").", "",
        "See you soon,", MY_NAME
      ] : [
        "Hi " + firstName + ",", "",
        "Thank you for your request. Unfortunately I can't take the following " + (g.items.length > 1 ? "times" : "time") + ":", "",
        list, "",
        "Please choose another slot on my website, or reply to this email and we'll find a time that works.", "",
        "Best wishes,", MY_NAME
      ];
    MailApp.sendEmail({
      to: String(r[COL.email - 1]).trim(),
      replyTo: MY_EMAIL,
      name: MY_NAME,
      subject: g.status === "Approved" ? "Your tutoring session is confirmed" : "About your tutoring request",
      body: body.join("\n")
    });
    const stamp = g.status + " email sent " + Utilities.formatDate(new Date(), "Asia/Dubai", "d MMM HH:mm");
    g.items.forEach(x => sh.getRange(x.rowNum, COL.emailed).setValue(stamp));
  });
  return Object.keys(groups).length;
}
