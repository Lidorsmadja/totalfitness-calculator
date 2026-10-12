/**
 * Total Fitness · מחשבון מתחם כושר · שרת לידים על Google Sheets
 *
 * התקנה (פעם אחת):
 * 1. פותחים Google Sheet חדש → תוספים (Extensions) → Apps Script.
 * 2. מדביקים את כל הקובץ הזה במקום הקוד הקיים ושומרים.
 * 3. הגדרות הפרויקט (גלגל שיניים) → Script Properties → מוסיפים:
 *      ADMIN_KEY = קוד הכניסה לפאנל האחורי (לבחור קוד חזק)
 * 4. Deploy → New deployment → Web app
 *      Execute as: Me   ·   Who has access: Anyone
 * 5. מעתיקים את כתובת ה-Web app ומדביקים ב-index.html בשדה CONFIG.sheetUrl
 */

const LEADS_SHEET = "לידים";
const CRM_SHEET = "טיפול צוות";
const HEADERS = ["מספר הצעה", "עודכן", "נוצר", "סטטוס", "שם", "טלפון", "סוג מקום", "מ״ר", "פריטים", "תפוסה %", "ריצוף", "מחיר לפני הנחה", "אחרי 30%", "הציוד", "מכשיר", "JSON"];
const STATUS = { details: "השאיר פרטים", quote: "צפה במחיר", offer: "ראה את ההטבה", whatsapp: "לחץ קבל · וואטסאפ" };

function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (body.action === "save") out = saveLead(body.lead);
    else if (body.action === "list") { auth(body.key); out = { ok: true, leads: listLeads(), crm: listCrm() }; }
    else if (body.action === "crm") { auth(body.key); saveCrm(body.id, body.crm); out = { ok: true }; }
    else out = { ok: false, error: "unknown action" };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, service: "total-fitness-calculator" })).setMimeType(ContentService.MimeType.JSON);
}

function auth(key) {
  const real = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
  if (!real || String(key || "") !== real) throw new Error("auth");
}

function sheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sh.setRightToLeft(true);
  }
  return sh;
}

// Cells that start with = + - @ would run as formulas in Sheets; store them as plain text.
function safe(v) {
  const s = v == null ? "" : String(v).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function findRow(sh, id) {
  const hit = sh.getRange("A:A").createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function saveLead(lead) {
  if (!lead || !/^\d{6}$/.test(String(lead.id))) throw new Error("bad lead");
  const json = JSON.stringify(lead);
  if (json.length > 20000) throw new Error("too large");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet(LEADS_SHEET, HEADERS);
    const q = lead.quote || {};
    const items = Object.keys(lead.qty || {}).map(k => k + "×" + lead.qty[k]).join(", ");
    const row = [
      "'" + lead.id,
      new Date(Number(lead.updatedAt) || Date.now()),
      new Date(Number(lead.createdAt) || Date.now()),
      STATUS[lead.status] || safe(lead.status),
      safe(lead.name), "'" + safe(lead.phone).replace(/^'/, ""),
      safe(lead.spaceLabel), Number(lead.area) || 0, Number(lead.devices) || 0, Number(lead.utilization) || 0,
      lead.floor ? "כן" : "לא", Number(q.list) || 0, Number(q.best) || 0,
      safe(items), safe(lead.device), json,
    ];
    const r = findRow(sh, lead.id);
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function listLeads() {
  const sh = sheet(LEADS_SHEET, HEADERS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const start = Math.max(2, sh.getLastRow() - 499);
  const vals = sh.getRange(start, HEADERS.length, sh.getLastRow() - start + 1, 1).getValues();
  return vals.map(v => { try { return JSON.parse(v[0]); } catch (e) { return null; } }).filter(Boolean);
}

function listCrm() {
  const sh = sheet(CRM_SHEET, ["מספר הצעה", "שלב", "הערות", "עודכן", "JSON"]);
  const n = sh.getLastRow() - 1;
  if (n < 1) return {};
  const out = {};
  sh.getRange(2, 1, n, 5).getValues().forEach(r => { try { out[String(r[0]).replace(/^'/, "")] = JSON.parse(r[4]); } catch (e) {} });
  return out;
}

function saveCrm(id, crm) {
  if (!/^\d{6}$/.test(String(id))) throw new Error("bad id");
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet(CRM_SHEET, ["מספר הצעה", "שלב", "הערות", "עודכן", "JSON"]);
    const row = ["'" + id, safe(crm && crm.stage), safe(crm && crm.note), new Date(), JSON.stringify(crm || {})];
    const r = findRow(sh, id);
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
}
