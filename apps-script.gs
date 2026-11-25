// Paste into a Google Sheet: Extensions > Apps Script. Deploy > New deployment > Web app,
// Execute as: Me, Who has access: Anyone. Paste the /exec URL into the app's settings.
// shortcut: the secret /exec URL is the only access control; fine for personal stock data, add a key check if you share the sheet.
const COLS = ['id','brand','model','colour','sizeEU','sizeUS','sku','ean','boxed','cond','buy','list','listedOn','status','loc','platform','buyer','soldPrice','payout','img','notes','soldAt','updatedAt'];

function doGet() {
  return out_(file_().getBlob().getDataAsString() || 'null');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const { base, db } = JSON.parse(e.postData.contents);
    if (!db || !Array.isArray(db.pairs)) return out_('bad payload');
    const f = file_();
    const cur = JSON.parse(f.getBlob().getDataAsString() || 'null');
    if (cur && cur.savedAt !== base) return out_('conflict'); // someone saved since this client last synced
    f.setContent(JSON.stringify(db));
    writeSheet_(db);
    return out_('ok');
  } finally {
    lock.releaseLock();
  }
}

// Full data (incl. photos) lives in a Drive JSON file; the sheet is a readable copy for you and Make.com.
function file_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('fileId');
  if (id) return DriveApp.getFileById(id);
  const f = DriveApp.createFile('sneaker-inventory.json', 'null');
  props.setProperty('fileId', f.getId());
  return f;
}

function writeSheet_(db) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Inventory') || ss.insertSheet('Inventory');
  const rows = [COLS].concat(db.pairs.map(p => COLS.map(k =>
    /At$/.test(k) && p[k] ? new Date(p[k]) : Array.isArray(p[k]) ? p[k].join(' | ') : (p[k] == null ? '' : p[k]))));
  sh.clearContents();
  sh.getRange(1, 1, rows.length, COLS.length).setValues(rows);
}

function out_(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT);
}
