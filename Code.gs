// ── 3827 N Greenview Offer Tracker — Google Apps Script ──────────────────────
// Deploy: Execute as = Me, Who has access = Anyone
// After ANY code change, redeploy a NEW VERSION:
//   Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy

function doGet(e) {
  try { return respond(getData()); }
  catch (err) { return respond({ error: err.message }); }
}

// Handles: upsert offer (default), update config, delete offer
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'updateConfig') { setConfig(body.key, body.value); return respond({ ok: true }); }
    if (body.action === 'deleteOffer')  { deleteOffer(body.num);            return respond({ ok: true }); }
    upsertOffer(body); // default action: update offer by num, or append if new
    return respond({ ok: true, num: body.num });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function findSheet(ss, wanted, keyword) {
  const sheets = ss.getSheets();
  for (const s of sheets) if (s.getName().toLowerCase() === wanted.toLowerCase()) return s;
  for (const s of sheets) if (s.getName().toLowerCase().indexOf(keyword) !== -1) return s;
  return null;
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = {};
  const cfgSheet = findSheet(ss, 'Config', 'config');
  if (cfgSheet) cfgSheet.getDataRange().getValues().forEach(function (r) { if (r[0]) config[r[0]] = r[1]; });

  const offers = [];
  const offersSheet = findSheet(ss, 'Offers', 'offer');
  if (offersSheet) {
    const rows = offersSheet.getDataRange().getValues();
    const headers = rows[0];
    rows.slice(1).filter(function (row) { return row[0] !== '' && row[0] !== null; }).forEach(function (row) {
      const o = {};
      headers.forEach(function (h, i) { o[h] = row[i]; });
      o.flags = String(o.flags || '').split('||').filter(Boolean).map(function (f) {
        const c = f.indexOf(':'); return { type: f.slice(0, c).trim(), text: f.slice(c + 1).trim() };
      });
      offers.push(o);
    });
  }
  return { config: config, offers: offers, updatedAt: new Date().toISOString() };
}

// Update an existing offer row (matched by num) in place, or append if new.
// Only the keys provided in `offer` are changed — others are left intact.
function upsertOffer(offer) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSheet(ss, 'Offers', 'offer');
  if (!sheet) throw new Error('No Offers tab found');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (Array.isArray(offer.flags)) {
    offer.flags = offer.flags.map(function (f) { return f.type + ':' + f.text; }).join('||');
  }

  const lastRow = sheet.getLastRow();
  const numCol = headers.indexOf('num');
  let targetRow = -1;
  if (numCol !== -1 && offer.num != null && lastRow > 1) {
    const nums = sheet.getRange(2, numCol + 1, lastRow - 1, 1).getValues().flat();
    for (let i = 0; i < nums.length; i++) {
      if (String(nums[i]) === String(offer.num)) { targetRow = i + 2; break; }
    }
  }

  if (targetRow === -1) {
    // append new
    const row = headers.map(function (h) { return offer[h] !== undefined ? offer[h] : ''; });
    sheet.appendRow(row);
  } else {
    // update only provided fields
    headers.forEach(function (h, i) {
      if (offer[h] !== undefined) sheet.getRange(targetRow, i + 1).setValue(offer[h]);
    });
  }
}

function deleteOffer(num) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSheet(ss, 'Offers', 'offer');
  if (!sheet) throw new Error('No Offers tab found');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const numCol = headers.indexOf('num');
  const lastRow = sheet.getLastRow();
  const nums = sheet.getRange(2, numCol + 1, lastRow - 1, 1).getValues().flat();
  for (let i = 0; i < nums.length; i++) {
    if (String(nums[i]) === String(num)) { sheet.deleteRow(i + 2); return; }
  }
}

function setConfig(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSheet(ss, 'Config', 'config');
  if (!sheet) throw new Error('No Config tab found');
  const rows = sheet.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(key)) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]); // add if missing
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
