// ── 3827 N Greenview Offer Tracker — Google Apps Script ──────────────────────
// Deploy as: Execute as = Me, Who has access = Anyone
// After ANY code change you MUST redeploy a NEW VERSION:
//   Deploy → Manage deployments → (pencil/Edit) → Version: New version → Deploy

function doGet(e) {
  try {
    return respond(getData());
  } catch (err) {
    return respond({ error: err.message });
  }
}

function doPost(e) {
  try {
    const offer = JSON.parse(e.postData.contents);
    appendOffer(offer);
    return respond({ success: true, num: offer.num });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

// Find a tab by name (case-insensitive). Falls back to a guess so a renamed
// or CSV-imported tab ("sheet_offers", "Sheet1", etc.) never breaks the app.
function findSheet(ss, wanted, keywordFallback) {
  const sheets = ss.getSheets();
  // 1. exact (case-insensitive) match
  for (const s of sheets) {
    if (s.getName().toLowerCase() === wanted.toLowerCase()) return s;
  }
  // 2. keyword contained in the tab name
  for (const s of sheets) {
    if (s.getName().toLowerCase().indexOf(keywordFallback) !== -1) return s;
  }
  return null;
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = {};

  const cfgSheet = findSheet(ss, 'Config', 'config');
  if (cfgSheet) {
    cfgSheet.getDataRange().getValues().forEach(function (r) {
      if (r[0]) config[r[0]] = r[1];
    });
  }

  const offers = [];
  const offersSheet = findSheet(ss, 'Offers', 'offer');
  if (offersSheet) {
    const rows = offersSheet.getDataRange().getValues();
    const headers = rows[0];
    rows.slice(1)
      .filter(function (row) { return row[0] !== '' && row[0] !== null; })
      .forEach(function (row) {
        const o = {};
        headers.forEach(function (h, i) { o[h] = row[i]; });
        o.flags = String(o.flags || '').split('||').filter(Boolean).map(function (f) {
          const c = f.indexOf(':');
          return { type: f.slice(0, c).trim(), text: f.slice(c + 1).trim() };
        });
        offers.push(o);
      });
  }

  return {
    config: config,
    offers: offers,
    updatedAt: new Date().toISOString(),
    _debug: { tabs: ss.getSheets().map(function (s) { return s.getName(); }) }
  };
}

function appendOffer(offer) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = findSheet(ss, 'Offers', 'offer');
  if (!sheet) throw new Error('No Offers tab found');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (Array.isArray(offer.flags)) {
    offer.flags = offer.flags.map(function (f) { return f.type + ':' + f.text; }).join('||');
  }

  const lastRow = sheet.getLastRow();
  if (!offer.num && lastRow > 1) {
    const nums = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Number);
    offer.num = nums.length ? Math.max.apply(null, nums) + 1 : 1;
  }

  const row = headers.map(function (h) { return offer[h] !== undefined ? offer[h] : ''; });
  sheet.appendRow(row);
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
