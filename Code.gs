// ── 3827 N Greenview Offer Tracker — Google Apps Script ──────────────────────
// Deploy as: Execute as = Me, Who has access = Anyone
// After deploying, copy the Web App URL into index.html → APPS_SCRIPT_URL

function doGet(e) {
  return respond(getData());
}

function doPost(e) {
  try {
    const offer = JSON.parse(e.postData.contents);
    appendOffer(offer);
    return respond({ success: true });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config tab → simple key:value pairs
  const cfgSheet = ss.getSheetByName('Config');
  const config = {};
  cfgSheet.getDataRange().getValues().forEach(([key, val]) => {
    if (key) config[key] = val;
  });

  // Offers tab → one row per offer
  const offersSheet = ss.getSheetByName('Offers');
  const rows = offersSheet.getDataRange().getValues();
  const headers = rows[0];
  const offers = rows.slice(1)
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      const o = {};
      headers.forEach((h, i) => { o[h] = row[i]; });
      // Parse flags from "type:text||type:text" string
      o.flags = String(o.flags || '').split('||').filter(Boolean).map(f => {
        const colon = f.indexOf(':');
        return { type: f.slice(0, colon).trim(), text: f.slice(colon + 1).trim() };
      });
      return o;
    });

  return { config, offers, updatedAt: new Date().toISOString() };
}

function appendOffer(offer) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Offers');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Serialize flags array back to string
  if (Array.isArray(offer.flags)) {
    offer.flags = offer.flags.map(f => `${f.type}:${f.text}`).join('||');
  }

  // Assign next offer number
  const lastRow = sheet.getLastRow();
  if (!offer.num) {
    const nums = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Number);
    offer.num = nums.length ? Math.max(...nums) + 1 : 1;
  }

  const row = headers.map(h => offer[h] !== undefined ? offer[h] : '');
  sheet.appendRow(row);
}

function respond(data) {
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
