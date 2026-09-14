/**
 * GoZayaan Retail Dashboard — Sheets JSON API
 *
 * Backwards compatible with the existing contract:
 *   ?action=health
 *   ?sheet=Dashboard            -> {ok, name, rows}
 *
 * Adds batch mode, which is the point of this file:
 *   ?sheets=Dashboard,DV,All%20Tour  -> {ok, sheets: {name: rows}}
 *
 * Why batch matters: /exec answers with a 302 to a single-use
 * script.googleusercontent.com URL. Requesting 13 tabs = 13 redirect
 * hops, and under concurrency Google invalidates some content keys
 * before the client follows them, producing transient 404s. One request
 * for all tabs = one hop = the failure mode disappears, and the
 * dashboard's cold start drops from minutes to seconds.
 *
 * Deploy: Deploy > New deployment > Web app
 *   Execute as:        Me
 *   Who has access:    Anyone   <- required, or you get an HTML login page
 * Take the NEW /exec URL and update GOZAAYAN_SHEETS_API_URL on Render.
 */

var SPREADSHEET_ID = '14108t-FKtE9xV0rrpyCokxUqwS5TkHQZ8oPVwmjF0ag';

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function readSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    throw new Error("Sheet not found: " + name);
  }
  // getDisplayValues keeps dates/currency as the sheet renders them,
  // which is what the Python normaliser already expects.
  return sh.getDataRange().getDisplayValues();
}

function doGet(e) {
  var p = (e && e.parameter) || {};

  try {
    if (p.action === 'health') {
      return json_({ ok: true, ts: new Date().toISOString() });
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ---- batch mode -------------------------------------------------
    if (p.sheets) {
      var names = p.sheets.split(',').map(function (n) { return n.trim(); })
                          .filter(function (n) { return n.length > 0; });

      var out = {};
      var errors = {};

      for (var i = 0; i < names.length; i++) {
        try {
          out[names[i]] = readSheet_(ss, names[i]);
        } catch (err) {
          errors[names[i]] = String(err);
        }
      }

      return json_({
        ok: true,
        sheets: out,
        errors: errors,
        count: names.length
      });
    }

    // ---- single-sheet mode (legacy) ---------------------------------
    if (p.sheet) {
      return json_({
        ok: true,
        name: p.sheet,
        rows: readSheet_(ss, p.sheet)
      });
    }

    return json_({ ok: false, error: 'Specify ?sheet= or ?sheets= or ?action=health' });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
