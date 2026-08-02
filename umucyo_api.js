// =============================================================================
//  UMUCYO OCDS LIVE API — Feature-flagged integration
// =============================================================================
// Add to doPost switch (before the "default:" line):
//   case "getUmucyoSummary": result = getUmucyoSummary(params.year || ""); break;
//   case "syncTendersFromLiveAPI": result = syncTendersFromLiveAPI(params.dateFrom || "", params.limit || 50); break;
//   case "lookupTenderOCDS": result = lookupTenderOCDS(params.ocid || ""); break;
//
// ScriptProperty to enable: UMUCYO_API_ENABLED = "true"

var UMUCYO_API_BASE = "https://ocds.umucyo.gov.rw/opendata";

function umucyoEnabled() {
  return PropertiesService.getScriptProperties().getProperty("UMUCYO_API_ENABLED") === "true";
}

function getUmucyoSummary(year) {
  if (!umucyoEnabled()) return { error: "Umucyo API not enabled yet." };
  var y = parseInt(year) || new Date().getFullYear();
  try {
    var res = UrlFetchApp.fetch(UMUCYO_API_BASE + "/api/v1/ui/summaries/" + y, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { error: "API returned " + res.getResponseCode() };
    var data = JSON.parse(res.getContentText());
    return { year: y + "/" + (y + 1), awarded_value: data.awarded_procurement_value || 0, published_entities: data.published_procurement_entities || 0, registered_suppliers: data.registered_suppliers || 0, awarded_suppliers: data.awarded_suppliers || 0, published_tenders: data.published_tenders || 0, procurement_plans: data.procurement_plans || 0, awards: data.awards || 0, source: "RPPA Official" };
  } catch(e) { return { error: String(e) }; }
}

function syncTendersFromLiveAPI(dateFrom, limit) {
  if (!umucyoEnabled()) return { error: "Umucyo API not enabled.", skipped: true };
  var df = dateFrom || new Date(Date.now() - 7 * 24 * 3600 * 1000).toISOString().substring(0, 10);
  var lim = Math.min(limit || 50, 300);
  try {
    var res = UrlFetchApp.fetch(UMUCYO_API_BASE + "/api/v1/releases/all?date_from=" + df + "&limit=" + lim + "&sort_direction=desc", { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { error: "Releases endpoint returned " + res.getResponseCode() + " — not production-ready yet.", skipped: true };
    var pkg = JSON.parse(res.getContentText());
    var releases = pkg.releases || [];
    var upserted = 0;
    var tSheet = getSheet().getSheetByName("TENDERS_FLAT");
    if (!tSheet || tSheet.getLastRow() < 2) return { error: "TENDERS_FLAT not ready" };
    var headers = tSheet.getDataRange().getValues()[0];
    var existingOcids = {};
    for (var i = 1; i < tSheet.getLastRow(); i++) { existingOcids[String(tSheet.getRange(i + 1, headers.indexOf("ocid") + 1).getValue() || "").toLowerCase()] = true; }
    for (var j = 0; j < releases.length; j++) {
      var r = releases[j]; var t = r.tender || {};
      var ocid = String(r.ocid || "").toLowerCase();
      if (existingOcids[ocid]) continue;
      var row = [];
      row[headers.indexOf("ocid")] = r.ocid || "";
      row[headers.indexOf("title")] = t.title || "";
      row[headers.indexOf("entity")] = (t.procuringEntity || {}).name || "";
      row[headers.indexOf("est_value")] = ((t.value || {}).amount) || 0;
      row[headers.indexOf("method")] = t.procurementMethod || "";
      row[headers.indexOf("deadline")] = (t.tenderPeriod || {}).endDate || "";
      row[headers.indexOf("num_tenderers")] = t.numberOfTenderers || 0;
      row[headers.indexOf("status")] = t.status || "";
      row[headers.indexOf("publish_date")] = r.date || "";
      row[headers.indexOf("description")] = t.description || "";
      row[headers.indexOf("source")] = "umucyo_live";
      tSheet.appendRow(row);
      existingOcids[ocid] = true;
      upserted++;
    }
    return { success: true, total: releases.length, upserted: upserted };
  } catch(e) { return { error: String(e), skipped: true }; }
}
