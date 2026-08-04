// =============================================================================
// RPPA Umucyo Assistant — Phase 1: Foundation
// =============================================================================
// Script Properties (set in Apps Script editor → Project Settings):
//   SHEET_ID              — Google Sheet ID for all data storage
//   OPENROUTER_API_KEY    — AI API key (Phase 3+)
//   BIGQUERY_PROJECT_ID   — BigQuery project (Phase 5+)
//   KB_ROOT_FOLDER_ID     — Google Drive KB folder (Phase 4+)
//   STRIPE_SECRET_KEY     — Payment (Phase 7)
//   MOMO_SUBSCRIPTION_KEY, MOMO_API_USER, MOMO_API_KEY
//   AIRTEL_CLIENT_ID, AIRTEL_CLIENT_SECRET
//   BANK_NAME, BANK_ACCOUNT_NAME, BANK_ACCOUNT_NUMBER
// =============================================================================

// ── SPREADSHEET MENU ──────────────────────────────────────────────────────────

function createMenu() {
  SpreadsheetApp.getUi()
    .createMenu("RPPA Assistant")
    .addItem("Setup Sheets", "setupSheets")
    .addItem("Seed Sample Tenders", "seedSampleTenders")
    .addSeparator()
    .addItem("Sync Tenders from BigQuery", "syncTendersFromBigQuery")
    .addSeparator()
    .addItem("Upload Scraped Tenders to BigQuery", "uploadScrapedToBigQuery")
    .addItem("Install Daily Sync Trigger", "installDailySync")
    .addSeparator()
    .addItem("Sync Latest OCP Registry Data", "syncFromOCPRegistry")
    .addSeparator()
    .addItem("Setup Scraped Import Sheet", "setupScrapedImportSheet")
    .addItem("Import Scraped Umucyo Tenders", "importScrapedTenders")
    .addItem("Import from Drive (umucyo_scraped_tenders.csv)", "importScrapedFromDrive")
    .addSeparator()
    .addItem("Extract Bidder Contacts for Marketing", "extractBidderContacts")
    .addItem("Deduplicate Bidder Contacts", "deduplicateBidderContacts")
    .addItem("Diagnose Contact Counts", "diagnoseContactCounts")
    .addItem("Diagnose rppc_main Columns", "diagnoseMainColumns")
    .addItem("Diagnose Latest Data Month", "diagnoseLatestMonth")
    .addItem("Install Weekly Digest Trigger", "installWeeklyDigest")
    .addSeparator()
    .addItem("Scrape JobInRwanda Tenders", "scrapeJobInRwanda")
    .addItem("Install Daily JIR Scraper", "installJIRScraper")
    .addSeparator()
    .addItem("Generate Matches (No AI)", "generateMatchesNoAI")
    .addItem("Build Buyer Profiles", "buildBuyerProfiles")
    .addSeparator()
    .addItem("Run Phase 1 Tests", "runPhase1Tests")
    .addItem("Run Phase 2 Tests", "testPhase2Functions")
    .addItem("Run Phase 3 Tests", "testPhase3Functions")
    .addItem("Run Phase 4 Tests", "testPhase4Functions")
    .addItem("Run Phase 5 Tests", "testPhase5Functions")
    .addSeparator()
    .addItem("KB Setup", "kbSetup")
    .addItem("KB Install Weekly Scan", "kbInstallTrigger")
    .addSeparator()
    .addItem("Downgrade Expired Subs", "downgradeExpired")
    .addSeparator()
    .addItem("Setup MoMo Sandbox Credentials", "setupMomoSandbox")
    .addSeparator()
    .addItem("Count Active Tenders in BigQuery", "countActiveTenders")
    .addToUi();
}

function setupMomoSandbox() {
  var props = PropertiesService.getScriptProperties();
  var subKey = props.getProperty("MOMO_SUBSCRIPTION_KEY");
  if (!subKey) {
    Logger.log("STEP 1: First, get your Subscription Key.");
    Logger.log("Go to https://momodeveloper.mtn.com → Profile → copy 'Primary Key'");
    Logger.log("Then paste it in Script Properties as MOMO_SUBSCRIPTION_KEY and re-run this.");
    return;
  }

  var baseUrl = "https://sandbox.momodeveloper.mtn.com";

  try {
    var apiUser = Utilities.getUuid();
    Logger.log("Generated API User UUID: " + apiUser);
    Logger.log("Using Subscription Key: " + subKey.substring(0, 8) + "...");

    var endpoints = [
      { url: baseUrl + "/v1_0/apiuser", label: "Collections v1" },
      { url: baseUrl + "/apiuser", label: "Provisioning" },
      { url: baseUrl + "/collection/v1_0/apiuser", label: "Collection Widget v1" }
    ];

    var created = false;
    for (var ei = 0; ei < endpoints.length; ei++) {
      var ep = endpoints[ei];
      Logger.log("Trying endpoint [" + ep.label + "]: " + ep.url);
      var res1 = UrlFetchApp.fetch(ep.url, {
        method: "post",
        headers: {
          "X-Reference-Id": apiUser,
          "Ocp-Apim-Subscription-Key": subKey,
          "Content-Type": "application/json"
        },
        payload: JSON.stringify({ providerCallbackHost: "script.google.com" }),
        muteHttpExceptions: true
      });
      Logger.log("  → HTTP " + res1.getResponseCode() + " — " + res1.getContentText().substring(0, 200));
      if (res1.getResponseCode() === 201) {
        created = true;

        Logger.log("Creating API Key for " + apiUser + "...");
        var res2 = UrlFetchApp.fetch(ep.url + "/" + apiUser + "/apikey", {
          method: "post",
          headers: {
            "Ocp-Apim-Subscription-Key": subKey,
            "Content-Type": "application/json"
          },
          muteHttpExceptions: true
        });
        Logger.log("  → HTTP " + res2.getResponseCode() + " — " + res2.getContentText().substring(0, 200));
        if (res2.getResponseCode() === 201) {
          var data = JSON.parse(res2.getContentText());
          var apiKey = data.apiKey || "";
          props.setProperty("MOMO_API_USER", apiUser);
          props.setProperty("MOMO_API_KEY", apiKey);
          Logger.log("=== MOMO SANDBOX CREDENTIALS SET (via " + ep.label + ") ===");
          Logger.log("MOMO_API_USER: " + apiUser);
          Logger.log("MOMO_API_KEY: " + apiKey);
          break;
        }
      }
    }
    if (!created) {
      Logger.log("ERROR: All endpoints returned non-201.");
      Logger.log("→ Make sure you subscribed to 'Collections' (not 'Collection Widget') in the MTN portal.");
      Logger.log("→ Go to Products → find 'Collections' → Subscribe → use THAT Primary Key.");
    }
  } catch(e) {
    Logger.log("ERROR: " + e);
  }
}

function onOpen() {
  try { createMenu(); } catch(e) { Logger.log("onOpen: " + e); }
}

// ── SPREADSHEET ACCESS ────────────────────────────────────────────────────────

var _appSheet = null;

function getSheet() {
  if (_appSheet) return _appSheet;
  var id = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  if (!id) throw new Error("SHEET_ID not set in Script Properties");
  _appSheet = SpreadsheetApp.openById(id);
  return _appSheet;
}

function getCurrentUserId(uid) {
  if (uid && String(uid).trim() !== "") return String(uid).trim().toLowerCase();
  try {
    var email = Session.getActiveUser().getEmail();
    if (email && email.trim() !== "") return email.trim().toLowerCase();
    email = Session.getEffectiveUser().getEmail();
    if (email && email.trim() !== "") return email.trim().toLowerCase();
  } catch(e) {}
  return "test_user";
}

// ── SETUP ─────────────────────────────────────────────────────────────────────

function setupSheets() {
  var sheets = [
    {
      name: "TENDERS_FLAT",
      headers: [
        "ocid","year","entity","title","description","method","status",
        "est_value","currency","publish_date","deadline",
        "days_to_deadline","num_tenderers","award_value","contract_value",
        "item_classifications","lot_count","lot_value","has_amendments"
      ]
    },
    {
      name: "USER_PROFILES",
      headers: [
        "user_id","company_name","tin","year_established","company_size",
        "is_women_led","is_youth_led","rdb_registered",
        "primary_sector","capabilities","past_gov_contracts","past_gov_count","past_gov_total_value",
        "certifications","min_value","max_value","preferred_methods","regions","timeline_urgency",
        "risk_appetite","growth_orientation","decision_speed","learning_mindset",
        "competitive_posture","price_sensitivity",
        "max_team_size","equipment","cash_flow_limit","pref_single_bid",
        "sectors_of_interest","language_preference","created_at"
      ]
    },
    {
      name: "MATCHES",
      headers: [
        "user_id","ocid","entity","title","description","method","est_value",
        "publish_date","deadline","num_tenderers",
        "match_score","win_chance","win_signals","buyer_snapshot",
        "competitive_context","decision_guidance","risk_flags",
        "checklist","ai_explainer","tag","generated_at",
        "win_probability","expected_bidders","recommended_bid",
        "top_competitors","buyer_contact_phone","buyer_contact_email",
        "buyer_amendment_pct","price_ci_low","price_ci_median","price_ci_high",
        "buyer_integrity_score","buyer_integrity_label",
        "cost_overrun_avg","cost_overrun_count",
        "timeline_risk_pct","contract_completion_pct",
        "award_contract_gap_pct","contract_signing_lag_days",
        "amendment_rationale","amendment_trend",
        "bid_no_bid","bid_no_bid_confidence","bid_no_bid_reasons",
        "single_bidder_pct"
      ]
    },
    {
      name: "BUYER_PROFILES",
      headers: [
        "entity","entity_category","total_tenders","active_years",
        "avg_tenderers","pct_single_bidder","pct_direct_award",
        "avg_est_value","median_est_value",
        "top_sectors","common_methods","last_active",
        "competition_label","generated_at"
      ]
    },
    {
      name: "SAVED_TENDERS",
      headers: ["user_id","ocid","saved_at"]
    },
    {
      name: "USER_ACTIVITY",
      headers: ["user_id","action","detail","timestamp"]
    },
    {
      name: "SUBSCRIPTIONS",
      headers: ["user_id","tier","status","match_limit","features_json","provider","payment_ref","amount","currency","created_at","expires_at","auto_renew"]
    }
  ];

  var ss = getSheet();

  sheets.forEach(function(s) {
    var sheet = ss.getSheetByName(s.name);
    if (!sheet) sheet = ss.insertSheet(s.name);

    var firstRow = sheet.getRange(1, 1, 1, s.headers.length).getValues()[0];
    var isEmpty = firstRow.every(function(cell) { return cell === ""; });

    if (isEmpty) {
      sheet.getRange(1, 1, 1, s.headers.length).setValues([s.headers]);
      sheet.setFrozenRows(1);
    }
  });

  Logger.log("All sheets verified.");
}

// ── SAMPLE DATA ──────────────────────────────────────────────────────────────

function seedSampleTenders() {
  var ss = getSheet();
  var sheet = ss.getSheetByName("TENDERS_FLAT");
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName("TENDERS_FLAT"); }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = function(name) { return headers.indexOf(name); };

  var tenders = [
    ["ocds-a44gk2-2025-000101","2025","MININFRA","Construction of Gakenke District Hospital","Build a 120-bed district hospital with outpatient wing, maternity ward, and staff quarters in Gakenke District, Northern Province","open","active",950000000,"RWF","2025-05-01","2025-07-15",62,6,895000000,895000000,"Construction works",1,950000000,"no"],
    ["ocds-a44gk2-2025-000102","2025","RURA","IT Network Infrastructure Upgrade","Design, supply, and install a fiber optic network backbone connecting 47 public institutions across Kigali City","selective","active",450000000,"RWF","2025-05-10","2025-07-30",77,4,420000000,420000000,"IT equipment and services",1,450000000,"no"],
    ["ocds-a44gk2-2025-000103","2025","MINAGRI","Supply of Agricultural Fertilizers — NPK 17-17-17","Supply 15,000 metric tonnes of NPK 17-17-17 fertilizer for distribution to smallholder farmers across all 30 districts","open","active",2800000000,"RWF","2025-05-15","2025-08-15",92,8,2710000000,2710000000,"Agricultural supplies",1,2800000000,"no"],
    ["ocds-a44gk2-2025-000104","2025","MINEDUC","Classroom Construction — 12 Schools in Eastern Province","Construct 8-classroom blocks at each of 12 primary schools in Bugesera, Ngoma, and Kayonza Districts","open","active",1800000000,"RWF","2025-05-02","2025-06-30",43,7,1760000000,1760000000,"Construction works",12,150000000,"yes"],
    ["ocds-a44gk2-2025-000105","2025","MINISANTE","Medical Equipment Supply — 5 District Hospitals","Supply, deliver, and install medical equipment including X-ray machines, ultrasound units, and laboratory analyzers for 5 district hospitals","selective","active",650000000,"RWF","2025-06-01","2025-08-01",61,5,620000000,620000000,"Medical equipment",1,650000000,"no"],
    ["ocds-a44gk2-2025-000106","2025","Rwanda Energy Group","Solar Mini-Grid Installation — 8 Rural Sites","Design, supply, and install 8 solar photovoltaic mini-grid systems with battery storage in off-grid rural communities across Eastern and Southern Provinces","open","active",1200000000,"RWF","2025-06-15","2025-09-15",92,5,1150000000,1150000000,"Energy equipment and services",8,150000000,"no"],
    ["ocds-a44gk2-2025-000107","2025","Kigali City","Road Rehabilitation — 15 km in Nyarugenge District","Rehabilitation of 15 kilometers of urban roads including drainage, sidewalks, street lighting, and road markings in Nyarugenge District","open","active",2100000000,"RWF","2025-05-20","2025-08-20",92,9,1980000000,0,"Civil works",1,2100000000,"yes"],
    ["ocds-a44gk2-2025-000108","2025","Rwanda Biomedical Centre","Laboratory Reagents and Consumables — Framework Contract","Framework contract for supply of laboratory reagents, test kits, and consumables for national reference laboratory network over 12 months","selective","active",350000000,"RWF","2025-04-10","2025-06-10",31,3,340000000,0,"Medical supplies",1,350000000,"no"],
    ["ocds-a44gk2-2025-000109","2025","WASAC","Water Supply System Extension — Rubavu District","Extension of water supply network including 25 km of distribution pipes, 3 pumping stations, and 500 household connections in Rubavu District","open","active",780000000,"RWF","2025-06-20","2025-09-01",73,6,750000000,750000000,"Water and sanitation works",1,780000000,"no"],
    ["ocds-a44gk2-2025-000110","2025","MINICT","Digital Government Platform Development","Design, develop, and deploy a unified e-government services platform with citizen portal, mobile app, payment gateway, and data analytics dashboard","selective","active",520000000,"RWF","2025-07-01","2025-10-01",92,5,510000000,510000000,"IT software and services",1,520000000,"no"],
    ["ocds-a44gk2-2025-000111","2025","Rwanda Housing Authority","Affordable Housing Construction — 200 Units in Bugesera","Design and build 200 affordable housing units with associated infrastructure including access roads, water, electricity, and sewage systems","open","active",3500000000,"RWF","2025-06-01","2025-10-15",136,10,3420000000,0,"Construction works",1,3500000000,"yes"],
    ["ocds-a44gk2-2025-000112","2025","MINAGRI","Irrigation System Construction — 500 Hectares in Nyagatare","Construction of drip irrigation system covering 500 hectares with water storage reservoirs and pumping stations in Nyagatare District","open","active",1600000000,"RWF","2025-05-25","2025-08-25",92,6,1540000000,1540000000,"Agricultural infrastructure",1,1600000000,"no"],
    ["ocds-a44gk2-2025-000113","2025","REB","ICT Equipment Supply for 30 Secondary Schools","Supply and install 300 desktop computers, 30 projectors, networking equipment, and teacher training for ICT integration across 30 secondary schools","open","active",420000000,"RWF","2025-06-10","2025-09-10",92,7,405000000,405000000,"IT equipment and training",1,420000000,"no"],
    ["ocds-a44gk2-2025-000114","2025","Rwanda Transport Development Agency","Feeder Road Maintenance — Southern Province","Routine and periodic maintenance of 85 km of feeder roads in Huye, Gisagara, and Nyanza Districts including gravelling, culvert installation, and drainage works","open","active",240000000,"RWF","2025-04-15","2025-06-15",27,4,235000000,235000000,"Road maintenance works",1,240000000,"no"],
    ["ocds-a44gk2-2025-000115","2025","MINALOC","Community Development Project — Office Equipment and Furniture","Supply and delivery of office furniture, computers, printers, and solar backup systems for 25 newly established sector offices across Northern Province","open","active",180000000,"RWF","2025-05-18","2025-07-18",61,8,172000000,172000000,"General supplies",1,180000000,"no"]
  ];

  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();

  tenders.forEach(function(row) {
    sheet.appendRow(row);
  });

  SpreadsheetApp.flush();
  var dataRows = sheet.getLastRow() - 1;
  Logger.log("Seeded " + tenders.length + " tenders. Sheet has " + sheet.getLastRow() + " total rows (" + dataRows + " data rows).");
}

function syncTendersFromBigQuery() {
  try {
    var sql = "SELECT ocid, buyer_name AS entity, tender_title AS title, tender_description AS description, tender_status AS status, tender_procurementMethod AS method, FORMAT_TIMESTAMP('%Y', date) AS year, FORMAT_TIMESTAMP('%Y-%m-%d', date) AS publish_date, CASE WHEN tender_tenderPeriod_endDate IS NOT NULL THEN FORMAT_TIMESTAMP('%Y-%m-%d', tender_tenderPeriod_endDate) ELSE NULL END AS deadline, COALESCE(tender_value_amount, 0) AS est_value, COALESCE(tender_value_currency, 'RWF') AS currency, tender_numberOfTenderers AS num_tenderers, tender_mainProcurementCategory AS item_classifications FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 365 DAY) AND tender_status = 'active' AND ocid IS NOT NULL AND buyer_name IS NOT NULL ORDER BY date DESC LIMIT 5000";
    Logger.log("Syncing tenders from BigQuery...");
    var rows = queryBigQuery(sql);
    if (!rows.length) { Logger.log("No rows returned from BigQuery."); return; }
    Logger.log("Fetched " + rows.length + " tenders from rppc_main");

    var sheet = getSheet().getSheetByName("TENDERS_FLAT");
    if (!sheet) { setupSheets(); sheet = getSheet().getSheetByName("TENDERS_FLAT"); }
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var col = function(name) { return headers.indexOf(name); };

    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();

    var batchRows = [];
    rows.forEach(function(r) {
      var row = [];
      for (var c = 0; c < headers.length; c++) row.push("");
      function put(colName, val) { var idx = col(colName); if (idx >= 0) row[idx] = val; }
      put("ocid", String(r.ocid || ""));
      put("year", String(r.year || ""));
      put("entity", String(r.entity || ""));
      put("title", String(r.title || r.entity || "Tender"));
      put("description", String(r.description || ""));
      put("method", String(r.method || "open"));
      put("status", String(r.status || "active"));
      put("est_value", Number(r.est_value || 0));
      put("currency", String(r.currency || "RWF"));
      put("publish_date", String(r.publish_date || ""));
      put("deadline", String(r.deadline || ""));
      put("days_to_deadline", "");
      put("num_tenderers", Number(r.num_tenderers || 0));
      put("award_value", 0);
      put("contract_value", 0);
      put("item_classifications", String(r.item_classifications || ""));
      put("lot_count", 1);
      put("lot_value", Number(r.est_value || 0));
      put("has_amendments", "no");
      batchRows.push(row);
    });

    var batchSize = 500;
    for (var i = 0; i < batchRows.length; i += batchSize) {
      var chunk = batchRows.slice(i, Math.min(i + batchSize, batchRows.length));
      sheet.getRange(i + 2, 1, chunk.length, chunk[0].length).setValues(chunk);
      SpreadsheetApp.flush();
    }

    Logger.log("Synced " + batchRows.length + " tenders to TENDERS_FLAT from BigQuery.");
  } catch(e) {
    Logger.log("syncTendersFromBigQuery ERROR: " + e);
  }
}

/**
 * Setup a "SCRAPED_RAW" sheet with CSV headers for pasting scraped Umucyo tenders.
 * Run this first, then paste CSV data into the sheet, then run importScrapedTenders().
 */
function setupScrapedImportSheet() {
  var ss = getSheet();
  var existing = ss.getSheetByName("SCRAPED_RAW");
  if (existing) { Logger.log("SCRAPED_RAW already exists."); return; }
  var sheet = ss.insertSheet("SCRAPED_RAW");
  var headers = ["tender_reference_no","tender_name","pe_code","pe_abbreviation","tender_type","tender_type_label","tender_method","tender_method_label","status_code","status","stage_code","stage_type","advertising_date","deadline_submitting","planed_open_date","submission_deadline_date","opening_datetime"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  Logger.log("SCRAPED_RAW sheet created. Paste your CSV data starting at row 2.");
}

/**
 * Read scraped tenders from SCRAPED_RAW, check against TENDERS_FLAT and BigQuery,
 * and append new ones to TENDERS_FLAT for use in matching.
 */
function importScrapedTenders() {
  var ss = getSheet();
  var raw = ss.getSheetByName("SCRAPED_RAW");
  if (!raw || raw.getLastRow() < 2) { throw new Error("SCRAPED_RAW is empty or missing. Run setupScrapedImportSheet first, then paste CSV data."); }

  // 1) Read and parse the raw CSV-like data
  var rawData = raw.getDataRange().getValues();
  var rawHeaders = rawData[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var scraped = [];
  for (var i = 1; i < rawData.length; i++) {
    var row = rawData[i];
    if (!row[0]) continue; // skip empty rows
    var obj = {};
    rawHeaders.forEach(function(h, j) { obj[h] = String(row[j] || "").trim(); });
    scraped.push(obj);
  }
  Logger.log("Loaded " + scraped.length + " scraped tenders from SCRAPED_RAW");

  // 2) Get existing TENDERS_FLAT data (ocid + title for dedup)
  var tSheet = ss.getSheetByName("TENDERS_FLAT");
  if (!tSheet) { setupSheets(); tSheet = ss.getSheetByName("TENDERS_FLAT"); }
  var tData = tSheet.getDataRange().getValues();
  var tHeaders = tData[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var existingOcids = {};
  var existingTitles = {};
  for (var ti = 1; ti < tData.length; ti++) {
    var tr = tData[ti];
    var o = tr[tHeaders.indexOf("ocid")] || "";
    var tl = String(tr[tHeaders.indexOf("title")] || "").toLowerCase().trim();
    if (o) existingOcids[String(o).trim()] = true;
    if (tl) existingTitles[tl] = true;
  }
  Logger.log("TENDERS_FLAT has " + Object.keys(existingOcids).length + " existing ocids");

  // 3) Map scraped → TENDERS_FLAT format, skipping duplicates
  var newRows = [];
  var skipped = 0;
  scraped.forEach(function(s) {
    var candidateOcid = s.tender_reference_no || ("umucyo-" + s.advertising_date + "-" + s.tender_reference_no);
    var candidateTitle = s.tender_name || "Tender";
    var titleLower = candidateTitle.toLowerCase().trim();

    // Skip if ocid or title already exists
    if (existingOcids[candidateOcid]) { skipped++; return; }
    if (existingTitles[titleLower]) { skipped++; return; }

    // Parse date formats (DD/MM/YYYY → YYYY-MM-DD)
    function parseDate(d) {
      if (!d) return "";
      var m = d.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) return m[3] + "-" + m[2] + "-" + m[1];
      return d;
    }

    var publishDate = parseDate(s.advertising_date);
    var deadline = parseDate(s.deadline_submitting || s.submission_deadline_date);

    // Compute days_to_deadline
    var days = "";
    if (deadline) {
      try {
        var dl = new Date(deadline);
        var now = new Date();
        days = Math.ceil((dl.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
        if (days < 0) days = 0;
      } catch(e) {}
    }

    var row = [];
    for (var c = 0; c < tHeaders.length; c++) row.push("");
    function put(name, val) { var idx = tHeaders.indexOf(name.toLowerCase()); if (idx >= 0) row[idx] = val; }
    put("ocid", candidateOcid);
    put("year", String(extractYear(publishDate) || ""));
    put("entity", s.pe_abbreviation || s.pe_code || "");
    put("title", candidateTitle);
    put("description", candidateTitle + (s.tender_type_label ? " · " + s.tender_type_label : "") + (s.stage_type ? " · " + s.stage_type : ""));
    put("method", s.tender_method_label || s.tender_method || "open");
    put("status", s.status || "active");
    put("est_value", 0);
    put("currency", "RWF");
    put("publish_date", publishDate);
    put("deadline", deadline);
    put("days_to_deadline", days);
    put("num_tenderers", 0);
    put("award_value", 0);
    put("contract_value", 0);
    put("item_classifications", s.tender_type_label || s.tender_type || "");
    put("lot_count", 1);
    put("lot_value", 0);
    put("has_amendments", "no");
    newRows.push(row);
  });

  Logger.log("Skipped " + skipped + " existing tenders, new: " + newRows.length);

  // 4) Append to TENDERS_FLAT
  if (newRows.length > 0) {
    var startRow = tSheet.getLastRow() + 1;
    tSheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
    SpreadsheetApp.flush();
    Logger.log("Appended " + newRows.length + " scraped tenders to TENDERS_FLAT (rows " + startRow + "-" + (startRow + newRows.length - 1) + ")");
  }

  return { imported: newRows.length, skipped: skipped, total: scraped.length };
}

/**
 * Alternative: upload the CSV file to Google Drive root, give it a fixed name,
 * and this function reads it directly.
 * Name the file: "umucyo_scraped_tenders.csv"
 */
function importScrapedFromDrive() {
  var files = DriveApp.getFilesByName("umucyo_scraped_tenders.csv");
  if (!files.hasNext()) throw new Error("No file named 'umucyo_scraped_tenders.csv' found in Google Drive. Upload it first.");
  var file = files.next();
  var csvText = file.getBlob().getDataAsString("UTF-8");
  Logger.log("Read " + csvText.length + " chars from Drive file: " + file.getName());
  return importScrapedCsvText(csvText);
}

/**
 * Inner function: parse CSV text, dedup, and append to TENDERS_FLAT.
 * Shared by importScrapedTenders() (reads from sheet) and importScrapedFromDrive() (reads from Drive).
 */
function importScrapedCsvText(csvText) {
  var lines = csvText.split("\n").map(function(l) { return l.trim(); }).filter(function(l) { return l; });
  if (lines.length < 2) throw new Error("CSV must have a header row and at least one data row.");

  var rawHeaders = lines[0].split(",").map(function(h) { return h.replace(/^"|"$/g, "").trim().toLowerCase(); });
  var scraped = [];
  for (var i = 1; i < lines.length; i++) {
    // Simple CSV parser (handles quoted commas)
    var vals = [];
    var cur = "";
    var inQ = false;
    for (var ci = 0; ci < lines[i].length; ci++) {
      var ch = lines[i][ci];
      if (ch === '"') { inQ = !inQ; continue; }
      if (ch === "," && !inQ) { vals.push(cur); cur = ""; continue; }
      cur += ch;
    }
    vals.push(cur);
    if (!vals[0]) continue;
    var obj = {};
    rawHeaders.forEach(function(h, j) { obj[h] = String(vals[j] || "").trim(); });
    scraped.push(obj);
  }
  Logger.log("Parsed " + scraped.length + " tenders from CSV text");

  // --- dedup + map logic (same as importScrapedTenders) ---
  var ss = getSheet();
  var tSheet = ss.getSheetByName("TENDERS_FLAT");
  if (!tSheet) { setupSheets(); tSheet = ss.getSheetByName("TENDERS_FLAT"); }
  var tData = tSheet.getDataRange().getValues();
  var tHeaders = tData[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var existingOcids = {};
  var existingTitles = {};
  for (var ti = 1; ti < tData.length; ti++) {
    var tr = tData[ti];
    var o = tr[tHeaders.indexOf("ocid")] || "";
    var tl = String(tr[tHeaders.indexOf("title")] || "").toLowerCase().trim();
    if (o) existingOcids[String(o).trim()] = true;
    if (tl) existingTitles[tl] = true;
  }

  var newRows = [];
  var skipped = 0;
  scraped.forEach(function(s) {
    var candidateOcid = s.tender_reference_no || ("umucyo-" + (s.advertising_date || "unknown") + "-" + (s.tender_reference_no || i));
    var candidateTitle = s.tender_name || "Tender";
    var titleLower = candidateTitle.toLowerCase().trim();

    if (existingOcids[candidateOcid]) { skipped++; return; }
    if (existingTitles[titleLower]) { skipped++; return; }

    function parseDate(d) {
      if (!d) return "";
      var m = d.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) return m[3] + "-" + m[2] + "-" + m[1];
      return d;
    }

    var publishDate = parseDate(s.advertising_date);
    var deadline = parseDate(s.deadline_submitting || s.submission_deadline_date);

    var days = "";
    if (deadline) {
      try {
        var dl = new Date(deadline);
        var now = new Date();
        days = Math.ceil((dl.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
        if (days < 0) days = 0;
      } catch(e) { days = ""; }
    }

    var row = [];
    for (var c = 0; c < tHeaders.length; c++) row.push("");
    function put(name, val) { var idx = tHeaders.indexOf(name.toLowerCase()); if (idx >= 0) row[idx] = val; }
    put("ocid", candidateOcid);
    put("year", String(extractYear(publishDate) || ""));
    put("entity", s.pe_abbreviation || s.pe_code || "");
    put("title", candidateTitle);
    put("description", candidateTitle + (s.tender_type_label ? " · " + s.tender_type_label : "") + (s.stage_type ? " · " + s.stage_type : ""));
    put("method", s.tender_method_label || s.tender_method || "open");
    put("status", s.status || "active");
    put("est_value", 0);
    put("currency", "RWF");
    put("publish_date", publishDate);
    put("deadline", deadline);
    put("days_to_deadline", days);
    put("num_tenderers", 0);
    put("award_value", 0);
    put("contract_value", 0);
    put("item_classifications", s.tender_type_label || s.tender_type || "");
    put("lot_count", 1);
    put("lot_value", 0);
    put("has_amendments", "no");
    newRows.push(row);
  });

  Logger.log("Skipped " + skipped + " duplicates, importing " + newRows.length + " new tenders");

  if (newRows.length > 0) {
    var startRow = tSheet.getLastRow() + 1;
    tSheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
    SpreadsheetApp.flush();
    Logger.log("Appended " + newRows.length + " tenders starting at row " + startRow);
  }

  return { imported: newRows.length, skipped: skipped, total: scraped.length };
}

// ════════════════════════════════════════════════════════════════════
//  LIVE TENDER PIPELINE — BigQuery table + daily sync
// ════════════════════════════════════════════════════════════════════

var LIVE_TABLE_ID = "umucyo_live_tenders";

function ensureLiveTendersTable() {
  var dataset = BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET;
  try {
    var existing = BigQuery.Tables.get(BIGQUERY_PROJECT_ID, BIGQUERY_DATASET, LIVE_TABLE_ID);
    Logger.log("Table " + LIVE_TABLE_ID + " already exists.");
    return existing;
  } catch(e) {
    Logger.log("Creating " + LIVE_TABLE_ID + "...");
  }
  var table = {
    tableReference: { projectId: BIGQUERY_PROJECT_ID, datasetId: BIGQUERY_DATASET, tableId: LIVE_TABLE_ID },
    schema: {
      fields: [
        { name: "tender_reference_no", type: "STRING", mode: "REQUIRED" },
        { name: "tender_name", type: "STRING" },
        { name: "pe_code", type: "STRING" },
        { name: "pe_abbreviation", type: "STRING" },
        { name: "tender_type", type: "STRING" },
        { name: "tender_type_label", type: "STRING" },
        { name: "tender_method", type: "STRING" },
        { name: "tender_method_label", type: "STRING" },
        { name: "status_code", type: "STRING" },
        { name: "status", type: "STRING" },
        { name: "stage_code", type: "STRING" },
        { name: "stage_type", type: "STRING" },
        { name: "advertising_date", type: "STRING" },
        { name: "deadline_submitting", type: "STRING" },
        { name: "submission_deadline_date", type: "STRING" },
        { name: "opening_datetime", type: "STRING" },
        { name: "ocid", type: "STRING" },
        { name: "entity", type: "STRING" },
        { name: "title", type: "STRING" },
        { name: "description", type: "STRING" },
        { name: "method", type: "STRING" },
        { name: "publish_date", type: "STRING" },
        { name: "deadline", type: "STRING" },
        { name: "days_to_deadline", type: "INTEGER" },
        { name: "item_classifications", type: "STRING" },
        { name: "currency", type: "STRING" },
        { name: "est_value", type: "FLOAT" },
        { name: "num_tenderers", type: "INTEGER" },
        { name: "synced_at", type: "TIMESTAMP" }
      ]
    }
  };
  BigQuery.Tables.insert(table, BIGQUERY_PROJECT_ID, BIGQUERY_DATASET);
  Logger.log("Created " + LIVE_TABLE_ID);
}

function getExistingLiveOcids() {
  try {
    var rows = queryBigQuery("SELECT tender_reference_no, ocid FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + "." + LIVE_TABLE_ID + " WHERE tender_reference_no IS NOT NULL");
    var refs = {};
    rows.forEach(function(r) {
      if (r.tender_reference_no) refs[String(r.tender_reference_no).trim()] = true;
      if (r.ocid) refs[String(r.ocid).trim()] = true;
    });
    return refs;
  } catch(e) {
    Logger.log("getExistingLiveOcids: " + e + " (table may not exist)");
    return {};
  }
}

/**
 * Read scraped tenders from SCRAPED_RAW sheet, dedup against BigQuery, and insert new ones.
 */
function uploadScrapedToBigQuery() {
  ensureLiveTendersTable();

  var ss = getSheet();
  var raw = ss.getSheetByName("SCRAPED_RAW");
  if (!raw || raw.getLastRow() < 2) throw new Error("SCRAPED_RAW is empty. Run Setup Scraped Import Sheet first and paste CSV data.");

  var rawData = raw.getDataRange().getValues();
  var rawHeaders = rawData[0].map(function(h) { return String(h).trim().toLowerCase(); });

  // Build scraped objects
  var scraped = [];
  for (var i = 1; i < rawData.length; i++) {
    var row = rawData[i];
    if (!row[0]) continue;
    var obj = {};
    rawHeaders.forEach(function(h, j) { obj[h] = String(row[j] || "").trim(); });
    scraped.push(obj);
  }

  // Get existing refs from BQ
  var existingRefs = getExistingLiveOcids();
  Logger.log("Existing live tenders in BQ: " + Object.keys(existingRefs).length);

  // Build insert rows (skip existing)
  var insertRows = [];
  var skipped = 0;
  scraped.forEach(function(s) {
    var ref = s.tender_reference_no || "";
    if (!ref) { skipped++; return; }
    if (existingRefs[ref]) { skipped++; return; }

    function parseDate(d) {
      if (!d) return "";
      var m = d.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) return m[3] + "-" + m[2] + "-" + m[1];
      return d;
    }

    var pd = parseDate(s.advertising_date);
    var dl = parseDate(s.deadline_submitting || s.submission_deadline_date);
    var days = null;
    if (dl) {
      try { var dt = new Date(dl); var n = new Date(); var d = Math.ceil((dt.getTime() - n.getTime()) / 86400000); days = d < 0 ? 0 : d; } catch(e) { days = null; }
    }

    insertRows.push({
      json: {
        tender_reference_no: ref,
        tender_name: s.tender_name || "",
        pe_code: s.pe_code || "",
        pe_abbreviation: s.pe_abbreviation || "",
        tender_type: s.tender_type || "",
        tender_type_label: s.tender_type_label || "",
        tender_method: s.tender_method || "",
        tender_method_label: s.tender_method_label || "",
        status_code: s.status_code || "",
        status: s.status || "active",
        stage_code: s.stage_code || "",
        stage_type: s.stage_type || "",
        advertising_date: s.advertising_date || "",
        deadline_submitting: s.deadline_submitting || "",
        submission_deadline_date: s.submission_deadline_date || "",
        opening_datetime: s.opening_datetime || "",
        ocid: ref,
        entity: s.pe_abbreviation || s.pe_code || "",
        title: s.tender_name || "Tender",
        description: (s.tender_name || "") + (s.tender_type_label ? " · " + s.tender_type_label : "") + (s.stage_type ? " · " + s.stage_type : ""),
        method: s.tender_method_label || s.tender_method || "open",
        publish_date: pd,
        deadline: dl,
        days_to_deadline: days,
        item_classifications: s.tender_type_label || s.tender_type || "",
        currency: "RWF",
        est_value: 0,
        num_tenderers: 0,
        synced_at: new Date().toISOString()
      }
    });
  });

  Logger.log("New tenders to insert: " + insertRows.length + ", skipped: " + skipped);

  if (insertRows.length > 0) {
    var batchSize = 50;
    for (var bi = 0; bi < insertRows.length; bi += batchSize) {
      var chunk = insertRows.slice(bi, Math.min(bi + batchSize, insertRows.length));
      var fields = Object.keys(chunk[0].json);
      var valuesSql = chunk.map(function(r) {
        return "(" + fields.map(function(f) {
          var v = r.json[f];
          if (v === null || v === undefined) return "NULL";
          if (typeof v === "number") return isFinite(v) ? String(v) : "NULL";
          var s = String(v).replace(/\\/g, "\\\\").replace(/'/g, "\\'");
          return "'" + s + "'";
        }).join(",") + ")";
      }).join(",\n");
      var sql = "INSERT INTO `" + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + "." + LIVE_TABLE_ID + "` (" + fields.join(",") + ")\nVALUES\n" + valuesSql;
      var req = { query: sql, useLegacySql: false };
      Logger.log("DML batch " + (bi / batchSize + 1) + ": ~" + sql.length + " chars for " + chunk.length + " rows");
      var res = BigQuery.Jobs.query(req, BIGQUERY_PROJECT_ID);
      var attempts = 0;
      while (!res.jobComplete && attempts < 20) { Utilities.sleep(2000); res = BigQuery.Jobs.getQueryResults(BIGQUERY_PROJECT_ID, res.jobReference.jobId); attempts++; }
      if (!res.jobComplete) throw new Error("DML job did not complete for batch " + (bi / batchSize + 1));
      Logger.log("Inserted batch " + (bi / batchSize + 1) + " (" + chunk.length + " rows)");
    }
    Logger.log("Total " + insertRows.length + " tenders uploaded to " + LIVE_TABLE_ID);
  } else {
    Logger.log("No new tenders to upload.");
  }

  return { uploaded: insertRows.length, skipped: skipped };
}

// ── MODIFIED syncTendersFromBigQuery — unions historical + live ──

function syncTendersFromBigQuery() {
  try {
    var ds = BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET;
    // Union historical active tenders (365 days) with live scraped tenders
    var sql = [
      "SELECT ocid, buyer_name AS entity, tender_title AS title, tender_description AS description, tender_status AS status, tender_procurementMethod AS method, FORMAT_TIMESTAMP('%Y', date) AS year, FORMAT_TIMESTAMP('%Y-%m-%d', date) AS publish_date, CASE WHEN tender_tenderPeriod_endDate IS NOT NULL THEN FORMAT_TIMESTAMP('%Y-%m-%d', tender_tenderPeriod_endDate) ELSE NULL END AS deadline, COALESCE(tender_value_amount, 0) AS est_value, COALESCE(tender_value_currency, 'RWF') AS currency, tender_numberOfTenderers AS num_tenderers, tender_mainProcurementCategory AS item_classifications FROM " + ds + ".rppc_main WHERE date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 365 DAY) AND tender_status = 'active' AND ocid IS NOT NULL AND buyer_name IS NOT NULL",
      "UNION ALL",
      "SELECT ocid, entity, title, description, status, method, FORMAT_TIMESTAMP('%Y', SAFE.PARSE_TIMESTAMP('%Y-%m-%d', publish_date)) AS year, publish_date, deadline, est_value, currency, num_tenderers, item_classifications FROM " + ds + "." + LIVE_TABLE_ID + " WHERE ocid IS NOT NULL AND entity IS NOT NULL"
    ].join("\n");
    sql += " ORDER BY publish_date DESC LIMIT 5000";

    Logger.log("Syncing tenders from BigQuery (historical + live)...");
    var rows = queryBigQuery(sql);
    if (!rows.length) { Logger.log("No rows returned from BigQuery."); return; }
    Logger.log("Fetched " + rows.length + " tenders (rppc_main historical + umucyo_live_tenders)");

    var sheet = getSheet().getSheetByName("TENDERS_FLAT");
    if (!sheet) { setupSheets(); sheet = getSheet().getSheetByName("TENDERS_FLAT"); }
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var col = function(name) { return headers.indexOf(name); };

    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();

    var batchRows = [];
    rows.forEach(function(r) {
      var row = [];
      for (var c = 0; c < headers.length; c++) row.push("");
      function put(colName, val) { var idx = col(colName); if (idx >= 0) row[idx] = val; }
      put("ocid", String(r.ocid || ""));
      put("year", String(r.year || ""));
      put("entity", String(r.entity || ""));
      put("title", String(r.title || r.entity || "Tender"));
      put("description", String(r.description || ""));
      put("method", String(r.method || "open"));
      put("status", String(r.status || "active"));
      put("est_value", Number(r.est_value || 0));
      put("currency", String(r.currency || "RWF"));
      put("publish_date", String(r.publish_date || ""));
      put("deadline", String(r.deadline || ""));
      put("days_to_deadline", "");
      put("num_tenderers", Number(r.num_tenderers || 0));
      put("award_value", 0);
      put("contract_value", 0);
      put("item_classifications", String(r.item_classifications || ""));
      put("lot_count", 1);
      put("lot_value", Number(r.est_value || 0));
      put("has_amendments", "no");
      batchRows.push(row);
    });

    var batchSize = 500;
    for (var i = 0; i < batchRows.length; i += batchSize) {
      var chunk = batchRows.slice(i, Math.min(i + batchSize, batchRows.length));
      sheet.getRange(i + 2, 1, chunk.length, chunk[0].length).setValues(chunk);
      SpreadsheetApp.flush();
    }

    Logger.log("Synced " + batchRows.length + " tenders to TENDERS_FLAT (historical + live).");
  } catch(e) {
    Logger.log("syncTendersFromBigQuery ERROR: " + e);
  }
}

// ── DAILY SYNC TRIGGER ──────────────────────────────────────────────

function installDailySync() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "syncDaily") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("syncDaily").timeBased().everyDays(1).atHour(6).create();
  Logger.log("Daily sync trigger installed at 6 AM.");
}

function syncDaily() {
  try {
    Logger.log("=== Daily sync: Upload scraped to BQ ===");
    // Check if SCRAPED_RAW has data to upload
    var ss = getSheet();
    var raw = ss.getSheetByName("SCRAPED_RAW");
    if (raw && raw.getLastRow() > 1) {
      uploadScrapedToBigQuery();
    } else {
      // Fallback: check Drive for the CSV file
      var files = DriveApp.getFilesByName("umucyo_scraped_tenders.csv");
      if (files.hasNext()) {
        var file = files.next();
        var csvText = file.getBlob().getDataAsString("UTF-8");
        Logger.log("Read " + csvText.length + " chars from Drive CSV");
        // Parse and upload manually
        var lines = csvText.split("\n").filter(function(l) { return l.trim(); });
        if (lines.length > 1) {
          var tempRaw = ss.getSheetByName("SCRAPED_RAW");
          if (!tempRaw) { tempRaw = ss.insertSheet("SCRAPED_RAW"); }
          tempRaw.clear();
          var csvRows = lines.map(function(l) { return l.split(",").map(function(v) { return v.replace(/^"|"$/g, "").trim(); }); });
          if (csvRows.length > 0) {
            tempRaw.getRange(1, 1, csvRows.length, csvRows[0].length).setValues(csvRows);
            SpreadsheetApp.flush();
            uploadScrapedToBigQuery();
          }
        }
      } else {
        Logger.log("No scraped data found in SCRAPED_RAW or Drive. Skipping BQ upload.");
      }
    }
    Logger.log("=== Daily sync: Refresh TENDERS_FLAT from BQ ===");
    syncTendersFromBigQuery();
    Logger.log("=== Daily sync: Regenerate matches ===");
    var allUids = {};
    ["USER_PROFILES","MATCHES","SUBSCRIPTIONS"].forEach(function(sn) {
      var sh = ss.getSheetByName(sn);
      if (!sh || sh.getLastRow() < 2) return;
      var d = sh.getDataRange().getValues();
      var hIdx = d[0].indexOf("user_id");
      if (hIdx < 0) return;
      for (var i = 1; i < d.length; i++) {
        var u = String(d[i][hIdx] || "").trim().toLowerCase();
        if (u) allUids[u] = true;
      }
    });
    var ids = Object.keys(allUids);
    ids.forEach(function(uid) {
      try { generateMatchesNoAI(uid); } catch(e) { Logger.log("Match fail " + uid + ": " + e); }
    });
    Logger.log("=== Daily sync complete — " + ids.length + " users updated ===");
  } catch(e) {
    Logger.log("syncDaily ERROR: " + e);
  }
}

function syncFromOCPRegistry() {
  var baseUrl = "https://data.open-contracting.org/en/publication/145/download?name=";
  var currentYear = new Date().getFullYear();
  Logger.log("=== OCP Data Registry — Rwanda RPPA ===");
  Logger.log("Data source: https://data.open-contracting.org/en/publication/145");
  Logger.log("Last updated: Monthly (latest: April 2026)");
  Logger.log("");
  Logger.log("Download links for " + currentYear + " data:");
  Logger.log("  JSON: " + baseUrl + currentYear + ".jsonl.gz");
  Logger.log("  CSV:  " + baseUrl + currentYear + ".csv.tar.gz");
  Logger.log("  Excel: " + baseUrl + currentYear + ".xlsx");
  Logger.log("");
  Logger.log("=== Manual Update Instructions ===");
  Logger.log("1. Download the CSV file from the link above");
  Logger.log("2. Extract the .tar.gz (use 7-Zip on Windows)");
  Logger.log("3. In BigQuery Console (https://console.cloud.google.com/bigquery):");
  Logger.log("   → Select dataset rppa_historical");
  Logger.log("   → Click 'Create Table'");
  Logger.log("   → Source: Upload, select the CSV files");
  Logger.log("   → Schema: Auto-detect");
  Logger.log("4. Run 'Sync Tenders from BigQuery' from the RPPA Assistant menu in this sheet");
  Logger.log("");
  Logger.log("=== Important Notes ===");
  Logger.log("• OCP registry updates monthly — check back at start of each month");
  Logger.log("• The full dataset is 329MB compressed — download on good internet");
  Logger.log("• The direct OCDS API (ocds.umucyo.gov.rw) is currently unreachable");
  Logger.log("• See INQUIRY_EMAIL.md in project repo for RPPA contact email");
}

// ── WEB APP ENTRY ────────────────────────────────────────────────────────────

function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  var uid = e && e.parameter && e.parameter.uid ? String(e.parameter.uid).trim().toLowerCase() : "";
  var lookup = e && e.parameter && e.parameter.lookup ? String(e.parameter.lookup).trim() : "";

  function serve(file) {
    var tpl = HtmlService.createTemplateFromFile(file);
    tpl.uid = uid || "";
    tpl.lookup = lookup || "";
    tpl.appUrl = ScriptApp.getService().getUrl();
    return tpl.evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === "dashboard") return serve("Dashboard");
  if (page === "dna")       return serve("BusinessDNA");
  if (page === "find")      return serve("FindOpportunities");
  if (page === "check")     return serve("CheckTender");
  if (page === "compliance") return serve("ReviewProposal");
  if (page === "learn")     return serve("LearnPractice");
  if (page === "insights")  return serve("MyInsights");
  if (page === "market")    return serve("MarketDashboard");
  if (page === "calendar")  return serve("ProcurementCalendar");
  if (page === "portfolio") return serve("SupplierPortfolio");
  if (page === "ratings")   return serve("BuyerRatings");
  if (page === "sector")    return serve("SectorHealth");
  if (page === "competitor") return serve("CompetitorIntel");
  if (page === "jv")        return serve("JVPartners");
  if (page === "upgrade")   return serve("UpgradePrompt");
  if (page === "admin")     return serve("Admin");
  if (page === "terms")    return serve("Terms");
  if (page === "privacy")  return serve("Privacy");
  if (page === "guide")    return serve("Guide");

  return serve("Dashboard");
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ── DASHBOARD DATA ────────────────────────────────────────────────────────────

function getDashboardData(uid) {
  try {
    var activeTenders = 0, newThisWeek = 0;
    var oneWeekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
    var userId = getCurrentUserId(uid);
    var profile = getUserProfile(userId);

    // If no profile, redirect to DNA
    if (!profile) {
      return { needsDNA: true };
    }

    var tSheet = getSheet().getSheetByName("TENDERS_FLAT");
    if (tSheet && tSheet.getLastRow() > 1) {
      var tData = tSheet.getDataRange().getValues();
      var tHeaders = tData[0];
      activeTenders = tData.length - 1;
      for (var i = 1; i < tData.length; i++) {
        var t = mapRow(tHeaders, tData[i]);
        var pub = t.publish_date ? new Date(t.publish_date) : null;
        if (pub && pub >= oneWeekAgo) newThisWeek++;
      }
    }

    var matchCount = 0;
    var mSheet = getSheet().getSheetByName("MATCHES");
    if (mSheet && mSheet.getLastRow() > 1) {
      var mData = mSheet.getDataRange().getValues();
      matchCount = mData.slice(1).filter(function(r) { return (r[0] || "").toLowerCase() === userId.toLowerCase(); }).length;
    }

    var savedCount = 0;
    var savedTenders = [];
    var sSheet = getSheet().getSheetByName("SAVED_TENDERS");
    if (sSheet && sSheet.getLastRow() > 1) {
      var sData = sSheet.getDataRange().getValues();
      var savedUserRows = sData.slice(1).filter(function(r) { return (r[0] || "").toLowerCase() === userId.toLowerCase(); });
      savedCount = savedUserRows.length;
      // Look up titles from TENDERS_FLAT
      savedUserRows.slice(-10).reverse().forEach(function(sr) {
        var ocid = sr[1] || "";
        var title = ocid;
        try {
          if (tSheet && tHeaders) {
            for (var ti = 1; ti < tData.length; ti++) {
              if (tData[ti][tHeaders.indexOf("ocid")] === ocid) {
                title = tData[ti][tHeaders.indexOf("title")] || ocid;
                break;
              }
            }
          }
        } catch(e2) {}
        savedTenders.push({ ocid: ocid, title: title });
      });
    }

    var recentActivity = [];
    var actSheet = getSheet().getSheetByName("USER_ACTIVITY");
    if (actSheet && actSheet.getLastRow() > 1) {
      var actData = actSheet.getDataRange().getValues();
      for (var ai = actData.length - 1; ai >= 1 && recentActivity.length < 10; ai--) {
        if (String(actData[ai][0] || "").toLowerCase() !== userId.toLowerCase()) continue;
        recentActivity.push({
          action: String(actData[ai][1] || ""),
          detail: String(actData[ai][2] || ""),
          timestamp: actData[ai][3] ? new Date(actData[ai][3]).toISOString() : ""
        });
      }
    }

    var sub = getSubscription(userId) || { tier: "free", status: "active", expires_at: null };

    // Add JIR tender count
    var jirSheet = getSheet().getSheetByName("JIR_TENDERS");
    if (jirSheet && jirSheet.getLastRow() > 1) {
      activeTenders += (jirSheet.getLastRow() - 1);
      var jirData = jirSheet.getDataRange().getValues();
      for (var ji = 1; ji < jirData.length; ji++) {
        var jirPub = jirData[ji][3] ? new Date(jirData[ji][3]) : null;
        if (jirPub && jirPub >= oneWeekAgo) newThisWeek++;
      }
    }

    // Auto-grant trial on first dashboard load if no subscription
    if (sub.status !== "trial" && sub.tier === "free") {
      try { grantTrial(userId); sub = getSubscription(userId) || sub; } catch(e) {}
    }

    return {
      activeTenders: activeTenders, newThisWeek: newThisWeek,
      savedCount: savedCount, matchCount: matchCount,
      savedTenders: savedTenders,
      recentActivity: recentActivity,
      tier: sub.tier || "free",
      status: sub.status || "active",
      expires_at: sub.expires_at || null,
      companyName: profile.company_name || "",
      email: userId,
      appUrl: ScriptApp.getService().getUrl(),
      pipelineWarnings: getPipelineWarning(userId),
      supplierScore: computeSupplierScore(userId)
    };
  } catch(e) {
    Logger.log("getDashboardData: " + e);
    return {
      activeTenders: 0, newThisWeek: 0, savedCount: 0, matchCount: 0,
      savedTenders: [],
      recentActivity: [], tier: "free", status: "error",
      expires_at: null, companyName: "", email: "",
      supplierScore: { score: 0, label: "Unknown" }
    };
  }
}

// ── UTILITIES ─────────────────────────────────────────────────────────────────

function mapRow(headers, row) {
  var obj = {};
  headers.forEach(function(h, i) { obj[h] = row[i]; });
  return obj;
}

function safeLower(val) {
  if (val == null) return "";
  if (typeof val === "string") return val.toLowerCase();
  return String(val).toLowerCase();
}

function formatMoney(val) {
  if (!val) return "0";
  var v = Number(val);
  if (v >= 1000000000) return (v / 1000000000).toFixed(1) + "B";
  if (v >= 1000000) return (v / 1000000).toFixed(1) + "M";
  return v.toLocaleString();
}

function mean(arr) {
  if (!arr.length) return 0;
  return arr.reduce(function(a, b) { return a + b; }, 0) / arr.length;
}

function median(arr) {
  if (!arr.length) return 0;
  var sorted = arr.slice().sort(function(a, b) { return a - b; });
  var mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function extractYear(dateStr) {
  if (!dateStr) return null;
  try { return new Date(dateStr).getFullYear(); } catch(e) { return null; }
}

function isValidTender(t) {
  if (!t || !t.ocid || !t.title) return false;
  return true;
}

function computeDaysToDeadline(deadlineStr) {
  if (!deadlineStr) return "";
  try {
    var deadline = new Date(deadlineStr);
    var now = new Date();
    var diff = deadline.getTime() - now.getTime();
    return Math.ceil(diff / (1000 * 60 * 60 * 24));
  } catch(e) {
    return "";
  }
}

function logActivity(userId, action, detail) {
  try {
    var sheet = getSheet().getSheetByName("USER_ACTIVITY");
    if (!sheet) return;
    sheet.appendRow([sanitizeSheet(userId), sanitizeSheet(action), sanitizeSheet(detail || ""), new Date()]);
  } catch(e) {}
}

function sanitizeSheet(val) {
  var s = String(val || "");
  if (s.length > 0 && (s.charAt(0) === '=' || s.charAt(0) === '+' || s.charAt(0) === '-' || s.charAt(0) === '@')) {
    return "'" + s;
  }
  return s;
}

function testEcho() {
  return { hello: "world", num: 42, sheetId: PropertiesService.getScriptProperties().getProperty("SHEET_ID") || "NOT SET" };
}

// ── PHASE 1 TESTS ────────────────────────────────────────────────────────────

function runPhase1Tests() {
  Logger.log("========================================");
  Logger.log("  PHASE 1 — FOUNDATION TESTS");
  Logger.log("========================================");

  var passed = 0, failed = 0;

  function test(name, condition, detail) {
    if (condition) { passed++; Logger.log("PASS: " + name + (detail ? " → " + detail : "")); }
    else { failed++; Logger.log("FAIL: " + name + (detail ? " → " + detail : "")); }
  }

  // 1. Script Properties
  var sheetId = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  test("SHEET_ID is set", !!sheetId, sheetId || "NOT SET");

  // 2. Sheet access
  try {
    var ss = getSheet();
    test("getSheet() returns spreadsheet", !!ss, ss.getName());
  } catch(e) { test("getSheet() returns spreadsheet", false, e.message); }

  // 3. Required sheets exist
  var required = ["TENDERS_FLAT","USER_PROFILES","MATCHES","BUYER_PROFILES","SAVED_TENDERS","USER_ACTIVITY"];
  required.forEach(function(name) {
    try {
      var s = ss.getSheetByName(name);
      test("Sheet: " + name, !!s, s.getLastRow() + " rows");
    } catch(e) { test("Sheet: " + name, false, e.message); }
  });

  // 4. Tenders data
  try {
    var tSheet = ss.getSheetByName("TENDERS_FLAT");
    var rows = tSheet.getLastRow() - 1;
    test("TENDERS_FLAT has data", rows > 0, rows + " tenders");
    test("TENDERS_FLAT ≥ 15 tenders", rows >= 15, rows + " of 15 expected");
  } catch(e) { test("TENDERS_FLAT has data", false, e.message); }

  // 5. User ID
  try {
    var uid = getCurrentUserId();
    test("getCurrentUserId()", !!uid, uid);
  } catch(e) { test("getCurrentUserId()", false, e.message); }

  // 6. Dashboard data
  try {
    var d = getDashboardData();
    test("getDashboardData() returns object", !!d, "typeof: " + typeof d);
    test("Dashboard has activeTenders", d.activeTenders > 0, d.activeTenders + " tenders");
    test("Dashboard has matchCount", typeof d.matchCount === "number", d.matchCount + " matches");
    test("Dashboard has savedCount", typeof d.savedCount === "number", d.savedCount + " saved");
    test("Dashboard has tier", !!d.tier, d.tier);
    test("Dashboard has recentActivity", Array.isArray(d.recentActivity), d.recentActivity.length + " items");
  } catch(e) { test("getDashboardData()", false, e.message); }

  // 7. Web app routing
  try {
    var r1 = doGet({ parameter: { page: "dashboard" } });
    test("doGet(?page=dashboard) returns HTML", !!r1, "type: " + typeof r1);
    var r2 = doGet({});
    test("doGet() default returns HTML", !!r2, "type: " + typeof r2);
  } catch(e) { test("doGet()", false, e.message); }

  // 8. Utilities
  test("mapRow() exists", typeof mapRow === "function");
  test("safeLower() exists", typeof safeLower === "function");
  test("formatMoney() exists", typeof formatMoney === "function");
  test("mean() exists", typeof mean === "function");
  test("median() exists", typeof median === "function");
  test("isValidTender() exists", typeof isValidTender === "function");

  Logger.log("========================================");
  Logger.log("  RESULTS: " + passed + " passed, " + failed + " failed");
  Logger.log("========================================");
}

// =============================================================================
//  PHASE 2 — BUSINESS DNA & MATCHING
// =============================================================================

// ── USER PROFILE ──────────────────────────────────────────────────────────────

function getUserProfile(userId) {
  try {
    var sheet = getSheet().getSheetByName("USER_PROFILES");
    if (!sheet || sheet.getLastRow() < 2) return null;
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0] || "").toLowerCase().trim() === userId.toLowerCase().trim()) {
        return mapRow(headers, data[i]);
      }
    }
  } catch(e) { Logger.log("getUserProfile: " + e); }
  return null;
}

// ── BUSINESS DNA ──────────────────────────────────────────────────────────────

function saveBusinessDNA(data, uid) {
  try {
    var userId = getCurrentUserId(uid);
    if (!userId || userId === "test_user") {
      throw new Error("Could not identify your account. Make sure you are signed into Google.");
    }

    var sheet = getSheet().getSheetByName("USER_PROFILES");
    if (!sheet) throw new Error("USER_PROFILES sheet not found. Run Setup from the menu first.");

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var allData = sheet.getDataRange().getValues();
    var filtered = allData.filter(function(r, i) { return i === 0 || (r[0] || "").toLowerCase() !== userId.toLowerCase(); });
    sheet.clearContents();
    if (filtered.length) sheet.getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);

    var extendedHeaders = [
      "user_id","company_name","tin","year_established","company_size",
      "is_women_led","is_youth_led","rdb_registered",
      "primary_sector","capabilities","past_gov_contracts","past_gov_count","past_gov_total_value",
      "certifications","min_value","max_value","preferred_methods","regions","timeline_urgency",
      "risk_appetite","growth_orientation","decision_speed","learning_mindset",
      "competitive_posture","price_sensitivity",
      "max_team_size","equipment","cash_flow_limit","pref_single_bid",
      "sectors_of_interest","language_preference","created_at"
    ];

    if (sheet.getLastRow() === 0 || headers.filter(function(h) { return h; }).length < 5) {
      sheet.getRange(1, 1, 1, extendedHeaders.length).setValues([extendedHeaders]);
      headers = extendedHeaders;
    }

    var rowArr = new Array(Math.max(headers.length, extendedHeaders.length)).fill("");
    function set(name, value) {
      var i = headers.indexOf(name);
      if (i >= 0) rowArr[i] = (typeof value === 'string') ? sanitizeSheet(value) : (value || "");
    }

    set("user_id", userId);
    set("company_name", data.company_name || "");
    set("tin", data.tin || "");
    set("year_established", data.year_established || "");
    set("company_size", data.company_size || "");
    set("is_women_led", data.is_women_led || "");
    set("is_youth_led", data.is_youth_led || "");
    set("rdb_registered", data.rdb_registered || "");
    set("primary_sector", data.primary_sector || "");
    set("capabilities", data.capabilities || "");
    set("past_gov_contracts", data.past_gov_contracts || "");
    set("past_gov_count", data.past_gov_count || 0);
    set("past_gov_total_value", data.past_gov_total_value || 0);
    set("certifications", data.certifications || "");
    set("min_value", data.min_value || 0);
    set("max_value", data.max_value || 0);
    set("preferred_methods", data.preferred_methods || "");
    set("regions", data.regions || "");
    set("timeline_urgency", data.timeline_urgency || "");
    set("risk_appetite", data.risk_appetite || "balanced");
    set("growth_orientation", data.growth_orientation || "steady");
    set("decision_speed", data.decision_speed || "careful");
    set("learning_mindset", data.learning_mindset || "improve");
    set("competitive_posture", data.competitive_posture || "comfortable");
    set("price_sensitivity", data.price_sensitivity || "competitive");
    set("max_team_size", data.max_team_size || "");
    set("equipment", data.equipment || "");
    set("cash_flow_limit", data.cash_flow_limit || "");
    set("pref_single_bid", data.pref_single_bid || "");
    set("sectors_of_interest", data.primary_sector || "");
    set("language_preference", "");
    set("created_at", new Date());

    sheet.appendRow(rowArr.slice(0, headers.length));
    logActivity(userId, "dna_saved");
    try { grantTrial(userId); } catch(e) {}
    try { generateMatchesNoAI(); } catch(e) { Logger.log("Match gen after DNA: " + e); }

    return "Profile saved! Redirecting to your dashboard...";
  } catch(e) {
    Logger.log("saveBusinessDNA: " + e);
    throw e;
  }
}

// ── MATCHING ENGINE ───────────────────────────────────────────────────────────

function computeCheapScore(t, profile) {
  var score = 0;
  var cleanTitle = getCleanTitle(t);

  if (profile.sectors_of_interest && cleanTitle) {
    var sectors = profile.sectors_of_interest.toLowerCase().split(",").map(function(s) { return s.trim(); });
    var title = cleanTitle.toLowerCase();
    var desc = String(t.description || "").toLowerCase();
    var cat = String(t.item_classifications || "").toLowerCase();
    var combined = title + " " + desc + " " + cat + " " + (t.entity || "").toLowerCase();
    var matched = sectors.filter(function(s) { return s && combined.indexOf(s) !== -1; });
    if (matched.length > 0) score += 40;
    else score += 5;
  } else {
    score += 15;
  }

  if (profile.preferred_methods && t.method) {
    var prefM = profile.preferred_methods.toLowerCase();
    var tMethod = t.method.toLowerCase();
    if (prefM.indexOf(tMethod) !== -1 || tMethod.indexOf(prefM.split(",")[0].trim()) !== -1) score += 15;
  }

  if (profile.regions && t.entity) {
    var regions = profile.regions.toLowerCase().split(",").map(function(s) { return s.trim(); });
    var entity = (t.entity || "").toLowerCase();
    if (regions.some(function(r) { return entity.indexOf(r) !== -1; })) score += 8;
  }

  var val = Number(t.est_value) || 0;
  var minV = Number(profile.min_value) || 0;
  var maxV = Number(profile.max_value) || 0;
  if (maxV > 0 && val >= minV && val <= maxV) score += 12;
  else if (minV > 0 && val >= minV * 0.5) score += 5;

  var tenderers = Number(t.num_tenderers);
  if (!isNaN(tenderers)) {
    if (tenderers === 1) score += 20;
    else if (tenderers <= 2) score += 15;
    else if (tenderers <= 4) score += 8;
    else if (tenderers >= 8) score -= 5;
  }

  var days = Number(t.days_to_deadline);
  if (!isNaN(days) && days >= 0 && days <= 7) score += 8;

  return Math.max(5, Math.min(98, score));
}

function generateMatchesNoAI(uid) {
  try {
    var userId = getCurrentUserId(uid);
    var profile = getUserProfile(userId);
    if (!profile) throw new Error("No user profile found. Complete your Business DNA first.");

    var tSheet = getSheet().getSheetByName("TENDERS_FLAT");
    var tData = tSheet.getDataRange().getValues();
    if (tData.length < 2) throw new Error("No tender data. Run Seed Sample Tenders first.");

    var tHeaders = tData[0];
    var rows = tData.slice(1);
    var minVal = parseFloat(profile.min_value) || 0;
    var maxVal = parseFloat(profile.max_value) || 0;

    var candidates = [];
    for (var i = 0; i < rows.length; i++) {
      var t = mapRow(tHeaders, rows[i]);
      if (!isValidTender(t)) continue;
      var tValue = parseFloat(t.est_value) || 0;
      if (tValue > 0 && minVal > 0 && maxVal > 0 && (tValue < minVal || tValue > maxVal)) continue;
      t.source = "umucyo";
      candidates.push(t);
      if (candidates.length >= 500) break;
    }

    // Also include JobInRwanda scraped tenders
    var jirSheet = getSheet().getSheetByName("JIR_TENDERS");
    if (jirSheet && jirSheet.getLastRow() > 1) {
      var jirData = jirSheet.getDataRange().getValues();
      var jirAdded = 0;
      for (var ji = 1; ji < jirData.length && jirAdded < 30; ji++) {
        var jr = jirData[ji];
        var jirTitle = String(jr[0] || "");
        var jirEmployer = String(jr[1] || "");
        var jirLocation = String(jr[2] || "");
        var jirPublished = String(jr[3] || "");
        var jirDeadline = String(jr[4] || "");
        var jirType = String(jr[5] || "Tender");
        var jirLink = String(jr[6] || "");
        if (!jirTitle) continue;
        candidates.push({
          ocid: jirLink, title: jirTitle, entity: jirEmployer, description: "JobInRwanda " + jirType + " · Location: " + jirLocation,
          method: "", est_value: 0, publish_date: jirPublished, deadline: jirDeadline,
          num_tenderers: 0, item_classifications: "", lot_count: 1, lot_value: 0,
          source: "jobinrwanda"
        });
      }
    }

    var scored = candidates.map(function(t) {
      return { t: t, score: computeCheapScore(t, profile) };
    });
    scored.sort(function(a, b) { return b.score - a.score; });

    var MAX = Math.min(100, scored.length);
    var matches = [];
    for (var j = 0; j < MAX; j++) {
      var ct = scored[j].t;
      var score = scored[j].score;
      var tag = score >= 6 ? "Top Match" : (score >= 3 ? "Good Fit" : "Explore");
      matches.push({
        ocid: ct.ocid, entity: ct.entity, title: ct.title,
        description: ct.description || "", method: ct.method,
        est_value: ct.est_value, publish_date: ct.publish_date,
        deadline: ct.deadline, num_tenderers: ct.num_tenderers,
        match_score: score, win_chance: score + "%",
        win_signals: "Cheap score: " + score, buyer_snapshot: "",
        competitive_context: "", decision_guidance: "",
        risk_flags: "", checklist: "", ai_explainer: "",
        tag: tag, generated_at: new Date()
      });
    }

    writeMatches(userId, matches);
    Logger.log("Generated " + matches.length + " matches (no AI)");
    return matches.length;
  } catch(e) {
    Logger.log("generateMatchesNoAI: " + e);
    return 0;
  }
}

function writeMatches(userId, matches) {
  var sheet = getSheet().getSheetByName("MATCHES");
  clearUserMatches(userId);

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIndex = function(name) { return headers.indexOf(name); };

  var rows = matches.map(function(m) {
    var row = new Array(headers.length).fill("");
    row[colIndex("user_id")] = userId;
    row[colIndex("ocid")] = m.ocid;
    row[colIndex("entity")] = m.entity;
    row[colIndex("title")] = m.title;
    row[colIndex("description")] = m.description || "";
    row[colIndex("method")] = m.method;
    row[colIndex("est_value")] = m.est_value;
    row[colIndex("publish_date")] = m.publish_date;
    row[colIndex("deadline")] = m.deadline;
    row[colIndex("num_tenderers")] = m.num_tenderers;
    row[colIndex("match_score")] = m.match_score;
    row[colIndex("win_chance")] = m.win_chance || "";
    row[colIndex("win_signals")] = m.win_signals || "";
    row[colIndex("buyer_snapshot")] = m.buyer_snapshot || "";
    row[colIndex("competitive_context")] = m.competitive_context || "";
    row[colIndex("decision_guidance")] = m.decision_guidance || "";
    row[colIndex("risk_flags")] = m.risk_flags || "";
    row[colIndex("checklist")] = m.checklist || "";
    row[colIndex("ai_explainer")] = m.ai_explainer || "";
    row[colIndex("tag")] = m.tag || "";
    row[colIndex("generated_at")] = m.generated_at;
    row[colIndex("win_probability")] = m.win_probability || "";
    row[colIndex("expected_bidders")] = m.expected_bidders || "";
    row[colIndex("recommended_bid")] = m.recommended_bid || "";
    row[colIndex("top_competitors")] = m.top_competitors || "";
    row[colIndex("buyer_contact_phone")] = m.buyer_contact_phone || "";
    row[colIndex("buyer_contact_email")] = m.buyer_contact_email || "";
    row[colIndex("buyer_amendment_pct")] = m.buyer_amendment_pct || "";
    return row;
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }
}

function clearUserMatches(userId) {
  var sheet = getSheet().getSheetByName("MATCHES");
  var data = sheet.getDataRange().getValues();
  var filtered = data.filter(function(r, i) { return i === 0 || r[0] !== userId; });
  sheet.clearContents();
  var maxCols = filtered.length > 0 ? filtered[0].length : 0;
  if (filtered.length > 0 && maxCols > 0) {
    sheet.getRange(1, 1, filtered.length, maxCols).setValues(filtered);
  }
  if (sheet.getMaxRows() > filtered.length + 5) {
    sheet.deleteRows(filtered.length + 1, sheet.getMaxRows() - filtered.length);
  }
}

// ── WEB DATA FUNCTIONS ───────────────────────────────────────────────────────

function getMatchesForWeb(userId) {
  try {
    if (!userId) userId = getCurrentUserId();
    var sheet = getSheet().getSheetByName("MATCHES");
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = function(name) { return headers.indexOf(name); };

    var now = new Date();
    var rows = data.slice(1)
      .filter(function(r) { return String(r[col("user_id")] || "").toLowerCase() === userId.toLowerCase(); })
      .sort(function(a, b) {
        var dlA = a[col("deadline")] ? new Date(a[col("deadline")]) : null;
        var dlB = b[col("deadline")] ? new Date(b[col("deadline")]) : null;
        var activeA = dlA && dlA >= now;
        var activeB = dlB && dlB >= now;
        if (activeA && !activeB) return -1;
        if (!activeA && activeB) return 1;
        if (activeA && activeB) {
          var urgencyA = (dlA - now) / (1000 * 3600 * 24);
          var urgencyB = (dlB - now) / (1000 * 3600 * 24);
          if (urgencyA <= 7 && urgencyB > 7) return -1;
          if (urgencyB <= 7 && urgencyA > 7) return 1;
        }
        var scoreA = parseInt(b[col("match_score")]) || 0;
        var scoreB = parseInt(a[col("match_score")]) || 0;
        return scoreA - scoreB;
      })
      .slice(0, 100);

    return rows.map(function(r) {
      return {
        ocid: String(r[col("ocid")] || ""),
        title: String(r[col("title")] || ""),
        entity: String(r[col("entity")] || ""),
        method: String(r[col("method")] || ""),
        est_value: Number(r[col("est_value")] || 0),
        deadline: String(r[col("deadline")] || ""),
        num_tenderers: Number(r[col("num_tenderers")] || 0),
        relevance: parseInt(r[col("match_score")]) || 0,
        match_score: Number(r[col("match_score")] || 0),
        win_chance: String(r[col("win_chance")] || ""),
        tag: String(r[col("tag")] || ""),
        win_probability: parseFloat(r[col("win_probability")]) || null,
        expected_bidders: parseFloat(r[col("expected_bidders")]) || null,
        recommended_bid: parseFloat(r[col("recommended_bid")]) || null,
        top_competitors: String(r[col("top_competitors")] || ""),
        buyer_contact_phone: String(r[col("buyer_contact_phone")] || ""),
        buyer_contact_email: String(r[col("buyer_contact_email")] || ""),
        buyer_amendment_pct: parseFloat(r[col("buyer_amendment_pct")]) || null
      };
    });
  } catch(e) {
    Logger.log("getMatchesForWeb: " + e);
    return [];
  }
}

function searchTenders(query, filters, uid) {
  try {
    var userId = getCurrentUserId(uid);
    var profile = getUserProfile(userId);

    var sheet = getSheet().getSheetByName("TENDERS_FLAT");
    if (!sheet || sheet.getLastRow() < 2) return [];

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var useDNA = filters && filters.use_dna;
    var minVal = parseFloat((filters && filters.min_value) || 0) || 0;
    var maxVal = parseFloat((filters && filters.max_value) || 0) || 0;
    var method = ((filters && filters.method) || "").toLowerCase().trim();
    var maxDays = parseInt((filters && filters.max_days) || 0) || 0;

    var queryLower = (query || "").toLowerCase().trim();
    var sectors = [];
    if (useDNA && profile) {
      if (!minVal) minVal = parseFloat(profile.min_value) || 0;
      if (!maxVal) maxVal = parseFloat(profile.max_value) || 0;
      if (!method) method = safeLower(profile.preferred_methods || "").split(",")[0].trim();
      sectors = safeLower(profile.sectors_of_interest || profile.primary_sector || "").split(",").map(function(s) { return s.trim(); }).filter(Boolean);
    }

    var results = [];
    for (var i = 1; i < data.length; i++) {
      var t = mapRow(headers, data[i]);
      if (!isValidTender(t)) continue;

      var val = parseFloat(t.est_value) || 0;
      if (minVal > 0 && val < minVal) continue;
      if (maxVal > 0 && val > maxVal) continue;
      if (method && t.method && safeLower(t.method).indexOf(method) === -1) continue;

      var days = Number(t.days_to_deadline);
      if (maxDays > 0 && (!isNaN(days)) && days > maxDays) continue;

      var combined = safeLower(t.title || "") + " " + safeLower(t.description || "") + " " + safeLower(t.entity || "");

      var textMatch = !queryLower || combined.indexOf(queryLower) !== -1;
      var dnaMatch = sectors.length === 0 || sectors.some(function(s) { return s && combined.indexOf(s) !== -1; });

      if (!textMatch && !dnaMatch) continue;

      var relevance = 40;
      if (queryLower && combined.indexOf(queryLower) !== -1) relevance += 20;
      if (sectors.some(function(s) { return s && combined.indexOf(s) !== -1; })) relevance += 20;
      var tenderers = Number(t.num_tenderers);
      if (!isNaN(tenderers) && tenderers <= 2) relevance += 15;
      else if (!isNaN(tenderers) && tenderers >= 8) relevance -= 10;
      if (minVal > 0 && maxVal > 0 && val >= minVal && val <= maxVal) relevance += 10;

      results.push({
        ocid: String(t.ocid || ""), title: String(t.title || ""), entity: String(t.entity || ""),
        est_value: Number(t.est_value || 0), method: String(t.method || ""),
        deadline: String(t.deadline || ""), num_tenderers: Number(t.num_tenderers || 0),
        relevance: Math.min(95, Math.max(5, relevance))
      });

      if (results.length >= 50) break;
    }

    // Also search JIR_TENDERS
    try {
      var jirSheet = getSheet().getSheetByName("JIR_TENDERS");
      if (jirSheet && jirSheet.getLastRow() > 1) {
        var jirData = jirSheet.getDataRange().getValues();
        var jirAdded = 0;
        for (var ji = 1; ji < jirData.length && jirAdded < 15; ji++) {
          var jr = jirData[ji];
          var jirTitle = String(jr[0] || "");
          var jirEmployer = String(jr[1] || "");
          var jirDeadline = String(jr[4] || "");
          if (!jirTitle) continue;
          var combined = safeLower(jirTitle) + " " + safeLower(jirEmployer);
          var textMatch = !queryLower || combined.indexOf(queryLower) !== -1;
          var dnaMatch = sectors.length === 0 || sectors.some(function(s) { return s && combined.indexOf(s) !== -1; });
          if (!textMatch && !dnaMatch) continue;
          var relevance = 30;
          if (queryLower && combined.indexOf(queryLower) !== -1) relevance += 20;
          if (sectors.some(function(s) { return s && combined.indexOf(s) !== -1; })) relevance += 20;
          results.push({
            ocid: String(jr[6] || ""), title: jirTitle, entity: jirEmployer,
            est_value: 0, method: "", deadline: jirDeadline, num_tenderers: 0,
            relevance: Math.min(90, Math.max(5, relevance)), source: "jobinrwanda"
          });
        }
      }
    } catch(e) {}

    results.sort(function(a, b) { return b.relevance - a.relevance; });
    logActivity(userId, "search");
    return results.slice(0, 30);
  } catch(e) {
    Logger.log("searchTenders: " + e);
    return [];
  }
}

// ── TOGGLE SAVE ──────────────────────────────────────────────────────────────

function toggleSave(ocid, uid) {
  var userId = getCurrentUserId(uid);
  var ss = getSheet();

  var sheet = ss.getSheetByName("SAVED_TENDERS");
  if (!sheet) {
    sheet = ss.insertSheet("SAVED_TENDERS");
    sheet.getRange(1, 1, 1, 3).setValues([["user_id","ocid","saved_at"]]);
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || "").toLowerCase() === userId.toLowerCase() && data[i][1] === ocid) {
      sheet.deleteRow(i + 1);
      return "unsaved";
    }
  }

  sheet.appendRow([userId, ocid, new Date()]);
  logActivity(userId, "saved");
  return "saved";
}

// ── CLEAN TITLE ──────────────────────────────────────────────────────────────

function getCleanTitle(t) {
  var title = t.title || "";
  var desc = t.description || "";
  if (title.length > 10 && title.indexOf(" ") !== -1) return title;
  return desc.length > 10 ? desc : title;
}

function getSavedTendersForWeb(uid) {
  try {
    var userId = getCurrentUserId(uid);
    var sSheet = getSheet().getSheetByName("SAVED_TENDERS");
    if (!sSheet || sSheet.getLastRow() < 2) return [];
    var sData = sSheet.getDataRange().getValues();
    var ocids = sData.slice(1).filter(function(r) { return String(r[0] || "").toLowerCase() === userId.toLowerCase(); }).map(function(r) { return String(r[1] || ""); });
    if (!ocids.length) return [];

    var results = [];
    var sources = [
      { sheet: getSheet().getSheetByName("TENDERS_FLAT"), headers: null }
    ];
    var jirSheet = getSheet().getSheetByName("JIR_TENDERS");

    sources.forEach(function(src) {
      if (!src.sheet || src.sheet.getLastRow() < 2) return;
      var data = src.sheet.getDataRange().getValues();
      var headers = data[0];
      for (var i = 1; i < data.length && results.length < 50; i++) {
        var row = data[i];
        var rowOcid = String(row[headers.indexOf("ocid")] || row[0] || "");
        var link = String(row[headers.indexOf("Link")] || row[6] || "");
        if (ocids.indexOf(rowOcid) === -1 && ocids.indexOf(link) === -1) continue;
        results.push({
          ocid: rowOcid || link, title: String(row[1] || row[headers.indexOf("title")] || "Saved"),
          entity: String(row[2] || row[headers.indexOf("entity")] || ""),
          est_value: Number(row[headers.indexOf("est_value")] || row[7] || 0),
          method: String(row[headers.indexOf("method")] || ""),
          deadline: String(row[headers.indexOf("deadline")] || row[4] || ""),
          num_tenderers: Number(row[headers.indexOf("num_tenderers")] || 0),
          relevance: 70, tag: "saved"
        });
      }
    });

    if (!results.length) {
      return sData.slice(1).filter(function(r) { return String(r[0] || "").toLowerCase() === userId.toLowerCase(); }).map(function(r) {
        return { ocid: String(r[1] || ""), title: String(r[2] || "Saved tender"), entity: "", est_value: 0, method: "", deadline: "", num_tenderers: 0, relevance: 50, tag: "saved" };
      }).slice(0, 50);
    }

    return results;
  } catch(e) { Logger.log("getSavedTendersForWeb: " + e); return []; }
}

function getNewThisWeekForWeb(uid) {
  try {
    var userId = getCurrentUserId(uid);
    var oneWeekAgo = new Date(Date.now() - 7 * 24 * 3600 * 1000);
    var results = [];
    var sources = [
      { sheet: getSheet().getSheetByName("TENDERS_FLAT"), isRppa: true },
      { sheet: getSheet().getSheetByName("JIR_TENDERS"), isRppa: false }
    ];
    sources.forEach(function(src) {
      if (!src.sheet || src.sheet.getLastRow() < 2) return;
      var data = src.sheet.getDataRange().getValues();
      var headers = data[0];
      for (var i = 1; i < data.length && results.length < 50; i++) {
        var row = data[i];
        var pubDate = src.isRppa ? (row[headers.indexOf("publish_date")] || "") : (row[3] || "");
        if (!pubDate) continue;
        var pub = new Date(pubDate);
        if (pub < oneWeekAgo) continue;
        results.push({
          ocid: src.isRppa ? String(row[headers.indexOf("ocid")] || "") : String(row[6] || ""),
          title: src.isRppa ? String(row[headers.indexOf("title")] || "") : String(row[0] || ""),
          entity: src.isRppa ? String(row[headers.indexOf("entity")] || "") : String(row[1] || ""),
          est_value: src.isRppa ? Number(row[headers.indexOf("est_value")] || 0) : 0,
          method: src.isRppa ? String(row[headers.indexOf("method")] || "") : "",
          deadline: src.isRppa ? String(row[headers.indexOf("deadline")] || "") : String(row[4] || ""),
          num_tenderers: src.isRppa ? Number(row[headers.indexOf("num_tenderers")] || 0) : 0,
          relevance: 60
        });
      }
    });
    return results;
  } catch(e) { Logger.log("getNewThisWeekForWeb: " + e); return []; }
}

// ── PHASE 2 TESTS ────────────────────────────────────────────────────────────

function testPhase2Functions() {
  Logger.log("=== PHASE 2 DIAGNOSTICS ===");
  Logger.log("userId: " + getCurrentUserId());

  var profile = getUserProfile(getCurrentUserId());
  Logger.log("Profile: " + (profile ? JSON.stringify({ company: profile.company_name, sector: profile.primary_sector, budget: profile.min_value + "-" + profile.max_value }) : "NONE"));

  var matches = getMatchesForWeb("");
  Logger.log("Matches for web (empty user): " + matches.length);
  if (matches.length > 0) {
    Logger.log("  First match: " + JSON.stringify({ title: matches[0].title, relevance: matches[0].relevance }));
  }

  var search = searchTenders("hospital", {});
  Logger.log("Search 'hospital': " + search.length + " results");

  var search2 = searchTenders("", {});
  Logger.log("Search empty: " + search2.length + " results");

  var dnaSearch = searchTenders("", { use_dna: true });
  Logger.log("Search DNA: " + dnaSearch.length + " results");

  Logger.log("=== END DIAGNOSTICS ===");
}

// =============================================================================
//  PHASE 3 — TENDER LOOKUP & COMPLIANCE
// =============================================================================

// ── TENDER LOOKUP ────────────────────────────────────────────────────────────

function lookupTender(identifier, uid) {
  try {
    var userId = getCurrentUserId(uid);
    var profile = getUserProfile(userId);

    if (!identifier || String(identifier).trim() === "") {
      return { found: false, message: "Please enter an OCID, tender title, or keyword." };
    }

    var sheet = getSheet().getSheetByName("TENDERS_FLAT");
    if (!sheet || sheet.getLastRow() < 2) {
      return { found: false, message: "No tender data available." };
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var query = String(identifier).trim().toLowerCase();
    var found = null;

    for (var i = 1; i < data.length; i++) {
      var t = mapRow(headers, data[i]);
      var ocid = safeLower(String(t.ocid || ""));
      var title = safeLower(String(t.title || ""));
      var desc = safeLower(String(t.description || ""));

      if (ocid === query || ocid.indexOf(query) !== -1 ||
          title.indexOf(query) !== -1 || desc.indexOf(query) !== -1) {
        found = t;
        break;
      }
    }

    if (!found) {
      var jirSheet = getSheet().getSheetByName("JIR_TENDERS");
      if (jirSheet && jirSheet.getLastRow() > 1) {
        var jirData = jirSheet.getDataRange().getValues();
        for (var ji = 1; ji < jirData.length; ji++) {
          var jirTitle = safeLower(String(jirData[ji][0] || ""));
          var jirEmployer = safeLower(String(jirData[ji][1] || ""));
          var jirOcid = safeLower(String(jirData[ji][6] || ""));
          if (jirOcid === query || jirOcid.indexOf(query) !== -1 ||
              jirTitle.indexOf(query) !== -1 || jirEmployer.indexOf(query) !== -1) {
            found = {
              ocid: String(jirData[ji][6] || ""), title: String(jirData[ji][0] || ""),
              entity: String(jirData[ji][1] || ""), method: "", status: "active",
              est_value: 0, deadline: String(jirData[ji][4] || ""),
              num_tenderers: 0, description: "Source: JobInRwanda.com · Published: " + String(jirData[ji][3] || "") + " · Location: " + String(jirData[ji][2] || ""),
              source: "jobinrwanda"
            };
            break;
          }
        }
      }
    }

    if (!found) {
      return { found: false, message: "No tender found for \"" + identifier + "\". Try pasting the full title or OCID, or paste the tender text below for AI analysis." };
    }

    var relevance = 50;
    if (profile) {
      var sectors = safeLower(profile.sectors_of_interest || profile.primary_sector || "").split(",").map(function(s) { return s.trim(); });
      var combined = safeLower(String(found.title || "")) + " " + safeLower(String(found.description || ""));
      var sectorMatch = sectors.some(function(s) { return s && combined.indexOf(s) !== -1; });
      if (sectorMatch) relevance += 20;

      var val = Number(found.est_value) || 0;
      var minV = Number(profile.min_value) || 0;
      var maxV = Number(profile.max_value) || 0;
      if (maxV > 0 && val >= minV && val <= maxV) relevance += 15;
      else if (minV > 0 && val >= minV) relevance += 8;

      var tenderers = Number(found.num_tenderers);
      if (!isNaN(tenderers) && tenderers <= 2) relevance += 15;
    }
    relevance = Math.min(95, Math.max(5, relevance));

    var buyerSnapshot = formatBuyerSnapshot(String(found.entity || ""));
    var predictions = getFullTenderEnrichment(String(found.entity || ""), String(found.method || ""), Number(found.est_value) || 0);
    logActivity(userId, "lookup");

    return {
      found: true,
      tender: {
        ocid: String(found.ocid || ""),
        title: String(found.title || ""),
        entity: String(found.entity || ""),
        method: String(found.method || ""),
        status: String(found.status || ""),
        est_value: Number(found.est_value || 0),
        deadline: String(found.deadline || ""),
        num_tenderers: Number(found.num_tenderers || 0),
        description: String(found.description || "")
      },
      relevance: relevance,
      buyer_snapshot: buyerSnapshot,
      predictions: predictions
    };
  } catch(e) {
    Logger.log("lookupTender: " + e);
    return { found: false, message: "Error looking up tender: " + e };
  }
}

// ── AI TEXT ANALYSIS ─────────────────────────────────────────────────────────

function analyzeTenderText(text, uid) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY");
    if (!apiKey) return "AI not configured. Add OPENROUTER_API_KEY to Script Properties.";

    var userId = getCurrentUserId(uid);
    var profile = getUserProfile(userId);
    var profileContext = profile ? "Sector: " + (profile.primary_sector || "") + ", Budget: " + formatMoney(profile.min_value) + "-" + formatMoney(profile.max_value) + " RWF" : "";

    var prompt = [
      "You are a procurement advisor for an SME in Rwanda. Analyze this tender text and give a practical assessment.",
      "Return exactly this format:",
      "What it is: [1 sentence]",
      "Who should bid: [1-2 sentences]",
      "Key requirements: [3-5 bullet points with •]",
      "Relevance to this user: [1 sentence]",
      profileContext ? "User context: " + profileContext : "",
      "Recommendation: BID / MAYBE / SKIP with 1 sentence reason",
      "Keep under 200 words.",
      "Tender text:",
      String(text).substring(0, 2000)
    ].filter(Boolean).join("\n");

    var res = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "post",
      headers: { "Authorization": "Bearer " + apiKey, "Content-Type": "application/json" },
      payload: JSON.stringify({
        model: "openrouter/auto",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.3, max_tokens: 350
      }),
      muteHttpExceptions: true
    });

    var json = JSON.parse(res.getContentText());
    return String((json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || "AI unavailable. Try again.");
  } catch(e) {
    Logger.log("analyzeTenderText: " + e);
    return "AI analysis failed: " + e;
  }
}

// ── COMPLIANCE CHECKER ────────────────────────────────────────────────────────

function checkCompliance(proposal, context, uid) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY");
    if (!apiKey) return "AI not configured. Add OPENROUTER_API_KEY to Script Properties.";
    if (!proposal || String(proposal).trim().length < 20) return "Please paste your proposal text (at least 20 characters).";

    var prompt = [
      "You are a procurement compliance reviewer for Rwandan public tenders.",
      "Review this proposal against the tender context and flag any issues.",
      "Return EXACTLY this format (copy the labels exactly):",
      "Coverage: [percentage e.g. 75%]",
      "Missing: [list of missing requirements or 'None']",
      "Weak: [list of weak sections or 'None']",
      "Red flag: [list of red flags or 'None']",
      "Quick fix: [1-2 sentences on how to fix the biggest gap]",
      context ? "Tender context: " + String(context).substring(0, 500) : "",
      "Proposal:",
      String(proposal).substring(0, 3000),
      "Keep each section 1-3 lines. Under 300 words total."
    ].filter(Boolean).join("\n");

    var res = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "post",
      headers: { "Authorization": "Bearer " + apiKey, "Content-Type": "application/json" },
      payload: JSON.stringify({
        model: "openrouter/auto",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.2, max_tokens: 500
      }),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    if (code !== 200) {
      Logger.log("checkCompliance HTTP " + code + ": " + res.getContentText().substring(0, 200));
      return "Coverage: Error\nMissing: API returned HTTP " + code + ". Check OpenRouter credits at https://openrouter.ai/settings/keys\nWeak: N/A\nRed flag: N/A\nQuick fix: Add credits to your OpenRouter account.";
    }
    var json = JSON.parse(res.getContentText());
    logActivity(getCurrentUserId(uid), "compliance_check");
    return String((json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || "Coverage: No response\nMissing: AI returned empty\nWeak: N/A\nRed flag: N/A\nQuick fix: Try again.");
  } catch(e) {
    Logger.log("checkCompliance: " + e);
    return "Compliance check failed: " + e;
  }
}

// ── TENDER Q&A ────────────────────────────────────────────────────────────────

function askAboutTender(question, ocid, uid) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY");
    if (!apiKey) return "AI not configured. Add OPENROUTER_API_KEY to Script Properties.";
    if (!question || String(question).trim().length < 3) return "Please ask a question.";

    var tenderContext = "";
    if (ocid) {
      var sheet = getSheet().getSheetByName("TENDERS_FLAT");
      if (sheet && sheet.getLastRow() > 1) {
        var data = sheet.getDataRange().getValues();
        var headers = data[0];
        for (var i = 1; i < data.length; i++) {
          if (String(data[i][headers.indexOf("ocid")] || "") === ocid) {
            var t = mapRow(headers, data[i]);
            tenderContext = "Tender: " + String(t.title || "") + "\nEntity: " + String(t.entity || "") + "\nValue: " + formatMoney(t.est_value) + " RWF\nMethod: " + String(t.method || "") + "\nDeadline: " + String(t.deadline || "") + "\nDescription: " + String(t.description || "").substring(0, 300);
            break;
          }
        }
      }
    }

    var prompt = [
      "You are a procurement advisor helping an SME understand a Rwandan government tender.",
      "Answer the question clearly and practically. Under 200 words.",
      tenderContext ? "TENDER CONTEXT:\n" + tenderContext : "",
      "Question: " + String(question)
    ].filter(Boolean).join("\n\n");

    var res = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "post",
      headers: { "Authorization": "Bearer " + apiKey, "Content-Type": "application/json" },
      payload: JSON.stringify({
        model: "openrouter/auto",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.3, max_tokens: 300
      }),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    if (code !== 200) {
      Logger.log("askAboutTender HTTP " + code + ": " + res.getContentText().substring(0, 300));
      return "AI unavailable (HTTP " + code + "). Try again later.";
    }
    var json = JSON.parse(res.getContentText());
    logActivity(getCurrentUserId(uid), "ask_tender");
    return String((json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || "AI returned no response. Try rephrasing your question.");
  } catch(e) {
    Logger.log("askAboutTender: " + e);
    return "Could not answer: " + e;
  }
}

// ── BUYER PROFILE ─────────────────────────────────────────────────────────────

function formatBuyerSnapshot(entity) {
  try {
    var bpSheet = getSheet().getSheetByName("BUYER_PROFILES");
    if (!bpSheet || bpSheet.getLastRow() < 2) return String(entity || "");
    var data = bpSheet.getDataRange().getValues();
    var headers = data[0];
    var lookup = safeLower(String(entity || "")).trim();
    for (var i = 1; i < data.length; i++) {
      var name = safeLower(String(data[i][0] || "")).trim();
      if (name === lookup || name.indexOf(lookup) !== -1 || lookup.indexOf(name) !== -1) {
        var bp = mapRow(headers, data[i]);
        return String(bp.entity || entity) + " — " + String(bp.total_tenders || 0) + " tenders, " + String(bp.competition_label || "Unknown") + ", " + String(bp.pct_single_bidder || 0) + "% single-bidder";
      }
    }
  } catch(e) {}
  return String(entity || "");
}

function buildBuyerProfiles() {
  try {
    var tSheet = getSheet().getSheetByName("TENDERS_FLAT");
    var bpSheet = getSheet().getSheetByName("BUYER_PROFILES");
    if (!tSheet || tSheet.getLastRow() < 2) return 0;

    var tData = tSheet.getDataRange().getValues();
    var tHeaders = tData[0];
    var rows = tData.slice(1);
    var buyers = {};

    rows.forEach(function(r) {
      var t = mapRow(tHeaders, r);
      var entity = String(t.entity || "").trim();
      if (!entity) return;
      if (!buyers[entity]) buyers[entity] = { tenders: [], methods: {}, sectors: {}, years: {} };
      var b = buyers[entity];
      b.tenders.push(t);
      var method = safeLower(String(t.method || "")).trim();
      if (method) b.methods[method] = (b.methods[method] || 0) + 1;
      var year = t.year || extractYear(t.publish_date);
      if (year) b.years[year] = true;
      var combined = safeLower(String(t.title || "")) + " " + safeLower(String(t.description || ""));
      var sectorKeywords = ["construction","road","building","school","hospital","health","medical","it","ict","software","supplies","food","catering","consulting","audit","training","transport","water","sanitation","energy","electricity","agriculture","furniture","security","cleaning"];
      sectorKeywords.forEach(function(kw) { if (combined.indexOf(kw) !== -1) b.sectors[kw] = (b.sectors[kw] || 0) + 1; });
    });

    bpSheet.getRange(2, 1, bpSheet.getLastRow(), bpSheet.getLastColumn()).clearContent();
    var buyerRows = [];
    Object.keys(buyers).forEach(function(entity) {
      var b = buyers[entity];
      var tenderers = b.tenders.map(function(t) { return Number(t.num_tenderers); }).filter(function(n) { return !isNaN(n) && n > 0; });
      var avgTenderers = tenderers.length ? mean(tenderers) : 0;
      var singleBidderPct = tenderers.length ? (tenderers.filter(function(n) { return n === 1; }).length / tenderers.length) * 100 : 0;
      var values = b.tenders.map(function(t) { return Number(t.est_value) || 0; }).filter(function(v) { return v > 0; });
      var avgValue = values.length ? mean(values) : 0;
      var medValue = values.length ? median(values) : 0;
      var directPct = b.methods["direct"] || b.methods["direct procurement"] || b.methods["single source"] || 0;
      var totalMethodCount = Object.keys(b.methods).reduce(function(sum, k) { return sum + b.methods[k]; }, 0);
      var directAwardPct = totalMethodCount ? (directPct / totalMethodCount) * 100 : 0;
      var methodKeys = Object.keys(b.methods).sort(function(a, b2) { return b.methods[b2] - b.methods[a]; });
      var sortedMethods = methodKeys.slice(0, 3).map(function(k) { return k + " (" + b.methods[k] + ")"; }).join(", ");
      var sectorKeys = Object.keys(b.sectors).sort(function(a, b2) { return b.sectors[b2] - b.sectors[a]; });
      var sortedSectors = sectorKeys.slice(0, 3).join(", ");
      var activeYears = Object.keys(b.years).sort();
      var lastActive = activeYears[activeYears.length - 1] || "";
      var competitionLabel = avgTenderers <= 1.5 ? "Very Low Competition" : avgTenderers <= 3 ? "Low Competition" : avgTenderers <= 6 ? "Moderate Competition" : avgTenderers <= 10 ? "High Competition" : "Very High Competition";
      buyerRows.push([entity, "Other", b.tenders.length, activeYears.length, Math.round(avgTenderers * 10) / 10, Math.round(singleBidderPct), Math.round(directAwardPct), Math.round(avgValue), Math.round(medValue), sortedSectors, sortedMethods, lastActive, competitionLabel, new Date()]);
    });

    buyerRows.sort(function(a, b) { return b[2] - a[2]; });
    if (buyerRows.length) bpSheet.getRange(2, 1, buyerRows.length, buyerRows[0].length).setValues(buyerRows);
    Logger.log("Built profiles for " + buyerRows.length + " buyers");
    return buyerRows.length;
  } catch(e) {
    Logger.log("buildBuyerProfiles: " + e);
    return 0;
  }
}

// ── PHASE 3 TESTS ────────────────────────────────────────────────────────────

function testPhase3Functions() {
  Logger.log("=== PHASE 3 DIAGNOSTICS ===");

  // AI key check
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY");
  Logger.log("OpenRouter key: " + (apiKey ? "SET (" + apiKey.substring(0, 8) + "...)" : "NOT SET"));

  // Lookup test
  try {
    var result = lookupTender("ocds-a44gk2-2025-000101");
    Logger.log("Lookup OCDS: " + (result.found ? "FOUND: " + result.tender.title : "NOT FOUND: " + result.message));
  } catch(e) { Logger.log("Lookup error: " + e); }

  try {
    var result2 = lookupTender("hospital");
    Logger.log("Lookup 'hospital': " + (result2.found ? "FOUND: " + result2.tender.title : "NOT FOUND: " + result2.message));
  } catch(e) { Logger.log("Lookup error: " + e); }

  // Buyer profiles
  try {
    var count = buildBuyerProfiles();
    Logger.log("Buyer profiles: " + count + " buyers");
  } catch(e) { Logger.log("Buyer profiles error: " + e); }

  Logger.log("=== END ===");
}

// =============================================================================
//  PHASE 4 — AI TUTOR, KNOWLEDGE BASE & INSIGHTS
// =============================================================================

// ── AI TUTOR ─────────────────────────────────────────────────────────────────

function askTutor(question, context, uid) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY");
    if (!apiKey) return "AI not configured. Add OPENROUTER_API_KEY to Script Properties.";
    if (!question || String(question).trim().length < 3) return "Please ask a question.";

    var userId = getCurrentUserId(uid);
    var profile = getUserProfile(userId);

    var kbContext = "";
    try { kbContext = buildKBContext(String(question)); } catch(e) { Logger.log("KB context: " + e); }

    var profileHint = profile ? "User's sector: " + (profile.primary_sector || profile.sectors_of_interest || "N/A") + ", budget: " + formatMoney(profile.min_value) + "-" + formatMoney(profile.max_value) + " RWF" : "";

    var prompt = [
      "You are a procurement tutor helping a Rwandan SME understand public procurement.",
      "Use the KNOWLEDGE BASE below as your primary source. Cite which document your answer comes from.",
      "If the answer is not in the knowledge base, use your training knowledge but say: 'Note: this is general knowledge, not from the official KB.'",
      "Keep answers practical, under 200 words, and specific to Rwanda's RPPA system.",
      "Use numbered steps or bullet points where helpful.",
      profileHint ? "User context: " + profileHint : "",
      kbContext ? "--- KNOWLEDGE BASE ---\n" + kbContext + "\n--- END KNOWLEDGE BASE ---" : "",
      "Question: " + String(question),
      String(context || "")
    ].filter(Boolean).join("\n");

    var res = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "post",
      headers: { "Authorization": "Bearer " + apiKey, "Content-Type": "application/json" },
      payload: JSON.stringify({
        model: "openrouter/auto",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.3, max_tokens: 400
      }),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    if (code !== 200) {
      Logger.log("askTutor HTTP " + code);
      return "AI unavailable (HTTP " + code + "). Check OpenRouter credits.";
    }
    var json = JSON.parse(res.getContentText());
    logActivity(userId, "ask_tutor");
    return String((json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || "AI returned no response. Try rephrasing.");
  } catch(e) {
    Logger.log("askTutor: " + e);
    return "Could not answer: " + e;
  }
}

// ── KNOWLEDGE BASE ───────────────────────────────────────────────────────────

function kbSetup() {
  var rootId = PropertiesService.getScriptProperties().getProperty("KB_ROOT_FOLDER_ID");
  if (!rootId) { Logger.log("KB_ROOT_FOLDER_ID not set in Script Properties"); return; }
  try {
    var root = DriveApp.getFolderById(rootId);
    Logger.log("KB root folder: " + root.getName());
    var folders = root.getFolders();
    while (folders.hasNext()) {
      var folder = folders.next();
      Logger.log("  Folder: " + folder.getName());
      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        Logger.log("    File: " + file.getName() + " (" + Math.round(file.getSize() / 1024) + " KB)");
      }
    }
  } catch(e) { Logger.log("KB error: " + e); }
}

function getKBFolderMap() {
  var rootId = PropertiesService.getScriptProperties().getProperty("KB_ROOT_FOLDER_ID");
  if (!rootId) return null;
  try {
    var root = DriveApp.getFolderById(rootId);
    var map = { root: rootId };
    var folders = root.getFolders();
    while (folders.hasNext()) {
      var folder = folders.next();
      var name = folder.getName().toLowerCase().replace(/[^a-z0-9]/g, "_");
      map[name] = folder.getId();
    }
    return map;
  } catch(e) { return null; }
}

function readKBFile(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    return file.getBlob().getDataAsString("UTF-8");
  } catch(e) { return ""; }
}

function listFilesInFolder(folderId) {
  if (!folderId) return [];
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var result = [];
    while (files.hasNext()) {
      var f = files.next();
      var name = f.getName().toLowerCase();
      if (name.match(/\.(txt|md)$/)) result.push({ name: f.getName(), id: f.getId(), size: f.getSize() });
    }
    return result;
  } catch(e) { return []; }
}

function findRelevantKBFiles(question) {
  var folders = getKBFolderMap();
  if (!folders) return [];
  var q = String(question || "").toLowerCase();
  var targetFolderIds = {};
  var rules = [
    { keywords: ["register","umucyo","account","login","portal","digital certificate","pki","supplier registration","submit bid","upload bid"], folder: "how_to_guides" },
    { keywords: ["price","cost","boq","bill of quantities","vat","rwf","financial proposal","withholding","currency","lump sum","unit rate","bid form"], folder: "how_to_guides" },
    { keywords: ["technical proposal","methodology","cv","personnel","work plan","tor","terms of reference","section","approach","evaluation criteria"], folder: "how_to_guides" },
    { keywords: ["reject","disqualif","missing document","tax clearance","rssb","bid security","late submission","bid bond","eligibility"], folder: "how_to_guides" },
    { keywords: ["score","qcbs","least cost","financial score","technical score","weighting","evaluation","minimum score","combined score"], folder: "how_to_guides" },
    { keywords: ["glossary","mean","definition","what is","acronym","term","ocid","rdb","rra","rppa","rssb","eac","beb","jv","itb","rfp"], folder: "glossary" },
    { keywords: ["law","article","regulation","legal","section","chapter","031/2022","64/2021","ministerial order","provision"], folder: "official_docs" },
    { keywords: ["clarification","circular","guideline","rppa says","authority","advance payment","framework agreement","bid validity extension"], folder: "official_docs" }
  ];
  rules.forEach(function(rule) {
    if (folders[rule.folder] && rule.keywords.some(function(kw) { return q.indexOf(kw) !== -1; })) {
      targetFolderIds[folders[rule.folder]] = true;
    }
  });
  if (!Object.keys(targetFolderIds).length && folders["how_to_guides"]) targetFolderIds[folders["how_to_guides"]] = true;
  var matched = [];
  Object.keys(targetFolderIds).forEach(function(fid) {
    listFilesInFolder(fid).forEach(function(f) { if (matched.length < 3) matched.push(f); });
  });
  return matched;
}

function buildKBContext(question) {
  var files = findRelevantKBFiles(String(question || ""));
  if (!files.length) return "";
  var parts = [];
  files.forEach(function(f) {
    var content = readKBFile(f.id);
    if (!content || content.length < 50) return;
    if (content.indexOf("PLACEHOLDER") !== -1) return;
    parts.push("=== SOURCE: " + f.name + " ===\n" + extractRelevantChunk(content, String(question || ""), 3000) + "\n=== END ===");
  });
  return parts.join("\n\n");
}

function extractRelevantChunk(content, question, maxChars) {
  if (content.length <= maxChars) return content;
  var keywords = String(question || "").toLowerCase().split(" ").filter(function(w) { return w.length > 4; });
  var bestPos = -1, bestScore = 0;
  for (var i = 0; i < content.length - 500; i += 200) {
    var windowText = content.substring(i, i + 500).toLowerCase();
    var score = keywords.filter(function(kw) { return windowText.indexOf(kw) !== -1; }).length;
    if (score > bestScore) { bestScore = score; bestPos = i; }
  }
  if (bestPos > 0 && bestScore > 0) return "...\n" + content.substring(Math.max(0, bestPos - 500), Math.max(0, bestPos - 500) + maxChars) + "\n...";
  return content.substring(0, maxChars) + "\n...";
}

function kbInstallTrigger() {
  ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === "kbWeeklyScan"; }).forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("kbWeeklyScan").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  Logger.log("KB weekly scan trigger installed");
}

function kbWeeklyScan() {
  var props = PropertiesService.getScriptProperties();
  var folders = getKBFolderMap();
  if (!folders) return;
  var targets = [
    { url: "https://www.rppa.gov.rw/publications", fileName: "rppa_clarifications.txt", folder: "official_docs" },
    { url: "https://www.rppa.gov.rw/faqs", fileName: "how_to_register_on_umucyo.md", folder: "how_to_guides" }
  ];
  targets.forEach(function(t) {
    try {
      var res = UrlFetchApp.fetch(t.url, { muteHttpExceptions: true, followRedirects: true });
      if (res.getResponseCode() !== 200) return;
      var hashKey = "KB_HASH_" + t.fileName.replace(/[^a-zA-Z0-9]/g, "_");
      var freshHash = String(res.getContentText().length);
      if (props.getProperty(hashKey) === freshHash) return;
      if (folders[t.folder]) {
        var folder = DriveApp.getFolderById(folders[t.folder]);
        var files = folder.getFilesByName(t.fileName);
        var header = "# Source: " + t.url + "\n# Updated: " + new Date().toISOString() + "\n---\n\n";
        if (files.hasNext()) files.next().setContent(header + res.getContentText());
        else folder.createFile(t.fileName, header + res.getContentText(), MimeType.PLAIN_TEXT);
        props.setProperty(hashKey, freshHash);
      }
    } catch(e) {}
  });
}

// ── MY INSIGHTS ──────────────────────────────────────────────────────────────

function getInsights(uid) {
  try {
    var userId = getCurrentUserId(uid);
    var profile = getUserProfile(userId);
    var totalExplored = 0, complianceChecks = 0, questionsAsked = 0;
    var searchTerms = {};

    var actSheet = getSheet().getSheetByName("USER_ACTIVITY");
    if (actSheet && actSheet.getLastRow() > 1) {
      var actData = actSheet.getDataRange().getValues();
      var actHeaders = actData[0];
      var actionCol = actHeaders.indexOf("action");
      var detailCol = actHeaders.indexOf("detail");
      for (var i = 1; i < actData.length; i++) {
        if (String(actData[i][0] || "").toLowerCase() !== userId.toLowerCase()) continue;
        var action = String(actData[i][actionCol] || "");
        var detail = String(actData[i][detailCol] || "");
        if (action === "search" || action === "lookup") totalExplored++;
        if (action === "compliance_check") complianceChecks++;
        if (action === "ask_tutor" || action === "ask_tender") questionsAsked++;
        if (detail && action === "search") {
          var words = detail.toLowerCase().split(" ").filter(function(w) { return w.length > 3; });
          words.forEach(function(w) { searchTerms[w] = (searchTerms[w] || 0) + 1; });
        }
      }
    }

    var searchKeys = Object.keys(searchTerms).sort(function(a, b) { return searchTerms[b] - searchTerms[a]; }).slice(0, 5);

    var aiInsight = "";
    try {
      var apiKey = PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY");
      if (apiKey && profile) {
        var prompt = [
          "Give 2-3 short, specific insights for this SME based on their profile. Under 100 words. No generic advice.",
          "Company: " + (profile.company_name || "N/A"),
          "Sector: " + (profile.primary_sector || profile.sectors_of_interest || "N/A"),
          "Budget: " + formatMoney(profile.min_value) + "-" + formatMoney(profile.max_value) + " RWF",
          "Activity: " + totalExplored + " explored, " + questionsAsked + " questions"
        ].join("\n");
        var res = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", {
          method: "post",
          headers: { "Authorization": "Bearer " + apiKey, "Content-Type": "application/json" },
          payload: JSON.stringify({ model: "openrouter/auto", messages: [{ role: "user", content: prompt }], temperature: 0.4, max_tokens: 200 }),
          muteHttpExceptions: true
        });
        if (res.getResponseCode() === 200) {
          var json = JSON.parse(res.getContentText());
          aiInsight = String((json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || "").trim();
        }
      }
    } catch(e) { Logger.log("insights AI: " + e); }

    return {
      companyName: profile ? (profile.company_name || "") : "",
      sector: profile ? (profile.primary_sector || profile.sectors_of_interest || "") : "",
      totalExplored: totalExplored,
      complianceChecks: complianceChecks,
      questionsAsked: questionsAsked,
      topSearchTerms: searchKeys,
      aiInsight: aiInsight,
      bestSectors: [], bestBuyers: [], marketTrends: [],
      sectorEntry: (function() {
        if (!profile || !checkFeatureAccess(userId, "trust_intel")) return null;
        var s = (profile.primary_sector || profile.sectors_of_interest || "").split(",")[0].trim();
        if (!s) return null;
        try { return getSectorEntryDifficulty(s); } catch(e) { return null; }
      })()
    };
  } catch(e) {
    Logger.log("getInsights: " + e);
    return { companyName: "", sector: "", totalExplored: 0, complianceChecks: 0, questionsAsked: 0, topSearchTerms: [], aiInsight: "", bestSectors: [], bestBuyers: [], marketTrends: [] };
  }
}

function generateStrategyReport(uid) {
  try {
    var userId = getCurrentUserId(uid);
    var profile = getUserProfile(userId);
    if (!profile) return { success: false, message: "Complete your Business DNA first." };

    var apiKey = PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY");
    if (!apiKey) return { success: false, message: "AI not configured." };

    var prompt = [
      "You are a procurement strategy advisor. Generate a brief market strategy report for this SME.",
      "Format: Best sectors, Best buyers, Best value range, Best timing, Market position, Actionable advice.",
      "Company: " + (profile.company_name || ""),
      "Sector: " + (profile.primary_sector || ""),
      "Budget: " + formatMoney(profile.min_value) + "-" + formatMoney(profile.max_value) + " RWF",
      "Regions: " + (profile.regions || "Nationwide"),
      "Under 300 words."
    ].join("\n");

    var res = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "post",
      headers: { "Authorization": "Bearer " + apiKey, "Content-Type": "application/json" },
      payload: JSON.stringify({ model: "openrouter/auto", messages: [{ role: "user", content: prompt }], temperature: 0.4, max_tokens: 600 }),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) return { success: false, message: "AI unavailable (HTTP " + res.getResponseCode() + ")." };
    var json = JSON.parse(res.getContentText());
    var report = String((json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || "");

    var rSheet = getSheet().getSheetByName("STRATEGY_REPORT");
    if (rSheet) {
      rSheet.appendRow([userId, new Date(), "", "", "", "", "", report, new Date()]);
    }
    logActivity(userId, "strategy_report");
    return { success: true, message: report };
  } catch(e) {
    Logger.log("generateStrategyReport: " + e);
    return { success: false, message: "Error: " + e };
  }
}

// ── PHASE 4 TESTS ────────────────────────────────────────────────────────────

function testPhase4Functions() {
  Logger.log("=== PHASE 4 DIAGNOSTICS ===");
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENROUTER_API_KEY");
  Logger.log("OpenRouter: " + (apiKey ? "SET" : "NOT SET"));

  var kbId = PropertiesService.getScriptProperties().getProperty("KB_ROOT_FOLDER_ID");
  Logger.log("KB_ROOT_FOLDER_ID: " + (kbId || "NOT SET"));
  if (kbId) {
    try { Logger.log("KB folder: " + DriveApp.getFolderById(kbId).getName()); } catch(e) { Logger.log("KB folder error: " + e); }
  }

  try {
    var insights = getInsights();
    Logger.log("Insights: explored=" + insights.totalExplored + " questions=" + insights.questionsAsked);
  } catch(e) { Logger.log("Insights error: " + e); }

  Logger.log("=== END ===");
}

function testBigQuery() {
  var sql = "SELECT COUNT(*) AS n FROM rppa-umucyo.rppa_historical.rppc_main";
  Logger.log("Testing BigQuery: " + sql);
  try {
    var token = ScriptApp.getOAuthToken();
    var res = UrlFetchApp.fetch("https://bigquery.googleapis.com/bigquery/v2/projects/rppa-umucyo/queries", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify({ query: sql, useLegacySql: false, maxResults: 1 }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    Logger.log("HTTP " + code);
    if (code === 200) {
      var json = JSON.parse(res.getContentText());
      if (json.rows && json.rows[0]) Logger.log("Result: " + json.rows[0].f[0].v + " rows");
      else Logger.log("Empty response");
    } else {
      Logger.log("Error: " + res.getContentText().substring(0, 500));
    }
  } catch(e) {
    Logger.log("Exception: " + e);
  }
}

// =============================================================================
//  PHASE 5 — BIGQUERY PREDICTIONS
// =============================================================================

// ── BIGQUERY CORE ────────────────────────────────────────────────────────────

var BIGQUERY_PROJECT_ID = PropertiesService.getScriptProperties().getProperty("BIGQUERY_PROJECT_ID") || "rppa-umucyo";
var BIGQUERY_DATASET = "rppa_historical";

function queryBigQuery(sql, maxResults) {
  var request = { query: sql, useLegacySql: false, maxResults: maxResults || 5000 };
  var allRows = [];
  var pageToken = null;
  var pages = 0;
  do {
    var result = pageToken
      ? BigQuery.Jobs.query(request, BIGQUERY_PROJECT_ID, pageToken)
      : BigQuery.Jobs.query(request, BIGQUERY_PROJECT_ID);
    pages++;
    if (!result.jobComplete) { Utilities.sleep(2000); continue; }
    if (result.rows && result.schema) {
      var fields = result.schema.fields.map(function(f) { return f.name; });
      result.rows.forEach(function(row) {
        var obj = {};
        fields.forEach(function(f, i) { obj[f] = row.f[i] ? row.f[i].v : null; });
        allRows.push(obj);
      });
    }
    pageToken = result.pageToken || null;
  } while (pageToken && pages < 10);
  Logger.log("BQ returned " + allRows.length + " rows (" + pages + " pages)");
  return allRows;
}

// ── PREDICTION ENRICHMENT ────────────────────────────────────────────────────

function getTenderPredictionEnrichment(entity, method, value) {
  try {
    var result = {};
    var e = String(entity || "").trim();
    var eSafe = e.replace(/'/g, "''");
    var v = Number(value) || 0;

    // Buyer competitors — top winners for this buyer (LIKE for fuzzy match)
    try {
      var compSql = "SELECT award_supplier, win_count FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_buyer_competitors WHERE LOWER(entity) LIKE LOWER('%" + eSafe + "%') ORDER BY win_count DESC LIMIT 5";
      var compRows = queryBigQuery(compSql);
      if (compRows.length > 0) {
        result.competitors = compRows.map(function(r) { return String(r.award_supplier) + " (" + r.win_count + " wins)"; });
      }
    } catch(e2) {}

    // Buyer risk — amendment percentage
    try {
      var riskSql = "SELECT pct_amended, total_tenders, avg_competition FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_buyer_risk WHERE LOWER(entity) LIKE LOWER('%" + eSafe + "%') LIMIT 1";
      var riskRows = queryBigQuery(riskSql);
      if (riskRows.length > 0) {
        result.amendment_pct = parseFloat(riskRows[0].pct_amended) || 0;
        result.total_tenders = parseInt(riskRows[0].total_tenders) || 0;
        result.avg_competition_buyer = parseFloat(riskRows[0].avg_competition) || null;
      }
    } catch(e2) {}

    // Buyer contacts
    try {
      var contactSql = "SELECT contact_phone, contact_email, contact_name FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_buyer_contacts WHERE LOWER(entity) LIKE LOWER('%" + eSafe + "%') LIMIT 1";
      var contactRows = queryBigQuery(contactSql);
      if (contactRows.length > 0) {
        result.contact_name = String(contactRows[0].contact_name || "");
        result.contact_phone = String(contactRows[0].contact_phone || "");
        result.contact_email = String(contactRows[0].contact_email || "");
      }
    } catch(e2) {}

    // Prediction signals — win probability, competition, price ratio
    try {
      var predSql = "SELECT win_probability, avg_competition, median_price_ratio FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_prediction_signals WHERE LOWER(entity) LIKE LOWER('%" + eSafe + "%') AND LOWER(method) LIKE LOWER('%" + String(method || "").replace(/'/g, "''") + "%') LIMIT 1";
      var predRows = queryBigQuery(predSql);
      if (predRows.length > 0) {
        result.win_probability = parseFloat(predRows[0].win_probability) || null;
        result.expected_bidders = parseFloat(predRows[0].avg_competition) || null;
        if (predRows[0].median_price_ratio) result.recommended_bid = Math.round(v * parseFloat(predRows[0].median_price_ratio));
      }
    } catch(e2) {}

    return result;
  } catch(e) { Logger.log("getTenderPrediction: " + e); return {}; }
}

function getPersonalMarketInsights(profile) {
  try {
    var result = { bestSectors: [], bestBuyers: [], marketTrends: [] };
    var sector = String(profile.primary_sector || profile.sectors_of_interest || "").split(",")[0].trim().toLowerCase();
    if (!sector) return result;
    var sSafe = sector.replace(/'/g, "''");
    try {
      var buyerSql = "SELECT entity, COUNT(*) AS tender_count FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(title) LIKE '%" + sSafe + "%' GROUP BY entity ORDER BY tender_count DESC LIMIT 5";
      result.bestBuyers = queryBigQuery(buyerSql).map(function(r) { return { entity: String(r.entity), tender_count: Number(r.tender_count) }; });
    } catch(e2) {}
    return result;
  } catch(e) { return { bestSectors: [], bestBuyers: [], marketTrends: [] }; }
}

// ── BID/NO-BID ENGINE ────────────────────────────────────────────────────────

function computeBidNoBid(match) {
  try {
    var wp = parseFloat(match.win_probability) || 0;
    var risk = parseFloat(match.buyer_amendment_pct) || 0;
    var competition = parseFloat(match.expected_bidders) || 5;
    var score = Math.round(20 + wp * 100 * 2 - risk * 0.2 - Math.max(0, (competition - 6) * 4));
    score = Math.min(100, Math.max(0, score));
    var rec = score >= 50 ? "BID" : score >= 25 ? "MAYBE" : "SKIP";
    var reasons = [];
    if (wp * 100 >= 20) reasons.push(Math.round(wp * 100) + "% historical win rate");
    else if (wp * 100 >= 10) reasons.push("Moderate win rate (" + Math.round(wp * 100) + "%)");
    else reasons.push("Low win probability (" + Math.round(wp * 100) + "%)");
    if (risk >= 30) reasons.push("High amendment risk (" + risk + "%)");
    if (competition <= 4) reasons.push("Low competition (" + competition + " avg bidders)");
    else if (competition >= 8) reasons.push("High competition (" + competition + " avg bidders)");
    return { recommendation: rec, confidence: score, reasons: reasons.join("; ") };
  } catch(e) { return { recommendation: "MAYBE", confidence: 50, reasons: "" }; }
}

// ── PHASE 5 TESTS ────────────────────────────────────────────────────────────

function testPhase5Functions() {
  Logger.log("=== PHASE 5 DIAGNOSTICS ===");
  var tables = ["rppc_main","rppc_prediction_signals","rppc_buyer_competitors","rppc_buyer_risk","rppc_buyer_contacts"];
  tables.forEach(function(t) {
    try {
      var rows = queryBigQuery("SELECT COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + "." + t);
      Logger.log("Table " + t + ": " + (rows.length > 0 ? rows[0].n + " rows" : "EMPTY"));
    } catch(e) { Logger.log("Table " + t + ": ERROR - " + e); }
  });
  try {
    var pred = getTenderPredictionEnrichment("MININFRA", "open", 950000000);
    Logger.log("Prediction MININFRA: " + JSON.stringify(pred));
  } catch(e) { Logger.log("Prediction error: " + e); }
  Logger.log("=== END ===");
}

function debugBQSchema() {
  var tables = ["rppc_prediction_signals","rppc_buyer_competitors","rppc_buyer_risk","rppc_buyer_contacts"];
  tables.forEach(function(t) {
    try {
      var rows = queryBigQuery("SELECT * FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + "." + t + " LIMIT 1");
      if (rows.length > 0) {
        Logger.log(t + " columns: " + Object.keys(rows[0]).join(", "));
        Logger.log(t + " sample: " + JSON.stringify(rows[0]));
      } else { Logger.log(t + ": EMPTY"); }
    } catch(e) { Logger.log(t + ": ERROR - " + e); }
  });
}

function findMININFRA() {
  try {
    var rows = queryBigQuery("SELECT DISTINCT entity FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_prediction_signals WHERE LOWER(entity) LIKE '%mininfra%' LIMIT 5");
    Logger.log("pred signals like mininfra: " + JSON.stringify(rows));
  } catch(e) { Logger.log("err: " + e); }
  try {
    var rows2 = queryBigQuery("SELECT DISTINCT entity FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_prediction_signals LIMIT 5");
    Logger.log("pred signals sample entities: " + JSON.stringify(rows2));
  } catch(e) {}
}

// =============================================================================
//  PHASE 6 — PREMIUM PAGES
// =============================================================================

function getMarketDashboardData() {
  try {
    var sql = "SELECT buyer_name, COUNT(*) AS tender_count, AVG(tender_numberOfTenderers) AS avg_bidders FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE buyer_name IS NOT NULL GROUP BY buyer_name ORDER BY tender_count DESC LIMIT 20";
    var topBuyers = queryBigQuery(sql).map(function(r) { return { entity: String(r.buyer_name || ""), tender_count: Number(r.tender_count || 0), avg_bidders: Math.round(Number(r.avg_bidders) || 0) }; });
    var sql2 = "SELECT tender_mainProcurementCategory, COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE tender_mainProcurementCategory IS NOT NULL GROUP BY tender_mainProcurementCategory ORDER BY n DESC LIMIT 10";
    var topSectors = queryBigQuery(sql2).map(function(r) { return { sector: String(r.tender_mainProcurementCategory || ""), count: Number(r.n || 0) }; });
    return { topBuyers: topBuyers, topSectors: topSectors };
  } catch(e) { Logger.log("marketDashboard: " + e); return { topBuyers: [], topSectors: [] }; }
}

function getProcurementCalendar() {
  try {
    var sql = "SELECT buyer_name AS entity, EXTRACT(MONTH FROM date) AS month, COUNT(*) AS tender_count FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE date IS NOT NULL GROUP BY entity, month ORDER BY entity, month LIMIT 500";
    var rows = queryBigQuery(sql).map(function(r) { return { entity: String(r.entity || ""), month: Number(r.month || 0), tender_count: Number(r.tender_count || 0) }; });
    return { months: rows, topBuyers: rows };
  } catch(e) { Logger.log("calendar: " + e); return { months: [], topBuyers: [] }; }
}

function getBuyerPaymentRatings() {
  try {
    var sql = "SELECT entity, pct_amended, total_tenders, avg_competition, avg_tender_value FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_buyer_risk ORDER BY pct_amended DESC LIMIT 50";
    return queryBigQuery(sql).map(function(r) { return { 
      entity: String(r.entity || ""), 
      amendment_pct: Number(r.pct_amended || 0), 
      contracts: Number(r.total_tenders || 0), 
      avg_duration_days: Math.round(Number(r.avg_competition || 0)), 
      avg_contract_value: Math.round(Number(r.avg_tender_value || 0)) 
    }; });
  } catch(e) { Logger.log("ratings: " + e); return []; }
}

function getSupplierPortfolio(name) {
  try {
    if (!name || String(name).trim().length < 2) return { error: "Enter a company name." };
    var n = String(name).replace(/'/g, "''");
    var awardsSql = "SELECT a.title, a.value_amount, a.date, m.buyer_name AS entity FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_award_suppliers s ON a.main_ocid = s.main_ocid AND a.id = s.awards_id JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON a.main_ocid = m.ocid WHERE LOWER(s.name) LIKE LOWER('%" + n + "%') ORDER BY a.date DESC LIMIT 20";
    var rows = queryBigQuery(awardsSql);
    var contracts = rows.map(function(r) { return { title: String(r.title || ""), entity: String(r.entity || ""), award_value: Number(r.value_amount || 0), est_value: Number(r.value_amount || 0), year: String(r.date || "").substring(0,4) }; });
    var totalValue = contracts.reduce(function(s, c) { return s + c.award_value; }, 0);
    var buyers = {};
    contracts.forEach(function(c) { var b = c.entity; buyers[b] = (buyers[b] || 0) + 1; });
    var sectors = {};
    var creditScore = null;
    try { creditScore = getSupplierCreditScore(name); } catch(e) { Logger.log("portfolio credit: " + e); }
    return { supplier: name, wins: contracts.length, totalValue: totalValue, buyers: buyers, sectors: sectors, contracts: contracts, credit_score: creditScore };
  } catch(e) { Logger.log("portfolio: " + e); return { error: "Error: " + e }; }
}

function getSectorHealth(sector) {
  try {
    if (!sector) return { overview: {} };
    var s = String(sector).replace(/'/g, "''");
    var sql = "SELECT COUNT(*) AS n, AVG(tender_value_amount) AS avg_unit_value, SUM(tender_value_amount) AS total_quantity FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(tender_mainProcurementCategory) LIKE LOWER('%" + s + "%')";
    var rows = queryBigQuery(sql);
    var overview = rows.length > 0 ? { category: sector, tender_count: Number(rows[0].n || 0), avg_unit_value: Math.round(Number(rows[0].avg_unit_value || 0)), total_quantity: Math.round(Number(rows[0].total_quantity || 0)) } : {};
    var trendSql = "SELECT EXTRACT(YEAR FROM date) AS year, COUNT(*) AS tender_count, AVG(tender_value_amount) AS avg_value FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(tender_mainProcurementCategory) LIKE LOWER('%" + s + "%') AND date IS NOT NULL GROUP BY year ORDER BY year DESC LIMIT 8";
    var trends = queryBigQuery(trendSql).map(function(r) { return { year: String(r.year || ""), tender_count: Number(r.tender_count || 0), avg_value: Math.round(Number(r.avg_value || 0)) }; });
    var buyerSql = "SELECT buyer_name AS entity, COUNT(*) AS tender_count, AVG(tender_numberOfTenderers) AS avg_bidders FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(tender_mainProcurementCategory) LIKE LOWER('%" + s + "%') AND buyer_name IS NOT NULL GROUP BY entity ORDER BY tender_count DESC LIMIT 10";
    var buyers = queryBigQuery(buyerSql).map(function(r) { return { entity: String(r.entity || ""), tender_count: Number(r.tender_count || 0), avg_bidders: Math.round(Number(r.avg_bidders || 0)) }; });
    var entryBarrier = null;
    if (overview.tender_count > 0) entryBarrier = getSectorEntryDifficulty(sector);
    return { overview: overview, trends: trends, buyers: buyers, entry_barrier: entryBarrier };
  } catch(e) { Logger.log("sectorHealth: " + e); return { overview: {} }; }
}

function getCompetitorIntel(name) {
  try {
    if (!name || String(name).trim().length < 2) return { error: "Enter a company name." };
    var n = String(name).replace(/'/g, "''");
    var winsSql = "SELECT s.name AS supplier_name, a.title, a.value_amount, a.date, m.buyer_name AS entity FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_award_suppliers s ON a.main_ocid = s.main_ocid AND a.id = s.awards_id JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON a.main_ocid = m.ocid WHERE LOWER(s.name) LIKE LOWER('%" + n + "%') ORDER BY a.date DESC LIMIT 30";
    var rows = queryBigQuery(winsSql);
    var recentWins = rows.map(function(r) { return { title: String(r.title || ""), entity: String(r.entity || ""), award_value: Number(r.value_amount || 0), est_value: Number(r.value_amount || 0), year: String(r.date || "").substring(0,4) }; });
    var winByBuyer = {};
    rows.forEach(function(r) { var b = String(r.entity || ""); winByBuyer[b] = (winByBuyer[b] || 0) + 1; });
    var winByYear = {};
    rows.forEach(function(r) { var y = String(r.date || "").substring(0,4); if (y) winByYear[y] = (winByYear[y] || 0) + 1; });
    return { name: name, wins: rows.length, winByBuyer: winByBuyer, winByYear: winByYear, recentWins: recentWins };
  } catch(e) { Logger.log("competitorIntel: " + e); return { error: "Error: " + e }; }
}

function findJVPartners(name) {
  try {
    if (!name || String(name).trim().length < 2) return { error: "Enter a company name." };
    var n = String(name).replace(/'/g, "''");
    var myBuyersSql = "SELECT DISTINCT m.buyer_name AS entity FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_award_suppliers s ON a.main_ocid = s.main_ocid AND a.id = s.awards_id JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON a.main_ocid = m.ocid WHERE LOWER(s.name) LIKE LOWER('%" + n + "%') AND m.buyer_name IS NOT NULL LIMIT 30";
    var myBuyers = queryBigQuery(myBuyersSql).map(function(r) { return String(r.entity || "").replace(/'/g, "''"); });
    if (!myBuyers.length) return { partners: [] };
    var partnerSql = "SELECT s.name AS partner, m.buyer_name AS buyer, COUNT(*) AS wins, AVG(a.value_amount) AS avg_award FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_award_suppliers s ON a.main_ocid = s.main_ocid AND a.id = s.awards_id JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON a.main_ocid = m.ocid WHERE m.buyer_name IN ('" + myBuyers.join("','") + "') AND LOWER(s.name) NOT LIKE LOWER('%" + n + "%') GROUP BY partner, buyer ORDER BY wins DESC LIMIT 100";
    var rows = queryBigQuery(partnerSql);
    var partners = rows.map(function(r) { return { partner: String(r.partner || ""), buyer: String(r.buyer || ""), wins: Number(r.wins || 0), avg_award: Math.round(Number(r.avg_award || 0)) }; });
    return { partners: partners };
  } catch(e) { Logger.log("JVPartners: " + e); return { partners: [] }; }
}

function debugBQCategories() {
  var sql = "SELECT tender_mainProcurementCategory, COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE tender_mainProcurementCategory IS NOT NULL GROUP BY tender_mainProcurementCategory ORDER BY n DESC LIMIT 30";
  var rows = queryBigQuery(sql);
  rows.forEach(function(r) { Logger.log(String(r.tender_mainProcurementCategory) + " = " + r.n); });
}

function testPhase6Functions() {
  Logger.log("=== PHASE 6 DIAGNOSTICS ===");
  var fns = [
    { name: "market", fn: function() { var d = getMarketDashboardData(); return d.topBuyers.length + " buyers, " + d.topSectors.length + " sectors"; } },
    { name: "calendar", fn: function() { var d = getProcurementCalendar(); return d.length + " rows"; } },
    { name: "portfolio", fn: function() { var d = getSupplierPortfolio("construction"); return JSON.stringify(d).substring(0, 100); } },
    { name: "ratings", fn: function() { var d = getBuyerPaymentRatings(); return d.length + " buyers"; } },
    { name: "sector", fn: function() { var d = getSectorHealth("construction"); return JSON.stringify(d.stats); } },
    { name: "competitor", fn: function() { var d = getCompetitorIntel("construction"); return JSON.stringify(d).substring(0, 100); } },
    { name: "jv", fn: function() { var d = findJVPartners("construction"); return JSON.stringify(d).substring(0, 100); } }
  ];
  fns.forEach(function(f) {
    try { Logger.log(f.name + ": " + f.fn()); } catch(e) { Logger.log(f.name + ": ERROR - " + e); }
  });
  // Check if rppc_awards exists
  try {
    var c = queryBigQuery("SELECT COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards");
    Logger.log("rppc_awards: " + (c.length > 0 ? c[0].n + " rows" : "EMPTY/MISSING"));
  } catch(e) { Logger.log("rppc_awards: ERROR - " + e); }
  Logger.log("=== END ===");
}

function debugBQPhase6() {
  var tables = ["rppc_main","rppc_awards","rppc_award_suppliers","rppc_tender_items","rppc_buyers"];
  tables.forEach(function(t) {
    try {
      var rows = queryBigQuery("SELECT * FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + "." + t + " LIMIT 1");
      if (rows.length > 0) {
        Logger.log(t + " columns: " + Object.keys(rows[0]).join(", "));
      } else { Logger.log(t + ": EMPTY"); }
    } catch(e) { Logger.log(t + ": ERROR - " + e.message); }
  });
}

// =============================================================================
//  PHASE 7 — MONETIZATION
// =============================================================================

var TIER_FEATURES = {
  free:     { win_prob: false, competitors: false, buyer_contacts: false, price_rec: false, calendar: false, email_alerts: false, portfolio: false, buyer_ratings: false, sector_health: false, competitor_intel: false, jv_partners: false, trust_intel: false, match_limit: 25 },
  pro:      { win_prob: true, competitors: true, buyer_contacts: true, price_rec: true, calendar: true, email_alerts: true, portfolio: true, buyer_ratings: false, sector_health: false, competitor_intel: false, jv_partners: false, trust_intel: false, pipeline_warnings: true, match_limit: 500 },
  premium:  { win_prob: true, competitors: true, buyer_contacts: true, price_rec: true, calendar: true, email_alerts: true, portfolio: true, buyer_ratings: true, sector_health: true, competitor_intel: true, jv_partners: true, trust_intel: true, match_limit: 99999 },
  trial:    { win_prob: true, competitors: true, buyer_contacts: true, price_rec: true, calendar: true, email_alerts: false, portfolio: false, buyer_ratings: false, sector_health: false, competitor_intel: false, jv_partners: false, trust_intel: true, match_limit: 100 }
};

var TIER_PRICES = { free: { usd: 0, rwf: 0 }, pro: { usd: 20, rwf: 25000 }, premium: { usd: 75, rwf: 90000 } };

// ── SUBSCRIPTION SYSTEM ──────────────────────────────────────────────────────

function getSubscription(userId) {
  if (!userId) return { tier: "free", status: "active", features: TIER_FEATURES.free, match_limit: 25 };
  try {
    var sheet = getSheet().getSheetByName("SUBSCRIPTIONS");
    if (!sheet || sheet.getLastRow() < 2) return { tier: "free", status: "active", features: TIER_FEATURES.free, match_limit: 25 };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = function(n) { return headers.indexOf(n); };
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][col("user_id")] || "").toLowerCase() !== userId.toLowerCase()) continue;
      var tier = String(data[i][col("tier")] || "free");
      var status = String(data[i][col("status")] || "active");
      var expiresRaw = data[i][col("expires_at")];
      if (status === "expired" || (expiresRaw && new Date(expiresRaw) < new Date())) { status = "expired"; tier = "free"; }
      return {
        tier: tier, status: status,
        provider: String(data[i][col("provider")] || ""),
        expires_at: String(expiresRaw || ""),
        match_limit: Number(data[i][col("match_limit")] || 25),
        features: TIER_FEATURES[tier] || TIER_FEATURES.free
      };
    }
  } catch(e) { Logger.log("getSubscription: " + e); }
  return { tier: "free", status: "active", features: TIER_FEATURES.free, match_limit: 25 };
}

function grantTrial(userId) {
  if (!userId) userId = getCurrentUserId();
  try {
    var sheet = getSheet().getSheetByName("SUBSCRIPTIONS");
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").toLowerCase() === userId.toLowerCase()) return;
    }
    var now = new Date();
    var expiry = new Date(now.getTime() + 14 * 24 * 3600 * 1000);
    sheet.appendRow([userId, "free", "trial", 100, JSON.stringify(TIER_FEATURES.trial), "trial", "", 0, "USD", now, expiry, false]);
    logActivity(userId, "trial_started");
  } catch(e) { Logger.log("grantTrial: " + e); }
}

function checkFeatureAccess(userId, feature) {
  var sub = getSubscription(userId);
  var features = sub.features || TIER_FEATURES.free;
  return features[feature] === true;
}

function upgradeTier(userId, tier, provider, ref, amount, currency) {
  if (!userId) userId = getCurrentUserId();
  try {
    var sheet = getSheet().getSheetByName("SUBSCRIPTIONS");
    if (!sheet) { setupSheets(); sheet = getSheet().getSheetByName("SUBSCRIPTIONS"); }
    if (!TIER_FEATURES[tier]) throw new Error("Invalid tier: " + tier);
    var now = new Date();
    var expiry = new Date(now.getTime() + 30 * 24 * 3600 * 1000);

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = function(n) { return headers.indexOf(n); };
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][col("user_id")] || "").toLowerCase() !== userId.toLowerCase()) continue;
      if (String(data[i][col("status")] || "") === "active" || String(data[i][col("status")] || "") === "trial") {
        sheet.getRange(i + 1, col("tier") + 1).setValue(tier);
        sheet.getRange(i + 1, col("status") + 1).setValue("active");
        sheet.getRange(i + 1, col("match_limit") + 1).setValue(TIER_FEATURES[tier].match_limit);
        sheet.getRange(i + 1, col("features_json") + 1).setValue(JSON.stringify(TIER_FEATURES[tier]));
        sheet.getRange(i + 1, col("provider") + 1).setValue(provider);
        sheet.getRange(i + 1, col("payment_ref") + 1).setValue(ref || "");
        sheet.getRange(i + 1, col("amount") + 1).setValue(amount);
        sheet.getRange(i + 1, col("currency") + 1).setValue(currency || "USD");
        sheet.getRange(i + 1, col("created_at") + 1).setValue(now);
        sheet.getRange(i + 1, col("expires_at") + 1).setValue(expiry);
        logActivity(userId, "subscription_upgrade");
        return { success: true, tier: tier, expires_at: expiry };
      }
    }
    sheet.appendRow([userId, tier, "active", TIER_FEATURES[tier].match_limit, JSON.stringify(TIER_FEATURES[tier]), provider, ref, amount, currency || "USD", now, expiry, false]);
    logActivity(userId, "subscription_upgrade");
    try { MailApp.sendEmail(userId, "[Umucyo] Subscription Upgraded!", "You are now on the " + tier.toUpperCase() + " plan. Manage your subscription at " + ScriptApp.getService().getUrl() + "?page=upgrade"); } catch(e) {}
    return { success: true, tier: tier, expires_at: expiry };
  } catch(e) { Logger.log("upgradeTier: " + e); return { success: false, error: String(e) }; }
}

function downgradeExpired() {
  try {
    var sheet = getSheet().getSheetByName("SUBSCRIPTIONS");
    if (!sheet) return 0;
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = function(n) { return headers.indexOf(n); };
    var now = new Date(); var count = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      var status = String(data[i][col("status")] || "");
      var tier = String(data[i][col("tier")] || "");
      var expiresRaw = data[i][col("expires_at")];
      if ((status === "active" || status === "trial") && tier !== "free" && expiresRaw && new Date(expiresRaw) < now) {
        sheet.getRange(i + 1, col("status") + 1).setValue("expired");
        sheet.getRange(i + 1, col("tier") + 1).setValue("free");
        count++;
      }
    }
    Logger.log("Downgraded " + count + " expired");
    return count;
  } catch(e) { Logger.log("downgradeExpired: " + e); return 0; }
}

function getSubscriptionForWeb(uid) {
  var userId = getCurrentUserId(uid);
  var sub = getSubscription(userId);
  return { tier: sub.tier, status: sub.status, provider: sub.provider, expires_at: sub.expires_at, features: sub.features, match_limit: sub.match_limit };
}

// ── PAYMENT: FLUTTERWAVE ─────────────────────────────────────────────────────

function getFlutterwavePublicKey() {
  return PropertiesService.getScriptProperties().getProperty("FLUTTERWAVE_PUBLIC_KEY") || "";
}

function prepareFlutterwavePayment(uid, tier, currency) {
  try {
    var userId = getCurrentUserId(uid);
    if (!TIER_PRICES[tier]) return { success: false, error: "Invalid tier" };
    var amount = currency === "RWF" ? TIER_PRICES[tier].rwf : TIER_PRICES[tier].usd;
    var txRef = "UMU-" + Utilities.getUuid().substring(0, 12).toUpperCase();
    PropertiesService.getScriptProperties().setProperty("FLW_PENDING_" + txRef, JSON.stringify({
      userId: userId, tier: tier, currency: currency, amount: amount
    }));
    return { success: true, tx_ref: txRef, amount: amount, currency: currency };
  } catch(e) { return { success: false, error: String(e) }; }
}

function verifyFlutterwavePayment(txRef, transactionId) {
  try {
    var secretKey = PropertiesService.getScriptProperties().getProperty("FLUTTERWAVE_SECRET_KEY");
    if (!secretKey) return { success: false, error: "Flutterwave not configured." };

    var res = UrlFetchApp.fetch("https://api.flutterwave.com/v3/transactions/" + encodeURIComponent(transactionId) + "/verify", {
      headers: { "Authorization": "Bearer " + secretKey },
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) return { success: false, error: "Verification failed: " + res.getResponseCode() };

    var data = JSON.parse(res.getContentText());
    if (data.status === "success" && data.data && data.data.status === "successful") {
      var pending = PropertiesService.getScriptProperties().getProperty("FLW_PENDING_" + txRef);
      var tier = "pro", uid = getCurrentUserId(), currency = "RWF", amount = 0;
      if (pending) {
        try { var p = JSON.parse(pending); tier = p.tier || "pro"; uid = p.userId || uid; currency = p.currency || "RWF"; amount = p.amount || 0; } catch(e) {}
        PropertiesService.getScriptProperties().deleteProperty("FLW_PENDING_" + txRef);
      }
      upgradeTier(uid, tier, "flutterwave", transactionId, amount || data.data.amount, currency);
      return { success: true, message: "Payment confirmed! " + tier.toUpperCase() + " features unlocked." };
    }
    return { success: false, error: "Payment not completed (status: " + (data.data ? data.data.status : "unknown") + ")" };
  } catch(e) { return { success: false, error: String(e).substring(0, 100) }; }
}

// ── PAYMENT: STRIPE ──────────────────────────────────────────────────────────

function createStripeSession(uid, tier, currency) {
  try {
    var userId = getCurrentUserId(uid);
    if (!TIER_PRICES[tier]) return { success: false, error: "Invalid tier" };
    var amount = currency === "RWF" ? TIER_PRICES[tier].rwf : TIER_PRICES[tier].usd;
    var unit = currency === "RWF" ? "rwf" : "usd";

    var apiKey = PropertiesService.getScriptProperties().getProperty("STRIPE_SECRET_KEY");
    if (!apiKey) return { success: false, error: "Stripe not configured." };

    var payload = {
      "line_items[0][price_data][currency]": unit === "rwf" ? "rwf" : "usd",
      "line_items[0][price_data][product_data][name]": "Umucyo " + tier.charAt(0).toUpperCase() + tier.slice(1) + " Plan",
      "line_items[0][price_data][unit_amount]": amount * (unit === "rwf" ? 1 : 100),
      "line_items[0][quantity]": "1",
      "mode": "payment",
      "success_url": ScriptApp.getService().getUrl() + "?page=dashboard&uid=" + encodeURIComponent(userId) + "&session_id={CHECKOUT_SESSION_ID}",
      "cancel_url": ScriptApp.getService().getUrl() + "?page=upgrade",
      "metadata[user_id]": userId,
      "metadata[tier]": tier,
      "metadata[currency]": currency
    };

    var formData = Object.keys(payload).map(function(k) { return encodeURIComponent(k) + "=" + encodeURIComponent(payload[k]); }).join("&");

    var res = UrlFetchApp.fetch("https://api.stripe.com/v1/checkout/sessions", {
      method: "post",
      headers: { "Authorization": "Bearer " + apiKey, "Content-Type": "application/x-www-form-urlencoded" },
      payload: formData,
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) return { success: false, error: "Stripe: " + res.getContentText().substring(0, 200) };

    var data = JSON.parse(res.getContentText());
    return { success: true, url: data.url, sessionId: data.id };
  } catch(e) { return { success: false, error: String(e) }; }
}

function verifyStripeSession(sessionId) {
  for (var attempt = 0; attempt < 3; attempt++) {
    try {
      var apiKey = PropertiesService.getScriptProperties().getProperty("STRIPE_SECRET_KEY");
      if (!apiKey) return { success: false, error: "Stripe not configured." };

      var res = UrlFetchApp.fetch("https://api.stripe.com/v1/checkout/sessions/" + encodeURIComponent(sessionId), {
        headers: { "Authorization": "Bearer " + apiKey },
        muteHttpExceptions: true
      });

      if (res.getResponseCode() !== 200) {
        Logger.log("Stripe API error attempt " + (attempt + 1) + ": " + res.getResponseCode() + " " + res.getContentText().substring(0, 200));
        if (attempt < 2) { Utilities.sleep(2000); continue; }
        return { success: false, error: "Stripe API error: " + res.getResponseCode() };
      }

      var data = JSON.parse(res.getContentText());
      Logger.log("Stripe attempt " + (attempt + 1) + "/3: status=" + data.status + " payment_status=" + data.payment_status + " metadata=" + JSON.stringify(data.metadata));

      if ((data.payment_status === "paid" || data.status === "complete") && data.metadata) {
        var uid = data.metadata.user_id || getCurrentUserId();
        var tier = data.metadata.tier;
        if (!tier) { Logger.log("Stripe: no tier in metadata"); if (attempt < 2) { Utilities.sleep(2000); continue; } return { success: false, error: "Missing tier in payment data." }; }
        var currency = data.metadata.currency || "USD";
        var amount = currency === "RWF" ? TIER_PRICES[tier].rwf : TIER_PRICES[tier].usd;
        Logger.log("Stripe: upgrading uid=" + uid + " tier=" + tier + " amount=" + amount + " currency=" + currency);
        upgradeTier(uid, tier, "stripe", sessionId, amount, currency);
        return { success: true, tier: tier, message: "Payment confirmed! " + tier.toUpperCase() + " features unlocked." };
      }

      if (attempt < 2) { Utilities.sleep(2000); continue; }
      Logger.log("Stripe: payment not completed after 3 attempts. payment_status=" + data.payment_status + " status=" + data.status);
      return { success: false, error: "Payment not completed (status: " + data.payment_status + "). Try again or check your Stripe dashboard." };
    } catch(e) {
      Logger.log("Stripe verify error attempt " + (attempt + 1) + ": " + e);
      if (attempt < 2) { Utilities.sleep(2000); continue; }
      return { success: false, error: "Stripe error: " + String(e).substring(0, 100) };
    }
  }
}

// ── PAYMENT: MTN MoMo ────────────────────────────────────────────────────────

function getMoMoToken() {
  try {
    var props = PropertiesService.getScriptProperties();
    var apiUser = props.getProperty("MOMO_API_USER");
    var apiKey = props.getProperty("MOMO_API_KEY");
    var subKey = props.getProperty("MOMO_SUBSCRIPTION_KEY");
    if (!apiUser || !apiKey) return { error: "MoMo API credentials not set." };
    var auth = Utilities.base64Encode(apiUser + ":" + apiKey);
    var res = UrlFetchApp.fetch("https://sandbox.momodeveloper.mtn.com/collection/token/", {
      method: "post",
      headers: { "Authorization": "Basic " + auth, "Ocp-Apim-Subscription-Key": subKey, "Content-Type": "application/json" },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) {
      Logger.log("MoMo token error: " + res.getContentText());
      return { error: "MoMo auth failed: " + res.getContentText().substring(0, 150) };
    }
    return JSON.parse(res.getContentText());
  } catch(e) { return { error: String(e) }; }
}

function createMoMoPayment(uid, phone, tier, currency) {
  try {
    var userId = getCurrentUserId(uid);
    if (!TIER_PRICES[tier]) return { success: false, error: "Invalid tier" };
    var amount = currency === "RWF" ? TIER_PRICES[tier].rwf : TIER_PRICES[tier].usd;
    var token = getMoMoToken();
    if (token.error) return { success: false, error: "MoMo auth: " + token.error };
    var refId = Utilities.getUuid();
    var rawPhone = String(phone).replace(/\D/g, "");
    if (rawPhone.length === 10) rawPhone = "250" + rawPhone;
    if (rawPhone.length < 11) return { success: false, error: "Phone must be 10 digits (078xxxxxxx) or 12 with country code (25078xxxxxxx)." };
    PropertiesService.getScriptProperties().setProperty("MOMO_PENDING_" + refId, JSON.stringify({ userId: userId, tier: tier, currency: currency }));
    var subKey = PropertiesService.getScriptProperties().getProperty("MOMO_SUBSCRIPTION_KEY");
    var res = UrlFetchApp.fetch("https://sandbox.momodeveloper.mtn.com/collection/v1_0/requesttopay", {
      method: "post",
      headers: { "Authorization": "Bearer " + token.access_token, "X-Reference-Id": refId, "X-Target-Environment": "sandbox", "Ocp-Apim-Subscription-Key": subKey, "Content-Type": "application/json" },
      payload: JSON.stringify({ amount: String(amount), currency: "EUR", externalId: refId, payer: { partyIdType: "MSISDN", partyId: rawPhone }, payerMessage: "Umucyo " + tier.toUpperCase(), payeeNote: "Subscription" }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() === 202) return { success: true, refId: refId, message: "Check your phone for a MoMo payment prompt." };
    return { success: false, error: "MoMo: " + res.getContentText().substring(0, 200) };
  } catch(e) { return { success: false, error: String(e) }; }
}

function pollMoMoStatus(refId) {
  try {
    var userId = getCurrentUserId();
    var sub = getSubscription(userId);
    if (sub.tier !== "free" && sub.status !== "trial") return { success: true, message: "Already upgraded." };
    var sheet = getSheet().getSheetByName("SUBSCRIPTIONS");
    if (sheet && sheet.getLastRow() > 1) {
      var rows = sheet.getDataRange().getValues();
      var hdrs = rows[0];
      var refCol = hdrs.indexOf("payment_ref");
      if (refCol >= 0) {
        for (var ri = rows.length - 1; ri >= 1; ri--) {
          if (String(rows[ri][refCol] || "") === refId) return { success: true, message: "Already processed." };
        }
      }
    }
    var token = getMoMoToken();
    if (token.error) return { success: false, error: token.error };
    var subKey = PropertiesService.getScriptProperties().getProperty("MOMO_SUBSCRIPTION_KEY");
    var res = UrlFetchApp.fetch("https://sandbox.momodeveloper.mtn.com/collection/v1_0/requesttopay/" + encodeURIComponent(refId), {
      headers: { "Authorization": "Bearer " + token.access_token, "X-Target-Environment": "sandbox", "Ocp-Apim-Subscription-Key": subKey },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return { success: false };
    var data = JSON.parse(res.getContentText());
    if (data.status === "SUCCESSFUL") {
      var pending = PropertiesService.getScriptProperties().getProperty("MOMO_PENDING_" + refId);
      var tier = "pro", currency = "USD", uid = userId;
      if (pending) {
        try { var p = JSON.parse(pending); tier = p.tier || "pro"; currency = p.currency || "USD"; uid = p.userId || userId; } catch(e) {}
        PropertiesService.getScriptProperties().deleteProperty("MOMO_PENDING_" + refId);
      }
      upgradeTier(uid, tier, "momo", refId, TIER_PRICES[tier].usd, currency);
      return { success: true, message: "Payment confirmed! " + tier.toUpperCase() + " features unlocked." };
    }
    if (data.status === "FAILED") return { success: false, error: "Payment failed." };
    return { success: false };
  } catch(e) { return { success: false }; }
}

// ── PAYMENT: Airtel Money ────────────────────────────────────────────────────

function createAirtelPayment(phone, tier, currency) {
  try {
    var userId = getCurrentUserId();
    Logger.log("[MOCK AIRTEL] " + userId + " phone=" + phone + " tier=" + tier);
    return { success: true, refId: "airtel_" + Date.now(), message: "Mock Airtel initiated." };
  } catch(e) { return { success: false, error: String(e) }; }
}

function pollAirtelStatus(refId) {
  upgradeTier(getCurrentUserId(), "pro", "airtel", refId, 20, "USD");
  return { status: "SUCCESSFUL", message: "Payment confirmed (mock)." };
}

// ── PAYMENT: BANK TRANSFER ─────────────────────────────────────────────────────

function submitBankProof(tier, reference, fileUrl, currency) {
  try {
    var userId = getCurrentUserId();
    if (!TIER_PRICES[tier]) throw new Error("Invalid tier");
    var sheet = getSheet().getSheetByName("SUBSCRIPTIONS");
    if (!sheet) { setupSheets(); sheet = getSheet().getSheetByName("SUBSCRIPTIONS"); }
    var amount = currency === "RWF" ? TIER_PRICES[tier].rwf : TIER_PRICES[tier].usd;
    sheet.appendRow([userId, tier, "pending_bank", TIER_FEATURES[tier].match_limit, JSON.stringify(TIER_FEATURES[tier]), "bank", reference || "", amount, currency || "USD", new Date(), null, false]);
    logActivity(userId, "bank_transfer_submitted");
    return { success: true, message: "Bank transfer submitted. Awaiting admin approval." };
  } catch(e) { return { success: false, error: String(e) }; }
}

function getBankDetails() {
  return {
    bank_name: PropertiesService.getScriptProperties().getProperty("BANK_NAME") || "—",
    account_name: PropertiesService.getScriptProperties().getProperty("BANK_ACCOUNT_NAME") || "—",
    account_number: PropertiesService.getScriptProperties().getProperty("BANK_ACCOUNT_NUMBER") || "—",
    reference_prefix: "UMU"
  };
}

// ── ADMIN ────────────────────────────────────────────────────────────────────

function isAdmin(uid) { 
  var userId = getCurrentUserId(uid); 
  return userId === "geraldnoria@gmail.com"; 
}

function getAdminDashboard() {
  if (!isAdmin()) return { error: "Access denied. Admin only." };
  try {
    var sheet = getSheet().getSheetByName("SUBSCRIPTIONS");
    var pending = [];
    var active = [];
    var stats = { total: 0, pro: 0, premium: 0, revenue: 0 };
    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var col = function(n) { return headers.indexOf(n); };
      for (var i = 1; i < data.length; i++) {
        var uid = String(data[i][col("user_id")] || "");
        var tier = String(data[i][col("tier")] || "");
        var status = String(data[i][col("status")] || "");
        if (status === "pending_bank") {
          pending.push({ user_id: uid, tier: tier, amount: Number(data[i][col("amount")] || 0), currency: String(data[i][col("currency")] || "USD"), payment_ref: String(data[i][col("payment_ref")] || ""), _rowIndex: i });
        }
        if (status === "active" && tier !== "free") {
          active.push({ user_id: uid, tier: tier, expires: String(data[i][col("expires_at")] || ""), amount: Number(data[i][col("amount")] || 0), provider: String(data[i][col("provider")] || "") });
          stats.total++;
          if (tier === "pro") stats.pro++;
          if (tier === "premium") stats.premium++;
          stats.revenue += Number(data[i][col("amount")] || 0);
        }
      }
    }
    return { pending: pending, active: active, stats: stats };
  } catch(e) { return { error: String(e) }; }
}

function approveBankPayment(userId, tier) {
  if (!isAdmin()) return { error: "Access denied" };
  upgradeTier(userId, tier, "bank_approved", "admin_approved", 0, "USD");
  return { success: true };
}

function rejectBankPayment(userId, reason) {
  if (!isAdmin()) return { error: "Access denied" };
  try {
    var sheet = getSheet().getSheetByName("SUBSCRIPTIONS");
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var col = function(n) { return headers.indexOf(n); };
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][col("user_id")] || "").toLowerCase() === userId.toLowerCase() && String(data[i][col("status")] || "") === "pending_bank") {
        sheet.getRange(i + 1, col("status") + 1).setValue("rejected");
        return { success: true };
      }
    }
  } catch(e) {}
  return { error: "Not found" };
}

function retryStripeVerification(sessionId) { return verifyStripeSession(sessionId); }
function createStripePortal() { return { url: "#", message: "Stripe portal (mock)." }; }

function refreshAndRegenerateForAll() {
  try {
    Logger.log("=== Refreshing tenders from BigQuery ===");
    syncTendersFromBigQuery();
    Logger.log("=== Regenerating matches for all users ===");
    var allUids = {};
    ["USER_PROFILES","MATCHES","SUBSCRIPTIONS"].forEach(function(sn) {
      var sh = getSheet().getSheetByName(sn);
      if (!sh || sh.getLastRow() < 2) return;
      var d = sh.getDataRange().getValues();
      var hIdx = d[0].indexOf("user_id");
      if (hIdx < 0) return;
      for (var i = 1; i < d.length; i++) {
        var u = String(d[i][hIdx] || "").trim().toLowerCase();
        if (u) allUids[u] = true;
      }
    });
    var ids = Object.keys(allUids);
    Logger.log("Regenerating matches for " + ids.length + " users...");
    for (var ui = 0; ui < ids.length; ui++) {
      try { generateMatchesNoAI(ids[ui]); } catch(e) { Logger.log("Failed for " + ids[ui] + ": " + e); }
    }
    Logger.log("=== Refresh complete ===");
  } catch(e) { Logger.log("refreshAndRegenerateForAll: " + e); }
}

function sendWeeklyDigest() {
  try {
    refreshAndRegenerateForAll();
    var shortUrl = "https://shorturl.at/Mztej";
    var users = {};
    var sheet = getSheet().getSheetByName("SUBSCRIPTIONS");
    if (sheet && sheet.getLastRow() > 1) {
      var sData = sheet.getDataRange().getValues();
      var sHeaders = sData[0];
      var sc = function(n) { return sHeaders.indexOf(n); };
      for (var si = 1; si < sData.length; si++) {
        var uid = String(sData[si][sc("user_id")] || "");
        var sStatus = String(sData[si][sc("status")] || "");
        if (sStatus === "active" || sStatus === "trial") users[uid] = { email: uid, tier: String(sData[si][sc("tier")] || "free"), status: sStatus };
      }
    }

    var mSheet = getSheet().getSheetByName("MATCHES");
    if (!mSheet || mSheet.getLastRow() < 2) { Logger.log("No matches."); return; }
    var mData = mSheet.getDataRange().getValues();
    var mHeaders = mData[0];
    var mc = function(n) { return mHeaders.indexOf(n); };

    var tSheet = getSheet().getSheetByName("TENDERS_FLAT");
    var tLookup = {};
    if (tSheet && tSheet.getLastRow() > 1) {
      var tData = tSheet.getDataRange().getValues();
      var tHeaders = tData[0];
      var ti = function(n) { return tHeaders.indexOf(n); };
      for (var ti_ = 1; ti_ < tData.length; ti_++) {
        var o = String(tData[ti_][ti("ocid")] || "");
        if (o) tLookup[o] = { status: String(tData[ti_][ti("status")] || ""), deadline: String(tData[ti_][ti("deadline")] || "") };
      }
    }

    var jLookup = {};
    var jirSheet = getSheet().getSheetByName("JIR_TENDERS");
    if (jirSheet && jirSheet.getLastRow() > 1) {
      var jirData = jirSheet.getDataRange().getValues();
      for (var ji = 1; ji < jirData.length; ji++) {
        var jo = String(jirData[ji][6] || "");
        if (jo) jLookup[jo] = { status: "active", deadline: String(jirData[ji][4] || "") };
      }
    }

    for (var mi = 1; mi < mData.length; mi++) {
      var muid = String(mData[mi][mc("user_id")] || "");
      if (muid && !users[muid]) users[muid] = { email: muid, tier: "free", status: "active" };
    }

    Logger.log("Users found: " + Object.keys(users).length + " | Match rows: " + (mData.length - 1));
    var now = new Date();
    var sent = 0;
    for (var email in users) {
      var user = users[email];
      var userName = email.split("@")[0].replace(/[^a-zA-Z]/g, " ").trim();
      userName = userName.charAt(0).toUpperCase() + userName.slice(1);
      var isFree = user.tier === "free" && user.status !== "trial";

      var allMatches = [];
      for (var ri = 1; ri < mData.length; ri++) {
        var r = mData[ri];
        if (String(r[mc("user_id")] || "").toLowerCase() !== email.toLowerCase()) continue;
        var ocid = String(r[mc("ocid")] || "");
        var tender = ocid ? (tLookup[ocid] || jLookup[ocid]) : null;
        var tenderStatus = tender && tender.status ? String(tender.status).toLowerCase() : "";
        var isActive = tenderStatus === "active";
        allMatches.push({ row: r, isActive: isActive });
        if (allMatches.length >= 10) break;
      }

      if (!allMatches.length) { Logger.log("User " + email + " has 0 matches total"); continue; }

      allMatches.sort(function(a, b) {
        if (a.isActive && !b.isActive) return -1;
        if (!a.isActive && b.isActive) return 1;
        return 0;
      });

      var activeCount = allMatches.filter(function(m) { return m.isActive; }).length;
      var closedCount = allMatches.length - activeCount;

      // Find "closing this week" tenders (deadline within 7 days)
      var closingThisWeek = allMatches.filter(function(m) {
        if (!m.isActive) return false;
        var dl = String(m.row[mc("deadline")] || "");
        if (!dl) return false;
        try { var dt = new Date(dl); return !isNaN(dt) && dt > now && dt.getTime() - now.getTime() <= 7 * 86400000; } catch(e) { return false; }
      });

      var matchCards = "";

      // "Closing This Week" banner
      if (closingThisWeek.length > 0) {
        matchCards += '<tr><td style="background:#FFF3CD;color:#856404;padding:10px 14px;border-radius:8px;font-size:13px;font-weight:600;margin-bottom:10px">⚠ ' + closingThisWeek.length + ' tender' + (closingThisWeek.length > 1 ? 's' : '') + ' closing this week!</td></tr><tr><td style="height:8px"></td></tr>';
        closingThisWeek.forEach(function(entry, idx) {
          var m = entry.row;
          var title = String(m[mc("title")] || "Tender");
          var entity = String(m[mc("entity")] || "");
          var score = parseInt(m[mc("match_score")]) || 0;
          var dl = String(m[mc("deadline")] || "");
          var dlStr = "";
          if (dl) {
            try { var dt = new Date(dl); var daysLeft = Math.ceil((dt - now) / 86400000); dlStr = daysLeft + " day" + (daysLeft > 1 ? "s" : "") + " left"; } catch(e) { dlStr = dl; }
          }
          matchCards += '<tr><td style="padding:12px 14px;border:1px solid #FFEAA7;border-radius:8px;background:#FFFCF0">';
          matchCards += '<div style="font-size:13px;font-weight:600;color:#856404;margin-bottom:2px">' + (idx + 1) + '. ' + title.replace(/</g,'&lt;') + '</div>';
          matchCards += '<div style="font-size:11px;color:#888"><b>' + entity.replace(/</g,'&lt;') + '</b> · <b>' + score + '%</b> match · ' + dlStr + '</div>';
          matchCards += '<a href="' + shortUrl + '" style="font-size:11px;color:#2060C0;font-weight:600;text-decoration:none">View →</a>';
          matchCards += '</td></tr><tr><td style="height:6px"></td></tr>';
        });
        matchCards += '<tr><td style="height:8px"></td></tr>';
      }

      allMatches.forEach(function(entry, idx) {
        var m = entry.row;
        var title = String(m[mc("title")] || "Tender");
        var entity = String(m[mc("entity")] || "");
        var score = parseInt(m[mc("match_score")]) || 0;
        var wp = parseFloat(m[mc("win_probability")]);
        var wpPct = wp ? Math.round(wp * 100) : null;
        var estValue = parseInt(m[mc("est_value")]) || 0;
        var valStr = estValue >= 1e9 ? (estValue / 1e9).toFixed(1) + "B" : (estValue >= 1e6 ? (estValue / 1e6).toFixed(0) + "M" : estValue.toLocaleString());
        var statusBadge = entry.isActive
          ? '<span style="background:#d4edda;color:#155724;padding:2px 8px;border-radius:6px;font-size:11px;font-weight:600;margin-left:4px">ACTIVE</span>'
          : '<span style="background:#f8d7da;color:#721c24;padding:2px 8px;border-radius:6px;font-size:11px;font-weight:600;margin-left:4px">CLOSED</span>';
        var winBadge = wpPct ? '<span style="background:#d4edda;color:#155724;padding:2px 8px;border-radius:6px;font-size:11px;font-weight:600">' + wpPct + '% win</span>' : '';
        matchCards += '<tr><td style="padding:14px;border:1px solid #e8eaed;border-radius:10px;background:#fff">';
        matchCards += '<div style="font-size:14px;font-weight:600;color:#1C2833;margin-bottom:4px">' + (idx + 1) + '. ' + title.replace(/</g,'&lt;') + statusBadge + '</div>';
        matchCards += '<div style="font-size:12px;color:#888;margin-bottom:6px"><b>' + entity.replace(/</g,'&lt;') + '</b> · ' + valStr + ' RWF · <b>' + score + '%</b> match ' + winBadge + '</div>';
        matchCards += ' <a href="' + shortUrl + '" style="font-size:11px;color:#2060C0;font-weight:600;text-decoration:none">View in app →</a>';
        matchCards += '</td></tr><tr><td style="height:8px"></td></tr>';
      });

      var summaryLine = "Here are <b>" + allMatches.length + " tenders</b> matching your profile";
      if (activeCount > 0) summaryLine += " — <b style='color:#155724;'>" + activeCount + " active</b>";
      if (closedCount > 0) summaryLine += " <span style='color:#888;'>(" + closedCount + " closed)</span>";
      summaryLine += ". Each one links to your full analysis.";

      var upgradeCta = "";
      if (isFree) {
        upgradeCta = '<tr><td style="padding:16px;background:linear-gradient(135deg,#2060C0,#F4A940);border-radius:10px;text-align:center">';
        upgradeCta += '<div style="font-size:14px;font-weight:700;color:#fff;margin-bottom:6px">Upgrade to Pro for RWF 25,000/mo</div>';
        upgradeCta += '<div style="font-size:11px;color:rgba(255,255,255,0.85);margin-bottom:10px">Unlock win probabilities, competitor names, buyer contacts & bid recommendations</div>';
        upgradeCta += '<a href="' + shortUrl + '" style="display:inline-block;padding:8px 20px;background:#F4A940;color:#1C2833;border-radius:8px;font-size:12px;font-weight:600;text-decoration:none">View Plans →</a>';
        upgradeCta += '</td></tr><tr><td style="height:12px"></td></tr>';
      }

      var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;font-family:Arial,Helvetica,sans-serif;background:#F4F6F9">';
      html += '<table width="100%" cellpadding="0" cellspacing="0" style="background:#F4F6F9;padding:20px 0"><tr><td align="center">';
      html += '<table width="580" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden">';
      html += '<tr><td style="background:#1C2833;padding:24px 28px;text-align:center">';
      html += '<div style="font-size:22px;font-weight:700;color:#fff">Umucyo<span style="color:#F4A940">Assistant</span></div>';
      html += '<div style="font-size:12px;color:rgba(255,255,255,0.6);margin-top:4px">Your weekly tender intelligence</div>';
      html += '</td></tr>';
      html += '<tr><td style="padding:24px 28px">';
      html += '<div style="font-size:15px;color:#1C2833;line-height:1.6;margin-bottom:20px">' + userName + ',</div>';
      html += '<div style="font-size:15px;color:#1C2833;line-height:1.7;margin-bottom:20px">' + summaryLine + '</div>';
      html += '<table width="100%" cellpadding="0" cellspacing="0">' + matchCards + '</table>';
      if (upgradeCta) html += '<table width="100%" cellpadding="0" cellspacing="0">' + upgradeCta + '</table>';
      html += '<div style="font-size:14px;color:#555;line-height:1.7;margin-top:16px">Every match in the app includes details this email can not show: <b>who your competitors are</b>, <b>what price range actually wins</b>, <b>the buyer phone number and email</b>, and <b>whether this buyer changes contracts after signing</b>. Open the app to see it all.</div>';
      html += '<div style="text-align:center;margin-top:20px"><a href="' + shortUrl + '" style="display:inline-block;padding:12px 28px;background:#2060C0;color:#fff !important;border-radius:8px;font-size:14px;font-weight:600;text-decoration:none">See All My Matches →</a></div>';
      html += '</td></tr>';
      html += '<tr><td style="padding:20px 28px;border-top:1px solid #e8eaed;background:#F4F6F9">';
      html += '<div style="font-size:13px;color:#1C2833;line-height:1.8"><b>Gerald Denor</b></div>';
      html += '<div style="font-size:12px;color:#888">EazyPickins Ltd · Kigali, RW</div>';
      html += '<div style="font-size:12px;color:#888">TIN: 155977574 · Phone/WhatsApp: 0793378679</div>';
      html += '<div style="font-size:12px;margin-top:6px"><a href="' + shortUrl + '" style="color:#2060C0;font-weight:600;text-decoration:none">UmucyoAssistant</a></div>';
      html += '</td></tr>';
      html += '<tr><td style="padding:12px 28px;background:#1C2833;text-align:center;font-size:10px;color:rgba(255,255,255,0.5)">You are receiving this because you have an active Umucyo Assistant account. Tender insights are based on RPPA historical data — always verify with official Umucyo sources before bidding.</td></tr>';
      html += '</table></td></tr></table></body></html>';

      try {
        var subj = userName + ", " + allMatches.length + " tenders";
        if (activeCount > 0) subj += " — " + activeCount + " active";
        MailApp.sendEmail({ to: email, subject: subj, htmlBody: html });
        sent++;
      } catch(e) { Logger.log("Failed to email " + email + ": " + e); }
    }
    Logger.log("Weekly digest sent to " + sent + " users.");
  } catch(e) { Logger.log("sendWeeklyDigest: " + e); }
}

function countActiveTenders() {
  try {
    Logger.log("=== Counting tenders in TENDERS_FLAT ===");
    var tSheet = getSheet().getSheetByName("TENDERS_FLAT");
    if (!tSheet || tSheet.getLastRow() < 2) { Logger.log("TENDERS_FLAT is empty."); return; }
    var tData = tSheet.getDataRange().getValues();
    var tHeaders = tData[0];
    var ti = function(n) { return tHeaders.indexOf(n); };
    var now = new Date();
    var active = 0, closed = 0, noDl = 0, total = 0;
    var statusCounts = {};
    for (var i = 1; i < tData.length; i++) {
      total++;
      var s = String(tData[i][ti("status")] || "(empty)");
      statusCounts[s] = (statusCounts[s] || 0) + 1;
      var dl = String(tData[i][ti("deadline")] || "");
      if (dl) {
        var d = new Date(dl);
        if (!isNaN(d.getTime()) && d >= now) active++;
        else closed++;
      } else {
        noDl++;
      }
    }
    Logger.log("TENDERS_FLAT: " + total + " total, " + active + " active (future deadline), " + closed + " closed (past deadline), " + noDl + " no deadline");
    Logger.log("Status distribution: " + JSON.stringify(statusCounts));

    Logger.log("=== Counting from BigQuery ===");
    Logger.log("Running: SELECT COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE tender_tenderPeriod_endDate >= TIMESTAMP(CURRENT_DATE())");
    var futureDl = queryBigQuery("SELECT COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE tender_tenderPeriod_endDate >= TIMESTAMP(CURRENT_DATE())");
    Logger.log("BigQuery active (future deadline): " + JSON.stringify(futureDl));
    var allStatuses = queryBigQuery("SELECT tender_status AS status, COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 180 DAY) GROUP BY tender_status ORDER BY n DESC LIMIT 20");
    Logger.log("BigQuery status distribution (last 180 days): " + JSON.stringify(allStatuses));
    var activeByStatus = queryBigQuery("SELECT tender_status AS status, COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE tender_tenderPeriod_endDate >= TIMESTAMP(CURRENT_DATE()) GROUP BY tender_status ORDER BY n DESC LIMIT 10");
    Logger.log("BigQuery active-by-status (future deadline): " + JSON.stringify(activeByStatus));
    Logger.log("=== Done ===");
  } catch(e) { Logger.log("countActiveTenders ERROR: " + e); }
}

function installWeeklyDigest() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { if (t.getHandlerFunction() === "sendWeeklyDigest") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("sendWeeklyDigest").timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  Logger.log("Weekly digest trigger installed. Emails will go out every Monday at 8 AM.");
}

function installJIRScraper() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { if (t.getHandlerFunction() === "scrapeJobInRwanda") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("scrapeJobInRwanda").timeBased().everyDays(1).atHour(7).create();
  Logger.log("Daily JIR scraper installed. Will run every morning at 7 AM.");
}

function scrapeJobInRwanda() {
  Logger.log("=== SCRAPING JOBINRWANDA ===");
  var pages = [
    { url: "https://www.jobinrwanda.com/jobs/tender", type: "Tender" },
    { url: "https://www.jobinrwanda.com/jobs/consultancy", type: "Consultancy" }
  ];

  var tenders = [];
  pages.forEach(function(page) {
    try {
      var res = UrlFetchApp.fetch(page.url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) { Logger.log(page.url + " HTTP " + res.getResponseCode()); return; }
      var html = res.getContentText();

      var blocks = html.match(/<article[^>]*class="[^"]*node--type-job[^"]*"[^>]*>[\s\S]*?<\/article>/g);
      if (!blocks) { Logger.log("No article blocks found on " + page.url); return; }
      Logger.log(page.url + ": found " + blocks.length + " article blocks");

      blocks.forEach(function(block) {
        var titleMatch = block.match(/<span[^>]*class="[^"]*field--name-title[^"]*"[^>]*>([^<]+)<\/span>/);
        var employerMatch = block.match(/<a[^>]*href="\/employer\/[^"]*"[^>]*>([^<]+)<\/a>/);
        var dateMatch = block.match(/Published on (\d{2}-\d{2}-\d{4})/);
        var deadlineMatch = block.match(/<time[^>]*datetime="[^"]*"[^>]*>(\d{2}-\d{2}-\d{4})<\/time>/);
        var locationMatch = block.match(/fa-map-marker-alt"><\/i>\s*([^|]+)\s*\|/);

        if (!titleMatch) return;
        var title = titleMatch[1].replace(/�/g, "'").trim();
        if (!title || title.length < 8) return;
        var linkMatch = block.match(/<a[^>]*href="(\/job\/[^"]*)"[^>]*>/);
        var link = linkMatch ? "https://www.jobinrwanda.com" + linkMatch[1] : "";
        var employer = employerMatch ? employerMatch[1].trim() : "";
        var published = dateMatch ? dateMatch[1] : "";
        var deadline = deadlineMatch ? deadlineMatch[1] : "";
        var location = locationMatch ? locationMatch[1].trim() : "";

        tenders.push({
          title: title, employer: employer, location: location,
          published: published, deadline: deadline, type: page.type,
          link: link, scraped: new Date().toISOString()
        });
      });
    } catch(e) { Logger.log(page.url + " error: " + e); }
  });

  if (!tenders.length) { Logger.log("No tenders found."); return; }

  var seen = {};
  tenders = tenders.filter(function(t) {
    var key = t.title.toLowerCase().trim() + "|" + t.employer.toLowerCase().trim();
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });

  var sheet = getSheet().getSheetByName("JIR_TENDERS");
  if (!sheet) { var ss = getSheet(); sheet = ss.insertSheet("JIR_TENDERS"); }
  var existing = {};
  if (sheet.getLastRow() > 1) {
    var existingData = sheet.getDataRange().getValues();
    for (var ei = 1; ei < existingData.length; ei++) { existing[String(existingData[ei][0] || "")] = true; }
  }

  var newOnly = tenders.filter(function(t) { return !existing[t.title]; });
  if (!newOnly.length) { Logger.log("No new tenders. All " + tenders.length + " already in sheet."); return; }

  var headers = ["Title","Employer","Location","Published","Deadline","Type","Link","Scraped"];
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var batch = newOnly.map(function(t) { return [t.title, t.employer, t.location, t.published, t.deadline, t.type, t.link, t.scraped]; });
  sheet.getRange(sheet.getLastRow() + 1, 1, batch.length, headers.length).setValues(batch);
  SpreadsheetApp.flush();

  Logger.log("Added " + newOnly.length + " new tenders. Sheet total: " + (sheet.getLastRow() - 1));
  Logger.log("Sample:");
  newOnly.slice(0, 3).forEach(function(t) { Logger.log(t.title + " | " + t.employer + " | Deadline: " + t.deadline); });
  Logger.log("=== DONE ===");
}

function discoverTables() {
  var candidates = ["rppc_contracts","rppc_contracts_amendments","rppc_tender_lots","rppc_tender_documents","rppc_related_processes","rppc_party_classifications","rppc_tenders_view","rppc_win_signals","rppc_buyer_signals","rppc_tender_predictions"];
  candidates.forEach(function(t) {
    try {
      var c = queryBigQuery("SELECT COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + "." + t);
      Logger.log(t + ": " + (c.length > 0 ? c[0].n + " rows" : "EMPTY") + " — EXISTS");
    } catch(e) { Logger.log(t + ": MISSING"); }
  });
}

function debugPhase8Schema() {
  var tables = ["rppc_contracts","rppc_contracts_amendments"];
  tables.forEach(function(t) {
    try {
      var rows = queryBigQuery("SELECT * FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + "." + t + " LIMIT 1");
      if (rows.length > 0) Logger.log(t + " cols: " + Object.keys(rows[0]).join(", "));
    } catch(e) { Logger.log(t + ": ERROR - " + e.message); }
  });
}

function discoverBidderContacts() {
  var dataset = BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET;
  Logger.log("=== BIDDER CONTACT DISCOVERY ===");

  var tables = [
    { name: "rppc_buyers", desc: "Parties (all entities)" },
    { name: "rppc_award_suppliers", desc: "Supplier names linked to awards" },
    { name: "rppc_bidders", desc: "Who bid on each tender" },
    { name: "rppc_tender_tenderers", desc: "Bidders per tender" },
    { name: "rppc_party_classifications", desc: "Supplier classification codes" },
    { name: "rppc_buyer_contacts", desc: "Buyer-only contacts (already built)" }
  ];

  tables.forEach(function(t) {
    try {
      var c = queryBigQuery("SELECT COUNT(*) AS n FROM " + dataset + "." + t.name);
      var n = c.length > 0 ? c[0].n : 0;
      Logger.log(t.name + ": " + n + " rows — " + t.desc);
    } catch(e) { Logger.log(t.name + ": MISSING"); }
  });

  Logger.log("---");
  Logger.log("Checking rppc_buyers columns...");
  try {
    var rows = queryBigQuery("SELECT * FROM " + dataset + ".rppc_buyers LIMIT 1");
    if (rows.length > 0) Logger.log("rppc_buyers cols: " + Object.keys(rows[0]).join(", "));
  } catch(e) { Logger.log("rppc_buyers schema: " + e); }

  Logger.log("---");
  Logger.log("Sample rppc_buyers rows:");
  try {
    var sample = queryBigQuery("SELECT * FROM " + dataset + ".rppc_buyers WHERE name IS NOT NULL ORDER BY RAND() LIMIT 5");
    sample.forEach(function(r) { Logger.log(JSON.stringify(r).substring(0, 800)); });
  } catch(e) { Logger.log("Sample: " + e.message); }

  Logger.log("---");
  Logger.log("Join test: suppliers with contact details...");
  try {
    var joinSql = "SELECT DISTINCT s.name, p.contact_phone, p.contact_email, p.streetAddress, p.locality, COUNT(a.main_ocid) AS awards FROM " + dataset + ".rppc_award_suppliers s JOIN " + dataset + ".rppc_awards a ON s.main_ocid = a.main_ocid AND s.awards_id = a.id LEFT JOIN " + dataset + ".rppc_buyers p ON LOWER(s.name) = LOWER(p.name) WHERE (p.contact_phone IS NOT NULL OR p.contact_email IS NOT NULL) GROUP BY s.name, p.contact_phone, p.contact_email, p.streetAddress, p.locality HAVING COUNT(a.main_ocid) > 0 ORDER BY awards DESC LIMIT 10";
    var joinRows = queryBigQuery(joinSql);
    Logger.log("Matched suppliers: " + joinRows.length);
    joinRows.forEach(function(r) { Logger.log(JSON.stringify(r).substring(0, 600)); });
  } catch(e) { Logger.log("Join: " + e.message); }

  Logger.log("=== DONE ===");
}

function extractBidderContacts() {
  var dataset = BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET;
  Logger.log("=== EXTRACTING BIDDER CONTACTS ===");

  var sql = "WITH unique_suppliers AS (SELECT DISTINCT name, identifier_legalName AS legal_name, identifier_id AS tin, contactPoint_name AS contact_person, contactPoint_email AS email, contactPoint_telephone AS phone, address_streetAddress AS street, address_locality AS locality, address_countryName AS country, address_region AS region, address_postalCode AS postal_code FROM " + dataset + ".rppc_buyers WHERE LOWER(roles) LIKE '%tenderer%' AND (contactPoint_email IS NOT NULL OR contactPoint_telephone IS NOT NULL)), award_counts AS (SELECT s.name AS sname, COUNT(a.main_ocid) AS awards, SUM(a.value_amount) AS total_value FROM " + dataset + ".rppc_award_suppliers s JOIN " + dataset + ".rppc_awards a ON s.main_ocid = a.main_ocid AND s.awards_id = a.id GROUP BY s.name) SELECT u.*, COALESCE(ac.awards, 0) AS awards, COALESCE(ac.total_value, 0) AS total_value FROM unique_suppliers u LEFT JOIN award_counts ac ON LOWER(u.name) = LOWER(ac.sname) ORDER BY awards DESC, total_value DESC LIMIT 50000";

  Logger.log("Running combined query...");
  var rows = queryBigQuery(sql, 50000);
  Logger.log("Extracted " + rows.length + " suppliers with contacts and award data");

  var sheet = getSheet().getSheetByName("BIDDER_CONTACTS");
  if (!sheet) { var ss = getSheet(); sheet = ss.insertSheet("BIDDER_CONTACTS"); }
  sheet.clear();
  var headers = ["Name","Legal Name","TIN","Contact Person","Email","Phone","Street","Locality","Country","Region","Postal Code","Awards Won","Total Award Value RWF"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var batchSize = 500;
  for (var i = 0; i < rows.length; i += batchSize) {
    var batch = rows.slice(i, Math.min(i + batchSize, rows.length)).map(function(r) {
      return [String(r.name||""), String(r.legal_name||""), String(r.tin||""), String(r.contact_person||""), String(r.email||""), String(r.phone||""), String(r.street||""), String(r.locality||""), String(r.country||""), String(r.region||""), String(r.postal_code||""), parseInt(r.awards)||0, Math.round(parseFloat(r.total_value)||0)];
    });
    sheet.getRange(i + 2, 1, batch.length, headers.length).setValues(batch);
  }
  SpreadsheetApp.flush();
  Logger.log("Saved " + rows.length + " contacts. Top 5:");
  rows.slice(0, 5).forEach(function(r) { Logger.log(r.name + " | " + r.email + " | " + r.awards + " awards"); });
  Logger.log("=== DONE ===");
}

function deduplicateBidderContacts() {
  var sheet = getSheet().getSheetByName("BIDDER_CONTACTS");
  if (!sheet || sheet.getLastRow() < 2) { Logger.log("BIDDER_CONTACTS sheet not found or empty."); return; }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var col = function(n) { return headers.indexOf(n); };

  var unique = {};
  var dups = 0;
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][col("Name")] || "").toLowerCase().trim();
    var email = String(data[i][col("Email")] || "").toLowerCase().trim();
    var phone = String(data[i][col("Phone")] || "").replace(/\D/g, "");
    var awards = parseInt(data[i][col("Awards Won")]) || 0;
    var key = [name, email, phone].join("|||");

    if (!unique[key] || unique[key].awards < awards) {
      unique[key] = { row: data[i], awards: awards };
      if (unique[key]) dups++;
    }
  }

  var deduped = Object.values(unique).map(function(v) { return v.row; });
  var removed = data.length - 1 - deduped.length;

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var batchSize = 500;
  for (var j = 0; j < deduped.length; j += batchSize) {
    var batch = deduped.slice(j, Math.min(j + batchSize, deduped.length));
    sheet.getRange(j + 2, 1, batch.length, headers.length).setValues(batch);
  }
  SpreadsheetApp.flush();
  Logger.log("Deduplication complete: " + deduped.length + " unique contacts kept, " + removed + " duplicates removed.");
}

function diagnoseMainColumns() {
  var dataset = BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET;
  Logger.log("=== rppc_main COLUMNS ===");
  try {
    var rows = queryBigQuery("SELECT * FROM " + dataset + ".rppc_main LIMIT 1");
    if (rows.length > 0) Logger.log("Columns: " + Object.keys(rows[0]).join(", "));
  } catch(e) { Logger.log("Error: " + e); }
  var cols = ["tender_procurementMethod","tender_tenderPeriod_endDate"];
  cols.forEach(function(c) {
    try { var r = queryBigQuery("SELECT " + c + " FROM " + dataset + ".rppc_main LIMIT 1"); Logger.log(c + ": ✅"); } catch(e) { Logger.log(c + ": ❌"); }
  });
  Logger.log("=== DONE ===");
}

function diagnoseLatestMonth() {
  var d = BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET;
  Logger.log("=== LATEST DATA MONTH ===");
  try {
    var r = queryBigQuery("SELECT FORMAT_TIMESTAMP('%Y-%m', MAX(date)) AS latest_month, COUNT(*) AS record_count FROM " + d + ".rppc_main");
    if (r.length > 0) Logger.log("rppc_main latest: " + r[0].latest_month + " (" + r[0].record_count + " records total)");
  } catch(e) { Logger.log("rppc_main: " + e); }
  try {
    var r2 = queryBigQuery("SELECT FORMAT_TIMESTAMP('%Y-%m', MAX(date)) AS latest_month, COUNT(*) AS record_count FROM " + d + ".rppc_main WHERE date IS NOT NULL GROUP BY FORMAT_TIMESTAMP('%Y-%m', date) ORDER BY latest_month DESC LIMIT 5");
    Logger.log("Recent months in rppc_main:");
    r2.forEach(function(rr) { Logger.log("  " + rr.latest_month + ": " + rr.record_count + " records"); });
  } catch(e) { Logger.log("Breakdown: " + e); }
  try {
    var r3 = queryBigQuery("SELECT FORMAT_TIMESTAMP('%Y-%m', MAX(date)) AS latest, COUNT(*) AS cnt FROM " + d + ".rppc_awards");
    if (r3.length > 0) Logger.log("rppc_awards latest: " + r3[0].latest + " (" + r3[0].cnt + " records)");
  } catch(e) {}
  try {
    var r4 = queryBigQuery("SELECT FORMAT_TIMESTAMP('%Y-%m', MAX(date)) AS latest, COUNT(*) AS cnt FROM " + d + ".rppc_contracts");
    if (r4.length > 0) Logger.log("rppc_contracts latest: " + r4[0].latest + " (" + r4[0].cnt + " records)");
  } catch(e) {}
  Logger.log("=== DONE ===");
}

function diagnoseContactCounts() {
  var dataset = BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET;
  Logger.log("=== CONTACT DIAGNOSTICS ===");

  try {
    var sql1 = "SELECT COUNT(*) AS n FROM " + dataset + ".rppc_buyers";
    var r1 = queryBigQuery(sql1);
    Logger.log("Total parties in rppc_buyers: " + (r1[0] ? r1[0].n : "?"));
  } catch(e) { Logger.log("A: " + e); }

  try {
    var sql2 = "SELECT roles, COUNT(*) AS n FROM " + dataset + ".rppc_buyers WHERE roles IS NOT NULL GROUP BY roles ORDER BY n DESC LIMIT 10";
    var r2 = queryBigQuery(sql2);
    Logger.log("Role breakdown:");
    r2.forEach(function(r) { Logger.log("  " + r.roles + ": " + r.n); });
  } catch(e) { Logger.log("B: " + e); }

  try {
    var sql3 = "SELECT COUNT(*) AS n FROM " + dataset + ".rppc_buyers WHERE LOWER(roles) LIKE '%tenderer%'";
    var r3 = queryBigQuery(sql3);
    Logger.log("Parties with 'tenderer' role: " + (r3[0] ? r3[0].n : "?"));
  } catch(e) { Logger.log("C: " + e); }

  try {
    var sql4 = "SELECT COUNT(*) AS n FROM " + dataset + ".rppc_buyers WHERE LOWER(roles) LIKE '%tenderer%' AND (contactPoint_email IS NOT NULL OR contactPoint_telephone IS NOT NULL)";
    var r4 = queryBigQuery(sql4);
    Logger.log("Tenderers WITH email/phone: " + (r4[0] ? r4[0].n : "?"));
  } catch(e) { Logger.log("D: " + e); }

  try {
    var sql5 = "SELECT COUNT(*) AS n FROM " + dataset + ".rppc_buyers WHERE LOWER(roles) LIKE '%tenderer%' AND contactPoint_email IS NOT NULL";
    var r5 = queryBigQuery(sql5);
    Logger.log("Tenderers with email: " + (r5[0] ? r5[0].n : "?"));
  } catch(e) { Logger.log("E: " + e); }

  try {
    var sql6 = "SELECT COUNT(*) AS n FROM " + dataset + ".rppc_buyers WHERE LOWER(roles) LIKE '%tenderer%' AND contactPoint_telephone IS NOT NULL";
    var r6 = queryBigQuery(sql6);
    Logger.log("Tenderers with phone: " + (r6[0] ? r6[0].n : "?"));
  } catch(e) { Logger.log("F: " + e); }

  try {
    var sql7 = "SELECT COUNT(DISTINCT name) AS n FROM " + dataset + ".rppc_award_suppliers";
    var r7 = queryBigQuery(sql7);
    Logger.log("Unique award_supplier names: " + (r7[0] ? r7[0].n : "?"));
  } catch(e) { Logger.log("G: " + e); }

  try {
    var sql8 = "SELECT COUNT(*) AS n FROM " + dataset + ".rppc_buyers p JOIN (SELECT DISTINCT name FROM " + dataset + ".rppc_award_suppliers) s ON LOWER(p.name) = LOWER(s.name) WHERE LOWER(p.roles) LIKE '%tenderer%'";
    var r8 = queryBigQuery(sql8);
    Logger.log("Tenderers that match award_suppliers by name: " + (r8[0] ? r8[0].n : "?"));
  } catch(e) { Logger.log("H: " + e); }

  Logger.log("=== DONE ===");
}

// =============================================================================
//  PHASE 8 — TRUST & INTELLIGENCE (Premium)
// =============================================================================

// ── 1. BID PRICE CONFIDENCE INTERVAL ──────────────────────────────────────────

function getPriceConfidence(entity, method, estValue) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var v = Number(estValue) || 0;
    if (!v) return { low: 0, median: 0, high: 0, lowRatio: 0, medianRatio: 0, highRatio: 0 };
    var sql = "SELECT a.value_amount * 1.0 / NULLIF(m.tender_value_amount, 0) AS ratio FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON a.main_ocid = m.ocid WHERE LOWER(m.buyer_name) LIKE LOWER('%" + e + "%') AND m.tender_value_amount > 0 AND a.value_amount > 0 ORDER BY ratio LIMIT 1000";
    var rows = queryBigQuery(sql);
    if (rows.length >= 10) {
      var ratios = rows.map(function(r) { return parseFloat(r.ratio) || 0; }).filter(function(x) { return x > 0.1 && x < 2; }).sort(function(a, b) { return a - b; });
      if (ratios.length >= 10) {
        var p10 = ratios[Math.floor(ratios.length * 0.1)];
        var p50 = ratios[Math.floor(ratios.length * 0.5)];
        var p90 = ratios[Math.floor(ratios.length * 0.9)];
        return { low: Math.round(v * p10), median: Math.round(v * p50), high: Math.round(v * p90), lowRatio: Math.round(p10 * 100), medianRatio: Math.round(p50 * 100), highRatio: Math.round(p90 * 100) };
      }
    }
  } catch(e) { Logger.log("priceConfidence: " + e); }
  var v2 = Number(estValue) || 0;
  return { low: Math.round(v2 * 0.85), median: Math.round(v2 * 0.92), high: Math.round(v2 * 0.98), lowRatio: 85, medianRatio: 92, highRatio: 98 };
}

// ── 2. BUYER INTEGRITY SCORE ─────────────────────────────────────────────────

function getBuyerIntegrityScore(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var riskSql = "SELECT pct_amended, total_tenders, avg_competition FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_buyer_risk WHERE LOWER(entity) LIKE LOWER('%" + e + "%') LIMIT 1";
    var rows = queryBigQuery(riskSql);
    if (rows.length > 0) {
      var pct = parseFloat(rows[0].pct_amended) || 0;
      var cancellationRate = 0;
      try {
        var cancelSql = "SELECT COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(buyer_name) LIKE LOWER('%" + e + "%') AND LOWER(status) LIKE '%cancel%'";
        var cancelRows = queryBigQuery(cancelSql);
        if (cancelRows.length > 0) cancellationRate = Math.round((parseInt(cancelRows[0].n) / Math.max(1, parseInt(rows[0].total_tenders) || 1)) * 100);
      } catch(e2) {}
      var singleBidderPct = 0;
      try {
        var sbSql = "SELECT COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(buyer_name) LIKE LOWER('%" + e + "%') AND tender_numberOfTenderers = 1";
        var sbRows = queryBigQuery(sbSql);
        if (sbRows.length > 0) singleBidderPct = Math.round((parseInt(sbRows[0].n) / Math.max(1, parseInt(rows[0].total_tenders) || 1)) * 100);
      } catch(e2) {}
      var score = Math.max(0, Math.round(100 - pct * 1.2 - cancellationRate * 0.8 - singleBidderPct * 0.5));
      var label = score >= 80 ? "Excellent" : score >= 60 ? "Good" : score >= 40 ? "Fair" : score >= 20 ? "Poor" : "Critical";
      return { score: score, label: label, amendment_pct: pct, cancellation_pct: cancellationRate, single_bidder_pct: singleBidderPct, total_tenders: parseInt(rows[0].total_tenders) || 0 };
    }
  } catch(e) {}
  return { score: 50, label: "Unknown", amendment_pct: 0, cancellation_pct: 0, single_bidder_pct: 0, total_tenders: 0 };
}

// ── 3. COST OVERRUN PROBABILITY ───────────────────────────────────────────────

function getCostOverrun(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var sql = "SELECT COUNT(DISTINCT am.id) AS amendment_count, COUNT(DISTINCT c.id) AS contract_count FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts_amendments am JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts c ON am.main_ocid = c.main_ocid JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON c.main_ocid = m.ocid WHERE LOWER(m.buyer_name) LIKE LOWER('%" + e + "%')";
    var rows = queryBigQuery(sql);
    if (rows.length > 0) {
      var amCnt = parseInt(rows[0].amendment_count) || 0;
      var cCnt = parseInt(rows[0].contract_count) || 1;
      return { amendment_count: amCnt, contract_count: cCnt, avg_amendments_per_contract: parseFloat((amCnt / cCnt).toFixed(1)), label: (amCnt / cCnt) > 3 ? "High" : (amCnt / cCnt) > 1.5 ? "Medium" : "Low" };
    }
  } catch(e) { Logger.log("costOverrun: " + e); }
  return null;
}

// ── 4. TIMELINE RISK ─────────────────────────────────────────────────────────

function getTimelineRisk(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var sql = "SELECT AVG(c.period_durationInDays) AS avg_contract_days, COUNT(DISTINCT c.id) AS contract_count, COUNT(DISTINCT am.id) AS amendment_count FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts c LEFT JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts_amendments am ON c.main_ocid = am.main_ocid JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON c.main_ocid = m.ocid WHERE LOWER(m.buyer_name) LIKE LOWER('%" + e + "%') AND c.period_durationInDays > 0";
    var rows = queryBigQuery(sql);
    if (rows.length > 0) {
      var avgDays = parseFloat(rows[0].avg_contract_days) || 0;
      var amCnt = parseInt(rows[0].amendment_count) || 0;
      var cCnt = parseInt(rows[0].contract_count) || 1;
      var amPerContract = amCnt / cCnt;
      var riskPct = Math.round(Math.min(100, amPerContract * 30));
      return { avg_contract_days: Math.round(avgDays), amendment_count: amCnt, contract_count: cCnt, amendments_per_contract: parseFloat(amPerContract.toFixed(1)), risk_pct: riskPct, label: riskPct > 60 ? "High" : riskPct > 30 ? "Medium" : "Low" };
    }
  } catch(e) { Logger.log("timelineRisk: " + e); }
  return null;
}

// ── 5. SECTOR ENTRY DIFFICULTY ────────────────────────────────────────────────

function getSectorEntryDifficulty(sector) {
  try {
    if (!sector) return null;
    var s = String(sector).replace(/'/g, "''");
    var sql = "SELECT AVG(tender_numberOfTenderers) AS avg_bidders, COUNT(*) AS total_tenders FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(tender_mainProcurementCategory) LIKE LOWER('%" + s + "%') AND tender_numberOfTenderers > 0";
    var rows = queryBigQuery(sql);
    if (!rows.length) return null;
    var avgBidders = Math.round(parseFloat(rows[0].avg_bidders) || 0);
    var totalTenders = parseInt(rows[0].total_tenders) || 0;
    var incumbentWinPct = 0;
    try {
      var incSql = "SELECT COUNT(DISTINCT a.main_ocid) AS wins FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_award_suppliers s ON a.main_ocid = s.main_ocid AND a.id = s.awards_id JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON a.main_ocid = m.ocid WHERE LOWER(m.tender_mainProcurementCategory) LIKE LOWER('%" + s + "%') AND LOWER(s.name) IN (SELECT LOWER(s2.name) FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a2 JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_award_suppliers s2 ON a2.main_ocid = s2.main_ocid AND a2.id = s2.awards_id JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m2 ON a2.main_ocid = m2.ocid WHERE LOWER(m2.tender_mainProcurementCategory) LIKE LOWER('%" + s + "%') GROUP BY s2.name HAVING COUNT(DISTINCT a2.main_ocid) >= 2)";
      var incRows = queryBigQuery(incSql);
      if (incRows.length > 0) incumbentWinPct = Math.round((parseInt(incRows[0].wins) / Math.max(1, totalTenders)) * 100);
    } catch(e2) {}
    var barrier = avgBidders >= 8 ? "High" : avgBidders >= 5 ? "Medium" : "Low";
    return { sector: sector, avg_bidders: avgBidders, total_tenders: totalTenders, incumbent_win_pct: incumbentWinPct, new_entrant_win_pct: Math.max(0, 100 - incumbentWinPct), barrier: barrier };
  } catch(e) { Logger.log("sectorEntry: " + e); }
  return null;
}

// ── 6. CONTRACT COMPLETION RATE ───────────────────────────────────────────────

function getContractCompletionRate(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var sql = "SELECT c.status, COUNT(*) AS cnt FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts c JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON c.main_ocid = m.ocid WHERE LOWER(m.buyer_name) LIKE LOWER('%" + e + "%') GROUP BY c.status";
    var rows = queryBigQuery(sql);
    var result = { completed: 0, active: 0, cancelled: 0, terminated: 0, total: 0, statuses: {} };
    rows.forEach(function(r) { var s = String(r.status || "unknown").toLowerCase(); var n = parseInt(r.cnt) || 0; result.total += n; result.statuses[s] = n; if (s.indexOf("complete") !== -1 || s.indexOf("execut") !== -1 || s.indexOf("close") !== -1 || s.indexOf("finish") !== -1 || s.indexOf("final") !== -1 || s === "closed") result.completed += n; else if (s.indexOf("cancel") !== -1) result.cancelled += n; else if (s.indexOf("terminat") !== -1) result.terminated += n; else result.active += n; });
    result.completion_pct = result.total > 0 ? Math.round((result.completed / result.total) * 100) : 0;
    result.label = result.completion_pct >= 80 ? "Excellent" : result.completion_pct >= 60 ? "Good" : result.completion_pct >= 40 ? "Fair" : "Poor";
    return result;
  } catch(e) { Logger.log("contractCompletion: " + e); }
  return null;
}

// ── 7. AWARD-TO-CONTRACT VALUE GAP ────────────────────────────────────────────

function getAwardContractGap(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var sql = "SELECT AVG(c.value_amount * 1.0 / NULLIF(a.value_amount, 0)) AS avg_ratio, COUNT(DISTINCT c.id) AS cnt FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts c JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a ON c.main_ocid = a.main_ocid JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON c.main_ocid = m.ocid WHERE LOWER(m.buyer_name) LIKE LOWER('%" + e + "%') AND c.value_amount > 0 AND a.value_amount > 0";
    var rows = queryBigQuery(sql);
    if (rows.length > 0 && rows[0].avg_ratio) {
      var ratio = parseFloat(rows[0].avg_ratio) || 1;
      return { avg_ratio: parseFloat(ratio.toFixed(2)), gap_pct: Math.round((ratio - 1) * 100), count: parseInt(rows[0].cnt) || 0 };
    }
  } catch(e) { Logger.log("awardContractGap: " + e); }
  return null;
}

// ── 8. CONTRACT SIGNING LAG ──────────────────────────────────────────────────

function getContractSigningLag(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var sql = "SELECT AVG(DATE_DIFF(c.dateSigned, a.date, DAY)) AS avg_lag_days, COUNT(DISTINCT c.id) AS cnt FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts c JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a ON c.main_ocid = a.main_ocid JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON c.main_ocid = m.ocid WHERE LOWER(m.buyer_name) LIKE LOWER('%" + e + "%') AND c.dateSigned IS NOT NULL AND a.date IS NOT NULL";
    var rows = queryBigQuery(sql);
    if (rows.length > 0) {
      var lag = Math.round(parseFloat(rows[0].avg_lag_days) || 0);
      return { avg_lag_days: lag, label: lag > 60 ? "Slow" : lag > 30 ? "Moderate" : "Fast", count: parseInt(rows[0].cnt) || 0 };
    }
  } catch(e) { Logger.log("contractSigningLag: " + e); }
  return null;
}

// ── 9. AMENDMENT RATIONALE ANALYSIS ───────────────────────────────────────────

function getAmendmentRationale(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var sql = "SELECT am.rationale, am.description, am.date FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts_amendments am JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts c ON am.main_ocid = c.main_ocid JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON c.main_ocid = m.ocid WHERE LOWER(m.buyer_name) LIKE LOWER('%" + e + "%') LIMIT 100";
    var rows = queryBigQuery(sql);
    var breakdown = { cost_overrun: 0, scope_change: 0, time_extension: 0, other: 0, total: 0 };
    rows.forEach(function(r) { var rt = String(r.rationale || "").toLowerCase(); var desc = String(r.description || "").toLowerCase(); breakdown.total++; if (rt.indexOf("cost") !== -1 || rt.indexOf("price") !== -1 || rt.indexOf("amount") !== -1 || rt.indexOf("increase") !== -1 || desc.indexOf("cost") !== -1 || desc.indexOf("increas") !== -1 || desc.indexOf("addition") !== -1) breakdown.cost_overrun++; else if (rt.indexOf("scope") !== -1 || rt.indexOf("change") !== -1 || rt.indexOf("variation") !== -1 || desc.indexOf("scope") !== -1) breakdown.scope_change++; else if (rt.indexOf("time") !== -1 || rt.indexOf("extension") !== -1 || rt.indexOf("duration") !== -1 || rt.indexOf("period") !== -1 || desc.indexOf("extension") !== -1 || desc.indexOf("delay") !== -1) breakdown.time_extension++; else breakdown.other++; });
    return breakdown.total > 0 ? breakdown : null;
  } catch(e) { Logger.log("amendmentRationale: " + e); }
  return null;
}

// ── 10. AMENDMENT FREQUENCY TREND ─────────────────────────────────────────────

function getAmendmentTrend(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var sql = "SELECT EXTRACT(YEAR FROM am.date) AS yr, COUNT(*) AS cnt FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts_amendments am JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts c ON am.main_ocid = c.main_ocid JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON c.main_ocid = m.ocid WHERE LOWER(m.buyer_name) LIKE LOWER('%" + e + "%') AND am.date IS NOT NULL GROUP BY yr ORDER BY yr DESC LIMIT 10";
    return queryBigQuery(sql).map(function(r) { return { year: parseInt(r.yr) || 0, count: parseInt(r.cnt) || 0 }; });
  } catch(e) { return []; }
}

// ── 11. AMENDMENT VALUE RATIO ─────────────────────────────────────────────────

function getAmendmentValueRatio(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var sql = "SELECT COUNT(DISTINCT am.id) AS amendment_count, COUNT(DISTINCT c.id) AS contract_count FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts_amendments am JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts c ON am.main_ocid = c.main_ocid JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON c.main_ocid = m.ocid WHERE LOWER(m.buyer_name) LIKE LOWER('%" + e + "%')";
    var rows = queryBigQuery(sql);
    if (rows.length > 0) {
      var amCnt = parseInt(rows[0].amendment_count) || 0;
      var cCnt = parseInt(rows[0].contract_count) || 1;
      return { amendment_count: amCnt, contract_count: cCnt, amendments_per_contract: parseFloat((amCnt / Math.max(1, cCnt)).toFixed(1)), pct_amended: Math.round(Math.min(100, (amCnt / Math.max(1, cCnt)) * 10)), count: cCnt };
    }
  } catch(e) { Logger.log("amendmentValueRatio: " + e); }
  return null;
}

// ── 12. SINGLE-BIDDER RISK ───────────────────────────────────────────────────

function getSingleBidderRisk(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var sql = "SELECT COUNT(*) AS sb_count FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(buyer_name) LIKE LOWER('%" + e + "%') AND tender_numberOfTenderers = 1";
    var sbRows = queryBigQuery(sql);
    var totalSql = "SELECT COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(buyer_name) LIKE LOWER('%" + e + "%')";
    var totalRows = queryBigQuery(totalSql);
    if (sbRows.length > 0 && totalRows.length > 0) {
      var sb = parseInt(sbRows[0].sb_count) || 0;
      var tot = parseInt(totalRows[0].n) || 1;
      var pct = Math.round((sb / tot) * 100);
      return { single_bidder_count: sb, total_tenders: tot, pct: pct, label: pct > 40 ? "High Risk" : pct > 20 ? "Medium Risk" : "Low Risk" };
    }
  } catch(e) { Logger.log("singleBidderRisk: " + e); }
  return null;
}

// ── 13. SUPPLIER CREDIT SCORE ─────────────────────────────────────────────────

function getSupplierCreditScore(name) {
  try {
    if (!name || String(name).trim().length < 2) return null;
    var n = String(name).replace(/'/g, "''");
    var sql = "SELECT COUNT(*) AS contract_count, SUM(c.value_amount) AS total_value, AVG(c.period_durationInDays) AS avg_duration FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts c JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a ON c.main_ocid = a.main_ocid JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_award_suppliers s ON a.main_ocid = s.main_ocid AND a.id = s.awards_id WHERE LOWER(s.name) LIKE LOWER('%" + n + "%') AND c.value_amount > 0";
    var rows = queryBigQuery(sql);
    if (!rows.length || !parseInt(rows[0].contract_count)) return null;
    var contractCount = parseInt(rows[0].contract_count) || 0;
    var totalValue = Math.round(parseFloat(rows[0].total_value) || 0);
    var avgDuration = Math.round(parseFloat(rows[0].avg_duration) || 0);
    var amendmentCount = 0;
    try {
      var amSql = "SELECT COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts_amendments am JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_contracts c ON am.main_ocid = c.main_ocid JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a ON c.main_ocid = a.main_ocid JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_award_suppliers s ON a.main_ocid = s.main_ocid AND a.id = s.awards_id WHERE LOWER(s.name) LIKE LOWER('%" + n + "%')";
      var amRows = queryBigQuery(amSql);
      if (amRows.length > 0) amendmentCount = parseInt(amRows[0].n) || 0;
    } catch(e2) {}
    var reliability = Math.min(100, Math.max(10, Math.round(50 + contractCount * 5 - (amendmentCount / Math.max(1, contractCount)) * 30)));
    return { name: name, contract_count: contractCount, total_value: totalValue, avg_duration_days: avgDuration, amendment_count: amendmentCount, reliability_score: reliability, label: reliability >= 70 ? "Strong" : reliability >= 40 ? "Moderate" : "Limited" };
  } catch(e) { Logger.log("supplierCreditScore: " + e); }
  return null;
}

// ── 14. BEST MONTH TO BID ────────────────────────────────────────────────────

function getBestMonthToBid(entity) {
  try {
    var e = String(entity || "").replace(/'/g, "''");
    var sql = "SELECT EXTRACT(MONTH FROM date) AS month, COUNT(*) AS tender_count, AVG(tender_numberOfTenderers) AS avg_bidders FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(buyer_name) LIKE LOWER('%" + e + "%') AND date IS NOT NULL AND tender_numberOfTenderers > 0 GROUP BY month ORDER BY avg_bidders ASC LIMIT 3";
    return queryBigQuery(sql).map(function(r) { return { month: parseInt(r.month) || 0, count: parseInt(r.tender_count) || 0, avg_bidders: Math.round(parseFloat(r.avg_bidders) || 0) }; });
  } catch(e) { return []; }
}

// ── 15. PIPELINE OVERLAP WARNING ─────────────────────────────────────────────

function getPipelineWarning(userId) {
  try {
    if (!userId) userId = getCurrentUserId();
    var sSheet = getSheet().getSheetByName("SAVED_TENDERS");
    if (!sSheet || sSheet.getLastRow() < 2) return null;
    var sData = sSheet.getDataRange().getValues();
    var ocids = sData.slice(1).filter(function(r) { return String(r[0] || "").toLowerCase() === userId.toLowerCase(); }).map(function(r) { return String(r[1] || ""); });
    if (ocids.length < 2) return null;
    var tSheet = getSheet().getSheetByName("TENDERS_FLAT");
    if (!tSheet) return null;
    var tData = tSheet.getDataRange().getValues();
    var tHeaders = tData[0];
    var deadlines = [];
    for (var i = 1; i < tData.length; i++) {
      if (ocids.indexOf(String(tData[i][tHeaders.indexOf("ocid")] || "")) === -1) continue;
      var dl = tData[i][tHeaders.indexOf("deadline")];
      if (dl) deadlines.push({ ocid: String(tData[i][tHeaders.indexOf("ocid")] || ""), title: String(tData[i][tHeaders.indexOf("title")] || ""), deadline: new Date(dl) });
    }
    var week = 7 * 24 * 3600 * 1000;
    var clusters = [];
    for (var j = 0; j < deadlines.length; j++) {
      var cluster = [deadlines[j]];
      for (var k = j + 1; k < deadlines.length; k++) {
        if (Math.abs(deadlines[j].deadline - deadlines[k].deadline) <= week) cluster.push(deadlines[k]);
      }
      if (cluster.length >= 3) clusters.push({ count: cluster.length, weekOf: cluster[0].deadline.toISOString().substring(0,10), tenders: cluster });
    }
    return clusters.length > 0 ? clusters : null;
  } catch(e) { return null; }
}

// ── 16. HISTORICAL WIN RATE (SECTOR-LEVEL) ────────────────────────────────────

function getHistoricalWinRate(profile) {
  try {
    if (!profile) return null;
    var sector = String(profile.primary_sector || profile.sectors_of_interest || "").split(",")[0].trim().toLowerCase();
    if (!sector) return null;
    var sSafe = sector.replace(/'/g, "''");
    var sql = "SELECT COUNT(*) AS total_awards, AVG(a.value_amount) AS avg_award FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_award_suppliers s ON a.main_ocid = s.main_ocid AND a.id = s.awards_id JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON a.main_ocid = m.ocid WHERE LOWER(m.tender_mainProcurementCategory) LIKE LOWER('%" + sSafe + "%') AND a.value_amount > 0";
    var rows = queryBigQuery(sql);
    if (rows.length > 0) {
      var totalAwards = parseInt(rows[0].total_awards) || 0;
      var tenderSql = "SELECT COUNT(*) AS n FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main WHERE LOWER(tender_mainProcurementCategory) LIKE LOWER('%" + sSafe + "%')";
      var tRows = queryBigQuery(tenderSql);
      var totalTenders = tRows.length > 0 ? parseInt(tRows[0].n) || 1 : 1;
      return { sector: sector, awards: totalAwards, tenders: totalTenders, winRate: Math.round((totalAwards / totalTenders) * 100), avgAward: Math.round(parseFloat(rows[0].avg_award) || 0) };
    }
  } catch(e) {}
  return null;
}

// ── COMPREHENSIVE ENRICHMENT (Phase 5 + 8 combined) ──────────────────────────

function getFullTenderEnrichment(entity, method, estValue) {
  var result = getTenderPredictionEnrichment(entity, method, estValue);
  try { var price = getPriceConfidence(entity, method, estValue); if (price) result.price_ci = price; } catch(e) {}
  try { var integrity = getBuyerIntegrityScore(entity); if (integrity) result.buyer_integrity = integrity; } catch(e) {}
  try { var overrun = getCostOverrun(entity); if (overrun) result.cost_overrun = overrun; } catch(e) {}
  try { var timeline = getTimelineRisk(entity); if (timeline) result.timeline_risk = timeline; } catch(e) {}
  try { var completion = getContractCompletionRate(entity); if (completion) result.contract_completion = completion; } catch(e) {}
  try { var gap = getAwardContractGap(entity); if (gap) result.award_contract_gap = gap; } catch(e) {}
  try { var lag = getContractSigningLag(entity); if (lag) result.signing_lag = lag; } catch(e) {}
  try { var rationale = getAmendmentRationale(entity); if (rationale) result.amendment_rationale = rationale; } catch(e) {}
  try { var trend = getAmendmentTrend(entity); if (trend && trend.length) result.amendment_trend = trend; } catch(e) {}
  try { var amRatio = getAmendmentValueRatio(entity); if (amRatio) result.amendment_value_ratio = amRatio; } catch(e) {}
  try { var sbRisk = getSingleBidderRisk(entity); if (sbRisk) result.single_bidder_risk = sbRisk; } catch(e) {}
  try { var bestMonth = getBestMonthToBid(entity); if (bestMonth && bestMonth.length) result.best_months = bestMonth; } catch(e) {}
  return result;
}

// ── WEB-FACING WRAPPERS ──────────────────────────────────────────────────────

function getMyTierDebug(uid) {
  var uid = getCurrentUserId(uid);
  var sub = getSubscription(uid);
  var profile = getUserProfile(uid);
  var sector = profile ? (profile.primary_sector || profile.sectors_of_interest || "") : "";
  var entry = null;
  if (sector) { try { entry = getSectorEntryDifficulty(sector); } catch(e) {} }
  return { userId: uid, tier: sub.tier, status: sub.status, trust_intel: checkFeatureAccess(getCurrentUserId(uid), "trust_intel"), sector: sector, sectorEntry: entry };
}

function getTrustIntel(entity, method, estValue, uid) {
  try {
    if (!checkFeatureAccess(getCurrentUserId(uid), "trust_intel")) return { blocked: true, message: "Upgrade to Premium for Trust & Intelligence reports." };
    var result = {};
    try { result.price_ci = getPriceConfidence(entity, method, estValue); } catch(e) {}
    try { result.buyer_integrity = getBuyerIntegrityScore(entity); } catch(e) {}
    try { result.cost_overrun = getCostOverrun(entity); } catch(e) {}
    try { result.timeline_risk = getTimelineRisk(entity); } catch(e) {}
    try { result.contract_completion = getContractCompletionRate(entity); } catch(e) {}
    try { result.award_contract_gap = getAwardContractGap(entity); } catch(e) {}
    try { result.signing_lag = getContractSigningLag(entity); } catch(e) {}
    try { result.amendment_rationale = getAmendmentRationale(entity); } catch(e) {}
    try { result.amendment_trend = getAmendmentTrend(entity); } catch(e) {}
    try { result.amendment_value_ratio = getAmendmentValueRatio(entity); } catch(e) {}
    try { result.single_bidder_risk = getSingleBidderRisk(entity); } catch(e) {}
    try { result.best_months = getBestMonthToBid(entity); } catch(e) {}
    try {
      var wp = 0.1;
      try {
        var predSql = "SELECT win_probability FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_prediction_signals WHERE LOWER(entity) LIKE LOWER('%" + String(entity || "").replace(/'/g, "''") + "%') LIMIT 1";
        var predRows = queryBigQuery(predSql);
        if (predRows.length > 0) wp = parseFloat(predRows[0].win_probability) || 0.1;
      } catch(e2) {}
      var amPct = result.buyer_integrity ? result.buyer_integrity.amendment_pct : 0;
      var comp = 5;
      try { if (result.buyer_integrity && result.buyer_integrity.total_tenders > 0) comp = Math.round((parseInt(result.buyer_integrity.single_bidder_count) || 0) * 3 + 2); } catch(e2) {}
      result.bid_no_bid = computeBidNoBid({ win_probability: wp, buyer_amendment_pct: amPct, expected_bidders: comp });
    } catch(e) {}
    return result;
  } catch(e) { Logger.log("getTrustIntel: " + e); return { error: e.message }; }
}

function getSectorDifficulty(sector, uid) {
  try {
    if (!checkFeatureAccess(getCurrentUserId(uid), "trust_intel")) return { blocked: true };
    return getSectorEntryDifficulty(sector);
  } catch(e) { Logger.log("getSectorDifficulty: " + e); return { error: e.message }; }
}

function getCreditScore(name, uid) {
  try {
    if (!checkFeatureAccess(getCurrentUserId(uid), "trust_intel")) return { blocked: true };
    return getSupplierCreditScore(name);
  } catch(e) { Logger.log("getCreditScore: " + e); return { error: e.message }; }
}

function getMyPipelineWarnings(uid) {
  try {
    return getPipelineWarning(getCurrentUserId(uid));
  } catch(e) { Logger.log("getMyPipelineWarnings: " + e); return null; }
}

// ── PHASE 8 TESTS ────────────────────────────────────────────────────────────

function testPhase8Functions() {
  Logger.log("=== PHASE 8 DIAGNOSTICS ===");
  try { Logger.log("1. Price CI: " + JSON.stringify(getPriceConfidence("MININFRA", "open", 950000000))); } catch(e) { Logger.log("Price: " + e); }
  try { Logger.log("2. Integrity: " + JSON.stringify(getBuyerIntegrityScore("MININFRA"))); } catch(e) {}
  try { Logger.log("3. Cost Overrun: " + JSON.stringify(getCostOverrun("MININFRA"))); } catch(e) {}
  try { Logger.log("4. Timeline Risk: " + JSON.stringify(getTimelineRisk("MININFRA"))); } catch(e) {}
  try { Logger.log("5. Sector Entry: " + JSON.stringify(getSectorEntryDifficulty("works"))); } catch(e) {}
  try { Logger.log("6. Contract Completion: " + JSON.stringify(getContractCompletionRate("MININFRA"))); } catch(e) {}
  try { Logger.log("7. Award-Contract Gap: " + JSON.stringify(getAwardContractGap("MININFRA"))); } catch(e) {}
  try { Logger.log("8. Signing Lag: " + JSON.stringify(getContractSigningLag("MININFRA"))); } catch(e) {}
  try { Logger.log("9. Amendment Rationale: " + JSON.stringify(getAmendmentRationale("MININFRA"))); } catch(e) {}
  try { Logger.log("10. Amendment Trend: " + JSON.stringify(getAmendmentTrend("MININFRA"))); } catch(e) {}
  try { Logger.log("11. Amendment Ratio: " + JSON.stringify(getAmendmentValueRatio("MININFRA"))); } catch(e) {}
  try { Logger.log("12. Single Bidder: " + JSON.stringify(getSingleBidderRisk("MININFRA"))); } catch(e) {}
  try { Logger.log("13. Credit Score: " + JSON.stringify(getSupplierCreditScore("greenbuild"))); } catch(e) {}
  try { Logger.log("14. Best Month: " + JSON.stringify(getBestMonthToBid("MININFRA"))); } catch(e) {}
  try { Logger.log("15. Win Rate: " + JSON.stringify(getHistoricalWinRate({ primary_sector: "works" }))); } catch(e) {}
  try { Logger.log("16. Full Enrichment: " + JSON.stringify(getFullTenderEnrichment("MININFRA", "open", 950000000)).substring(0,300)); } catch(e) {}
  Logger.log("=== END ===");
}

function testPhase8Quick() {
  Logger.log("=== PHASE 8 QUICK ===");
  try { Logger.log("Trust Intel: " + JSON.stringify(getTrustIntel("MININFRA", "open", 950000000)).substring(0,400)); } catch(e) { Logger.log("Err: " + e); }
  Logger.log("=== END ===");
}

// =============================================================================
//  MOBILE APP REST API — doPost JSON endpoint (43 actions)
// =============================================================================

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = (body.action || "").trim();
    var params = body.params || {};
    var uid = getCurrentUserId(params.uid || "");
    var result;
    switch (action) {
      case "requestMagicLink": result = requestMagicLink(params.email || ""); break;
      case "verifyToken": result = verifyToken(params.token || ""); break;
      case "getDashboardData": result = getDashboardData(uid); break;
      case "saveBusinessDNA": result = saveBusinessDNA(params.data || {}, uid); break;
      case "getUserProfile": result = getUserProfile(uid); break;
      case "searchTenders": result = searchTenders(params.query || "", params.filters || {}, uid); break;
      case "getMatchesForWeb": result = getMatchesForWeb(uid); break;
      case "getSavedTendersForWeb": result = getSavedTendersForWeb(uid); break;
      case "getNewThisWeekForWeb": result = getNewThisWeekForWeb(uid); break;
      case "generateMatchesNoAI": result = generateMatchesNoAI(uid); break;
      case "toggleSave": result = toggleSave(params.ocid || "", uid); break;
      case "lookupTender": result = lookupTender(params.identifier || "", uid); break;
      case "getTrustIntel": result = getTrustIntel(params.entity || "", params.method || "", Number(params.estValue) || 0, uid); break;
      case "askAboutTender": result = askAboutTender(params.question || "", params.ocid || "", uid); break;
      case "analyzeTenderText": result = analyzeTenderText(params.text || "", uid); break;
      case "askTutor": result = askTutor(params.question || "", params.context || "", uid); break;
      case "isAdmin": result = isAdmin(uid); break;
      case "getAppUrl": result = getAppUrl(); break;
      case "computeSupplierScore": result = computeSupplierScore(uid); break;
      case "getBidPackTeaser": result = getBidPackTeaser(params.ocid || "", uid); break;
      case "createBidPackPayment": result = createBidPackPayment(params.ocid || "", uid, params.method || "stripe", params.phone || ""); break;
      case "verifyBidPackPayment": result = verifyBidPackPayment(params.refId || ""); break;
      case "getBidPack": result = getBidPack(params.ocid || "", uid); break;
      case "getMarketDashboardData": result = getMarketDashboardData(); break;
      case "getProcurementCalendar": result = getProcurementCalendar(); break;
      case "getBuyerPaymentRatings": result = getBuyerPaymentRatings(); break;
      case "getSectorHealth": result = getSectorHealth(params.sector || ""); break;
      case "getSupplierPortfolio": result = getSupplierPortfolio(params.name || ""); break;
      case "getCreditScore": result = getCreditScore(params.name || "", uid); break;
      case "getCompetitorIntel": result = getCompetitorIntel(params.name || ""); break;
      case "findJVPartners": result = findJVPartners(params.name || ""); break;
      case "getInsights": result = getInsights(uid); break;
      case "generateStrategyReport": result = generateStrategyReport(uid); break;
      case "getSubscriptionForWeb": result = getSubscriptionForWeb(uid); break;
      case "createStripeSession": result = createStripeSession(uid, params.tier || "pro", params.currency || "USD"); break;
      case "createMoMoPayment": result = createMoMoPayment(uid, params.phone || "", params.tier || "pro", params.currency || "RWF"); break;
      case "pollMoMoStatus": result = pollMoMoStatus(params.refId || ""); break;
      case "getBankDetails": result = getBankDetails(); break;
      case "submitBankProof": result = submitBankProof(params.tier || "pro", params.reference || "", params.fileUrl || "", params.currency || "RWF"); break;
      case "getAdminDashboard": result = getAdminDashboard(); break;
      case "approveBankPayment": result = approveBankPayment(params.userId || "", params.tier || "pro"); break;
      case "rejectBankPayment": result = rejectBankPayment(params.userId || "", params.reason || ""); break;
      case "getUmucyoSummary": result = getUmucyoSummary(params.year || ""); break;
      case "syncTendersFromLiveAPI": result = syncTendersFromLiveAPI(params.dateFrom || "", params.limit || 50); break;
      case "lookupTenderOCDS": result = lookupTenderOCDS(params.ocid || ""); break;
      default: return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Unknown action: " + action })).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =============================================================================
//  MOBILE APP REST API
function ensureUserTokensSheet() {
  var sheets = getSheet().getSheets().map(function(s) { return s.getName(); });
  if (sheets.indexOf("USER_TOKENS") === -1) {
    getSheet().insertSheet("USER_TOKENS").appendRow(["email", "token", "created_at", "used", "used_at"]);
  }
}

function generateToken() {
  var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  var t = ""; for (var i = 0; i < 48; i++) t += chars.charAt(Math.floor(Math.random() * chars.length));
  return t;
}

function requestMagicLink(email) {
  ensureUserTokensSheet();
  var e = String(email || "").trim().toLowerCase();
  if (!e || e.indexOf("@") === -1) return { success: false, error: "Valid email required" };
  var token = generateToken();
  getSheet().getSheetByName("USER_TOKENS").appendRow([e, token, new Date().toISOString(), false, ""]);
  return { success: true, message: "Sign-in ready", token: token };
}

function verifyToken(token) {
  ensureUserTokensSheet();
  var t = String(token || "").trim();
  if (!t) return { error: "No token provided" };
  var sheet = getSheet().getSheetByName("USER_TOKENS");
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var tc = h.indexOf("token"), ec = h.indexOf("email"), cc = h.indexOf("created_at"), uc = h.indexOf("used"), ua = h.indexOf("used_at");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][tc] || "") === t) {
      if (String(data[i][uc] || "").toLowerCase() === "true") return { error: "Token already used" };
      if ((new Date() - new Date(data[i][cc] || 0)) > 15 * 60 * 1000) return { error: "Token expired" };
      var emailAddr = String(data[i][ec] || "").trim().toLowerCase();
      sheet.getRange(i + 1, uc + 1).setValue(true);
      sheet.getRange(i + 1, ua + 1).setValue(new Date().toISOString());
      var profile = getUserProfile(emailAddr);
      return { success: true, user: { email: emailAddr, companyName: (profile ? (profile.company_name || "") : "") }, token: emailAddr };
    }
  }
  return { error: "Invalid token" };
}

// =============================================================================
//  PRE-QUALIFICATION SCORE
// =============================================================================

function computeSupplierScore(uid) {
  var userId = getCurrentUserId(uid);
  var profile = getUserProfile(userId);
  if (!profile) return { score: 0, label: "No profile", breakdown: {}, recommendations: ["Complete your Business DNA profile first."] };

  var breakdown = {};
  var score = 0;
  var recommendations = [];

  // 1. Profile basics (0-10)
  var name = String(profile.company_name || "").trim();
  var tin = String(profile.tin || "").trim();
  var basics = 0;
  if (name) basics += 5;
  if (tin) basics += 5;
  breakdown.profile_basics = basics;
  score += basics;

  // 2. Certifications (0-40)
  var certs = String(profile.certifications || "").toLowerCase();
  var certScore = 0;
  if (certs.indexOf("tax clearance") >= 0) certScore += 10;
  if (certs.indexOf("rssb") >= 0) certScore += 10;
  if (certs.indexOf("rdb") >= 0) certScore += 10;
  if (certs.indexOf("iso") >= 0) certScore += 10;
  breakdown.certifications = certScore;
  score += certScore;
  if (certScore < 20) recommendations.push("Add more certifications (Tax Clearance, RSSB, RDB, ISO) to boost your score.");

  // 3. Past contracts count (0-20, 5 per contract, max 20)
  var govCount = parseInt(profile.past_gov_count) || 0;
  var expScore = Math.min(20, govCount * 5);
  breakdown.past_contracts = expScore;
  score += expScore;
  if (govCount === 0) recommendations.push("Winning your first government contract will significantly increase your score.");

  // 4. Past contract value (0-10, log scale)
  var govValue = parseInt(profile.past_gov_total_value) || 0;
  var valueScore = 0;
  if (govValue >= 1000000000) valueScore = 10;
  else if (govValue >= 500000000) valueScore = 8;
  else if (govValue >= 100000000) valueScore = 6;
  else if (govValue >= 50000000) valueScore = 4;
  else if (govValue > 0) valueScore = 2;
  breakdown.past_value = valueScore;
  score += valueScore;

  // 5. RDB registration (0-10)
  var rdb = String(profile.rdb_registered || "").toLowerCase();
  var rdbScore = rdb === "yes" ? 10 : rdb === "in_progress" ? 5 : 0;
  breakdown.rdb_registration = rdbScore;
  score += rdbScore;
  if (rdbScore < 10) recommendations.push("Complete your RDB registration for full procurement eligibility.");

  // 6. Company age (0-5)
  var yearEst = parseInt(profile.year_established) || 0;
  var nowYear = new Date().getFullYear();
  var age = yearEst > 0 ? nowYear - yearEst : 0;
  var ageScore = age >= 10 ? 5 : age >= 5 ? 3 : age > 0 ? 2 : 0;
  breakdown.company_age = ageScore;
  score += ageScore;
  if (ageScore < 2) recommendations.push("Older companies tend to have higher win rates. Consider highlighting team experience instead.");

  // 7. Profile completeness (0-15)
  var keyFields = ["company_name","tin","year_established","company_size","is_women_led","is_youth_led","rdb_registered","primary_sector","capabilities","certifications","preferred_methods","regions","max_team_size","equipment","cash_flow_limit"];
  var filled = keyFields.filter(function(f) { return Boolean(String(profile[f] || "").trim()); }).length;
  var completeScore = Math.round((filled / keyFields.length) * 15);
  breakdown.profile_completeness = completeScore;
  score += completeScore;
  if (completeScore < 10) recommendations.push("Complete more fields in your Business DNA profile to improve match quality.");

  var label = score >= 70 ? "Strong" : score >= 40 ? "Moderate" : "Developing";
  return { score: score, label: label, breakdown: breakdown, recommendations: recommendations };
}

// =============================================================================
//  BID PACK SYSTEM
// =============================================================================

var BID_PACK_PRICE = { usd: 15, rwf: 19000 };

function ensureBidPacksSheet() {
  var sheets = getSheet().getSheets().map(function(s) { return s.getName(); });
  if (sheets.indexOf("BID_PACKS") === -1) {
    var sheet = getSheet().insertSheet("BID_PACKS");
    sheet.appendRow(["user_id","ocid","tender_title","tender_entity","amount","currency","payment_method","payment_ref","status","created_at","pack_json","paid_at"]);
    sheet.setFrozenRows(1);
  }
}

function getBidPackTeaser(ocid, uid) {
  var userId = getCurrentUserId(uid);
  var lookup = lookupTender(ocid, userId);
  if (!lookup || !lookup.found) return { error: "Tender not found" };

  var t = lookup.tender;
  var predictions = lookup.predictions || {};
  var isJIR = (String(t.source || "").toLowerCase() === "jobinrwanda");

  var teaser = {
    tender_title: t.title,
    tender_entity: t.entity,
    tender_value: t.est_value,
    tender_method: t.method,
    is_jir: isJIR,
    bid_no_bid: null,
    win_probability: predictions.win_probability || null,
    buyer_amendment_pct: predictions.buyer_amendment_pct || null,
    price: BID_PACK_PRICE,
    already_purchased: false
  };

  // Check if already purchased
  ensureBidPacksSheet();
  var sheet = getSheet().getSheetByName("BID_PACKS");
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var uCol = h.indexOf("user_id"), oCol = h.indexOf("ocid"), sCol = h.indexOf("status");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uCol] || "").toLowerCase() === userId.toLowerCase() && String(data[i][oCol] || "") === ocid) {
      teaser.already_purchased = true;
      if (String(data[i][sCol] || "") === "generated") teaser.pack = tryParseJSON(data[i][h.indexOf("pack_json")]);
      break;
    }
  }

  // Compute bid/no-bid
  if (isJIR) {
    teaser.bid_no_bid = { recommendation: "MAYBE", confidence: 50, reasons: "JobInRwanda tender — limited historical data. AI analysis can still help with compliance and strategy." };
    teaser.win_probability = null;
    teaser.buyer_amendment_pct = null;
  } else {
    try {
      var enrich = getTrustIntel(t.entity, t.method, t.est_value, userId);
      if (enrich && !enrich.blocked && enrich.bid_no_bid) {
        teaser.bid_no_bid = enrich.bid_no_bid;
      }
    } catch(e) {}
  }

  return teaser;
}

function createBidPackPayment(ocid, uid, method, phone) {
  var userId = getCurrentUserId(uid);
  var lookup = lookupTender(ocid, userId);
  if (!lookup || !lookup.found) return { success: false, error: "Tender not found" };

  ensureBidPacksSheet();
  var sheet = getSheet().getSheetByName("BID_PACKS");
  var now = new Date();

  var t = lookup.tender;
  var amount = method === "stripe" ? BID_PACK_PRICE.usd : BID_PACK_PRICE.rwf;
  var currency = method === "stripe" ? "USD" : "RWF";

  switch (method) {
    case "stripe":
      return createStripeSessionForBidPack(userId, ocid, t.title, amount, currency, sheet, now);
    case "bank":
      return { success: true, message: "Bank transfer details", bank_details: getBankDetails(), instructions: "Submit proof after transfer to complete payment." };
    default:
      return { success: false, error: "Unknown payment method: " + method };
  }
}

// MoMo/Airtel BidPack functions removed — replaced by Flutterwave

function createStripeSessionForBidPack(userId, ocid, title, amount, currency, sheet, now) {
  try {
    var key = PropertiesService.getScriptProperties().getProperty("STRIPE_SECRET_KEY");
    if (!key) return { success: false, error: "Stripe not configured yet" };
    var price = amount * 100;
    var appUrl = getAppUrl();
    var sessionRes = UrlFetchApp.fetch("https://api.stripe.com/v1/checkout/sessions", {
      method: "post", headers: { "Authorization": "Bearer " + key, "Content-Type": "application/x-www-form-urlencoded" },
      payload: "payment_method_types[]=card&mode=payment&line_items[0][price_data][currency]=" + currency.toLowerCase() + "&line_items[0][price_data][product_data][name]=Bid Pack: " + encodeURIComponent(title.substring(0, 50)) + "&line_items[0][price_data][unit_amount]=" + price + "&line_items[0][quantity]=1&success_url=" + encodeURIComponent(appUrl + "?page=check&lookup=" + ocid) + "&cancel_url=" + encodeURIComponent(appUrl + "?page=check&lookup=" + ocid),
      muteHttpExceptions: true
    });
    var session = JSON.parse(sessionRes.getContentText());
    if (session.url) {
      sheet.appendRow([userId, ocid, title, "", amount, currency, "stripe", session.id, "pending", now.toISOString(), "", ""]);
      return { success: true, url: session.url, method: "stripe", ref_id: session.id };
    }
    return { success: false, error: "Could not create payment session" };
  } catch(e) { return { success: false, error: "Stripe error: " + e }; }
}

function verifyBidPackPayment(refId) {
  ensureBidPacksSheet();
  var sheet = getSheet().getSheetByName("BID_PACKS");
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var rCol = h.indexOf("payment_ref"), mCol = h.indexOf("payment_method"), sCol = h.indexOf("status"), oCol = h.indexOf("ocid"), uCol = h.indexOf("user_id");

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][rCol] || "") !== refId) continue;
    if (String(data[i][sCol] || "") === "generated") return { success: true, pack: tryParseJSON(data[i][h.indexOf("pack_json")]) };

    var method = String(data[i][mCol] || "");
    if (method === "stripe") {
      try {
        var ss = verifyStripeSession(refId);
        if (ss && ss.paid) {
          sheet.getRange(i + 1, sCol + 1).setValue("paid");
          sheet.getRange(i + 1, h.indexOf("paid_at") + 1).setValue(new Date().toISOString());
          var pack = generateBidPack(data[i][oCol], data[i][uCol]);
          sheet.getRange(i + 1, h.indexOf("pack_json") + 1).setValue(JSON.stringify(pack));
          sheet.getRange(i + 1, sCol + 1).setValue("generated");
          return { success: true, pack: pack };
        }
        return { success: false, message: "Payment not confirmed yet." };
      } catch(e) { return { success: false, error: "Could not verify: " + e }; }
    }
  }
  return { success: false, error: "Payment reference not found" };
}

function verifyBidPackBankPayment(userId, ocid) {
  ensureBidPacksSheet();
  var sheet = getSheet().getSheetByName("BID_PACKS");
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][h.indexOf("user_id")] || "").toLowerCase() === userId.toLowerCase() && String(data[i][h.indexOf("ocid")] || "") === ocid && String(data[i][h.indexOf("payment_method")] || "") === "bank" && String(data[i][h.indexOf("status")] || "") === "pending") {
      sheet.getRange(i + 1, h.indexOf("status") + 1).setValue("paid");
      sheet.getRange(i + 1, h.indexOf("paid_at") + 1).setValue(new Date().toISOString());
      var pack = generateBidPack(ocid, userId);
      sheet.getRange(i + 1, h.indexOf("pack_json") + 1).setValue(JSON.stringify(pack));
      sheet.getRange(i + 1, h.indexOf("status") + 1).setValue("generated");
      return { success: true, pack: pack };
    }
  }
  return { success: false, error: "Pending bank payment not found" };
}

function generateBidPack(ocid, uid) {
  var userId = getCurrentUserId(uid);
  var lookup = lookupTender(ocid, userId);
  if (!lookup || !lookup.found) return { error: "Tender not found" };

  var t = lookup.tender;
  var entity = t.entity, method = t.method, estValue = t.est_value;
  var isJIR = (String(t.source || "").toLowerCase() === "jobinrwanda");

  var pack = {
    tender_title: t.title,
    tender_entity: entity,
    tender_value: estValue,
    tender_method: method,
    tender_ocid: ocid,
    source: isJIR ? "jobinrwanda" : "rppa",
    generated_at: new Date().toISOString(),
    disclaimer: "This analysis is based on 14 years of historical RPPA procurement data. Past patterns do not guarantee future outcomes. This is not a guarantee of winning any tender. Prices, competition, and buyer behavior may vary. Always verify tender details on official sources before bidding. This report is for informational purposes only and does not constitute financial or legal advice." + (isJIR ? " Note: This tender is from JobInRwanda — limited historical data available. Analysis is AI-generated only." : "")
  };

  if (isJIR) {
    // JIR tenders: skip BigQuery, use AI-only analysis
    pack.bid_no_bid = { recommendation: "MAYBE", confidence: 50, reasons: "This tender is from JobInRwanda — limited historical data available for automated analysis. Review the tender manually and use the compliance checklist below." };
    pack.competitors = [];
    pack.best_months = [];
  } else {
    // RPPA tenders: full BigQuery enrichment
    var trust = getTrustIntel(entity, method, estValue, userId);
    if (trust && !trust.blocked) {
      pack.bid_no_bid = trust.bid_no_bid;
      pack.price_ci = trust.price_ci;
      pack.buyer_integrity = trust.buyer_integrity;
      pack.timeline_risk = trust.timeline_risk;
      pack.cost_overrun = trust.cost_overrun;
      pack.amendment_value_ratio = trust.amendment_value_ratio;
      pack.single_bidder_risk = trust.single_bidder_risk;
      pack.contract_completion = trust.contract_completion;
      pack.signing_lag = trust.signing_lag;
      pack.amendment_rationale = trust.amendment_rationale;
    } else {
      try { pack.bid_no_bid = computeBidNoBid({ win_probability: 0.1, buyer_amendment_pct: 0, expected_bidders: 5 }); } catch(e) {}
    }

    pack.competitors = [];
    try {
      var compSql = "SELECT s.name, COUNT(*) AS wins FROM " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_awards a JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_award_suppliers s ON a.main_ocid = s.main_ocid AND a.id = s.awards_id JOIN " + BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET + ".rppc_main m ON a.main_ocid = m.ocid WHERE LOWER(m.buyer_name) LIKE LOWER('%" + String(entity || "").replace(/'/g, "''") + "%') GROUP BY s.name ORDER BY wins DESC LIMIT 5";
      var compRows = queryBigQuery(compSql);
      (compRows || []).forEach(function(r) { pack.competitors.push({ name: r.name, wins: parseInt(r.wins) || 0 }); });
    } catch(e) {}

    try { pack.best_months = getBestMonthToBid(entity); } catch(e) { pack.best_months = []; }
  }

  // 4. Compliance checklist (AI) — works for both RPPA and JIR
  try {
    var desc = String(t.description || t.title || "").substring(0, 500);
    var sourceNote = isJIR ? " This tender is from JobInRwanda (not on Umucyo). " : "";
    var compliancePrompt = "You are analyzing a Rwandan tender" + (isJIR ? " (from JobInRwanda)" : " (on Umucyo RPPA)") + ": \"" + t.title + "\" by \"" + entity + "\". Based on the title and type, generate a compliance checklist for a bidder. List 5-10 items a bidder should prepare (documents, certifications, technical requirements). Start each line with a dash. Keep it practical and relevant to Rwandan procurement. Do NOT fabricate specific tender requirements — base it on common procurement practices." + sourceNote;
    var aiRes = callOpenRouter(compliancePrompt, "You are a procurement compliance expert familiar with Rwandan RPPA requirements.", userId);
    pack.compliance_checklist = aiRes || "Generate compliance checklist based on tender type.";
  } catch(e) { pack.compliance_checklist = "Unable to generate checklist at this time."; }

  // 5. Recommended strategy (AI) — works for both
  try {
    var bidRec = pack.bid_no_bid ? pack.bid_no_bid.recommendation : "Unknown";
    var winPct = pack.bid_no_bid ? pack.bid_no_bid.confidence : "N/A";
    var jiNote = isJIR ? " IMPORTANT: This tender is from JobInRwanda, not Umucyo. Historical RPPA data is not available. Focus on general bidding best practices." : "";
    var strategyPrompt = "Based on the following tender analysis, recommend a bidding strategy. Tender: \"" + t.title + "\" by \"" + entity + "\", value: " + (estValue || "unknown") + " RWF. Bid/No-Bid says: " + bidRec + " with " + winPct + "% confidence. Provide 3-4 actionable recommendations. Keep each under 2 sentences. Use the data available — do NOT fabricate numbers." + jiNote;
    var strategyRes = callOpenRouter(strategyPrompt, "You are a procurement strategy advisor for businesses bidding on Rwandan government tenders.", userId);
    pack.strategy_recommendations = strategyRes || "Prepare thoroughly: review the tender documents, ensure all certifications are current, and price competitively.";
  } catch(e) { pack.strategy_recommendations = "Prepare thoroughly for this tender."; }

  return pack;
}

function getBidPack(ocid, uid) {
  var userId = getCurrentUserId(uid);
  ensureBidPacksSheet();
  var sheet = getSheet().getSheetByName("BID_PACKS");
  var data = sheet.getDataRange().getValues();
  var h = data[0];
  var oCol = h.indexOf("ocid"), uCol = h.indexOf("user_id"), sCol = h.indexOf("status"), pCol = h.indexOf("pack_json");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][uCol] || "").toLowerCase() === userId.toLowerCase() && String(data[i][oCol] || "") === ocid && String(data[i][sCol] || "") === "generated") {
      return tryParseJSON(data[i][pCol]);
    }
  }
  return null;
}

function tryParseJSON(str) {
  try { return JSON.parse(str); } catch(e) { return null; }
}

// =============================================================================
//  UMUCYO OCDS LIVE API INTEGRATION — Feature-flagged
// =============================================================================
// ScriptProperty: UMUCYO_API_ENABLED = "true" to activate
// Base: https://ocds.umucyo.gov.rw/opendata (production server)

var UMUCYO_API_BASE = "https://ocds.umucyo.gov.rw/opendata";

function umucyoEnabled() {
  return PropertiesService.getScriptProperties().getProperty("UMUCYO_API_ENABLED") === "true";
}

// ── WORKING: Fetch annual procurement summary ──────────────────────────────

function getUmucyoSummary(year) {
  if (!umucyoEnabled()) return { error: "Umucyo API not enabled yet. Set UMUCYO_API_ENABLED=true in Script Properties." };
  var y = parseInt(year) || new Date().getFullYear();
  try {
    var res = UrlFetchApp.fetch(UMUCYO_API_BASE + "/api/v1/ui/summaries/" + y, { muteHttpExceptions: true, method: "get", headers: { "Accept": "application/json" } });
    if (res.getResponseCode() !== 200) return { error: "API returned " + res.getResponseCode() };
    var data = JSON.parse(res.getContentText());
    return {
      year: String(y) + "/" + String(y + 1),
      awarded_value: data.awarded_procurement_value || 0,
      published_entities: data.published_procurement_entities || 0,
      registered_suppliers: data.registered_suppliers || 0,
      awarded_suppliers: data.awarded_suppliers || 0,
      published_tenders: data.published_tenders || 0,
      procurement_plans: data.procurement_plans || 0,
      awards: data.awards || 0,
      source: "RPPA Official"
    };
  } catch(e) { return { error: String(e) }; }
}

// ── WORKING: Download monthly dataset JSON ────────────────────────────────

function downloadMonthlyDataset(year, month) {
  if (!umucyoEnabled()) return { error: "Umucyo API not enabled yet." };
  var y = String(year || "2025");
  var m = String(month || "01");
  var lookup = y + "/" + m + "-";
  try {
    var listRes = UrlFetchApp.fetch(UMUCYO_API_BASE + "/api/v1/ui/data_set/available_datasets", { muteHttpExceptions: true });
    if (listRes.getResponseCode() !== 200) return { error: "Dataset list failed: " + listRes.getResponseCode() };
    var listData = JSON.parse(listRes.getContentText());
    var datasets = listData.datasets || {};
    var files = datasets[y] || [];
    var jsonFile = null;
    for (var i = 0; i < files.length; i++) {
      if (files[i].indexOf(lookup) !== -1 && files[i].indexOf("-json") !== -1) { jsonFile = files[i]; break; }
    }
    if (!jsonFile) return { error: "No JSON dataset found for " + y + "/" + m + ". Available: " + (files || []).join(", ") };

    var dlRes = UrlFetchApp.fetch(UMUCYO_API_BASE + "/api/v1/ui/data_set/download?year=" + y + "&month_file=" + encodeURIComponent(jsonFile), { muteHttpExceptions: true });
    if (dlRes.getResponseCode() !== 200) return { error: "Download failed: " + dlRes.getResponseCode() };

    var blob = dlRes.getBlob();
    var unzipped = Utilities.unzip(blob);
    var releases = [];
    unzipped.forEach(function(zipEntry) {
      if (zipEntry.getName().indexOf(".json") !== -1) {
        try {
          var content = zipEntry.getDataAsString();
          var pkg = JSON.parse(content);
          if (pkg.releases) releases = releases.concat(pkg.releases);
        } catch(e2) {}
      }
    });
    return { success: true, file: jsonFile, release_count: releases.length, releases: releases };
  } catch(e) { return { error: String(e) }; }
}

// ── NOT YET WORKING: Live releases endpoint (disabled by feature flag) ─────

function syncTendersFromLiveAPI(dateFrom, limit) {
  if (!umucyoEnabled()) return { error: "Umucyo API not enabled.", skipped: true };
  var df = dateFrom || new Date(Date.now() - 7 * 24 * 3600 * 1000).toISOString().substring(0, 10);
  var lim = Math.min(limit || 50, 300);
  try {
    var res = UrlFetchApp.fetch(
      UMUCYO_API_BASE + "/api/v1/releases/all?date_from=" + df + "&limit=" + lim + "&sort_direction=desc",
      { muteHttpExceptions: true, method: "get", headers: { "Accept": "application/json" } }
    );
    if (res.getResponseCode() !== 200) return { error: "Releases API returned " + res.getResponseCode() + ". The endpoint may not be production-ready yet.", skipped: true };
    var pkg = JSON.parse(res.getContentText());
    var releases = pkg.releases || [];
    var upserted = 0;
    var tSheet = getSheet().getSheetByName("TENDERS_FLAT");
    if (!tSheet) { tSheet = getSheet().insertSheet("TENDERS_FLAT"); setupTendersFlatHeaders(tSheet); }
    var headers = tSheet.getDataRange().getValues()[0];
    var existingOcids = {};
    if (tSheet.getLastRow() > 1) {
      var existingData = tSheet.getDataRange().getValues();
      for (var i = 1; i < existingData.length; i++) existingOcids[String(existingData[i][headers.indexOf("ocid")] || "").toLowerCase()] = true;
    }
    for (var j = 0; j < releases.length; j++) {
      var r = releases[j];
      var ocid = String(r.ocid || "").toLowerCase();
      if (existingOcids[ocid]) continue;
      var t = r.tender || {};
      var row = [];
      row[headers.indexOf("ocid")] = r.ocid || "";
      row[headers.indexOf("title")] = t.title || "";
      row[headers.indexOf("entity")] = (t.procuringEntity || {}).name || (r.buyer || {}).name || "";
      row[headers.indexOf("est_value")] = (t.value || {}).amount || 0;
      row[headers.indexOf("method")] = t.procurementMethod || "";
      row[headers.indexOf("deadline")] = t.tenderPeriod ? (t.tenderPeriod.endDate || "") : "";
      row[headers.indexOf("num_tenderers")] = t.numberOfTenderers || 0;
      row[headers.indexOf("status")] = t.status || "";
      row[headers.indexOf("publish_date")] = r.date || "";
      row[headers.indexOf("description")] = t.description || "";
      row[headers.indexOf("source")] = "umucyo_live";
      tSheet.appendRow(row);
      existingOcids[ocid] = true;
      upserted++;
    }
    return { success: true, total: releases.length, upserted: upserted, skipped: releases.length - upserted };
  } catch(e) { return { error: String(e), skipped: true }; }
}

// ── NOT YET WORKING: Live tender lookup by OCID ──────────────────────────

function lookupTenderOCDS(ocid) {
  if (!umucyoEnabled()) return { error: "Umucyo API not enabled." };
  try {
    var res = UrlFetchApp.fetch(UMUCYO_API_BASE + "/api/v1/releases/?ocid=" + encodeURIComponent(ocid), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { error: "Lookup returned " + res.getResponseCode() + ". Endpoint not production-ready yet." };
    var release = JSON.parse(res.getContentText());
    return { success: true, release: release };
  } catch(e) { return { error: String(e) }; }
}
