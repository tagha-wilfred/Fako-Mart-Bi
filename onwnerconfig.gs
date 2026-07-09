// ============================================================
//  ownerConfig.gs  —  Woldenga Property Services
//  Data Layer  |  Owner Report Automation
//
//  SEPARATION OF CONCERNS
//  ┌──────────────────────────────────────────────────────┐
//  │  ownerConfig.gs   →  data only (this file)          │
//  │  ownerReport.html →  presentation only              │
//  └──────────────────────────────────────────────────────┘
//
//  SOURCE SHEETS (all in the same Google Spreadsheet)
//  ┌─────────────────────┬────────────────────────────────┐
//  │ Sheet               │ Purpose                        │
//  ├─────────────────────┼────────────────────────────────┤
//  │ Final DataSet       │ Appointments (one row = 1 job) │
//  │ Quality Complaints  │ Complaint events               │
//  │ Issues              │ Property issue events          │
//  │ Inventory           │ Inventory request events       │
//  │ Property Details    │ Property master list + Status  │
//  └─────────────────────┴────────────────────────────────┘
//
//  ENTRY POINT
//    generateOwnerReport(ownerName, month, year)
//    e.g. generateOwnerReport("Selah Stays", "January", 2026)
//
//  One PDF is produced per unique Company Name (Owner).
//  To generate for all owners in one run → generateAllOwnerReports()
// ============================================================


// ── GLOBAL CONFIG ────────────────────────────────────────────
var CONFIG_SHEET_NAME = "DASHBOARD";
var CONFIG_CELLS = {
  REPORT_YEAR  : { row: 4, col: 2 },
  REPORT_MONTH : { row: 4, col: 3 },
  OWNER_NAME   : { row: 4, col: 4 },
  STAFF_NAME   : { row: 4, col: 5 }
};

var TM_MARKET_IMAGE_ID = "1Q1OlvU39hHghTl6lkRocCTHdaJxHkPNI";
var SD_MARKET_IMAGE_ID = "1ePyWbFTD4oB7Qba2fMxHi1JG_fpQ8P98";
var LOGO_ID = "1h8dpv1IHjpSbQ3eqgg54DIjXAlopy7EQ";
var FOOTER_ID = "1_2zDbzzBQ3z5HIoQl3dDTVmUWNLUELq0"
var WATERMARK_ID = "1E2kkcOVe2DU_1e9Y_CWrNBPF2xeXIf1y"
var logo         = "https://drive.google.com/thumbnail?id=" + LOGO_ID + "&sz=s1000";
var footer   = "https://drive.google.com/thumbnail?id=" + FOOTER_ID + "&sz=s1000";

// Watermark reuses the same logo image — opacity/greyscale applied in CSS only.
var watermarkUrl = "https://drive.google.com/thumbnail?id=" + WATERMARK_ID + "&sz=s1000";;

/**
 * readConfig()
 * Reads the three inputs from the Configure Report sheet and returns
 * { ownerName, months, year }.
 */
function readConfig() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!configSheet) {
    Logger.log("⚠️  Sheet '" + CONFIG_SHEET_NAME + "' not found — reading from active sheet instead.");
    configSheet = ss.getActiveSheet();
  }

  var year      = configSheet.getRange(CONFIG_CELLS.REPORT_YEAR.row,  CONFIG_CELLS.REPORT_YEAR.col).getValue();
  var month     = configSheet.getRange(CONFIG_CELLS.REPORT_MONTH.row, CONFIG_CELLS.REPORT_MONTH.col).getValue();
  var ownerName = configSheet.getRange(CONFIG_CELLS.OWNER_NAME.row,   CONFIG_CELLS.OWNER_NAME.col).getValue();

  year      = parseInt(year);
  ownerName = String(ownerName).trim();

  var monthRaw    = String(month).trim();
  var validMonths = ["January","February","March","April","May","June",
                     "July","August","September","October","November","December"];
  var months = monthRaw.split(",").map(function(m) { return m.trim(); }).filter(Boolean);

  Logger.log("── readConfig() ──────────────────────────────────");
  Logger.log("  Sheet     : " + configSheet.getName());
  Logger.log("  Year raw  : " + year);
  Logger.log("  Month raw : " + monthRaw + "  -> parsed: " + JSON.stringify(months));
  Logger.log("  Owner raw : " + ownerName);
  Logger.log("──────────────────────────────────────────────────");

  if (isNaN(year) || year < 2000 || year > 2100) {
    throw new Error("Config error: Year cell (B4) contains '" + year + "' — expected a 4-digit number like 2026.");
  }
  if (months.length === 0) {
    throw new Error("Config error: Month cell (C4) is empty.");
  }
  months.forEach(function(m) {
    if (validMonths.indexOf(m) === -1) {
      throw new Error("Config error: '" + m + "' is not a valid month name. Use full names like 'January'.");
    }
  });
  if (!ownerName) {
    throw new Error("Config error: Owner Name cell (D4) is empty.");
  }

  return { ownerName: ownerName, months: months, year: year };
}

var SHEET = {
  FINAL           : "Final DataSet",
  COMPLAINTS      : "Quality Complaints",
  ISSUES          : "Issues",
  INVENTORY       : "Inventory",
  PROPERTY_DETAILS: "Property Details"   // ← NEW: master property list with Status column
};

var MARKET_LABELS = {
  "Temecula" : "Temecula Market",
  "San Diego": "San Diego Market"
};

var MARKET_IMAGES = {
  "Temecula" : "https://drive.google.com/thumbnail?id=" + TM_MARKET_IMAGE_ID + "&sz=s1000",
  "San Diego": "https://drive.google.com/thumbnail?id=" + SD_MARKET_IMAGE_ID + "&sz=s1000",
};


// ============================================================
//  ACTIVE PROPERTY FILTER
//  Reads the "Property Details" sheet and returns a Set of
//  property names whose Status column equals "Active" (case-
//  insensitive).  This is used to guard every data section so
//  Offboarded / Pending / Onboarding properties are excluded.
// ============================================================

/**
 * getActivePropertyNames()
 * Returns an object (used as a Set) of property names that have
 * Status === "Active" in the Property Details sheet.
 * Keys are trimmed property names; value is always true.
 *
 * Falls back to an empty object if the sheet cannot be found,
 * which will cause all properties to be excluded — intentional,
 * so bad config is visible rather than silently including everything.
 */
function getActivePropertyNames() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET.PROPERTY_DETAILS);

  if (!sheet) {
    Logger.log("⚠️  Sheet '" + SHEET.PROPERTY_DETAILS + "' not found. No active properties will be matched.");
    return {};
  }

  var values  = sheet.getDataRange().getValues();
  var headers = values[0].map(function(h) { return String(h).trim(); });

  // Find "Name" and "Status" columns — tolerant of minor variations
  var nameCol   = -1;
  var statusCol = -1;
  headers.forEach(function(h, i) {
    var lower = h.toLowerCase();
    if (lower === "name"   && nameCol   === -1) nameCol   = i;
    if (lower === "status" && statusCol === -1) statusCol = i;
  });

  if (nameCol === -1 || statusCol === -1) {
    Logger.log("⚠️  Property Details sheet is missing 'Name' or 'Status' column. Headers found: " + JSON.stringify(headers));
    return {};
  }

  var activeSet = {};
  values.slice(1).forEach(function(row) {
    var name   = String(row[nameCol]   || "").trim();
    var status = String(row[statusCol] || "").trim().toLowerCase();
    if (name && status === "active") {
      activeSet[name] = true;
    }
  });

  Logger.log("Active properties loaded: " + Object.keys(activeSet).length + " — " + Object.keys(activeSet).join(", "));
  return activeSet;
}

/**
 * getActiveOwnerNames()
 * Returns the unique list of Company values from Property Details
 * where Status === "Active".  Used to populate the sidebar dropdown.
 */
function getActiveOwnerNames() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET.PROPERTY_DETAILS);

  if (!sheet) return [];

  var values  = sheet.getDataRange().getValues();
  var headers = values[0].map(function(h) { return String(h).trim(); });

  var companyCol = -1;
  var statusCol  = -1;
  headers.forEach(function(h, i) {
    var lower = h.toLowerCase();
    // Accept "Company" or "Company Name"
    if ((lower === "company" || lower === "company name") && companyCol === -1) companyCol = i;
    if (lower === "status" && statusCol === -1) statusCol = i;
  });

  if (companyCol === -1 || statusCol === -1) {
    Logger.log("⚠️  Property Details: missing 'Company'/'Company Name' or 'Status' column.");
    return [];
  }

  var seen   = {};
  var owners = [];
  values.slice(1).forEach(function(row) {
    var company = String(row[companyCol] || "").trim();
    var status  = String(row[statusCol]  || "").trim().toLowerCase();
    if (company && status === "active" && !seen[company]) {
      seen[company] = true;
      owners.push(company);
    }
  });

  return owners.sort();
}


// ============================================================
//  AI — OpenAI API KEY SETUP  (one-time)
// ============================================================
function setupOpenAIApiKey() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    "OpenAI API Key Setup",
    "Enter your OpenAI API key (starts with sk-...).\n" +
    "This is stored securely in Script Properties, not in the spreadsheet.",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    var key = result.getResponseText().trim();
    if (!key) { ui.alert("No key entered."); return; }
    PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", key);
    ui.alert("✅ API key saved. You can now generate reports with AI summaries.");
  }
}


// ============================================================
//  AI — OPENAI API HELPER
// ============================================================
function callOpenAI(prompt, maxTokens) {
  var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    throw new Error("OpenAI API key not set. Run setupOpenAIApiKey() first.");
  }

  var payload = {
    model: "gpt-4o-mini",
    max_completion_tokens: maxTokens || 1024,
    messages: [{ role: "user", content: prompt }],
  };

  var options = {
    method:             "post",
    contentType:        "application/json",
    headers:            { "Authorization": "Bearer " + apiKey },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
  var body     = JSON.parse(response.getContentText());

  if (body.error) throw new Error("OpenAI error: " + body.error.message);

  return body.choices[0].message.content.trim();
}


// ============================================================
//  ENTRY POINTS
// ============================================================

function generateOwnerReport(ownerName, month, year) {
  var data     = getOwnerReportData(ownerName, month, year);
  var template = HtmlService.createTemplateFromFile("ownerReport");
  template.ownerData = data;
  template.logoUrl   = logo;
      template.watermarkUrl = watermarkUrl;
    template.footerUrl = footer;
  var html = template.evaluate().getContent();
  return html;
}

function previewOwnerReport() {
  var cfg;
  try {
    cfg = readConfig();
  } catch (e) {
    SpreadsheetApp.getUi().alert("Configuration Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ownerName = cfg.ownerName;
  var months    = cfg.months;
  var year      = cfg.year;

  Logger.log("previewOwnerReport() called with:");
  Logger.log("  ownerName = " + ownerName);
  Logger.log("  months    = " + JSON.stringify(months));
  Logger.log("  year      = " + year);

  var data;
  try {
    data = getOwnerReportData(ownerName, months, year);
  } catch (e) {
    Logger.log("ERROR in data build: " + e.message);
    SpreadsheetApp.getUi().alert("Data Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  Logger.log("Data snapshot:");
  Logger.log("  clientName       = " + data.clientName);
  Logger.log("  markets.length   = " + (data.markets ? data.markets.length : "MISSING"));
  Logger.log("  properties.length= " + (data.properties ? data.properties.length : "MISSING"));
  Logger.log("  services.length  = " + (data.services ? data.services.length : "MISSING"));
  Logger.log("  serviceColumns   = " + JSON.stringify(data.serviceColumns));
  Logger.log("  complaints.length= " + (data.complaints ? data.complaints.length : "MISSING"));
  Logger.log("  inventoryItems.length = " + (data.inventoryItems ? data.inventoryItems.length : "MISSING"));
  Logger.log("  inventoryColumns = " + JSON.stringify(data.inventoryColumns));
  Logger.log("  issueItems.length= " + (data.issueItems ? data.issueItems.length : "MISSING"));
  Logger.log("  issueColumns     = " + JSON.stringify(data.issueColumns));
  Logger.log("  reviewsByProperty.length = " + (data.reviewsByProperty ? data.reviewsByProperty.length : "MISSING"));

  var template = HtmlService.createTemplateFromFile("ownerReport");
  template.ownerData = data;
  template.logoUrl   = logo;
  template.watermarkUrl = watermarkUrl;
  template.footerUrl = footer;
  var output = template.evaluate()
    .setWidth(900)
    .setHeight(1150)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showModalDialog(
    output,
    "Owner Report — " + ownerName + " · " + months.join(", ") + " " + year
  );
}

function debugConfig() {
  try {
    var cfg = readConfig();
    var msg = "✅ Config read successfully:\n\n" +
              "Owner : " + cfg.ownerName + "\n" +
              "Months: " + cfg.months.join(", ") + "\n" +
              "Year  : " + cfg.year;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert("Config OK", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log("Config ERROR: " + e.message);
    SpreadsheetApp.getUi().alert("Config Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function debugDataOnly() {
  var cfg;
  try {
    cfg = readConfig();
  } catch (e) {
    SpreadsheetApp.getUi().alert("Config Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  try {
    var data = getOwnerReportData(cfg.ownerName, cfg.months, cfg.year);

    Logger.log("=== FULL DATA SNAPSHOT ===");
    Logger.log("clientName       : " + data.clientName);
    Logger.log("markets          : " + JSON.stringify(data.markets));
    Logger.log("properties count : " + data.properties.length);
    Logger.log("properties       : " + JSON.stringify(data.properties));
    Logger.log("serviceColumns   : " + JSON.stringify(data.serviceColumns));
    Logger.log("serviceByProperty: " + JSON.stringify(data.serviceByProperty));
    Logger.log("complaints       : " + JSON.stringify(data.complaints));
    Logger.log("inventoryColumns : " + JSON.stringify(data.inventoryColumns));
    Logger.log("inventoryByProp  : " + JSON.stringify(data.inventoryByProperty));
    Logger.log("issueColumns     : " + JSON.stringify(data.issueColumns));
    Logger.log("issuesByProperty : " + JSON.stringify(data.issuesByProperty));
    Logger.log("reviewsByProperty: " + JSON.stringify(data.reviewsByProperty));

    var summary = "Data built OK. Check View > Logs for full snapshot.\n\n" +
      "markets: "    + data.markets.length    + "\n" +
      "properties: " + data.properties.length + "\n" +
      "services: "   + data.services.length   + "\n" +
      "complaints: " + data.complaints.length + "\n" +
      "inventoryItems: " + data.inventoryItems.length + "\n" +
      "issueItems: " + data.issueItems.length + "\n" +
      "reviewsByProperty: " + data.reviewsByProperty.length;

    SpreadsheetApp.getUi().alert("Data OK", summary, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (e) {
    Logger.log("ERROR building data: " + e.message + "\n" + e.stack);
    SpreadsheetApp.getUi().alert("Data Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function debugRawSheet() {
  var cfg;
  try { cfg = readConfig(); }
  catch (e) {
    SpreadsheetApp.getUi().alert("Config Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var monthIdx = monthNameToIndex(cfg.months[0]);
  var summary  = "";

  function debugSheet(sheetKey, sheetName, dateColNames, propColNames, ownerColNames) {
    Logger.log("");
    Logger.log("════════════════════════════════════════");
    Logger.log("SHEET: " + sheetName);
    Logger.log("════════════════════════════════════════");

    var s = ss.getSheetByName(sheetName);
    if (!s) {
      Logger.log("  ERROR: Sheet not found!");
      summary += sheetName + ": NOT FOUND";
      return;
    }

    var values  = s.getDataRange().getValues();
    var headers = values[0];
    var rows    = values.slice(1);
    Logger.log("  Total data rows: " + rows.length);

    Logger.log("  Headers:");
    headers.forEach(function(h, i) {
      Logger.log("    col " + (i+1) + ": [" + String(h).trim() + "]");
    });

    var colDate = -1, colProp = -1, colOwner = -1;
    headers.forEach(function(h, i) {
      var c = String(h).trim();
      if (dateColNames.indexOf(c)  > -1 && colDate  === -1) colDate  = i;
      if (propColNames.indexOf(c)  > -1 && colProp  === -1) colProp  = i;
      if (ownerColNames.indexOf(c) > -1 && colOwner === -1) colOwner = i;
    });

    Logger.log("  Date col  [" + dateColNames.join("/") + "]: col index " + colDate  + (colDate  === -1 ? " ← NOT FOUND" : " ✅"));
    Logger.log("  Prop col  [" + propColNames.join("/") + "]: col index " + colProp  + (colProp  === -1 ? " ← NOT FOUND" : " ✅"));
    if (ownerColNames.length) {
      Logger.log("  Owner col [" + ownerColNames.join("/") + "]: col index " + colOwner + (colOwner === -1 ? " ← NOT FOUND" : " ✅"));
    }

    Logger.log("  First 5 rows:");
    rows.slice(0, 5).forEach(function(row, i) {
      Logger.log("    Row " + (i+2) + ":"
        + "  date=[" + (colDate > -1 ? row[colDate] : "N/A") + "]"
        + "  type=" + (colDate > -1 ? typeof row[colDate] : "N/A")
        + "  prop=[" + (colProp > -1 ? row[colProp] : "N/A") + "]"
        + (colOwner > -1 ? "  owner=[" + row[colOwner] + "]" : ""));
    });

    var passDate = 0, passBoth = 0, passOwner = 0;
    rows.forEach(function(row) {
      var dateOk  = dateMatchesPeriod(row[colDate], monthIdx, cfg.year);
      var ownerOk = colOwner > -1 ? String(row[colOwner]).trim() === cfg.ownerName : true;
      if (dateOk)            passDate++;
      if (ownerOk)           passOwner++;
      if (dateOk && ownerOk) passBoth++;
    });

    if (colOwner > -1) {
      Logger.log("  Filter: owner only=" + passOwner + "  date only=" + passDate + "  both=" + passBoth + (passBoth === 0 ? " ← EMPTY" : " ← OK"));
      summary += sheetName + ": both=" + passBoth + (passBoth === 0 ? " EMPTY" : " OK") + "\n";
    } else {
      Logger.log("  Filter: date match=" + passDate + (passDate === 0 ? " ← NO DATE MATCHES" : " ← OK"));
      summary += sheetName + ": date matches=" + passDate + (passDate === 0 ? " EMPTY" : " OK") + "\n";
    }
  }

  debugSheet("FINAL", SHEET.FINAL,
    ["Appt Date", "App Date", "Appointment Date"],
    ["Property Name"],
    ["Company Name"]
  );
  debugSheet("COMPLAINTS", SHEET.COMPLAINTS,
    ["Date"], ["Property", "Property Name"], []
  );
  debugSheet("ISSUES", SHEET.ISSUES,
    ["Date"], ["Property Name", "Property"], []
  );
  debugSheet("INVENTORY", SHEET.INVENTORY,
    ["Date"], ["Property Name", "Property"], []
  );

  Logger.log("════ SUMMARY ════");
  Logger.log(summary);

  SpreadsheetApp.getUi().alert(
    "Debug Complete — All 4 Sheets",
    "Check View > Logs for full detail.\n\n" + summary,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function generateAllOwnerReports(month, year) {
  month = month || "January";
  year  = year  || 2026;

  // Use active owners from Property Details sheet
  var owners = getActiveOwnerNames();
  Logger.log("Generating reports for " + owners.length + " active owners: " + owners.join(", "));

  owners.forEach(function(ownerName) {
    try {
      generateOwnerReport(ownerName, month, year);
      Logger.log("✓ Done: " + ownerName);
    } catch (e) {
      Logger.log("✗ Failed: " + ownerName + " — " + e.message);
    }
  });
}


// ============================================================
//  LIVE DATA READER
// ============================================================

function getOwnerReportData(ownerName, months, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var monthsArr    = Array.isArray(months) ? months : [months];
  var monthIndices = monthsArr.map(function(m) { return monthNameToIndex(m); });

  var sortedIndices   = monthIndices.slice().sort(function(a, b) { return a - b; });
  var validMonthNames = ["January","February","March","April","May","June",
                         "July","August","September","October","November","December"];
  var monthLabel = sortedIndices.length === 1
    ? validMonthNames[sortedIndices[0]]
    : validMonthNames[sortedIndices[0]] + " – " + validMonthNames[sortedIndices[sortedIndices.length - 1]];

  // ── Load active property filter ──────────────────────────────
  var activePropertyNames = getActivePropertyNames();

  var allAppts      = sheetToObjects(ss, SHEET.FINAL);
  var allComplaints = sheetToObjects(ss, SHEET.COMPLAINTS);
  var allIssues     = sheetToObjects(ss, SHEET.ISSUES);
  var allInventory  = sheetToObjects(ss, SHEET.INVENTORY);

  // ── Filter appointments: must match owner, date, AND be an active property
  var appts = allAppts.filter(function(r) {
    var owner   = trim(r["Company Name"]);
    var prop    = trim(r["Property Name"]);
    var dateVal = r["Appt Date"] || r["App Date"] || r["Appointment Date"];
    if (owner !== ownerName) return false;
    if (!activePropertyNames[prop]) return false;   // ← Active filter
    return monthIndices.some(function(mi) { return dateMatchesPeriod(dateVal, mi, year); });
  });

  Logger.log("=== APPOINTMENT FILTER RESULT ===");
  Logger.log("  Owner=[" + ownerName + "] months=[" + monthLabel + "] year=[" + year + "]");
  Logger.log("  Total rows in Final DataSet : " + allAppts.length);
  Logger.log("  Rows after filter (active)  : " + appts.length);

  // ownerProperties is now guaranteed to only contain active properties
  var ownerProperties = unique(appts.map(function(r) { return trim(r["Property Name"]); }));

  // ── Build a normalised (stripped + lowercased) version of ownerProperties
  //    for fuzzy-matching against the Quality Complaints sheet, which may
  //    have numeric prefixes (e.g. "123 - Palm Springs") or different casing.
  var ownerPropertiesNorm = ownerProperties.map(function(p) {
    return stripPropertyId(p).toLowerCase();
  });

  Logger.log("  ownerProperties (raw)  : " + JSON.stringify(ownerProperties));
  Logger.log("  ownerProperties (norm) : " + JSON.stringify(ownerPropertiesNorm));

  // ── Filter complaints by matching the normalised property name ────────────
  //    Also log a sample of complaint property values so mismatches are visible.
  var sampleComplaintProps = allComplaints.slice(0, 5).map(function(r) {
    return trim(r["Property"]);
  });
  Logger.log("  Sample complaint 'Property' values: " + JSON.stringify(sampleComplaintProps));

  var complaints = allComplaints.filter(function(r) {
    var rawProp  = trim(r["Property"]);
    var normProp = stripPropertyId(rawProp).toLowerCase();
    if (ownerPropertiesNorm.indexOf(normProp) === -1) return false;
    return monthIndices.some(function(mi) { return dateMatchesPeriod(r["Date"], mi, year); });
  });

  Logger.log("  Complaints matched: " + complaints.length);

  var issues = allIssues.filter(function(r) {
    var owner  = trim(r["Company Name"]);
    var dateOk = monthIndices.some(function(mi) { return dateMatchesPeriod(r["Date"], mi, year); });
    if (owner && owner !== "") return owner === ownerName && dateOk;
    var rawProp = trim(r["Property Name"] || r["Property"]);
    var prop    = stripPropertyId(rawProp);
    return ownerProperties.indexOf(prop) > -1 && dateOk;
  });

  var inventory = allInventory.filter(function(r) {
    var owner  = trim(r["Company Name"]);
    var dateOk = monthIndices.some(function(mi) { return dateMatchesPeriod(r["Date"], mi, year); });
    if (owner && owner !== "") return owner === ownerName && dateOk;
    var prop = trim(r["Property Name"]);
    return ownerProperties.indexOf(prop) > -1 && dateOk;
  });

  var serviceSplit   = buildServiceByProperty(appts, ownerProperties);
  var inventorySplit = buildInventoryByProperty(inventory, ownerProperties);
  var issuesSplit    = buildIssuesByProperty(issues, ownerProperties);

  // ── AI Summaries (with graceful fallbacks) ────────────────
  var executiveSummary         = buildExecutiveSummary(appts, complaints, issues, ownerName, monthLabel, year);
  var qualityComplaintsSummary = buildComplaintsSummary(complaints);
  var inventorySummary         = buildInventorySummary(inventory);
  var reviewSummary            = buildReviewSummary(appts, ownerName, monthLabel);

  return {
    month      : monthLabel,
    year       : String(year),
    date       : formatDate(new Date()),
    clientName : ownerName,

    executiveSummary : executiveSummary,

    markets    : buildMarkets(appts, ownerProperties),
    properties : buildPropertyBreakdown(appts, ownerProperties),

    services          : buildServiceBreakdown(appts),
    serviceColumns    : serviceSplit.columns,
    serviceByProperty : serviceSplit.rows,

    qualityComplaintsSummary : qualityComplaintsSummary,
    complaints               : buildComplaints(complaints),

    inventorySummary    : inventorySummary,
    inventoryItems      : buildInventoryChart(inventory),
    inventoryColumns    : inventorySplit.columns,
    inventoryByProperty : inventorySplit.rows,

    issueItems       : buildIssuesChart(issues),
    issueColumns     : issuesSplit.columns,
    issuesByProperty : issuesSplit.rows,

    reviewSummary     : reviewSummary,
    reviewsByProperty : buildReviewsByProperty(appts, ownerProperties)
  };
}


// ============================================================
//  SECTION BUILDERS
// ============================================================

// ── Executive Summary (AI) ────────────────────────────────────
function buildExecutiveSummary(appts, complaints, issues, ownerName, monthLabel, year) {
  var totalAppts      = appts.length;
  var reschedules     = appts.filter(function(r) { return isTruthy(r["Reschedules"] || r["Reschedule"]); }).length;
  var reschPct        = totalAppts > 0 ? ((reschedules / totalAppts) * 100).toFixed(1) : "0.0";
  var totalComplaints = complaints.length;
  var totalIssues     = issues.length;
  var avgRating       = calcAvgRating(appts);

  var props = unique(appts.map(function(r) { return trim(r["Property Name"]); }));

  var svcCounts = {};
  appts.forEach(function(r) {
    var s = trim(r["Service"]);
    if (s) svcCounts[s] = (svcCounts[s] || 0) + 1;
  });
  var topService = Object.keys(svcCounts).sort(function(a, b) { return svcCounts[b] - svcCounts[a]; })[0] || "N/A";

  var allServiceLines = Object.keys(svcCounts)
    .sort(function(a, b) { return svcCounts[b] - svcCounts[a]; })
    .map(function(s) { return s + ": " + svcCounts[s] + " job" + (svcCounts[s] > 1 ? "s" : ""); })
    .join(", ");

  var propApptLines = props.map(function(p) {
    var count = appts.filter(function(r) { return trim(r["Property Name"]) === p; }).length;
    return p + " (" + count + " appt" + (count > 1 ? "s" : "") + ")";
  }).join(", ");

  var propRatingLines = props.map(function(p) {
    var propAppts = appts.filter(function(r) { return trim(r["Property Name"]) === p; });
    var ratings   = propAppts.map(function(r) { return parseFloat(r["Reviews"]); }).filter(function(v) { return !isNaN(v) && v > 0; });
    if (!ratings.length) return p + ": no reviews";
    var avg = Math.round((ratings.reduce(function(t, v) { return t + v; }, 0) / ratings.length) * 10) / 10;
    return p + ": " + avg + "★ (" + ratings.length + ")";
  }).join(", ");

  var prompt =
    "You are the lead analyst at Woldenga Property Services, a professional " +
    "short-term rental cleaning and operations company. You are writing the " +
    "executive summary for " + ownerName + "'s monthly owner report — this is the " +
    "first thing the owner reads, so it must be thorough, engaging, and tell the " +
    "complete story of the month across every dimension of the portfolio.\n\n" +

    "REPORT PERIOD: " + monthLabel + " " + year + "\n\n" +

    "DATA:\n" +
    "  - Owner / Client            : " + ownerName + "\n" +
    "  - Properties managed        : " + props.length + " — " + propApptLines + "\n" +
    "  - Total appointments        : " + totalAppts + "\n" +
    "  - Service breakdown         : " + (allServiceLines || "N/A") + "\n" +
    "  - Reschedules               : " + reschedules + " (" + reschPct + "% — target is ≤ 5%)\n" +
    "  - Quality complaints        : " + totalComplaints + "\n" +
    "  - Property issues logged    : " + totalIssues + "\n" +
    "  - Portfolio avg guest rating: " + (avgRating !== null ? avgRating + " ★" : "No reviews this period") + "\n" +
    "  - Per-property ratings      : " + (propRatingLines || "N/A") + "\n\n" +

    "WRITING INSTRUCTIONS:\n" +
    "- Write exactly ONE flowing paragraph of 9–11 sentences. Do not break it into multiple paragraphs.\n" +
    "- Sentence 1 — Opening: Set the scene. Name the owner naturally and anchor the reader in the reporting period. " +
    "Reference the portfolio size and the overall volume of work delivered this month.\n" +
    "- Sentence 2–3 — Operations: Break down the appointment activity in detail. " +
    "Name the specific services performed, highlight the most dominant service type, " +
    "and describe how work was distributed across the properties.\n" +
    "- Sentence 4 — Scheduling: Discuss rescheduling performance. " +
    "State the rate clearly, compare it to the 5% target, and frame it positively or constructively depending on the result.\n" +
    "- Sentence 5–6 — Quality & Issues: Address quality complaints — how many occurred, what types, " +
    "and the resolution approach. Then cover property issues logged — volume, nature, and whether they were handled. " +
    "Be honest and accountable without being alarming.\n" +
    "- Sentence 7–8 — Guest Satisfaction: Discuss the guest rating results in depth. " +
    "Reference the portfolio average, call out any standout properties by name, " +
    "and flag any property that may need attention. If no reviews exist, acknowledge it and pivot to operational strengths.\n" +
    "- Sentence 9–10 — Closing: Reflect on the overall shape of the month — what went well, " +
    "what the team is focused on improving, and what the owner can look forward to. " +
    "End with a warm, confident, forward-looking statement that reinforces trust in the partnership.\n" +
    "- Tone: professional, warm, candid, and data-driven. Write as if speaking directly to the owner — " +
    "someone who trusts you with their properties and deserves a full, honest picture.\n" +
    "- Do NOT use bullet points, numbered lists, headers, subheadings, or any markdown formatting.\n" +
    "- Do NOT start with 'Dear' or any salutation.\n" +
    "- The output must be one single unbroken paragraph. No line breaks between sentences.\n" +
    "- Return ONLY the paragraph text. Nothing before it, nothing after it.\n";

  try {
    var result = callOpenAI(prompt, 1000);
    Logger.log("Executive summary generated (" + result.length + " chars).");
    return result;
  } catch (e) {
    Logger.log("Executive summary AI failed — using fallback: " + e.message);
    return monthLabel + " " + year + " saw " + totalAppts + " completed appointment(s) across " +
      ownerName + "'s portfolio of " + props.length + " propert" + (props.length === 1 ? "y" : "ies") +
      " — " + propApptLines + ". " +
      "The primary service delivered was " + topService + ", accounting for " + (svcCounts[topService] || 0) + " of the total jobs. " +
      "The rescheduling rate for the period was " + reschPct + "%, " +
      (parseFloat(reschPct) <= 5 ? "which came in at or below the 5% target — a strong result for the team. " : "which exceeded the 5% target, and the team is actively working to bring this back in line. ") +
      totalComplaints + " quality complaint(s) were recorded and " + totalIssues + " property issue(s) were logged, " +
      "all of which are documented with resolutions in the sections below. " +
      (avgRating !== null
        ? "Guests rated their experience an average of " + avgRating + " ★ across the portfolio this period — " +
          "a reflection of the consistent, high-quality service the team delivered. "
        : "No guest reviews were recorded this period, though the operational metrics reflect a strong month of service. ") +
      "We remain committed to continuous improvement and look forward to building on this month's performance for " + ownerName + ".";
  }
}


// ── Quality Complaints Summary (AI) ──────────────────────────
function buildComplaintsSummary(complaints) {
  if (!complaints.length) {
    return "No quality complaints were recorded this period — an excellent outcome for the portfolio.";
  }

  var resolved = complaints.filter(function(r) {
    var h = trim(r["How it Was Handled?"] || r["How it was Handled?"]);
    return h && h !== "";
  }).length;

  var descriptions = complaints
    .map(function(r) { return trim(r["Description"]); })
    .filter(Boolean)
    .slice(0, 15)
    .join(" | ");

  var prompt =
    "You are writing the Quality Complaints summary section for an owner operations report " +
    "at Woldenga Property Services, a short-term rental cleaning company.\n\n" +

    "DATA:\n" +
    "  - Total complaints      : " + complaints.length + "\n" +
    "  - Complaints resolved   : " + resolved + " of " + complaints.length + "\n" +
    "  - Complaint descriptions: " + (descriptions || "N/A") + "\n\n" +

    "WRITING INSTRUCTIONS:\n" +
    "- Write 2–3 concise sentences.\n" +
    "- State the total number of complaints and the most common themes you observe across the descriptions.\n" +
    "- Note how many were resolved and the general resolution approach.\n" +
    "- Tone: factual, accountable, and solution-oriented. Do not be alarming.\n" +
    "- Do NOT use bullet points, headers, or markdown formatting.\n" +
    "- Return ONLY the summary text. Nothing else.\n";

  try {
    var result = callOpenAI(prompt, 300);
    Logger.log("Complaints summary generated (" + result.length + " chars).");
    return result;
  } catch (e) {
    Logger.log("Complaints summary AI failed — using fallback: " + e.message);
    return complaints.length + " quality complaint(s) were recorded this period. " +
      resolved + " of " + complaints.length + " were resolved as documented below.";
  }
}


// ── Inventory Summary (AI) ────────────────────────────────────
function buildInventorySummary(inventory) {
  if (!inventory.length) {
    return "No inventory requests were recorded this period.";
  }

  var itemCounts = {};
  inventory.forEach(function(r) {
    var item = trim(r["Item"]);
    var qty  = parseInt(r["Quantity"]) || 1;
    if (item) itemCounts[item] = (itemCounts[item] || 0) + qty;
  });

  var totalRequests = inventory.length;
  var totalUnits    = Object.keys(itemCounts).reduce(function(s, k) { return s + itemCounts[k]; }, 0);

  var itemLines = Object.keys(itemCounts)
    .sort(function(a, b) { return itemCounts[b] - itemCounts[a]; })
    .map(function(i) { return i + " (" + itemCounts[i] + ")"; })
    .join(", ");

  var prompt =
    "You are writing the Inventory Requests summary section for an owner operations report " +
    "at Woldenga Property Services, a short-term rental cleaning company.\n\n" +

    "DATA:\n" +
    "  - Total requests : " + totalRequests + "\n" +
    "  - Total units    : " + totalUnits + "\n" +
    "  - Items & counts : " + itemLines + "\n\n" +

    "WRITING INSTRUCTIONS:\n" +
    "- Write 2 concise sentences.\n" +
    "- Mention the total number of requests and the top requested items.\n" +
    "- End by confirming all requests have been fulfilled.\n" +
    "- Tone: efficient, professional.\n" +
    "- Do NOT use bullet points, headers, or markdown formatting.\n" +
    "- Return ONLY the summary text. Nothing else.\n";

  try {
    var result = callOpenAI(prompt, 200);
    Logger.log("Inventory summary generated (" + result.length + " chars).");
    return result;
  } catch (e) {
    Logger.log("Inventory summary AI failed — using fallback: " + e.message);
    var topItems = Object.keys(itemCounts)
      .sort(function(a, b) { return itemCounts[b] - itemCounts[a]; })
      .slice(0, 3)
      .join(", ");
    return totalRequests + " inventory request(s) were logged this period. " +
      "The most requested items were: " + topItems + ". " +
      "All requests have been fulfilled and marked Closed.";
  }
}


// ── Review Summary (AI) ───────────────────────────────────────
function buildReviewSummary(appts, ownerName, monthLabel) {
  var rated = appts.filter(function(r) {
    var v = parseFloat(r["Reviews"]);
    return !isNaN(v) && v > 0;
  });

  if (!rated.length) {
    return "No guest reviews were recorded this period.";
  }

  var avg = calcAvgRating(appts);
  var min = Math.min.apply(null, rated.map(function(r) { return parseFloat(r["Reviews"]); }));
  var max = Math.max.apply(null, rated.map(function(r) { return parseFloat(r["Reviews"]); }));

  var propRatings = {};
  rated.forEach(function(r) {
    var prop = trim(r["Property Name"]);
    var val  = parseFloat(r["Reviews"]);
    if (!propRatings[prop]) propRatings[prop] = [];
    propRatings[prop].push(val);
  });

  var propLines = Object.keys(propRatings).map(function(p) {
    var vals    = propRatings[p];
    var propAvg = Math.round((vals.reduce(function(t, v) { return t + v; }, 0) / vals.length) * 10) / 10;
    return p + ": " + propAvg + "★ (" + vals.length + " review" + (vals.length > 1 ? "s" : "") + ")";
  }).join(", ");

  var prompt =
    "You are writing the Guest Reviews summary section for an owner operations report " +
    "at Woldenga Property Services, a short-term rental cleaning company.\n\n" +

    "DATA:\n" +
    "  - Owner / Client         : " + ownerName + "\n" +
    "  - Period                 : " + monthLabel + "\n" +
    "  - Total reviews          : " + rated.length + "\n" +
    "  - Portfolio avg rating   : " + avg + " ★\n" +
    "  - Highest single rating  : " + max + " ★\n" +
    "  - Lowest single rating   : " + min + " ★\n" +
    "  - Per-property breakdown : " + propLines + "\n\n" +

    "WRITING INSTRUCTIONS:\n" +
    "- Write 2–3 concise sentences.\n" +
    "- State the total reviews collected and overall average rating.\n" +
    "- Mention standout performers and flag any property below 4.5 ★ if applicable.\n" +
    "- End on a positive, encouraging note about guest satisfaction.\n" +
    "- Tone: professional, warm, data-driven.\n" +
    "- Do NOT use bullet points, headers, or markdown formatting.\n" +
    "- Return ONLY the summary text. Nothing else.\n";

  try {
    var result = callOpenAI(prompt, 300);
    Logger.log("Review summary generated (" + result.length + " chars).");
    return result;
  } catch (e) {
    Logger.log("Review summary AI failed — using fallback: " + e.message);
    return rated.length + " guest review(s) were collected this period with an overall " +
      "average rating of " + avg + " ★. The lowest individual property rating was " + min + " ★. " +
      "Full per-property detail is shown in the table below.";
  }
}


// ── Market Cards ──────────────────────────────────────────────
function buildMarkets(appts, ownerProperties) {
  var marketMap = {};
  appts.forEach(function(r) {
    var area = trim(r["Property Area"]);
    var prop = trim(r["Property Name"]);
    if (!area) return;
    if (!marketMap[area]) {
      marketMap[area] = { propSet: {}, appts: [], complaints: 0, reschedules: 0, reviews: [], issues: 0 };
    }
    marketMap[area].propSet[prop] = true;
    marketMap[area].appts.push(r);
    if (isTruthy(r["Quality Complaints"])) marketMap[area].complaints++;
    if (isTruthy(r["Resheduels"] || r["Reschedules"] || r["Reschedule"])) marketMap[area].reschedules++;
    var rating = parseFloat(r["Reviews"]);
    if (!isNaN(rating) && rating > 0) marketMap[area].reviews.push(rating);
  });

  return Object.keys(marketMap).map(function(area) {
    var m         = marketMap[area];
    var apptCount = m.appts.length;
    var avgRating = m.reviews.length
      ? Math.round((m.reviews.reduce(function(t, v) { return t + v; }, 0) / m.reviews.length) * 10) / 10
      : null;

    return {
      name              : MARKET_LABELS[area] || area + " Market",
      photoUrl          : MARKET_IMAGES[area] || null,
      properties        : Object.keys(m.propSet).length,
      avgRating         : avgRating || "—",
      reviewCount       : m.reviews.length,
      appointments      : apptCount,
      qualityComplaints : m.complaints,
      reschedules       : m.reschedules,
      rescheduleTarget  : Math.round(apptCount * 0.05 * 10) / 10,
      inventoryIssues   : 0
    };
  });
}


// ── Property Breakdown Table ──────────────────────────────────
function buildPropertyBreakdown(appts, ownerProperties) {
  return ownerProperties.map(function(prop) {
    var rows        = appts.filter(function(r) { return trim(r["Property Name"]) === prop; });
    var reschedules = rows.filter(function(r) { return isTruthy(r["Resheduels"] || r["Reschedules"] || r["Reschedule"]); }).length;
    var qc          = rows.filter(function(r) { return isTruthy(r["Quality Complaints"]); }).length;
    var ta          = rows.filter(function(r) { return isTruthy(r["T/A"]); }).length;
    var ratings     = rows
      .map(function(r) { return parseFloat(r["Reviews"]); })
      .filter(function(v) { return !isNaN(v) && v > 0; });
    var reviewCount   = ratings.length;
    var avgRating     = reviewCount
      ? Math.round((ratings.reduce(function(t, v) { return t + v; }, 0) / reviewCount) * 10) / 10
      : null;
    var reviewDisplay = reviewCount > 0
      ? reviewCount + " (" + avgRating.toFixed(1) + "\u2605)"
      : "\u2014";

    return {
      name              : prop,
      cleans            : rows.length,
      reschedules       : reschedules,
      qualityComplaints : qc,
      inventoryIssues   : 0,
      reviews           : reviewDisplay,
      reviewCount       : reviewCount,
      avgRating         : avgRating,
      turnArounds       : ta
    };
  });
}


// ── Service Breakdown Bar Chart ───────────────────────────────
var SERVICE_BAR_COLORS = [
  "#9fc5e8", "#e0973a", "#9b7ed9", "#6fa8d6",
  "#5db8d0", "#a3c87a", "#e07a7a", "#7ad4a3", "#c4a87a"
];

function buildServiceBreakdown(appts) {
  var counts = {};
  appts.forEach(function(r) {
    var svc = trim(r["Service"]);
    if (!svc) return;
    counts[svc] = (counts[svc] || 0) + 1;
  });

  var sorted = Object.keys(counts).sort(function(a, b) { return counts[b] - counts[a]; });

  return sorted.map(function(label, i) {
    return {
      label : label,
      count : counts[label],
      color : SERVICE_BAR_COLORS[i % SERVICE_BAR_COLORS.length]
    };
  });
}


// ── Service By Property Table ─────────────────────────────────
function buildServiceByProperty(appts, ownerProperties) {
  var counts = {};
  appts.forEach(function(r) {
    var svc = trim(r["Service"]);
    if (svc) counts[svc] = (counts[svc] || 0) + 1;
  });
  var serviceColumns = Object.keys(counts).sort(function(a, b) { return counts[b] - counts[a]; });

  var rows = ownerProperties.map(function(prop) {
    var propAppts = appts.filter(function(r) { return trim(r["Property Name"]) === prop; });
    var row = { name: prop };
    serviceColumns.forEach(function(svc) {
      row[svc] = propAppts.filter(function(r) { return trim(r["Service"]) === svc; }).length;
    });
    return row;
  });

  return { columns: serviceColumns, rows: rows };
}


// ── Quality Complaints Table ──────────────────────────────────
//
//  Two-stage AI strategy:
//  ─────────────────────────────────────────────────────────────
//  Stage 1 — Categorization (single call):
//    All complaint descriptions are sent together so the model
//    sees the full scope of issues before deciding on labels.
//    Returns JSON: [{ "index": 0, "category": "..." }, ...]
//    NOTE: token limit is set high (2000) because descriptions
//    can be multi-paragraph; a low limit truncates the JSON and
//    causes parse failures that silently fall through to the
//    Type column — which we never want to use for categorization.
//
//  Stage 2 — Resolution per category (one call per category):
//    For each derived category, the model reads the complaint
//    descriptions + any raw resolution notes and writes a clean,
//    professional 1–2 sentence resolution statement.
//
function buildComplaints(complaints) {
  if (!complaints.length) return [];

  // ── Stage 1: Categorize ──────────────────────────────────────
  //  Use only the Description field — never the Type column.
  //  Long descriptions are truncated to 400 chars each so the
  //  combined prompt stays within a safe token budget while still
  //  giving the model enough context to categorize accurately.
  var descLines = complaints.map(function(r, i) {
    var desc = trim(r["Description"]) || "No description provided";
    if (desc.length > 400) desc = desc.substring(0, 400) + "…";
    return (i + 1) + ". " + desc;
  }).join("\n");

  var categorizationPrompt =
    "You are categorizing quality complaints for Woldenga Property Services, " +
    "a short-term rental cleaning company.\n\n" +

    "Below are " + complaints.length + " complaint descriptions, each numbered.\n\n" +
    "COMPLAINTS:\n" + descLines + "\n\n" +

    "TASK:\n" +
    "1. Read ALL descriptions first to understand the full range of issues.\n" +
    "2. Derive a set of 4–8 concise, specific category labels that together cover " +
    "every complaint above. Base categories ONLY on the description content — " +
    "do not use any pre-existing type labels. " +
    "Good examples: \'Hair on Linens\', \'Missed Cleaning Areas\', \'Lock-Up Procedure\', " +
    "\'Kitchen & Appliances\', \'Presentation & Staging\', \'Guest Supplies\', " +
    "\'Mold & Moisture\', \'Floor & Surface Cleanliness\'. " +
    "Do NOT use vague labels like \'Multiple\', \'Other\', or \'Miscellaneous\'.\n" +
    "3. Assign each complaint to exactly ONE category from your derived list.\n\n" +

    "OUTPUT RULES:\n" +
    "- Return ONLY a valid JSON array. No explanation, no markdown, no code fences.\n" +
    "- Each element must have exactly two keys: \"index\" (0-based integer) and \"category\" (string).\n" +
    "- The array must contain exactly " + complaints.length + " elements — one per complaint.\n" +
    "- Example: [{\"index\":0,\"category\":\"Mold & Moisture\"},{\"index\":1,\"category\":\"Hair on Linens\"}]\n";

  var categoryMap = {};
  try {
    var raw     = callOpenAI(categorizationPrompt, 2000);  // high limit — descriptions can be multi-paragraph
    var cleaned = raw.replace(/```json|```/gi, "").trim();
    var parsed  = JSON.parse(cleaned);

    if (!Array.isArray(parsed)) throw new Error("Response is not an array");
    if (parsed.length !== complaints.length) {
      Logger.log("⚠️  AI returned " + parsed.length + " items for " + complaints.length + " complaints — partial map, gaps will be filled.");
    }

    parsed.forEach(function(item) {
      if (typeof item.index === "number" && typeof item.category === "string") {
        categoryMap[item.index] = item.category.trim();
      }
    });

    Logger.log("AI complaint categorization succeeded. Categories assigned: " + Object.keys(categoryMap).length);
  } catch (e) {
    Logger.log("AI complaint categorization failed — using description-keyword fallback: " + e.message);
    // Fallback: derive a rough category from the description content.
    // The Type column is intentionally never used here.
    complaints.forEach(function(r, i) {
      var desc = trim(r["Description"]);
      if (!desc) { categoryMap[i] = "General Cleanliness"; return; }
      var lower = desc.toLowerCase();
      if (lower.indexOf("mold") > -1 || lower.indexOf("mould") > -1) {
        categoryMap[i] = "Mold & Moisture";
      } else if (lower.indexOf("hair") > -1 && (lower.indexOf("linen") > -1 || lower.indexOf("sheet") > -1 || lower.indexOf("bed") > -1)) {
        categoryMap[i] = "Hair on Linens";
      } else if (lower.indexOf("lock") > -1 || lower.indexOf("door") > -1 || lower.indexOf("unlock") > -1) {
        categoryMap[i] = "Lock-Up Procedure";
      } else if (lower.indexOf("floor") > -1 || lower.indexOf("mop") > -1 || lower.indexOf("sweep") > -1) {
        categoryMap[i] = "Floor & Surface Cleanliness";
      } else if (lower.indexOf("kitchen") > -1 || lower.indexOf("stove") > -1 || lower.indexOf("dish") > -1 || lower.indexOf("fridge") > -1) {
        categoryMap[i] = "Kitchen & Appliances";
      } else if (lower.indexOf("towel") > -1 || lower.indexOf("supply") > -1 || lower.indexOf("supplies") > -1) {
        categoryMap[i] = "Guest Supplies";
      } else if (lower.indexOf("dust") > -1 || lower.indexOf("cobweb") > -1 || lower.indexOf("dirty") > -1) {
        categoryMap[i] = "Missed Cleaning Areas";
      } else {
        categoryMap[i] = "Presentation & Staging";
      }
    });
  }

  // ── Group complaints by category ─────────────────────────────
  //  Any complaint whose AI category is missing gets "Uncategorized".
  //  The Type column is never referenced here.
  var grouped = {};
  var categoryOrder = [];

  complaints.forEach(function(r, i) {
    var cat = categoryMap[i] || "Uncategorized";
    if (!grouped[cat]) {
      grouped[cat] = { descriptions: [], rawResolutions: [] };
      categoryOrder.push(cat);
    }
    var desc = trim(r["Description"]) || "";
    var res  = trim(r["How it Was Handled?"] || r["How it was Handled?"] || "");
    if (desc) grouped[cat].descriptions.push(desc);
    if (res)  grouped[cat].rawResolutions.push(res);
  });

  // ── Stage 2: AI resolution per category ──────────────────────
  return categoryOrder.map(function(cat) {
    var g         = grouped[cat];
    var frequency = complaints.filter(function(r, i) {
      return (categoryMap[i] || "Uncategorized") === cat;
    }).length;

    var descSummary = g.descriptions.slice(0, 5).join(" | ") || "No description available";
    var rawResSummary = g.rawResolutions.length
      ? g.rawResolutions.slice(0, 5).join(" | ")
      : "No resolution notes recorded";

    var resolutionPrompt =
      "You are writing the 'How We Resolved It' entry for a quality complaint category " +
      "in an owner operations report at Woldenga Property Services, a short-term rental " +
      "cleaning company.\n\n" +

      "COMPLAINT CATEGORY: " + cat + "\n" +
      "FREQUENCY: " + frequency + " occurrence" + (frequency > 1 ? "s" : "") + " this period\n\n" +

      "COMPLAINT DESCRIPTIONS:\n" + descSummary + "\n\n" +

      "RAW RESOLUTION NOTES (from the team, may be informal or incomplete):\n" +
      rawResSummary + "\n\n" +

      "TASK:\n" +
      "Write a clean, professional 1–2 sentence resolution statement that an owner would " +
      "read in a formal report. It should:\n" +
      "- Acknowledge what happened briefly (don't dwell on the problem).\n" +
      "- Describe the action taken or the corrective measure applied.\n" +
      "- Sound accountable, solution-oriented, and confident.\n" +
      "- If raw resolution notes are vague or missing, write a reasonable standard " +
      "response based on the complaint type (e.g. re-cleaning, staff coaching, " +
      "process review).\n\n" +

      "RULES:\n" +
      "- Do NOT use bullet points, headers, or markdown.\n" +
      "- Return ONLY the resolution text. Nothing else.\n";

    var resolution = "Pending";
    try {
      resolution = callOpenAI(resolutionPrompt, 200);
      Logger.log("AI resolution generated for [" + cat + "]: " + resolution.length + " chars.");
    } catch (e) {
      Logger.log("AI resolution failed for [" + cat + "] — using raw fallback: " + e.message);
      resolution = g.rawResolutions.length ? g.rawResolutions[0] : "Pending review by operations team.";
    }

    return {
      category  : cat,
      frequency : frequency,
      resolution: resolution
    };
  });
}


// ── Inventory Chart & By-Property Table ──────────────────────
function buildInventoryChart(inventory) {
  var itemCounts = {};
  inventory.forEach(function(r) {
    var item = trim(r["Item"]);
    var qty  = parseInt(r["Quantity"]) || 1;
    if (item) itemCounts[item] = (itemCounts[item] || 0) + qty;
  });

  var sorted = Object.keys(itemCounts).sort(function(a, b) { return itemCounts[b] - itemCounts[a]; });

  return sorted.map(function(label) {
    return { label: label, count: itemCounts[label] };
  });
}

function buildInventoryByProperty(inventory, ownerProperties) {
  var itemCounts = {};
  inventory.forEach(function(r) {
    var item = trim(r["Item"]);
    var qty  = parseInt(r["Quantity"]) || 1;
    if (item) itemCounts[item] = (itemCounts[item] || 0) + qty;
  });
  var itemColumns = Object.keys(itemCounts).sort(function(a, b) { return itemCounts[b] - itemCounts[a]; });

  var rows = ownerProperties.map(function(prop) {
    var propRows = inventory.filter(function(r) { return trim(r["Property Name"]) === prop; });
    var row = { name: prop };
    itemColumns.forEach(function(item) {
      var qty = propRows
        .filter(function(r) { return trim(r["Item"]) === item; })
        .reduce(function(t, r) { return t + (parseInt(r["Quantity"]) || 1); }, 0);
      row[item] = qty;
    });
    return row;
  });

  return { columns: itemColumns, rows: rows };
}


// ── Issues Chart & By-Property Table ─────────────────────────
function buildIssuesChart(issues) {
  var typeCounts = {};
  issues.forEach(function(r) {
    var type = trim(r["Type"]);
    if (type) typeCounts[type] = (typeCounts[type] || 0) + 1;
  });

  var sorted = Object.keys(typeCounts).sort(function(a, b) { return typeCounts[b] - typeCounts[a]; });

  return sorted.map(function(label) {
    return { label: label, count: typeCounts[label] };
  });
}

function buildIssuesByProperty(issues, ownerProperties) {
  var typeCounts = {};
  issues.forEach(function(r) {
    var type = trim(r["Type"]);
    if (type) typeCounts[type] = (typeCounts[type] || 0) + 1;
  });
  var typeColumns = Object.keys(typeCounts).sort(function(a, b) { return typeCounts[b] - typeCounts[a]; });

  var rows = ownerProperties.map(function(prop) {
    var propIssues = issues.filter(function(r) {
      return stripPropertyId(trim(r["Property Name"] || r["Property"])) === prop;
    });
    var row = { name: prop };
    typeColumns.forEach(function(type) {
      row[type] = propIssues.filter(function(r) { return trim(r["Type"]) === type; }).length;
    });
    return row;
  });

  return { columns: typeColumns, rows: rows };
}


// ── Reviews By Property ───────────────────────────────────────
function buildReviewsByProperty(appts, ownerProperties) {
  return ownerProperties.map(function(prop) {
    var propAppts = appts.filter(function(r) { return trim(r["Property Name"]) === prop; });
    var ratings   = propAppts
      .map(function(r) { return parseFloat(r["Reviews"]); })
      .filter(function(v) { return !isNaN(v) && v > 0; });

    var avg = ratings.length
      ? Math.round((ratings.reduce(function(t, v) { return t + v; }, 0) / ratings.length) * 10) / 10
      : null;

    return {
      name        : prop,
      reviewCount : ratings.length,
      avgRating   : avg,
      display     : avg !== null
        ? avg.toFixed(1) + "\u2605 (" + ratings.length + ")"
        : "\u2014"
    };
  });
}


// ============================================================
//  UTILITY — SHEET READER
// ============================================================

function sheetToObjects(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
    return [];
  }
  var range  = sheet.getDataRange();
  var values = range.getValues();
  if (values.length < 2) return [];

  var headers = values[0].map(function(h) { return String(h).trim(); });

  return values.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      obj[h] = row[i] !== undefined ? row[i] : "";
    });
    return obj;
  }).filter(function(r) {
    return Object.values(r).some(function(v) { return v !== "" && v !== null; });
  });
}


// ============================================================
//  UTILITY — OWNERS
//  NOTE: getUniqueOwners() now returns only owners that have at
//  least one Active property in the Property Details sheet.
//  getOwnerListForSidebar() delegates here for consistency.
// ============================================================

function getUniqueOwners() {
  // Primary: derive from Property Details "Active" rows
  var activeOwners = getActiveOwnerNames();
  if (activeOwners.length > 0) return activeOwners;

  // Fallback (if Property Details sheet is missing): derive from Final DataSet
  Logger.log("⚠️  Falling back to Final DataSet for owner list (Property Details unavailable).");
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var rows  = sheetToObjects(ss, SHEET.FINAL);
  var names = rows.map(function(r) { return trim(r["Company Name"]); }).filter(Boolean);
  return unique(names);
}


// ============================================================
//  UTILITY — GENERAL HELPERS
// ============================================================

function trim(v) {
  return v !== undefined && v !== null ? String(v).trim() : "";
}

function isTruthy(v) {
  if (v === true || v === 1) return true;
  var s = trim(v).toLowerCase();
  return s === "yes" || s === "y" || s === "true" || s === "1";
}

function unique(arr) {
  var seen = {};
  return arr.filter(function(v) {
    if (seen[v]) return false;
    seen[v] = true;
    return true;
  });
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  if (typeof value === "number") {
    var d = new Date((value - 25569) * 86400 * 1000);
    return isNaN(d.getTime()) ? null : d;
  }
  var d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function dateMatchesPeriod(value, monthIndex, year) {
  if (!value) return false;
  var d = parseDate(value);
  if (!d || isNaN(d.getTime())) return false;
  var tz         = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var localMonth = parseInt(Utilities.formatDate(d, tz, "M"), 10) - 1;
  var localYear  = parseInt(Utilities.formatDate(d, tz, "yyyy"), 10);
  return localMonth === monthIndex && localYear === year;
}

function formatDate(d) {
  var dd   = String(d.getDate()).padStart(2, "0");
  var mm   = String(d.getMonth() + 1).padStart(2, "0");
  var yyyy = d.getFullYear();
  return dd + "/" + mm + "/" + yyyy;
}

function monthNameToIndex(month) {
  var months = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  var idx = months.indexOf(month);
  if (idx === -1) throw new Error("Invalid month name: " + month);
  return idx;
}

function stripPropertyId(raw) {
  return raw.replace(/^\d+\s*-\s*/, "").trim();
}

function calcAvgRating(appts) {
  var ratings = appts
    .map(function(r) { return parseFloat(r["Reviews"]); })
    .filter(function(v) { return !isNaN(v) && v > 0; });
  if (!ratings.length) return null;
  var avg = ratings.reduce(function(t, v) { return t + v; }, 0) / ratings.length;
  return Math.round(avg * 10) / 10;
}


// ============================================================
//  SIDEBAR FUNCTIONS
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Owner Report")
    .addItem("Generate Owner Report", "openOwnerSidebar")
    .addSeparator()
    .addItem("⚙️ Set OpenAI API Key", "setupOpenAIApiKey")
    .addToUi();
}

function openOwnerSidebar() {
  var tmpl = HtmlService.createTemplateFromFile("OwnerSidebar");
  tmpl.logoUrl = logo;
  tmpl.footerUrl = footer;
  var html = tmpl.evaluate()
    .setTitle("Generate Owner Report")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function previewFromOwnerSidebar(ownerName, months, year) {
  var data;
  try {
    data = getOwnerReportData(ownerName, months, parseInt(year));
  } catch (e) {
    return { error: e.message };
  }

  try {
    var template = HtmlService.createTemplateFromFile("ownerReport");
    template.ownerData = data;
    template.logoUrl   = logo;
    template.footerUrl = footer;
    template.watermarkUrl = watermarkUrl;
    var output = template.evaluate()
      .setWidth(900)
      .setHeight(1150)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    SpreadsheetApp.getUi().showModalDialog(
      output,
      "Owner Report — " + ownerName + " · " + months.join(", ") + " " + year
    );
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * getOwnerListForSidebar()
 * Called by the sidebar via google.script.run.
 * Returns only owners with at least one Active property.
 */
function getOwnerListForSidebar() {
  try {
    return getUniqueOwners();   // already filters to Active via getActiveOwnerNames()
  } catch (e) {
    Logger.log("getOwnerListForSidebar error: " + e.message);
    return [];
  }
}
