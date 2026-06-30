// ============================================================
//  fakoConfig.gs  —  Fako Mart BI System
//  Data Layer  |  Employee Performance Report
//
//  SEPARATION OF CONCERNS
//  ┌──────────────────────────────────────────────────────┐
//  │  fakoConfig.gs        →  data only (this file)      │
//  │  employeeReport.html  →  presentation only          │
//  │  FakoSidebar.html     →  sidebar UI                 │
//  └──────────────────────────────────────────────────────┘
//
//  SOURCE SHEETS (all in the same Google Spreadsheet)
//  ┌──────────────────────┬─────────────────────────────────┐
//  │ Sheet                │ Purpose                         │
//  ├──────────────────────┼─────────────────────────────────┤
//  │ WORKERS              │ Worker master list + salaries   │
//  │ MANAGERS & LEADERS   │ Leader master list + scores     │
//  │ WORKER REVIEWS       │ Monthly 5-dim review scores     │
//  │ LEADER REVIEWS       │ Monthly 6-dim review scores     │
//  │ DAILY SALES          │ Individual sales transactions   │
//  │ BONUSES              │ Bonus awards per worker/month   │
//  │ ERRANDS              │ Errand assignments              │
//  │ REPAIRS & INCIDENTS  │ Incidents reported by workers   │
//  └──────────────────────┴─────────────────────────────────┘
//
//  ENTRY POINTS
//    generateEmployeeReport(employeeId, months, year)
//    previewFromFakoSidebar(employeeId, months, year)
//
//  PDF DOWNLOAD: handled entirely client-side. The report preview
//  (employeeReport.html) includes a "Download PDF" button that calls
//  window.print() — the browser's native print-to-PDF flow renders
//  the already-styled report using @page / @media print CSS rules.
//  No server-side conversion or Drive API access is required.
// ============================================================


// ── GLOBAL CONFIG ─────────────────────────────────────────────
var FM_CONFIG_SHEET_NAME = "DASHBOARD";
var FM_CONFIG_CELLS = {
  REPORT_YEAR    : { row: 4, col: 2 },
  REPORT_MONTH   : { row: 4, col: 3 },
  EMPLOYEE_ID    : { row: 4, col: 4 },
  EMPLOYEE_TYPE  : { row: 4, col: 5 }   // "Worker" or "Leader"
};

// ── Brand assets — replace IDs with your Drive file IDs ───────
var FM_LOGO_ID   = "YOUR_LOGO_FILE_ID_HERE";
var FM_FOOTER_ID = "YOUR_FOOTER_FILE_ID_HERE";
var fmLogo   = "https://drive.google.com/thumbnail?id=" + FM_LOGO_ID   + "&sz=s1000";
var fmFooter = "https://drive.google.com/thumbnail?id=" + FM_FOOTER_ID + "&sz=s1000";

// ── Sheet name constants ───────────────────────────────────────
var FM_SHEET = {
  WORKERS          : "WORKERS",
  LEADERS          : "MANAGERS & LEADERS",
  WORKER_REVIEWS   : "WORKER REVIEWS",
  LEADER_REVIEWS   : "LEADER REVIEWS",
  DAILY_SALES      : "DAILY SALES",
  BONUSES          : "BONUSES",
  ERRANDS          : "ERRANDS",
  REPAIRS          : "REPAIRS & INCIDENTS"
};

// ── Performance grade thresholds ──────────────────────────────
var FM_GRADE_THRESHOLDS = { A: 8, B: 6, C: 4 };   // D = below 4


// ============================================================
//  CONFIG READER
// ============================================================
function fmReadConfig() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(FM_CONFIG_SHEET_NAME);

  if (!configSheet) {
    Logger.log("⚠️  Sheet '" + FM_CONFIG_SHEET_NAME + "' not found — using active sheet.");
    configSheet = ss.getActiveSheet();
  }

  var year         = parseInt(configSheet.getRange(FM_CONFIG_CELLS.REPORT_YEAR.row,   FM_CONFIG_CELLS.REPORT_YEAR.col).getValue());
  var monthRaw     = String(configSheet.getRange(FM_CONFIG_CELLS.REPORT_MONTH.row,  FM_CONFIG_CELLS.REPORT_MONTH.col).getValue()).trim();
  var employeeId   = String(configSheet.getRange(FM_CONFIG_CELLS.EMPLOYEE_ID.row,   FM_CONFIG_CELLS.EMPLOYEE_ID.col).getValue()).trim();
  var employeeType = String(configSheet.getRange(FM_CONFIG_CELLS.EMPLOYEE_TYPE.row, FM_CONFIG_CELLS.EMPLOYEE_TYPE.col).getValue()).trim();

  var validMonths  = ["January","February","March","April","May","June",
                      "July","August","September","October","November","December"];
  var months = monthRaw.split(",").map(function(m) { return m.trim(); }).filter(Boolean);

  Logger.log("── fmReadConfig() ────────────────────────────────");
  Logger.log("  Year         : " + year);
  Logger.log("  Months       : " + JSON.stringify(months));
  Logger.log("  Employee ID  : " + employeeId);
  Logger.log("  Employee Type: " + employeeType);
  Logger.log("──────────────────────────────────────────────────");

  if (isNaN(year) || year < 2000 || year > 2100) {
    throw new Error("Config error: Year (B4) must be a 4-digit number like 2025.");
  }
  if (!months.length) {
    throw new Error("Config error: Month (C4) is empty.");
  }
  months.forEach(function(m) {
    if (validMonths.indexOf(m) === -1) {
      throw new Error("Config error: '" + m + "' is not a valid month name.");
    }
  });
  if (!employeeId) {
    throw new Error("Config error: Employee ID (D4) is empty.");
  }
  if (["Worker", "Leader"].indexOf(employeeType) === -1) {
    throw new Error("Config error: Employee Type (E4) must be 'Worker' or 'Leader'.");
  }

  return { employeeId: employeeId, employeeType: employeeType, months: months, year: year };
}


// ============================================================
//  ENTRY POINTS
// ============================================================

function generateEmployeeReport(employeeId, months, year, employeeType) {
  var data     = getEmployeeReportData(employeeId, months, year, employeeType);
  var template = HtmlService.createTemplateFromFile("employeeReport");
  template.reportData = data;
  template.logoUrl    = fmLogo;
  template.footerUrl  = fmFooter;
  var html = template.evaluate().getContent();
  return html;
}

function previewEmployeeReport() {
  var cfg;
  try {
    cfg = fmReadConfig();
  } catch (e) {
    SpreadsheetApp.getUi().alert("Configuration Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var data;
  try {
    data = getEmployeeReportData(cfg.employeeId, cfg.months, cfg.year, cfg.employeeType);
  } catch (e) {
    Logger.log("ERROR in data build: " + e.message);
    SpreadsheetApp.getUi().alert("Data Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var template = HtmlService.createTemplateFromFile("employeeReport");
  template.reportData = data;
  template.logoUrl    = fmLogo;
  template.footerUrl  = fmFooter;

  var output = template.evaluate()
    .setWidth(950)
    .setHeight(1150)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showModalDialog(
    output,
    "Employee Report — " + cfg.employeeId + " · " + cfg.months.join(", ") + " " + cfg.year
  );
}

function previewFromFakoSidebar(employeeId, months, year, employeeType) {
  var data;
  try {
    data = getEmployeeReportData(employeeId, months, parseInt(year), employeeType);
  } catch (e) {
    return { error: e.message };
  }

  try {
    var template = HtmlService.createTemplateFromFile("employeeReport");
    template.reportData = data;
    template.logoUrl    = fmLogo;
    template.footerUrl  = fmFooter;

    var output = template.evaluate()
      .setWidth(950)
      .setHeight(1150)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    SpreadsheetApp.getUi().showModalDialog(
      output,
      "Employee Report — " + employeeId + " · " + months.join(", ") + " " + year
    );
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}


// ============================================================
//  MAIN DATA BUILDER
// ============================================================

function getEmployeeReportData(employeeId, months, year, employeeType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var monthsArr    = Array.isArray(months) ? months : [months];
  var monthIndices = monthsArr.map(function(m) { return fmMonthNameToIndex(m); });

  var validMonthNames = ["January","February","March","April","May","June",
                         "July","August","September","October","November","December"];
  var sortedIdx  = monthIndices.slice().sort(function(a, b) { return a - b; });
  var monthLabel = sortedIdx.length === 1
    ? validMonthNames[sortedIdx[0]]
    : validMonthNames[sortedIdx[0]] + " – " + validMonthNames[sortedIdx[sortedIdx.length - 1]];

  // ── Build YYYY-MM strings for month-based filtering ────────
  var monthKeys = monthIndices.map(function(mi) {
    return year + "-" + String(mi + 1).padStart(2, "0");
  });

  var isWorker = employeeType === "Worker";

  // ── Load master record ────────────────────────────────────
  var profile = isWorker
    ? fmGetWorkerProfile(ss, employeeId)
    : fmGetLeaderProfile(ss, employeeId);

  if (!profile) {
    throw new Error("Employee ID '" + employeeId + "' not found in " +
      (isWorker ? FM_SHEET.WORKERS : FM_SHEET.LEADERS) + " sheet.");
  }

  // ── Load reviews for this period ──────────────────────────
  var reviews = isWorker
    ? fmGetWorkerReviews(ss, employeeId, monthKeys)
    : fmGetLeaderReviews(ss, employeeId, monthKeys);

  // ── Sales (workers only — leaders don't serve directly) ───
  var salesData = isWorker
    ? fmGetWorkerSales(ss, employeeId, monthIndices, year)
    : null;

  // ── Bonuses ───────────────────────────────────────────────
  var bonuses = fmGetBonuses(ss, employeeId, monthKeys);

  // ── Errands ───────────────────────────────────────────────
  var errands = fmGetErrands(ss, employeeId, monthIndices, year);

  // ── Incidents reported by this employee ──────────────────
  var incidents = fmGetIncidents(ss, employeeId, monthIndices, year);

  // ── Aggregate review scores ───────────────────────────────
  var scoreAgg = fmAggregateScores(reviews, isWorker);

  Logger.log("=== DATA BUILD COMPLETE ===");
  Logger.log("  Profile     : " + profile.fullName + " | " + profile.department);
  Logger.log("  Reviews     : " + reviews.length);
  Logger.log("  Bonuses     : " + bonuses.length);
  Logger.log("  Errands     : " + errands.length);
  Logger.log("  Incidents   : " + incidents.length);
  if (salesData) Logger.log("  Sales txns  : " + salesData.transactions);

  return {
    // Meta
    reportDate   : fmFormatDate(new Date()),
    month        : monthLabel,
    monthsList   : monthsArr,
    year         : String(year),
    employeeType : employeeType,

    // Profile
    profile      : profile,

    // Review scores & trend
    scoreAgg     : scoreAgg,
    reviews      : reviews,

    // Sales (null for leaders)
    sales        : salesData,

    // Bonuses
    bonuses      : bonuses,
    totalBonus   : bonuses.reduce(function(t, b) { return t + (b.amount || 0); }, 0),

    // Errands
    errands      : errands,

    // Incidents
    incidents    : incidents,

    // Branding
    companyName  : "Fako Mart"
  };
}


// ============================================================
//  PROFILE READERS
// ============================================================

function fmGetWorkerProfile(ss, workerId) {
  var rows = fmSheetToObjects(ss, FM_SHEET.WORKERS);
  var r    = rows.filter(function(row) { return fmTrim(row["Worker ID"]) === workerId; })[0];
  if (!r) return null;

  return {
    id           : fmTrim(r["Worker ID"]),
    fullName     : fmTrim(r["Full Name"]),
    gender       : fmTrim(r["Gender"]),
    age          : r["Age"],
    department   : fmTrim(r["Department"]),
    position     : fmTrim(r["Position"]),
    dateHired    : fmFormatDateVal(r["Date Hired"]),
    baseSalary   : r["Base Salary (XAF)"] || 0,
    phone        : fmTrim(r["Phone"]),
    status       : fmTrim(r["Employment Status (Active/Inactive)"])
  };
}

function fmGetLeaderProfile(ss, leaderId) {
  var rows = fmSheetToObjects(ss, FM_SHEET.LEADERS);
  var r    = rows.filter(function(row) { return fmTrim(row["Manager ID"]) === leaderId; })[0];
  if (!r) return null;

  return {
    id           : fmTrim(r["Manager ID"]),
    fullName     : fmTrim(r["Full Name"]),
    gender       : fmTrim(r["Gender"]),
    age          : null,
    department   : fmTrim(r["Department Managed"]),
    position     : "Manager / Team Leader",
    dateHired    : fmFormatDateVal(r["Date Appointed"]),
    baseSalary   : r["Monthly Salary (XAF)"] || 0,
    phone        : "",
    status       : "Active",
    perfScore    : r["Performance Score (1–10)"] || null,
    reportsTo    : fmTrim(r["Reports To"])
  };
}


// ============================================================
//  REVIEWS
// ============================================================

/**
 * Returns an array of monthly review objects for a worker.
 * Computes Overall score from the 5 raw dimensions since the
 * sheet stores Excel formulas rather than resolved values.
 */
function fmGetWorkerReviews(ss, workerId, monthKeys) {
  var rows = fmSheetToObjects(ss, FM_SHEET.WORKER_REVIEWS);
  return rows
    .filter(function(r) {
      return fmTrim(r["Worker ID"]) === workerId &&
             monthKeys.indexOf(fmTrim(r["Month"])) > -1;
    })
    .map(function(r) {
      var dims = {
        punctuality   : fmNum(r["Punctuality Score (1–10)"]),
        attendance    : fmNum(r["Attendance Score (1–10)"]),
        communication : fmNum(r["Communication Score (1–10)"]),
        effectiveness : fmNum(r["Effectiveness Score (1–10)"]),
        efficiency    : fmNum(r["Efficiency Score (1–10)"])
      };
      var scores  = Object.values(dims).filter(function(v) { return v !== null; });
      var overall = scores.length ? Math.round((scores.reduce(function(t,v){ return t+v;},0)/scores.length)*10)/10 : null;
      var grade   = fmGrade(overall);
      return {
        month         : fmTrim(r["Month"]),
        reviewedBy    : fmTrim(r["Reviewed By Name"]),
        dims          : dims,
        overall       : overall,
        grade         : grade,
        strengths     : fmTrim(r["Strengths (short text)"]),
        improvements  : fmTrim(r["Areas for Improvement (short text)"]),
        status        : fmTrim(r["Review Status (Submitted/Pending)"])
      };
    })
    .sort(function(a, b) { return a.month.localeCompare(b.month); });
}

/**
 * Returns an array of monthly review objects for a leader.
 * Leaders have 6 dimensions (adds Coordination).
 */
function fmGetLeaderReviews(ss, leaderId, monthKeys) {
  var rows = fmSheetToObjects(ss, FM_SHEET.LEADER_REVIEWS);
  return rows
    .filter(function(r) {
      return fmTrim(r["Leader ID"]) === leaderId &&
             monthKeys.indexOf(fmTrim(r["Month"])) > -1;
    })
    .map(function(r) {
      var dims = {
        punctuality   : fmNum(r["Punctuality Score (1–10)"]),
        attendance    : fmNum(r["Attendance Score (1–10)"]),
        communication : fmNum(r["Communication Score (1–10)"]),
        effectiveness : fmNum(r["Effectiveness Score (1–10)"]),
        efficiency    : fmNum(r["Efficiency Score (1–10)"]),
        coordination  : fmNum(r["Coordination Score (1–10)"])
      };
      var scores  = Object.values(dims).filter(function(v) { return v !== null; });
      var overall = scores.length ? Math.round((scores.reduce(function(t,v){ return t+v;},0)/scores.length)*10)/10 : null;
      var grade   = fmGrade(overall);
      return {
        month         : fmTrim(r["Month"]),
        reviewedBy    : fmTrim(r["Reviewed By Name"]),
        dims          : dims,
        overall       : overall,
        grade         : grade,
        strengths     : fmTrim(r["Strengths (short text)"]),
        improvements  : fmTrim(r["Areas for Improvement (short text)"]),
        status        : fmTrim(r["Review Status (Submitted/Pending)"])
      };
    })
    .sort(function(a, b) { return a.month.localeCompare(b.month); });
}

/**
 * Aggregates review scores across multiple months.
 * Returns averages per dimension, best/worst month, trend direction.
 */
function fmAggregateScores(reviews, isWorker) {
  if (!reviews.length) return null;

  var dimKeys = isWorker
    ? ["punctuality","attendance","communication","effectiveness","efficiency"]
    : ["punctuality","attendance","communication","effectiveness","efficiency","coordination"];

  var dimTotals = {};
  dimKeys.forEach(function(k) { dimTotals[k] = []; });
  var overalls = [];

  reviews.forEach(function(rv) {
    dimKeys.forEach(function(k) {
      if (rv.dims[k] !== null) dimTotals[k].push(rv.dims[k]);
    });
    if (rv.overall !== null) overalls.push({ month: rv.month, score: rv.overall });
  });

  var dimAvg = {};
  dimKeys.forEach(function(k) {
    var arr = dimTotals[k];
    dimAvg[k] = arr.length ? Math.round((arr.reduce(function(t,v){return t+v;},0)/arr.length)*10)/10 : null;
  });

  var avgOverall = overalls.length
    ? Math.round((overalls.reduce(function(t,o){return t+o.score;},0)/overalls.length)*10)/10
    : null;

  var bestMonth  = overalls.length ? overalls.reduce(function(best,o){return o.score>best.score?o:best;}) : null;
  var worstMonth = overalls.length ? overalls.reduce(function(worst,o){return o.score<worst.score?o:worst;}) : null;

  // Trend: compare last month to first month overall score
  var trend = "stable";
  if (overalls.length >= 2) {
    var diff = overalls[overalls.length-1].score - overalls[0].score;
    if (diff >= 0.3) trend = "improving";
    else if (diff <= -0.3) trend = "declining";
  }

  return {
    dimAvg      : dimAvg,
    dimKeys     : dimKeys,
    avgOverall  : avgOverall,
    avgGrade    : fmGrade(avgOverall),
    bestMonth   : bestMonth,
    worstMonth  : worstMonth,
    trend       : trend,
    monthCount  : reviews.length
  };
}


// ============================================================
//  SALES DATA (Workers only)
// ============================================================

function fmGetWorkerSales(ss, workerId, monthIndices, year) {
  var rows = fmSheetToObjects(ss, FM_SHEET.DAILY_SALES);
  var myRows = rows.filter(function(r) {
    var wid     = fmTrim(r["Served By (Worker ID)"]);
    var dateVal = r["Date"];
    if (wid !== workerId) return false;
    return monthIndices.some(function(mi) { return fmDateMatchesPeriod(dateVal, mi, year); });
  });

  if (!myRows.length) return { transactions: 0, totalRevenue: 0, avgPerSale: 0, byCategory: {}, byMonth: {} };

  var totalRev = 0;
  var byCategory = {};
  var byMonth    = {};

  myRows.forEach(function(r) {
    var qty   = fmNum(r["Quantity Sold"]) || 0;
    var price = fmNum(r["Unit Price (XAF)"]) || 0;
    var disc  = fmNum(r["Discount Amount (XAF)"]) || 0;
    var total = qty * price - disc;
    totalRev += total;

    var cat = fmTrim(r["Category"]) || "Other";
    byCategory[cat] = (byCategory[cat] || 0) + total;

    var d   = fmParseDate(r["Date"]);
    var key = d ? (d.getFullYear() + "-" + String(d.getMonth()+1).padStart(2,"0")) : "Unknown";
    if (!byMonth[key]) byMonth[key] = { count: 0, revenue: 0 };
    byMonth[key].count++;
    byMonth[key].revenue += total;
  });

  return {
    transactions : myRows.length,
    totalRevenue : Math.round(totalRev),
    avgPerSale   : myRows.length ? Math.round(totalRev / myRows.length) : 0,
    byCategory   : byCategory,
    byMonth      : byMonth
  };
}


// ============================================================
//  BONUSES
// ============================================================

function fmGetBonuses(ss, employeeId, monthKeys) {
  var rows = fmSheetToObjects(ss, FM_SHEET.BONUSES);
  return rows
    .filter(function(r) {
      var wid = fmTrim(r["Worker ID"]);
      var mon = fmTrim(r["Month Awarded"]);
      return wid === employeeId && monthKeys.indexOf(mon) > -1;
    })
    .map(function(r) {
      return {
        type     : fmTrim(r["Bonus Type (Performance/Holiday/Sales Target)"]),
        amount   : fmNum(r["Amount (XAF)"]) || 0,
        month    : fmTrim(r["Month Awarded"]),
        reason   : fmTrim(r["Reason"]),
        approvedBy: fmTrim(r["Approved By"])
      };
    })
    .sort(function(a, b) { return a.month.localeCompare(b.month); });
}


// ============================================================
//  ERRANDS
// ============================================================

function fmGetErrands(ss, employeeId, monthIndices, year) {
  var rows = fmSheetToObjects(ss, FM_SHEET.ERRANDS);
  return rows
    .filter(function(r) {
      var person  = fmTrim(r["Person Sent (Worker ID or Leader ID)"]);
      var dateVal = r["Date"];
      if (person !== employeeId) return false;
      return monthIndices.some(function(mi) { return fmDateMatchesPeriod(dateVal, mi, year); });
    })
    .map(function(r) {
      return {
        errandId    : fmTrim(r["Errand ID"]),
        date        : fmFormatDateVal(r["Date"]),
        type        : fmTrim(r["Errand Type (Delivery/Collection/Purchase/Bank Run/Client Visit/Other)"]),
        destination : fmTrim(r["Destination"]),
        reason      : fmTrim(r["Reason/Description"]),
        cost        : fmNum(r["Cost of Errand (XAF — transport, purchases, etc.)"]) || 0,
        status      : fmTrim(r["Status (Completed/Pending/Cancelled)"])
      };
    })
    .sort(function(a, b) { return a.date.localeCompare(b.date); });
}


// ============================================================
//  INCIDENTS REPORTED
// ============================================================

function fmGetIncidents(ss, employeeId, monthIndices, year) {
  var rows = fmSheetToObjects(ss, FM_SHEET.REPAIRS);
  return rows
    .filter(function(r) {
      var reporter = fmTrim(r["Reported By (Worker ID)"]);
      var dateVal  = r["Date Reported"];
      if (reporter !== employeeId) return false;
      return monthIndices.some(function(mi) { return fmDateMatchesPeriod(dateVal, mi, year); });
    })
    .map(function(r) {
      return {
        repairId    : fmTrim(r["Repair ID"]),
        date        : fmFormatDateVal(r["Date Reported"]),
        item        : fmTrim(r["Item/Equipment Affected (e.g. Freezer Unit 2, Generator, CCTV)"]),
        description : fmTrim(r["Issue Description"]),
        severity    : fmTrim(r["Severity (Low/Medium/High)"]),
        cost        : fmNum(r["Repair Cost (XAF)"]) || 0,
        status      : fmTrim(r["Status (Resolved/Pending/Ongoing)"])
      };
    })
    .sort(function(a, b) { return a.date.localeCompare(b.date); });
}


// ============================================================
//  SIDEBAR SUPPORT
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Fako Mart Reports")
    .addItem("Employee Performance Report", "openFakoSidebar")
    .addSeparator()
    .addItem("Preview (from Config Sheet)", "previewEmployeeReport")
    .addToUi();
}

function openFakoSidebar() {
  var tmpl     = HtmlService.createTemplateFromFile("sidebar");
  tmpl.logoUrl = fmLogo;
  var html = tmpl.evaluate()
    .setTitle("Employee Performance Report")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns sorted list of all active workers for the sidebar dropdown.
 * Format: [{ id: "WKR001", name: "Nkemdi Acha", type: "Worker" }, ...]
 */
function getFakoEmployeeList() {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var workers = fmSheetToObjects(ss, FM_SHEET.WORKERS)
      .filter(function(r) {
        return fmTrim(r["Employment Status (Active/Inactive)"]).toLowerCase() === "active";
      })
      .map(function(r) {
        return {
          id   : fmTrim(r["Worker ID"]),
          name : fmTrim(r["Full Name"]),
          dept : fmTrim(r["Department"]),
          type : "Worker"
        };
      });

    var leaders = fmSheetToObjects(ss, FM_SHEET.LEADERS)
      .map(function(r) {
        return {
          id   : fmTrim(r["Manager ID"]),
          name : fmTrim(r["Full Name"]),
          dept : fmTrim(r["Department Managed"]),
          type : "Leader"
        };
      });

    return workers.concat(leaders).sort(function(a, b) { return a.name.localeCompare(b.name); });
  } catch (e) {
    Logger.log("getFakoEmployeeList error: " + e.message);
    return [];
  }
}

function debugFakoData() {
  var cfg;
  try { cfg = fmReadConfig(); }
  catch (e) {
    SpreadsheetApp.getUi().alert("Config Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  try {
    var data = getEmployeeReportData(cfg.employeeId, cfg.months, cfg.year, cfg.employeeType);
    var msg =
      "Employee  : " + data.profile.fullName + "\n" +
      "Dept      : " + data.profile.department + "\n" +
      "Reviews   : " + data.reviews.length + "\n" +
      "Avg Score : " + (data.scoreAgg ? data.scoreAgg.avgOverall : "N/A") + "\n" +
      "Bonuses   : " + data.bonuses.length + " | Total: " + data.totalBonus.toLocaleString() + " XAF\n" +
      "Errands   : " + data.errands.length + "\n" +
      "Incidents : " + data.incidents.length + "\n" +
      (data.sales ? "Sales Txns: " + data.sales.transactions + " | Revenue: " + data.sales.totalRevenue.toLocaleString() + " XAF" : "");
    SpreadsheetApp.getUi().alert("Data OK", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(JSON.stringify(data, null, 2));
  } catch (e) {
    SpreadsheetApp.getUi().alert("Data Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


// ============================================================
//  UTILITY — SHEET READER
// ============================================================

function fmSheetToObjects(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
    return [];
  }
  var values  = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function(h) { return String(h).trim(); });
  return values.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i] !== undefined ? row[i] : ""; });
    return obj;
  }).filter(function(r) {
    return Object.values(r).some(function(v) { return v !== "" && v !== null; });
  });
}


// ============================================================
//  UTILITY — HELPERS
// ============================================================

function fmTrim(v) {
  return v !== undefined && v !== null ? String(v).trim() : "";
}

function fmNum(v) {
  if (v === null || v === undefined || v === "") return null;
  var n = parseFloat(v);
  return isNaN(n) ? null : n;
}

function fmGrade(score) {
  if (score === null || score === undefined) return "—";
  if (score >= FM_GRADE_THRESHOLDS.A) return "A";
  if (score >= FM_GRADE_THRESHOLDS.B) return "B";
  if (score >= FM_GRADE_THRESHOLDS.C) return "C";
  return "D";
}

function fmParseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  if (typeof value === "number") {
    var d = new Date((value - 25569) * 86400 * 1000);
    return isNaN(d.getTime()) ? null : d;
  }
  var d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function fmDateMatchesPeriod(value, monthIndex, year) {
  if (!value) return false;
  var d = fmParseDate(value);
  if (!d) return false;
  var tz         = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var localMonth = parseInt(Utilities.formatDate(d, tz, "M"), 10) - 1;
  var localYear  = parseInt(Utilities.formatDate(d, tz, "yyyy"), 10);
  return localMonth === monthIndex && localYear === year;
}

function fmFormatDate(d) {
  var dd   = String(d.getDate()).padStart(2, "0");
  var mm   = String(d.getMonth() + 1).padStart(2, "0");
  var yyyy = d.getFullYear();
  return dd + "/" + mm + "/" + yyyy;
}

function fmFormatDateVal(value) {
  var d = fmParseDate(value);
  return d ? fmFormatDate(d) : "—";
}

function fmMonthNameToIndex(month) {
  var months = ["January","February","March","April","May","June",
                "July","August","September","October","November","December"];
  var idx = months.indexOf(month);
  if (idx === -1) throw new Error("Invalid month name: " + month);
  return idx;
}
