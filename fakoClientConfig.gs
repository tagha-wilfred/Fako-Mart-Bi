// ============================================================
//  fakoClientConfig.gs  —  Fako Mart BI System
//  Data Layer  |  Big Client Report
//
//  SEPARATION OF CONCERNS
//  ┌──────────────────────────────────────────────────────┐
//  │  fakoClientConfig.gs  →  data only (this file)      │
//  │  clientReport.html    →  presentation only          │
//  │  FakoSidebar.html     →  sidebar UI (shared)        │
//  └──────────────────────────────────────────────────────┘
//
//  SOURCE SHEETS
//  ┌──────────────────────────┬────────────────────────────┐
//  │ Sheet                    │ Purpose                    │
//  ├──────────────────────────┼────────────────────────────┤
//  │ BIG CLIENTS              │ Client master + status     │
//  │ BIG CLIENT ORDERS        │ Order transactions         │
//  │ MONTHLY CLIENT INVOICES  │ Invoices + payment status  │
//  │ ERRANDS                  │ Delivery/errand costs      │
//  │ OUTREACH & PROMOTIONS    │ Campaigns linked to client │
//  └──────────────────────────┴────────────────────────────┘
//
//  ENTRY POINTS
//    getClientReportData(clientId, months, year)
//    previewClientReport()
//    previewClientFromSidebar(clientId, months, year)
// ============================================================


// ── SHEET NAME CONSTANTS ──────────────────────────────────────
var FM_CLIENT_SHEET = {
  BIG_CLIENTS  : "BIG CLIENTS",
  ORDERS       : "BIG CLIENT ORDERS",
  INVOICES     : "MONTHLY CLIENT INVOICES",
  ERRANDS      : "ERRANDS",
  OUTREACH     : "OUTREACH & PROMOTIONS"
};

// ── GRADE THRESHOLDS (payment behaviour) ────────────────────
// A = Fully paid  B = Partial (≤10% balance)  C = Partial >10%  D = Unpaid / Overdue
var FM_CLIENT_PAYMENT_GRADE = { A: 0, B: 0.10, C: 0.30 };

// ── TAX RATE ─────────────────────────────────────────────────
var FM_TAX_RATE = 0.09975;   // 9.975% standard — same rate used in the Invoices sheet

// ── BRAND ASSETS (already declared in fakoConfig.gs — reuse) ─
// fmLogo and fmFooter are declared in fakoConfig.gs and available
// in the same Apps Script project; no need to redeclare them.


// ============================================================
//  CONFIG READER (shared DASHBOARD cell layout)
//  For the client report we reuse the same DASHBOARD sheet:
//    B4 = Year   C4 = Month(s)   D4 = Client ID   E4 = "Client"
// ============================================================
function fmReadClientConfig() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(FM_CONFIG_SHEET_NAME);  // from fakoConfig.gs
  if (!configSheet) configSheet = ss.getActiveSheet();

  var year      = parseInt(configSheet.getRange(FM_CONFIG_CELLS.REPORT_YEAR.row,   FM_CONFIG_CELLS.REPORT_YEAR.col).getValue());
  var monthRaw  = String(configSheet.getRange(FM_CONFIG_CELLS.REPORT_MONTH.row,  FM_CONFIG_CELLS.REPORT_MONTH.col).getValue()).trim();
  var clientId  = String(configSheet.getRange(FM_CONFIG_CELLS.EMPLOYEE_ID.row,   FM_CONFIG_CELLS.EMPLOYEE_ID.col).getValue()).trim();

  var validMonths = ["January","February","March","April","May","June",
                     "July","August","September","October","November","December"];
  var months = monthRaw.split(",").map(function(m){ return m.trim(); }).filter(Boolean);

  if (isNaN(year) || year < 2000 || year > 2100)
    throw new Error("Config error: Year (B4) must be a 4-digit number.");
  if (!months.length)
    throw new Error("Config error: Month (C4) is empty.");
  months.forEach(function(m){
    if (validMonths.indexOf(m) === -1)
      throw new Error("Config error: '" + m + "' is not a valid month name.");
  });
  if (!clientId)
    throw new Error("Config error: Client ID (D4) is empty.");

  return { clientId: clientId, months: months, year: year };
}


// ============================================================
//  ENTRY POINTS
// ============================================================

function previewClientReport() {
  var cfg;
  try { cfg = fmReadClientConfig(); }
  catch (e) {
    SpreadsheetApp.getUi().alert("Config Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var data;
  try { data = getClientReportData(cfg.clientId, cfg.months, cfg.year); }
  catch (e) {
    SpreadsheetApp.getUi().alert("Data Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var template = HtmlService.createTemplateFromFile("clientReport");
  template.reportData = data;
  template.logoUrl    = fmLogo;
  template.footerUrl  = fmFooter;

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(960).setHeight(1150).setSandboxMode(HtmlService.SandboxMode.IFRAME),
    "Client Report — " + data.profile.businessName + " · " + data.month + " " + data.year
  );
}

function previewClientFromSidebar(clientId, months, year) {
  var data;
  try { data = getClientReportData(clientId, months, parseInt(year)); }
  catch (e) { return { error: e.message }; }

  try {
    var template = HtmlService.createTemplateFromFile("clientReport");
    template.reportData = data;
    template.logoUrl    = fmLogo;
    template.footerUrl  = fmFooter;

    SpreadsheetApp.getUi().showModalDialog(
      template.evaluate().setWidth(960).setHeight(1150).setSandboxMode(HtmlService.SandboxMode.IFRAME),
      "Client Report — " + data.profile.businessName + " · " + months.join(", ") + " " + year
    );
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}


// ============================================================
//  MAIN DATA BUILDER
// ============================================================
function getClientReportData(clientId, months, year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var monthsArr    = Array.isArray(months) ? months : [months];
  var monthIndices = monthsArr.map(function(m){ return fmMonthNameToIndex(m); });  // from fakoConfig.gs
  var monthKeys    = monthIndices.map(function(mi){
    return year + "-" + String(mi + 1).padStart(2, "0");
  });

  var validMonthNames = ["January","February","March","April","May","June",
                         "July","August","September","October","November","December"];
  var sortedIdx  = monthIndices.slice().sort(function(a,b){ return a-b; });
  var monthLabel = sortedIdx.length === 1
    ? validMonthNames[sortedIdx[0]]
    : validMonthNames[sortedIdx[0]] + " – " + validMonthNames[sortedIdx[sortedIdx.length-1]];

  // ── Profiles ──────────────────────────────────────────────
  var profile = fmGetClientProfile(ss, clientId);
  if (!profile) throw new Error("Client ID '" + clientId + "' not found in BIG CLIENTS sheet.");

  // ── Orders in period ──────────────────────────────────────
  var allOrders = fmSheetToObjects(ss, FM_CLIENT_SHEET.ORDERS);  // fmSheetToObjects from fakoConfig.gs
  var orders = allOrders.filter(function(r) {
    return fmTrim(r["Client ID"]) === clientId &&
           monthIndices.some(function(mi){ return fmDateMatchesPeriod(r["Order Date"], mi, year); });
  }).map(function(r) {
    var val  = fmNum(r["Total Order Value (XAF)"]) || 0;
    var paid = fmNum(r["Amount Paid (XAF)"]) || 0;
    var bal  = val - paid;
    return {
      orderId      : fmTrim(r["Order ID"]),
      orderDate    : fmFormatDateVal(r["Order Date"]),
      deliveryDate : fmFormatDateVal(r["Delivery Date"]),
      items        : fmTrim(r["Items Ordered (list as text)"]),
      orderValue   : val,
      amountPaid   : paid,
      balance      : bal,
      payStatus    : bal <= 0 ? "Paid" : paid === 0 ? "Pending" : "Partial",
      handledBy    : fmTrim(r["Handled By (Manager ID)"])
    };
  }).sort(function(a,b){ return a.orderDate.localeCompare(b.orderDate); });

  // ── Invoices in period ────────────────────────────────────
  var allInvoices = fmSheetToObjects(ss, FM_CLIENT_SHEET.INVOICES);
  var invoices = allInvoices.filter(function(r) {
    return fmTrim(r["Client ID"]) === clientId &&
           monthKeys.indexOf(fmTrim(r["Month Covered"])) > -1;
  }).map(function(r) {
    var ordersVal = fmNum(r["Total Orders Value (XAF)"]) || 0;
    var errandVal = fmNum(r["Total Errand Costs Attributed (XAF)"]) || 0;
    var subtotal  = ordersVal + errandVal;
    var tax       = Math.round(subtotal * FM_TAX_RATE);
    var totalInv  = subtotal + tax;
    var paid      = fmNum(r["Amount Paid (XAF)"]) || 0;
    var balance   = totalInv - paid;
    return {
      invoiceId    : fmTrim(r["Invoice ID"]),
      invoiceDate  : fmFormatDateVal(r["Invoice Date"]),
      monthCovered : fmTrim(r["Month Covered"]),
      ordersValue  : ordersVal,
      errandCost   : errandVal,
      subtotal     : subtotal,
      tax          : tax,
      totalInvoice : totalInv,
      amountPaid   : paid,
      balance      : balance,
      payStatus    : balance <= 0 ? "Paid" : paid === 0 ? "Unpaid" : "Partial",
      payDate      : fmFormatDateVal(r["Payment Date"]),
      payMethod    : fmTrim(r["Payment Method (Cash/Mobile Money/Bank Transfer)"]),
      processedBy  : fmTrim(r["Processed By (Manager ID)"]),
      notes        : fmTrim(r["Notes"])
    };
  }).sort(function(a,b){ return a.monthCovered.localeCompare(b.monthCovered); });

  // ── Errands linked to this client (by Destination matching business name) ─
  var allErrands = fmSheetToObjects(ss, FM_CLIENT_SHEET.ERRANDS);
  var bizName    = profile.businessName.toLowerCase();
  var clientErrands = allErrands.filter(function(r) {
    var dest = fmTrim(r["Destination"]).toLowerCase();
    var dateOk = monthIndices.some(function(mi){ return fmDateMatchesPeriod(r["Date"], mi, year); });
    return dateOk && (dest.indexOf(bizName.split(" ")[0].toLowerCase()) > -1 ||
                      dest.indexOf(bizName.split(" ")[1] ? bizName.split(" ")[1].toLowerCase() : "_") > -1);
  }).map(function(r) {
    return {
      errandId    : fmTrim(r["Errand ID"]),
      date        : fmFormatDateVal(r["Date"]),
      type        : fmTrim(r["Errand Type (Delivery/Collection/Purchase/Bank Run/Client Visit/Other)"]),
      destination : fmTrim(r["Destination"]),
      reason      : fmTrim(r["Reason/Description"]),
      cost        : fmNum(r["Cost of Errand (XAF — transport, purchases, etc.)"]) || 0,
      status      : fmTrim(r["Status (Completed/Pending/Cancelled)"])
    };
  });

  // ── Financial Aggregates ──────────────────────────────────
  var totalOrderValue  = orders.reduce(function(t,o){ return t + o.orderValue; }, 0);
  var totalPaid        = invoices.reduce(function(t,i){ return t + i.amountPaid; }, 0);
  var totalInvoiced    = invoices.reduce(function(t,i){ return t + i.totalInvoice; }, 0);
  var totalBalance     = invoices.reduce(function(t,i){ return t + i.balance; }, 0);
  var totalTax         = invoices.reduce(function(t,i){ return t + i.tax; }, 0);
  var totalErrandCost  = invoices.reduce(function(t,i){ return t + i.errandCost; }, 0);

  // ── Payment health ────────────────────────────────────────
  var paymentRate   = totalInvoiced > 0 ? (totalPaid / totalInvoiced) : 0;
  var paymentGrade  = paymentRate >= 1 ? "A" : paymentRate >= 0.90 ? "B" : paymentRate >= 0.70 ? "C" : "D";
  var paymentLabel  = paymentGrade === "A" ? "Fully Paid" : paymentGrade === "B" ? "Near Settled" :
                      paymentGrade === "C" ? "Partially Paid" : "Overdue / Unpaid";

  // ── Category breakdown of orders ─────────────────────────
  var categorySpend = {};
  orders.forEach(function(o) {
    var cats = (profile.categories || "").split(";").map(function(c){ return c.trim(); });
    cats.forEach(function(c) {
      if (c) categorySpend[c] = (categorySpend[c] || 0) + (o.orderValue / Math.max(cats.length, 1));
    });
  });

  // ── Month-by-month summary ────────────────────────────────
  var monthSummary = monthKeys.map(function(mk) {
    var mInvoices = invoices.filter(function(i){ return i.monthCovered === mk; });
    var mOrders   = orders.filter(function(o){ return o.orderDate.substring(0,7) === mk.replace("-","-"); });
    // actually match by year-month
    var mOrdersFiltered = allOrders.filter(function(r) {
      if (fmTrim(r["Client ID"]) !== clientId) return false;
      var d = fmParseDate(r["Order Date"]);
      if (!d) return false;
      var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
      var key = Utilities.formatDate(d, tz, "yyyy-MM");
      return key === mk;
    });
    var mOrderValue = mOrdersFiltered.reduce(function(t,r){ return t + (fmNum(r["Total Order Value (XAF)"]) || 0); }, 0);
    var mPaid       = mInvoices.reduce(function(t,i){ return t + i.amountPaid; }, 0);
    var mInvoiced   = mInvoices.reduce(function(t,i){ return t + i.totalInvoice; }, 0);
    return {
      month      : mk,
      orderCount : mOrdersFiltered.length,
      orderValue : mOrderValue,
      invoiced   : mInvoiced,
      paid       : mPaid,
      balance    : mInvoiced - mPaid
    };
  });

  Logger.log("=== CLIENT REPORT DATA ===");
  Logger.log("Client: " + profile.businessName + " (" + clientId + ")");
  Logger.log("Orders: " + orders.length + " | Invoices: " + invoices.length);
  Logger.log("Total invoiced: " + totalInvoiced + " | Paid: " + totalPaid + " | Balance: " + totalBalance);
  Logger.log("Payment grade: " + paymentGrade + " (" + (paymentRate*100).toFixed(1) + "%)");

  return {
    reportDate    : fmFormatDate(new Date()),
    month         : monthLabel,
    monthsList    : monthsArr,
    year          : String(year),

    profile       : profile,

    orders        : orders,
    invoices      : invoices,
    errands       : clientErrands,

    // Financials
    totalOrderValue : totalOrderValue,
    totalInvoiced   : totalInvoiced,
    totalPaid       : totalPaid,
    totalBalance    : totalBalance,
    totalTax        : totalTax,
    totalErrandCost : totalErrandCost,

    paymentRate     : Math.round(paymentRate * 100),
    paymentGrade    : paymentGrade,
    paymentLabel    : paymentLabel,

    categorySpend   : categorySpend,
    monthSummary    : monthSummary,

    companyName     : "Fako Mart"
  };
}


// ============================================================
//  PROFILE READER
// ============================================================
function fmGetClientProfile(ss, clientId) {
  var rows = fmSheetToObjects(ss, FM_CLIENT_SHEET.BIG_CLIENTS);
  var r = rows.filter(function(row){ return fmTrim(row["Client ID"]) === clientId; })[0];
  if (!r) return null;

  return {
    clientId      : fmTrim(r["Client ID"]),
    businessName  : fmTrim(r["Business Name"]),
    contactPerson : fmTrim(r["Contact Person"]),
    phone         : fmTrim(r["Phone"]),
    email         : fmTrim(r["Email"]),
    categories    : fmTrim(r["Category of Goods Purchased"]),
    monthlyVolume : fmNum(r["Monthly Order Volume (XAF)"]) || 0,
    paymentTerms  : fmTrim(r["Payment Terms (Cash/Credit/30days)"]),
    lastOrderDate : fmFormatDateVal(r["Last Order Date"]),
    status        : fmTrim(r["Account Status"])
  };
}


// ============================================================
//  SIDEBAR SUPPORT
// ============================================================

/**
 * Returns sorted list of all active Big Clients for the sidebar.
 * Format: [{ id: "CLT001", name: "Mboko Community Hostel", type: "Client" }, ...]
 */
function getFakoClientList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return fmSheetToObjects(ss, FM_CLIENT_SHEET.BIG_CLIENTS)
      .map(function(r) {
        return {
          id       : fmTrim(r["Client ID"]),
          name     : fmTrim(r["Business Name"]),
          dept     : fmTrim(r["Category of Goods Purchased"]),
          status   : fmTrim(r["Account Status"]),
          type     : "Client"
        };
      })
      .sort(function(a, b){ return a.name.localeCompare(b.name); });
  } catch (e) {
    Logger.log("getFakoClientList error: " + e.message);
    return [];
  }
}

function debugClientData() {
  var cfg;
  try { cfg = fmReadClientConfig(); }
  catch (e) {
    SpreadsheetApp.getUi().alert("Config Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  try {
    var data = getClientReportData(cfg.clientId, cfg.months, cfg.year);
    var msg =
      "Client    : " + data.profile.businessName + "\n" +
      "Status    : " + data.profile.status + "\n" +
      "Orders    : " + data.orders.length + "\n" +
      "Invoices  : " + data.invoices.length + "\n" +
      "Invoiced  : " + data.totalInvoiced.toLocaleString() + " XAF\n" +
      "Paid      : " + data.totalPaid.toLocaleString() + " XAF\n" +
      "Balance   : " + data.totalBalance.toLocaleString() + " XAF\n" +
      "Pay Grade : " + data.paymentGrade + " (" + data.paymentRate + "%) " + data.paymentLabel;
    SpreadsheetApp.getUi().alert("Client Data OK", msg, SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(JSON.stringify(data, null, 2));
  } catch (e) {
    SpreadsheetApp.getUi().alert("Data Error", e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
