// ============================================================
// FILE 1 - MAIN CRM SCRIPT - code.gs
// ============================================================

function onEdit(e) {
  var sheet = e.range.getSheet();
  var shName = sheet.getName();
  var row = e.range.getRow();
  if (row === 1) return;

 // ── AD_LEADS_SYNC AUTO PENDING ──
  if (shName === 'Ad_Leads_Sync') {
    var adHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var adStatusCol = adHeaders.indexOf('Status') + 1;
    var adNameCol = adHeaders.indexOf('Name') + 1;
    if (sheet.getRange(row, adNameCol).getValue() && !sheet.getRange(row, adStatusCol).getValue()) {
      sheet.getRange(row, adStatusCol).setValue('Pending');
    }
    return;
  }

  // ── AUTO-COLOR LOOKUPS ON EDIT ──
  if (shName === 'Lookups') {
    if (row > 1) {
      var role = sheet.getRange(row, 4).getValue();
      var range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
      if (role === 'Admin') range.setBackground('#c9daf8');
      else if (role === 'Team Leader') range.setBackground('#d9d2e9');
      else if (role === 'BDA') range.setBackground('#d9ead3');
      else range.setBackground(null);
    }
    return;
  }

  if (shName !== 'Leads_Master') return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var col = e.range.getColumn();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var userEmail = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail();
  var role = getUserRole(userEmail);
  var ui = SpreadsheetApp.getUi();

  var leadIDCol        = headers.indexOf('LeadID') + 1;
  var lastDateCol      = headers.indexOf('Last Date') + 1;
  var leadLockCol      = headers.indexOf('Lead Lock') + 1;
  var assignedEmailCol = headers.indexOf('Assigned Email') + 1;
  var callStatusCol    = headers.indexOf('Call Status') + 1;
  var unqualReasonCol  = headers.indexOf('Unqualified Reason') + 1;
  var pipelineCol      = headers.indexOf('Pipeline & Stage') + 1;
  var bdaNameCol       = headers.indexOf('BDA Name') + 1;
  var teamLeaderCol    = headers.indexOf('Team Leader') + 1;
  var assignedToCol    = headers.indexOf('Assigned To') + 1;
  var mobileCol        = headers.indexOf('Mobile') + 1;

  // ── DUPLICATE MOBILE BLOCKER ──
  if (col == mobileCol) {
    var newMobile = e.range.getValue().toString().trim();
    if (newMobile !== '') {
      var allMobiles = sheet.getRange(2, mobileCol, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < allMobiles.length; i++) {
        var existing = allMobiles[i][0].toString().trim();
        if ((i + 2) !== row && existing === newMobile) {
          ui.alert('⚠️ Duplicate Mobile Number! This lead already exists.');
          e.range.setValue(e.oldValue || '');
          return;
        }
      }
    }
  }

  // ── MANUAL ASSIGNMENT FROM Assigned To COLUMN ──
  if (col == assignedToCol) {
    var lookupSheet = ss.getSheetByName('Lookups');
    var empData = lookupSheet.getDataRange().getValues();
    var startRow = e.range.getRow();
    var numRows = e.range.getNumRows();

    for (var r = 0; r < numRows; r++) {
      var currentRow = startRow + r;
      var assignedName = sheet.getRange(currentRow, assignedToCol).getValue();
      if (!assignedName) continue;

      for (var i = 1; i < empData.length; i++) {
        if (empData[i][0] === assignedName) {
          var empEmail = empData[i][1];
          var managerEmail = empData[i][4];

          sheet.getRange(currentRow, assignedEmailCol).setValue(empEmail);
          sheet.getRange(currentRow, bdaNameCol).setValue(assignedName);

          for (var j = 1; j < empData.length; j++) {
           if (empData[j][1] === managerEmail || empData[j][0] === managerEmail) {
              sheet.getRange(currentRow, teamLeaderCol).setValue(empData[j][0]);
              break;
            }
          }

          var currentCounter = lookupSheet.getRange(i + 1, 3).getValue() || 0;
          lookupSheet.getRange(i + 1, 3).setValue(Number(currentCounter) + 1);
          // Write Assign Date
          var assignDateColCode = headers.indexOf('Assign Date') + 1;
          if (assignDateColCode > 0) {
            sheet.getRange(currentRow, assignDateColCode).setValue(new Date());
          }
          // Queue notification
          queueLeadNotification(empEmail, sheet.getRange(currentRow, leadIDCol).getValue());
          break;
        }
      }
    }
  }

  var editableCols = [
  'Call Status', 'Unqualified Reason', 'Remarks', 'Pipeline & Stage', 'Follow-up Date',
  'Name', 'Alternate Phone', 'Service Required'
].map(function(f) { return headers.indexOf(f) + 1; });

  // ── AUTO-ASSIGN LEADID ──
  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var leadID = data[leadIDCol - 1];
  if (!leadID) {
    leadID = "LD-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
    sheet.getRange(row, leadIDCol).setValue(leadID);
  }

  // ── AUTO-UPDATE LAST DATE ──
  sheet.getRange(row, lastDateCol).setValue(new Date());

  // ── LEAD LOCK CHECK ──
  Logger.log('LOCK CHECK — role: ' + role + ' | userEmail: ' + userEmail + ' | lockVal: ' + sheet.getRange(row, leadLockCol).getValue());
  if (sheet.getRange(row, leadLockCol).getValue() == 'Y' && role !== 'Admin') {
    ui.alert("🔒 Lead is locked. Only Admins can edit.");
    e.range.setValue(e.oldValue || '');
    return;
  }

  // ── BDA PERMISSIONS ──
  if (role == 'BDA') {
    var assignedEmail = sheet.getRange(row, assignedEmailCol).getValue();
    if (userEmail && assignedEmail != userEmail) {
      ui.alert("⛔ You can only view/edit your assigned leads.");
      e.range.setValue(e.oldValue || '');
      return;
    }
    if (editableCols.indexOf(col) === -1) {
      ui.alert("⛔ You can only edit: Call Status, Unqualified Reason, Remarks, Pipeline & Stage, Follow-up Date");
      e.range.setValue(e.oldValue || '');
      return;
    }
  }

  // ── TEAM LEADER PERMISSIONS ──
  if (role == 'Team Leader') {
    var allowedForLeader = ['BDA Name', 'Assigned To', 'Assigned Email', 'Call Status', 'Remarks', 'Pipeline & Stage', 'Follow-up Date'];
    var leaderCols = allowedForLeader.map(function(f) { return headers.indexOf(f) + 1; });
    if (leaderCols.indexOf(col) === -1) {
      ui.alert("⛔ Team Leaders can only assign BDAs and update lead status.");
      e.range.setValue(e.oldValue || '');
      return;
    }
  }

  // ── AUTO-FILL TEAM LEADER WHEN BDA NAME IS ENTERED ──
  if (col == bdaNameCol && e.range.getValue()) {
    var bdaVal = e.range.getValue();
    var lookupSheet2 = ss.getSheetByName('Lookups');
    var empData2 = lookupSheet2.getDataRange().getValues();
    for (var i = 1; i < empData2.length; i++) {
      if (empData2[i][0] === bdaVal) {
        var managerEmail2 = empData2[i][4];
        sheet.getRange(row, assignedEmailCol).setValue(empData2[i][1]);
        for (var j = 1; j < empData2.length; j++) {
          if (empData2[j][1] === managerEmail2) {
            sheet.getRange(row, teamLeaderCol).setValue(empData2[j][0]);
            break;
          }
        }
        break;
      }
    }
  }

  // ── CLEAR UNQUALIFIED REASON IF CALL STATUS CHANGES ──
  if (col == callStatusCol && e.range.getValue() !== 'Unqualified Lead') {
    sheet.getRange(row, unqualReasonCol).setValue('');
  }

  // ── EMAIL ALERT FOR KEY PIPELINE STAGES ──
  if (col == pipelineCol) {
    var pipelineVal = e.range.getValue();
    if (pipelineVal == "Won" || pipelineVal == "Unqualified" || pipelineVal.indexOf("Project (80%)") >= 0) {
      sendAdminEmail(leadID, pipelineVal);
    }
  }

  // ── AUTO COLOR + SORT ON STATUS CHANGE ──
  if (col == callStatusCol || col == pipelineCol) {
    applyLeadColors();
    sortLeadsByStatus();
  }
  // ── UPDATE LEAD SCORE ON EDIT ──
var scoringCols = [callStatusCol, pipelineCol, 
  headers.indexOf('Remarks') + 1,
  headers.indexOf('Follow-up Date') + 1,
  headers.indexOf('Service Required') + 1
];
if (scoringCols.indexOf(col) !== -1) {
  var scoreResult = calculateLeadScore(
    sheet.getRange(row, headers.indexOf('Remarks') + 1).getValue(),
    sheet.getRange(row, headers.indexOf('Pipeline & Stage') + 1).getValue(),
    sheet.getRange(row, headers.indexOf('Call Status') + 1).getValue(),
    sheet.getRange(row, headers.indexOf('Follow-up Date') + 1).getValue(),
    sheet.getRange(row, headers.indexOf('Last Date') + 1).getValue(),
    sheet.getRange(row, headers.indexOf('Service Required') + 1).getValue()
  );
  sheet.getRange(row, headers.indexOf('Lead Score') + 1).setValue(scoreResult.score);
  sheet.getRange(row, headers.indexOf('Lead Label') + 1).setValue(scoreResult.label);
}
}

// ── ROLE DETECTION ──
function getUserRole(email) {
  if (email === 'developernotifytechai@gmail.com' || email === 'saniyakhan1709@gmail.com') return 'Admin';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === email) return data[i][3];
    }
  } catch(e) {
    Logger.log('getUserRole error: ' + e.message);
  }
  return null;
}

// ── EMAIL ALERT ──
function sendAdminEmail(leadID, status) {
  MailApp.sendEmail({
    to: 'developernotifytechai@gmail.com',
    subject: 'Lead Status Update: ' + leadID,
    htmlBody: 'Lead <b>' + leadID + '</b> has been updated.<br><br>Pipeline & Stage: <b>' + status + '</b>'
  });
}

// ── CRM MENU ──
function onOpen() {
  SpreadsheetApp.getUi().createMenu('CRM')
    .addItem('Filter My Leads', 'filterAssignedLeads')
    .addItem('🧹 Clean Leads Master', 'cleanLeadsMaster')
    .addItem('⚡ Optimise for PowerBI', 'optimiseForPowerBI')
    .addItem('🔄 Sync Targets from Lookups', 'syncTargetsFromLookups')
    .addItem('🎨 Apply Colors + Sort', 'sortLeadsByStatus')
    .addItem('🔧 Fix Bulk Leads', 'fixBulkLeads')
    .addToUi();

  syncTargetsFromLookups();
}

// ── FILTER MY LEADS ──
function filterAssignedLeads() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads_Master');
  var email = Session.getActiveUser().getEmail();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var assignedToCol = headers.indexOf('Assigned To') + 1;
  sheet.getFilter() && sheet.getFilter().remove();
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
    .createFilter()
    .setColumnFilterCriteria(
      assignedToCol,
      SpreadsheetApp.newFilterCriteria().whenTextEqualTo(email).build()
    );
}

// ── AUTO CLEAN HELPER ──
function cleanText(val) {
  if (!val) return '';
  return val.toString().trim().replace(/\w\S*/g, function(word) {
    return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase();
  });
}

function cleanMobile(val) {
  if (!val) return '';
  return val.toString().replace(/[\s\-\(\)\+\.]/g, '');
}

// ── AD LEADS SYNC ──
function ImportAdLeadsSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName('Ad_Leads_Sync');
  const dst = ss.getSheetByName('Leads_Master');
  if (!src || !dst) return;

  const data = src.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];
  const statusCol = headers.indexOf('Status');

  for (let i = 1; i < data.length; i++) {
    if (data[i][statusCol] === 'Pending') {
      const leadId = 'LD-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
      const cleanedName     = cleanText(data[i][2]);
      const cleanedMobile   = cleanMobile(data[i][3]);
      const cleanedCompany  = cleanText(data[i][5]);
      const cleanedLocation = cleanText(data[i][8]);

      dst.appendRow([
        leadId, data[i][1], data[i][0],
        cleanedName, cleanedMobile, data[i][4],
        cleanedCompany, cleanedLocation, data[i][6], data[i][7],
        '', '', '', '', 'N',
        '', '', '', '', '', '',
        new Date()
      ]);
      autoAssignLead(dst.getLastRow());
     src.getRange(i + 1, statusCol + 1).setValue('Assigned');
    }
  }
  applyLeadColors();
}

// ── AUTO ASSIGN ──
function autoAssignLead(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var leadsSheet = ss.getSheetByName('Leads_Master');
  var lookupSheet = ss.getSheetByName('Lookups');

  var headers = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];
  var assignedToCol    = headers.indexOf('Assigned To') + 1;
  var assignedEmailCol = headers.indexOf('Assigned Email') + 1;
  var bdaNameCol       = headers.indexOf('BDA Name') + 1;
  var teamLeaderCol    = headers.indexOf('Team Leader') + 1;

  var empData = lookupSheet.getDataRange().getValues();
  var bdaOnly = empData.filter(function(r) { return r[3] === 'BDA'; });

  var minVal = Infinity, chosenIndex = -1;
  for (var i = 0; i < bdaOnly.length; i++) {
    if (bdaOnly[i][2] < minVal) { minVal = bdaOnly[i][2]; chosenIndex = i; }
  }
  if (chosenIndex === -1) return;

  var assignedName  = bdaOnly[chosenIndex][0];
  var assignedEmail = bdaOnly[chosenIndex][1];
  var managerEmail  = bdaOnly[chosenIndex][4];

  leadsSheet.getRange(row, assignedToCol).setValue(assignedName);
  leadsSheet.getRange(row, assignedEmailCol).setValue(assignedEmail);
  leadsSheet.getRange(row, bdaNameCol).setValue(assignedName);

  for (var j = 1; j < empData.length; j++) {
    if (empData[j][1] === managerEmail) {
      leadsSheet.getRange(row, teamLeaderCol).setValue(empData[j][0]);
      break;
    }
  }

  for (var k = 1; k < empData.length; k++) {
    if (empData[k][1] === assignedEmail) {
      lookupSheet.getRange(k + 1, 3).setValue(minVal + 1);
      break;
    }
  }

  // Write Assign Date
  var autoHeaders = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];
  var autoAssignDateCol = autoHeaders.indexOf('Assign Date') + 1;
  if (autoAssignDateCol > 0) {
    leadsSheet.getRange(row, autoAssignDateCol).setValue(new Date());
  }
}

// ── GENERATE MISSING LEAD IDS ──
function generateMissingLeadIDs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads_Master');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var leadIDCol = headers.indexOf('LeadID');
  var mobileCol = headers.indexOf('Mobile');

  for (var i = 1; i < data.length; i++) {
    if (!data[i][leadIDCol] && data[i][mobileCol]) {
      var newID = "LD-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss") + "-" + i;
      sheet.getRange(i + 1, leadIDCol + 1).setValue(newID);
    }
  }
}

// ── CLEAN LEADS MASTER ──
function cleanLeadsMaster() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads_Master');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var leadLockCol = headers.indexOf('Lead Lock');
  var dateCol = headers.indexOf('Date');
  var lastDateCol = headers.indexOf('Last Date');

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (!data[i][leadLockCol]) {
      sheet.getRange(i + 1, leadLockCol + 1).setValue('N');
    }
   if (!data[i][dateCol]) {
      var fallback = new Date(data[i][lastDateCol] || new Date());
      sheet.getRange(i + 1, dateCol + 1).setValue(new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate()));
    }
  }
  Logger.log('✅ Leads_Master cleaned successfully!');
}

// ── OPTIMISE FOR POWER BI ──
function optimiseForPowerBI() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads_Master');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var dateCol       = headers.indexOf('Date');
  var lastDateCol   = headers.indexOf('Last Date');
  var teamLeaderCol = headers.indexOf('Team Leader');
  var bdaNameCol    = headers.indexOf('BDA Name');
  var sourceCol     = headers.indexOf('Source');
  var pipelineCol   = headers.indexOf('Pipeline & Stage');
  var callStatusCol = headers.indexOf('Call Status');

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][dateCol]) {
      var d = new Date(data[i][dateCol]);
      if (!isNaN(d)) sheet.getRange(i + 1, dateCol + 1).setValue(new Date(d.getFullYear(), d.getMonth(), d.getDate()));
    }
    if (data[i][lastDateCol]) {
      var ld = new Date(data[i][lastDateCol]);
      if (!isNaN(ld)) sheet.getRange(i + 1, lastDateCol + 1).setValue(new Date(ld.getFullYear(), ld.getMonth(), ld.getDate()));
    }
    [teamLeaderCol, bdaNameCol, sourceCol].forEach(function(col) {
      if (col > -1 && data[i][col]) {
        var trimmed = data[i][col].toString().trim();
        if (trimmed !== data[i][col].toString()) sheet.getRange(i + 1, col + 1).setValue(trimmed);
      }
    });
    if (!data[i][pipelineCol]) sheet.getRange(i + 1, pipelineCol + 1).setValue('New Lead');
    if (!data[i][callStatusCol]) sheet.getRange(i + 1, callStatusCol + 1).setValue('Not Called');
  }
  SpreadsheetApp.getUi().alert('✅ Sheet optimised for PowerBI!');
}

// ── STALE LEAD REMINDER ──
function checkStaleLeads() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Leads_Master');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var callStatusCol    = headers.indexOf('Call Status');
  var lastDateCol      = headers.indexOf('Last Date');
  var createdDateCol = headers.indexOf('Date');
  var assignedEmailCol = headers.indexOf('Assigned Email');
  var assignedToCol    = headers.indexOf('Assigned To');
  var nameCol          = headers.indexOf('Name');
  var mobileCol        = headers.indexOf('Mobile');
  var serviceCol       = headers.indexOf('Service Required');
  var leadIDCol        = headers.indexOf('LeadID');

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var threeDaysAgo = new Date(today);
  threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);

  var staleByBDA = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var callStatus = data[i][callStatusCol];
    var lastDate   = new Date(data[i][lastDateCol]);
    var bdaEmail   = data[i][assignedEmailCol];
    var bdaName    = data[i][assignedToCol];

    if (callStatus === 'Not Called' && lastDate <= threeDaysAgo && bdaEmail) {
      if (!staleByBDA[bdaEmail]) staleByBDA[bdaEmail] = { name: bdaName, leads: [] };
      staleByBDA[bdaEmail].leads.push({
        leadID: data[i][leadIDCol], name: data[i][nameCol],
        mobile: data[i][mobileCol], service: data[i][serviceCol]
      });
    }
  }

  for (var email in staleByBDA) {
    var bda = staleByBDA[email];
    var rows = bda.leads.map(function(l) {
      return '<tr><td style="padding:8px;border:1px solid #ddd">' + l.leadID + '</td>'
           + '<td style="padding:8px;border:1px solid #ddd">' + l.name + '</td>'
           + '<td style="padding:8px;border:1px solid #ddd">' + l.mobile + '</td>'
           + '<td style="padding:8px;border:1px solid #ddd">' + l.service + '</td></tr>';
    }).join('');

    var html = '<p>Hi ' + bda.name + ',</p>'
      + '<p>You have <b>' + bda.leads.length + ' lead(s)</b> with no activity in the last 3 days. Please follow up today!</p>'
      + '<table style="border-collapse:collapse;width:100%;font-family:Arial;font-size:13px">'
      + '<thead><tr style="background:#1a73e8;color:white">'
      + '<th style="padding:8px">Lead ID</th><th style="padding:8px">Name</th>'
      + '<th style="padding:8px">Mobile</th><th style="padding:8px">Service</th>'
      + '</tr></thead><tbody>' + rows + '</tbody></table>'
      + '<br><p style="color:#666;font-size:12px">— NotifyTechAI CRM</p>';

    MailApp.sendEmail({
      to: email,
      subject: '⚠️ ' + bda.leads.length + ' Stale Lead(s) Need Your Attention — NotifyTechAI',
      htmlBody: html
    });
  }
}

// ── AUTO SYNC TARGETS FROM LOOKUPS ──
function syncTargetsFromLookups() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var lookupsSheet = ss.getSheetByName('Lookups');
  var targetsSheet = ss.getSheetByName('Targets');
  var lookupsData = lookupsSheet.getDataRange().getValues();
  var targetsData = targetsSheet.getDataRange().getValues();

  var now = new Date();
  var currentMonth = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');

  var existingKeys = {};
  for (var i = 1; i < targetsData.length; i++) {
    if (!targetsData[i][1]) continue;
    var rowEmail = String(targetsData[i][1]).trim();
    var rowMonth = String(targetsData[i][4]).trim();
    if (rowMonth.length !== 7 && targetsData[i][4]) {
      try { rowMonth = Utilities.formatDate(new Date(targetsData[i][4]), Session.getScriptTimeZone(), 'yyyy-MM'); } catch(e) {}
    }
    existingKeys[rowEmail + '|' + rowMonth] = true;
  }

  for (var i = 1; i < lookupsData.length; i++) {
    var name  = String(lookupsData[i][0] || '').trim();
    var email = String(lookupsData[i][1] || '').trim();
    var role  = String(lookupsData[i][3] || '').trim();
    if (!name || !email) continue;
    if (role !== 'BDA' && role !== 'Team Leader') continue;
    var key = email + '|' + currentMonth;
    if (existingKeys[key]) continue;
    targetsSheet.appendRow([name, email, 0, 0, currentMonth, role]);
    existingKeys[key] = true;
  }
    colorTargetsByMonth();
}

// ── COLOR CODING ──
function applyLeadColors() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Leads_Master');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  sheet.clearConditionalFormatRules();

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var callStatusCol = headers.indexOf('Call Status') + 1;
  var pipelineCol = headers.indexOf('Pipeline & Stage') + 1;
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var colorMap = {
    'Won':                '#b7e1cd',
    'Interested':         '#b6d7a8',
    'Connected':          '#d9ead3',
    'Called - Callback':  '#fce5cd',
    'Called - No Answer': '#fff2cc',
    'Not Called':         '#c9daf8',
    'Unqualified Lead':   '#f4cccc'
  };

  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) continue;
    var rowNum = i + 2;
    var callStatus = data[i][callStatusCol - 1];
    var pipeline = data[i][pipelineCol - 1];
    var colorKey = pipeline === 'Won' ? 'Won' : callStatus;
    var bg = colorMap[colorKey] || '#ffffff';
    sheet.getRange(rowNum, 1, 1, lastCol).setBackground(bg);
  }
}

// ── SORT LEADS BY CALL STATUS ──
function sortLeadsByStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Leads_Master');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 3) return;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var callStatusCol = headers.indexOf('Call Status') + 1;
  var pipelineCol = headers.indexOf('Pipeline & Stage') + 1;

  var sortOrder = {
    'Interested':         1,
    'Connected':          2,
    'Called - Callback':  3,
    'Called - No Answer': 4,
    'Not Called':         5,
    'Unqualified Lead':   6,
    'Won':                7
  };

  var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  var data = dataRange.getValues();

  data.sort(function(a, b) {
    var aKey = a[pipelineCol - 1] === 'Won' ? 'Won' : a[callStatusCol - 1];
    var bKey = b[pipelineCol - 1] === 'Won' ? 'Won' : b[callStatusCol - 1];
    return (sortOrder[aKey] || 5) - (sortOrder[bKey] || 5);
  });

  dataRange.setValues(data);
  applyLeadColors();
}
function queueLeadNotification(bdaEmail, leadID) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var qSheet = ss.getSheetByName('Notification_Queue');
    if (!qSheet) return;
    qSheet.appendRow([bdaEmail, leadID, new Date(), 'N']);
  } catch(e) {
    Logger.log('queueLeadNotification error: ' + e.message);
  }
}
// ── PROCESS NOTIFICATION QUEUE + UNTOUCHED LEAD ALERTS ──
function processNotifications() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var qSheet = ss.getSheetByName('Notification_Queue');
  var leadsSheet = ss.getSheetByName('Leads_Master');
  var lookupsSheet = ss.getSheetByName('Lookups');

  var now = new Date();

  // ── PART 1: BATCHED NEW LEAD NOTIFICATIONS ──
  if (qSheet && qSheet.getLastRow() > 1) {
    var qData = qSheet.getDataRange().getValues();
    var pendingByBDA = {};

    for (var i = 1; i < qData.length; i++) {
      if (qData[i][3] === 'Y') continue; // already notified
      var bdaEmail = qData[i][0];
      var leadID   = qData[i][1];
      if (!pendingByBDA[bdaEmail]) pendingByBDA[bdaEmail] = [];
      pendingByBDA[bdaEmail].push({ leadID: leadID, row: i + 1 });
    }

    // Get lead details for each pending notification
    var leadsData = leadsSheet.getDataRange().getValues();
    var leadsHeaders = leadsData[0];
    var leadIDCol    = leadsHeaders.indexOf('LeadID');
    var nameCol      = leadsHeaders.indexOf('Name');
    var mobileCol    = leadsHeaders.indexOf('Mobile');
    var serviceCol   = leadsHeaders.indexOf('Service Required');
    var sourceCol    = leadsHeaders.indexOf('Source');

    // Build lead lookup map
    var leadMap = {};
    for (var r = 1; r < leadsData.length; r++) {
      if (!leadsData[r][leadIDCol]) continue;
      leadMap[leadsData[r][leadIDCol]] = {
        name:    leadsData[r][nameCol],
        mobile:  leadsData[r][mobileCol],
        service: leadsData[r][serviceCol],
        source:  leadsData[r][sourceCol]
      };
    }

    // Get BDA names from Lookups
    var lookupsData = lookupsSheet.getDataRange().getValues();
    var nameMap = {};
    for (var l = 1; l < lookupsData.length; l++) {
      if (lookupsData[l][1]) nameMap[lookupsData[l][1]] = lookupsData[l][0];
    }

    // Send one email per BDA
    for (var email in pendingByBDA) {
      var items = pendingByBDA[email];
      var bdaName = nameMap[email] || email;

      var rows = items.map(function(item) {
        var lead = leadMap[item.leadID] || {};
        return '<tr>' +
          '<td style="padding:8px;border:1px solid #ddd">' + (item.leadID || '-') + '</td>' +
          '<td style="padding:8px;border:1px solid #ddd">' + (lead.name || '-') + '</td>' +
          '<td style="padding:8px;border:1px solid #ddd">' + (lead.mobile || '-') + '</td>' +
          '<td style="padding:8px;border:1px solid #ddd">' + (lead.service || '-') + '</td>' +
          '<td style="padding:8px;border:1px solid #ddd">' + (lead.source || '-') + '</td>' +
          '</tr>';
      }).join('');

      var html = '<p>Hi ' + bdaName + ',</p>' +
        '<p>You have been assigned <b>' + items.length + ' new lead(s)</b>. Please action them as soon as possible!</p>' +
        '<table style="border-collapse:collapse;width:100%;font-family:Arial;font-size:13px">' +
        '<thead><tr style="background:#1a73e8;color:white">' +
        '<th style="padding:8px">Lead ID</th>' +
        '<th style="padding:8px">Name</th>' +
        '<th style="padding:8px">Mobile</th>' +
        '<th style="padding:8px">Service</th>' +
        '<th style="padding:8px">Source</th>' +
        '</tr></thead><tbody>' + rows + '</tbody></table>' +
        '<br><p style="color:#666;font-size:12px">Please log in to the CRM portal and call these leads. — NotifyTechAI CRM</p>';

      try {
        MailApp.sendEmail({
          to: email,
          subject: '🆕 ' + items.length + ' New Lead(s) Assigned to You — NotifyTechAI',
          htmlBody: html
        });

        // Mark as notified
        items.forEach(function(item) {
          qSheet.getRange(item.row, 4).setValue('Y');
        });
      } catch(e) {
        Logger.log('Email send error for ' + email + ': ' + e.message);
      }
    }
  }

  // ── PART 2: UNTOUCHED LEAD ALERTS (2 hours, no action) ──
  var leadsData2 = leadsSheet.getDataRange().getValues();
  var leadsHeaders2 = leadsData2[0];
  var callStatusCol    = leadsHeaders2.indexOf('Call Status');
  var lastDateCol      = leadsHeaders2.indexOf('Last Date');
  var assignedEmailCol = leadsHeaders2.indexOf('Assigned Email');
  var assignedToCol    = leadsHeaders2.indexOf('Assigned To');
  var nameCol2         = leadsHeaders2.indexOf('Name');
  var mobileCol2       = leadsHeaders2.indexOf('Mobile');
  var serviceCol2      = leadsHeaders2.indexOf('Service Required');
  var leadIDCol2       = leadsHeaders2.indexOf('LeadID');

  var twoHoursAgo = new Date(now.getTime() - 2 * 60 * 60 * 1000);
  var fourHoursAgo = new Date(now.getTime() - 4 * 60 * 60 * 1000);

  // Only alert for leads assigned between 2-4 hours ago
  // (avoids re-alerting leads that are older)
  var untouchedByBDA = {};

  for (var i = 1; i < leadsData2.length; i++) {
    if (!leadsData2[i][0]) continue;
    if (leadsData2[i][callStatusCol] !== 'Not Called') continue;

    var lastDate = new Date(leadsData2[i][lastDateCol]);
    if (isNaN(lastDate)) continue;

    // Only leads where Last Date is between 2 and 4 hours ago
    if (lastDate > twoHoursAgo || lastDate < fourHoursAgo) continue;

    var bdaEmail2 = leadsData2[i][assignedEmailCol];
    var bdaName2  = leadsData2[i][assignedToCol];
    if (!bdaEmail2) continue;

    if (!untouchedByBDA[bdaEmail2]) {
      untouchedByBDA[bdaEmail2] = { name: bdaName2, leads: [] };
    }
    untouchedByBDA[bdaEmail2].leads.push({
      leadID:  leadsData2[i][leadIDCol2],
      name:    leadsData2[i][nameCol2],
      mobile:  leadsData2[i][mobileCol2],
      service: leadsData2[i][serviceCol2]
    });
  }

  for (var email2 in untouchedByBDA) {
    var bda = untouchedByBDA[email2];
    var rows2 = bda.leads.map(function(l) {
      return '<tr>' +
        '<td style="padding:8px;border:1px solid #ddd">' + l.leadID + '</td>' +
        '<td style="padding:8px;border:1px solid #ddd">' + l.name + '</td>' +
        '<td style="padding:8px;border:1px solid #ddd">' + l.mobile + '</td>' +
        '<td style="padding:8px;border:1px solid #ddd">' + l.service + '</td>' +
        '</tr>';
    }).join('');

    var html2 = '<p>Hi ' + bda.name + ',</p>' +
      '<p>You have <b>' + bda.leads.length + ' lead(s)</b> that have been assigned for over 2 hours with no action taken yet. Please call them now!</p>' +
      '<table style="border-collapse:collapse;width:100%;font-family:Arial;font-size:13px">' +
      '<thead><tr style="background:#f9ab00;color:white">' +
      '<th style="padding:8px">Lead ID</th>' +
      '<th style="padding:8px">Name</th>' +
      '<th style="padding:8px">Mobile</th>' +
      '<th style="padding:8px">Service</th>' +
      '</tr></thead><tbody>' + rows2 + '</tbody></table>' +
      '<br><p style="color:#666;font-size:12px">Please log in to the CRM portal and action these leads immediately. — NotifyTechAI CRM</p>';

    try {
      MailApp.sendEmail({
        to: email2,
        subject: '⚠️ ' + bda.leads.length + ' Lead(s) Untouched for 2 Hours — NotifyTechAI',
        htmlBody: html2
      });
    } catch(e) {
      Logger.log('Untouched email error for ' + email2 + ': ' + e.message);
    }
  }
}
function calculateLeadScore(remarks, pipeline, callStatus, followUpDate, lastDate, serviceRequired) {
  var breakdown = {};

  // ── PIPELINE STAGE (max 40) ──
  var pipelineScores = {
    'Won': 40, 'Project (80%)': 35, 'Negotiation': 30,
    'Proposal Sent': 25, 'Qualified': 20, 'Contacted': 10, 'New Lead': 5
  };
  breakdown.pipeline = pipelineScores[pipeline] || 0;

  // ── CALL STATUS (max 20) ──
  var callScores = {
    'Interested': 20, 'Connected': 15, 'Called - Callback': 10,
    'Called - No Answer': 5, 'Not Called': 0, 'Unqualified Lead': 0
  };
  breakdown.callStatus = callScores[callStatus] || 0;

  // ── FOLLOW-UP DATE (max 10) ──
  breakdown.followUp = 0;
  if (followUpDate) {
    var fDate = new Date(followUpDate);
    var today = new Date(); today.setHours(0,0,0,0);
    if (fDate >= today) breakdown.followUp = 10;
  }

  // ── DAYS SINCE LAST CONTACT (max 10) ──
  breakdown.recency = 0;
  if (lastDate) {
    var lDate = new Date(lastDate);
    var diffDays = Math.floor((new Date() - lDate) / (1000 * 60 * 60 * 24));
    if (diffDays < 1) breakdown.recency = 10;
    else if (diffDays <= 3) breakdown.recency = 5;
  }

  // ── SERVICE REQUIRED (max 10) ──
  breakdown.service = (serviceRequired && serviceRequired.toString().trim() !== '') ? 10 : 0;

  // ── SENTIMENT FROM REMARKS via Gemini (max 10) ──
  breakdown.sentiment = 5;
  var sentiment = 'Neutral';
 if (remarks && remarks.toString().trim().split(' ').length > 10) {
    try {
      var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + getGeminiKey();
      var payload = {
        contents: [{
          parts: [{ text: 'Analyse the sentiment of this sales remark in one word only — reply with exactly one of these three words: Positive, Neutral, Negative. Remark: "' + remarks + '"' }]
        }]
      };
      var response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
      });
      var json = JSON.parse(response.getContentText());
      var result = json.candidates[0].content.parts[0].text.trim();
      if (result.indexOf('Positive') !== -1) { sentiment = 'Positive'; breakdown.sentiment = 10; }
      else if (result.indexOf('Negative') !== -1) { sentiment = 'Negative'; breakdown.sentiment = 0; }
      else { sentiment = 'Neutral'; breakdown.sentiment = 5; }
    } catch(e) {
      Logger.log('Gemini error: ' + e.message);
      breakdown.sentiment = 5;
    }
  }

  var score = breakdown.pipeline + breakdown.callStatus + breakdown.followUp + breakdown.recency + breakdown.service + breakdown.sentiment;
  var label = score >= 70 ? '🔥 Hot' : score >= 40 ? '⚡ Warm' : '❄️ Cold';

  return { score: score, label: label, sentiment: sentiment, breakdown: breakdown };
}

// ── UPDATE ALL LEAD SCORES ──
function updateAllLeadScores() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Leads_Master');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var remarksCol    = headers.indexOf('Remarks');
  var pipelineCol   = headers.indexOf('Pipeline & Stage');
  var callStatusCol = headers.indexOf('Call Status');
  var followUpCol   = headers.indexOf('Follow-up Date');
  var lastDateCol   = headers.indexOf('Last Date');
  var serviceCol    = headers.indexOf('Service Required');
  var leadScoreCol  = headers.indexOf('Lead Score');
  var leadLabelCol  = headers.indexOf('Lead Label');

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    // ── SKIP GEMINI IF NO REMARKS ──
    if (!data[i][remarksCol] || data[i][remarksCol].toString().trim().length <= 3) {
      var result = calculateLeadScore(
        '',
        data[i][pipelineCol],
        data[i][callStatusCol],
        data[i][followUpCol],
        data[i][lastDateCol],
        data[i][serviceCol]
      );
      sheet.getRange(i + 1, leadScoreCol + 1).setValue(result.score);
      sheet.getRange(i + 1, leadLabelCol + 1).setValue(result.label);
      continue;
    }

    
   // ── HAS REMARKS — SKIP GEMINI IN BULK RUN ──
    var result = calculateLeadScore(
      '', // sentiment handled in onEdit only
      data[i][pipelineCol],
      data[i][callStatusCol],
      data[i][followUpCol],
      data[i][lastDateCol],
      data[i][serviceCol]
    );
    sheet.getRange(i + 1, leadScoreCol + 1).setValue(result.score);
    sheet.getRange(i + 1, leadLabelCol + 1).setValue(result.label);
  }
  Logger.log('Lead scores updated for ' + (data.length - 1) + ' rows');
}

function getGeminiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

function createMonthlyTrigger() {
  // Delete existing triggers for syncTargetsFromLookups first
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'syncTargetsFromLookups') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new monthly trigger — runs on 1st of every month at 8am
  ScriptApp.newTrigger('syncTargetsFromLookups')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();
    
  Logger.log('✅ Monthly trigger created!');
}

function colorTargetsByMonth() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Targets');
  var data = sheet.getDataRange().getValues();
  var lastCol = sheet.getLastColumn();

  // One color per month (Jan→Dec)
  var monthColors = {
    '01': '#cfe2f3', // Jan — blue
    '02': '#d9ead3', // Feb — green
    '03': '#fff2cc', // Mar — yellow
    '04': '#fce5cd', // Apr — orange
    '05': '#ead1dc', // May — pink
    '06': '#d0e0e3', // Jun — teal
    '07': '#e1d5e7', // Jul — purple
    '08': '#f4cccc', // Aug — red
    '09': '#d9d2e9', // Sep — lavender
    '10': '#c9daf8', // Oct — cornflower
    '11': '#b6d7a8', // Nov — sage
    '12': '#ffd966'  // Dec — gold
  };

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var rowMonth = String(data[i][4]).length === 7 ? String(data[i][4]) : '';
    if (!rowMonth && data[i][4]) {
      try {
        rowMonth = Utilities.formatDate(new Date(data[i][4]), Session.getScriptTimeZone(), 'yyyy-MM');
      } catch(e) {}
    }
    if (!rowMonth) continue;
    var mm = rowMonth.split('-')[1]; // extract month number
    var bg = monthColors[mm] || '#ffffff';
    sheet.getRange(i + 1, 1, 1, lastCol).setBackground(bg);
  }
  Logger.log('✅ Targets colored by month!');
}

function colorLookups() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Lookups');
  var data = sheet.getDataRange().getValues();

 // Add auto-filter only if not already present
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).createFilter();
  }

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var role = data[i][3];
    var range = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());

    if (role === 'Admin') {
      range.setBackground('#c9daf8'); // light blue
    } else if (role === 'Team Leader') {
      range.setBackground('#d9d2e9'); // light purple
    } else if (role === 'BDA') {
      range.setBackground('#d9ead3'); // light green
    } else {
      range.setBackground(null); // clear for unknown roles
    }
  }

  Logger.log('colorLookups complete');
}

function fixBulkLeads() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Leads_Master');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var leadIDCol   = headers.indexOf('LeadID');
  var dateCol     = headers.indexOf('Date');
  var lastDateCol = headers.indexOf('Last Date');
  var mobileCol   = headers.indexOf('Mobile');
  var leadLockCol = headers.indexOf('Lead Lock');

  for (var i = 1; i < data.length; i++) {
    if (!data[i][mobileCol]) continue;

    // Generate missing LeadID
    if (!data[i][leadIDCol]) {
      var newID = 'LD-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + i;
      sheet.getRange(i + 1, leadIDCol + 1).setValue(newID);
    }

    // Fill missing Date
    if (!data[i][dateCol]) {
      var fallback = new Date(data[i][lastDateCol] || new Date());
      sheet.getRange(i + 1, dateCol + 1).setValue(new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate()));
    }

    // Fix Lead Lock
    if (!data[i][leadLockCol]) {
      sheet.getRange(i + 1, leadLockCol + 1).setValue('N');
    }
  }

  applyLeadColors();
  sortLeadsByStatus();
  SpreadsheetApp.getUi().alert('✅ Bulk leads fixed! LeadIDs, dates and colors applied.');
}
var ARCHIVE_MONTHS_THRESHOLD = 6;
var ARCHIVE_SHEET_NAME = 'Leads_Archive';
 
// ── SHARED HELPER: get or create Leads_Archive sheet ──
function getOrCreateArchiveSheet_(ss) {
  var master  = ss.getSheetByName('Leads_Master');
  var archive = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (!archive) {
    archive = ss.insertSheet(ARCHIVE_SHEET_NAME);
    var headerRow = master.getRange(1, 1, 1, master.getLastColumn()).getValues();
    archive.getRange(1, 1, 1, headerRow[0].length).setValues(headerRow);
    archive.getRange(1, 1, 1, headerRow[0].length)
      .setBackground('#0b1120')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    archive.setFrozenRows(1);
    Logger.log('Created Leads_Archive sheet.');
  }
  return archive;
}
 
// ── SHARED HELPER: move rows matching a filter from master to archive ──
function moveRowsToArchive_(ss, matchFn) {
  var master     = ss.getSheetByName('Leads_Master');
  var archive    = getOrCreateArchiveSheet_(ss);
  var masterData = master.getDataRange().getValues();
 
  if (masterData.length <= 1) {
    return { success: true, message: 'No leads to archive.', count: 0 };
  }
 
  var headers         = masterData[0];
  var COL_PIPELINE    = headers.indexOf('Pipeline & Stage');
  var COL_LAST_DATE   = headers.indexOf('Last Date');
  var COL_CALL_STATUS = headers.indexOf('Call Status');
  var COL_LEAD_ID     = headers.indexOf('LeadID');
 
  var rowsToArchive   = [];
  var indicesToDelete = [];
 
  for (var i = 1; i < masterData.length; i++) {
    if (!masterData[i][COL_LEAD_ID]) continue;
 
    var pipeline    = String(masterData[i][COL_PIPELINE]   || '').trim();
    var callStatus  = String(masterData[i][COL_CALL_STATUS] || '').trim();
    var lastDateVal = masterData[i][COL_LAST_DATE];
 
    // Must be a closed lead
    var isClosed = (pipeline === 'Won' || pipeline === 'Unqualified' || callStatus === 'Unqualified Lead');
    if (!isClosed) continue;
 
    if (!lastDateVal) continue;
    var lastDate = new Date(lastDateVal);
    if (isNaN(lastDate.getTime())) continue;
 
    if (!matchFn(lastDate)) continue;
 
    rowsToArchive.push(masterData[i]);
    indicesToDelete.push(i + 1); // 1-based sheet row
  }
 
  if (rowsToArchive.length === 0) {
    return { success: true, message: 'No leads qualify for archiving.', count: 0 };
  }
 
  // Write to archive in one batch
  var archiveLastRow = archive.getLastRow();
  archive.getRange(archiveLastRow + 1, 1, rowsToArchive.length, headers.length)
    .setValues(rowsToArchive);
 
  // Delete from master bottom-up so row indices stay valid
  for (var d = indicesToDelete.length - 1; d >= 0; d--) {
    master.deleteRow(indicesToDelete[d]);
  }
 
  try { applyLeadColors(); } catch(e) {}
 
  return { success: true, count: rowsToArchive.length };
}
 
// ── AUTO ARCHIVE ──
// Archives all Won/Unqualified leads with Last Date older than 6 months.
// Called by monthly time trigger OR the "Auto Archive" button in the portal.
function archiveLeads() {
  var ss     = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - ARCHIVE_MONTHS_THRESHOLD);
  cutoff.setHours(0, 0, 0, 0);
 
  var result = moveRowsToArchive_(ss, function(lastDate) {
    var d = new Date(lastDate);
    d.setHours(0, 0, 0, 0);
    return d < cutoff;
  });
 
  if (!result.success) return result;
  if (result.count === 0) {
    return { success: true, message: 'No leads older than ' + ARCHIVE_MONTHS_THRESHOLD + ' months found.', count: 0 };
  }
  Logger.log('Auto-archived ' + result.count + ' leads.');
  return {
    success: true,
    message: 'Archived ' + result.count + ' lead(s) older than ' + ARCHIVE_MONTHS_THRESHOLD + ' months.',
    count: result.count
  };
}
 
// ── MONTH-SPECIFIC ARCHIVE ──
// Archives Won/Unqualified leads whose Last Date falls in the given month.
// month: 'yyyy-MM' string e.g. '2025-03'
function archiveLeadsByMonth(month) {
  if (!month || !/^\d{4}-\d{2}$/.test(month)) {
    return { success: false, message: 'Invalid month format. Expected yyyy-MM.' };
  }
 
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var tz = Session.getScriptTimeZone();
 
  var result = moveRowsToArchive_(ss, function(lastDate) {
    var rowMonth = Utilities.formatDate(new Date(lastDate), tz, 'yyyy-MM');
    return rowMonth === month;
  });
 
  if (!result.success) return result;
  if (result.count === 0) {
    return { success: true, message: 'No Won / Unqualified leads found for ' + month + '.', count: 0 };
  }
  Logger.log('Month-archived ' + result.count + ' leads for ' + month + '.');
  return {
    success: true,
    message: 'Archived ' + result.count + ' lead(s) from ' + month + '.',
    count: result.count
  };
}
 
// ── PRE-CHECK COUNT ──
// Returns how many leads WOULD be archived for a given month,
// used to show the count in the confirmation prompt before actually archiving.
function countArchivableByMonth(month) {
  if (!month || !/^\d{4}-\d{2}$/.test(month)) {
    return { success: false, message: 'Invalid month.' };
  }
 
  var ss         = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var master     = ss.getSheetByName('Leads_Master');
  var masterData = master.getDataRange().getValues();
  var headers    = masterData[0];
  var tz         = Session.getScriptTimeZone();
 
  var COL_PIPELINE    = headers.indexOf('Pipeline & Stage');
  var COL_LAST_DATE   = headers.indexOf('Last Date');
  var COL_CALL_STATUS = headers.indexOf('Call Status');
  var COL_LEAD_ID     = headers.indexOf('LeadID');
 
  var count = 0;
  for (var i = 1; i < masterData.length; i++) {
    if (!masterData[i][COL_LEAD_ID]) continue;
    var pipeline   = String(masterData[i][COL_PIPELINE]   || '').trim();
    var callStatus = String(masterData[i][COL_CALL_STATUS] || '').trim();
    var lastDateVal = masterData[i][COL_LAST_DATE];
 
    var isClosed = (pipeline === 'Won' || pipeline === 'Unqualified' || callStatus === 'Unqualified Lead');
    if (!isClosed) continue;
    if (!lastDateVal) continue;
 
    var lastDate = new Date(lastDateVal);
    if (isNaN(lastDate.getTime())) continue;
 
    var rowMonth = Utilities.formatDate(lastDate, tz, 'yyyy-MM');
    if (rowMonth === month) count++;
  }
 
  return { success: true, count: count, month: month };
}
 
// ── MONTHLY AUTO-ARCHIVE TRIGGER SETUP ──
// Run createArchiveTrigger() ONCE manually from the Apps Script editor.
function createArchiveTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'archiveLeads') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  // Runs 1st of every month at 6 AM — before snapshot trigger at 7 AM
  ScriptApp.newTrigger('archiveLeads')
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();
  Logger.log('Archive trigger created — runs 1st of month at 6 AM.');
}

// ============================================================
// HRMS — LEAVE MANAGEMENT
// ============================================================

const LEAVE_POLICY = {
  Casual:  { perMonth: 1,    carryForward: false },
  Sick:    { perMonth: 0.5,  carryForward: true  },
  Earned:  { perMonth: 1.25, carryForward: true  }
};

// ------------------------------------------------------------
// submitLeaveRequest(data)
// Called from Index.html when employee submits the leave form
// ------------------------------------------------------------
function submitLeaveRequest(data) {
  try {
   var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var requestSheet = ss.getSheetByName('Leave_Requests');
    var balanceSheet = ss.getSheetByName('Leave_Balances');
    var lookupSheet  = ss.getSheetByName('Lookups');

    if (!requestSheet || !balanceSheet || !lookupSheet) {
      return { success: false, message: 'One or more required sheets not found.' };
    }

  var userEmail = data.employeeEmail;
    var userName  = data.employeeName;
    var empId     = data.employeeId;

    // --- Calculate total days ---
    var from  = new Date(data.fromDate);
    var to    = new Date(data.toDate);
    var diffMs   = to - from;
    var totalDays = Math.floor(diffMs / (1000 * 60 * 60 * 24)) + 1;

    if (totalDays <= 0) {
      return { success: false, message: 'End date must be on or after start date.' };
    }

    // --- Check leave balance ---
    var balanceData = balanceSheet.getDataRange().getValues();
    var balanceRowIndex = -1;
    var currentBalance  = 0;
    var balanceCol      = data.leaveType === 'Casual' ? 2
                        : data.leaveType === 'Sick'   ? 3
                        : 4; // Earned

    for (var i = 1; i < balanceData.length; i++) {
      if (balanceData[i][1] === userEmail) { // col B = EmployeeEmail
        balanceRowIndex = i + 1;
        currentBalance  = balanceData[i][balanceCol];
        break;
      }
    }

    if (balanceRowIndex === -1) {
      return { success: false, message: 'Your balance record was not found. Please contact Admin.' };
    }

    if (currentBalance < totalDays) {
      return {
        success: false,
        message: 'Insufficient ' + data.leaveType + ' leave balance. Available: ' + currentBalance + ' day(s).'
      };
    }

    // --- Find approver from Lookups ---
    // Lookups: col A = Email, col B = Name, col D = Role, col E = ReportsTo (TL email or Admin)
    var approverEmail = '';
    var approverName  = '';
    var lookupData    = lookupSheet.getDataRange().getValues();

    for (var j = 1; j < lookupData.length; j++) {
      if (lookupData[j][1] === userEmail) { // col B = Email
        var role = lookupData[j][3]; // col D = Role
        if (role === 'BDA') {
  var tlEmailOrName = lookupData[j][4];
  for (var k = 1; k < lookupData.length; k++) {
    if ((lookupData[k][0] === tlEmailOrName || lookupData[k][1] === tlEmailOrName)
        && lookupData[k][3] === 'Team Leader') {
      approverEmail = lookupData[k][1]; // col B = email
      approverName  = lookupData[k][0]; // col A = name
      break;
    }
  }
} else {
          // TL or Admin — goes directly to Admin
          approverEmail = 'developernotifytechai@gmail.com';
          approverName  = 'Admin';
        }
        break;
      }
    }

    if (!approverEmail) {
      return { success: false, message: 'Could not determine your approver. Please contact Admin.' };
    }

    // --- Log to Leave_Requests sheet ---
    var timestamp = new Date();
    requestSheet.appendRow([
      timestamp,
      empId,
      userName,
      userEmail,
      data.leaveType,
      data.fromDate,
      data.toDate,
      totalDays,
      data.reason,
      data.emailBody,
      approverEmail,
      'Pending',
      '',   // ApproverRemarks — blank until actioned
      ''    // ActionTimestamp — blank until actioned
    ]);

    // --- Send email via GmailApp ---
    var subject = '[Leave Request] ' + userName + ' — ' + data.leaveType + ' Leave (' + data.fromDate + ' to ' + data.toDate + ')';
    var body    = data.emailBody
                + '\n\n---'
                + '\nLeave type: ' + data.leaveType
                + '\nFrom: '       + data.fromDate
                + '\nTo: '         + data.toDate
                + '\nTotal days: ' + totalDays
                + '\nReason: '     + data.reason
                + '\n\nPlease log in to the CRM to approve or reject this request.';

  // Build email options
var emailOptions = {
  to      : approverEmail,
  subject : subject,
  body    : body + '\n\n---\nSubmitted by: ' + userName + ' (' + userEmail + ')',
  name    : userName + ' via NotifyTech CRM',
  replyTo : userEmail
};

// If an attachment was uploaded, decode it and attach to the email
if (data.attachmentBase64 && data.attachmentMime && data.attachmentName) {
  try {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(data.attachmentBase64),
      data.attachmentMime,
      data.attachmentName
    );
    emailOptions.attachments = [blob];
  } catch (attachErr) {
    Logger.log('Attachment error (non-fatal): ' + attachErr.message);
    // We don't fail the whole request just because attachment failed
  }
}

MailApp.sendEmail(emailOptions);

    return {
      success: true,
      message: 'Leave request submitted successfully. Your email has been sent to ' + approverName + '.'
    };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}


// -----------------------------------------------------------
// getMyLeaves()
// Returns the current user's leave history + balances
// ------------------------------------------------------------
function getMyLeaves(userEmail) {
  try {
   var ss           = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var requestSheet = ss.getSheetByName('Leave_Requests');
    var balanceSheet = ss.getSheetByName('Leave_Balances');
    // var userEmail    = Session.getActiveUser().getEmail();

    var requests = requestSheet.getDataRange().getValues();
    var myLeaves = [];

    for (var i = 1; i < requests.length; i++) {
      if (requests[i][3] === userEmail) { // col D = EmployeeEmail
       myLeaves.push({
  timestamp      : requests[i][0] ? requests[i][0].toString() : '',
  leaveType      : requests[i][4] ? requests[i][4].toString() : '',
  fromDate       : requests[i][5] ? requests[i][5].toString() : '',
  toDate         : requests[i][6] ? requests[i][6].toString() : '',
  totalDays      : requests[i][7],
  reason         : requests[i][8] ? requests[i][8].toString() : '',
  approverEmail  : requests[i][10] ? requests[i][10].toString() : '',
  status         : requests[i][11] ? requests[i][11].toString() : '',
  remarks        : requests[i][12] ? requests[i][12].toString() : '',
  actionTimestamp: requests[i][13] ? requests[i][13].toString() : ''
});
      }
    }

    // --- Get balance --
    var balances    = balanceSheet.getDataRange().getValues();
    var myBalance   = { Casual: 0, Sick: 0, Earned: 0 };

    for (var j = 1; j < balances.length; j++) {
      if (balances[j][1] === userEmail) { // col B = EmployeeEmail
        myBalance.Casual  = balances[j][2];
        myBalance.Sick    = balances[j][3];
        myBalance.Earned  = balances[j][4];
        break;
      }
    }

    return { success: true, leaves: myLeaves, balance: myBalance };

  } catch (e) {
    return { success: false, message: e.message };
  }
}


// ------------------------------------------------------------
// getPendingLeaves()
// Returns pending requests visible to the logged-in TL or Admin
// ------------------------------------------------------------
function getPendingLeaves(userEmail) {
  try {
   var ss           = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var requestSheet = ss.getSheetByName('Leave_Requests');
      var role = getUserRole(userEmail);

    var requests = requestSheet.getDataRange().getValues();
    var pending  = [];

    for (var i = 1; i < requests.length; i++) {
      var status        = requests[i][11]; // col L
      var approverEmail = requests[i][10]; // col K

      if (status !== 'Pending') continue;

      // Admin sees all pending; TL sees only their own team's requests
      if (role === 'Admin' || approverEmail === userEmail) {
       pending.push({
  rowIndex      : i + 1,
  timestamp     : requests[i][0] ? requests[i][0].toString() : '',
  employeeId    : requests[i][1] ? requests[i][1].toString() : '',
  employeeName  : requests[i][2] ? requests[i][2].toString() : '',
  employeeEmail : requests[i][3] ? requests[i][3].toString() : '',
  leaveType     : requests[i][4] ? requests[i][4].toString() : '',
  fromDate      : requests[i][5] ? requests[i][5].toString() : '',
  toDate        : requests[i][6] ? requests[i][6].toString() : '',
  totalDays     : requests[i][7],
  reason        : requests[i][8] ? requests[i][8].toString() : '',
  emailBody     : requests[i][9] ? requests[i][9].toString() : ''
});
      }
    }

    return { success: true, pending: pending };

  } catch (e) {
    return { success: false, message: e.message };
  }
}


// ------------------------------------------------------------
// actionLeaveRequest(rowIndex, action, remarks)
// action = 'Approved' or 'Rejected'
// Called when TL/Admin clicks Approve or Reject in the CRM
// ------------------------------------------------------------
function actionLeaveRequest(rowIndex, action, remarks, userEmail) {
  try {
   var ss           = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var requestSheet = ss.getSheetByName('Leave_Requests');
    var balanceSheet = ss.getSheetByName('Leave_Balances');
    var role = getUserRole(userEmail);

    var row = requestSheet.getRange(rowIndex, 1, 1, 14).getValues()[0];

    // Security check — TL can only action their own team's requests
    if (role !== 'Admin' && row[10] !== userEmail) {
      return { success: false, message: 'You are not authorised to action this request.' };
    }

    // Prevent double-action
    if (row[11] !== 'Pending') {
      return { success: false, message: 'This request has already been ' + row[11] + '.' };
    }

    var employeeEmail = row[3];
    var employeeName  = row[2];
    var leaveType     = row[4];
    var totalDays     = row[7];
    var fromDate      = row[5];
    var toDate        = row[6];
    var actionTime    = new Date();

    // --- Update Leave_Requests row ---
    requestSheet.getRange(rowIndex, 11).setValue(userEmail);  // ApproverEmail = actioner's email
    requestSheet.getRange(rowIndex, 12).setValue(action);
    requestSheet.getRange(rowIndex, 13).setValue(remarks || '');
    requestSheet.getRange(rowIndex, 14).setValue(actionTime);

    // --- Deduct balance if Approved ---
    if (action === 'Approved') {
      var balances  = balanceSheet.getDataRange().getValues();
      var balCol    = leaveType === 'Casual' ? 3
                    : leaveType === 'Sick'   ? 4
                    : 5; // Earned (1-indexed for setCell)

      for (var i = 1; i < balances.length; i++) {
        if (balances[i][1] === employeeEmail) {
          var current = balances[i][balCol - 1];
          balanceSheet.getRange(i + 1, balCol).setValue(current - totalDays);
          break;
        }
      }
    }

    // --- Notify employee by email ---
    var approverName = getUserNameFromEmail(userEmail);
    var subject = '[Leave ' + action + '] ' + leaveType + ' Leave — ' + fromDate + ' to ' + toDate;
    var body    = 'Hi ' + employeeName + ',\n\n'
                + 'Your ' + leaveType + ' leave request from ' + fromDate + ' to ' + toDate
                + ' has been ' + action.toLowerCase() + ' by ' + approverName + '.\n\n'
                + (remarks ? 'Remarks: ' + remarks + '\n\n' : '')
                + 'You can view your leave history in the CRM under My Leaves.';

GmailApp.sendEmail(employeeEmail, subject, body, {
  name: approverName + ' via NotifyTech CRM'
});

    return { success: true, message: 'Leave ' + action.toLowerCase() + ' and employee notified.' };

  } catch (e) {
    return { success: false, message: e.message };
  }
}


// ------------------------------------------------------------
// getUserNameFromEmail(email) — helper
// Resolves an email to a display name from the Lookups sheet
// ------------------------------------------------------------
function getUserNameFromEmail(email) {
  var ss          = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var lookupSheet = ss.getSheetByName('Lookups');
  var data        = lookupSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === email) return data[i][0]; // col B = email, col A = name
  }

  // Fallback for admin emails
  if (email === 'developernotifytechai@gmail.com' || email === 'saniyakhan1709@gmail.com') {
    return 'Admin';
  }

  return email;
}


// ------------------------------------------------------------
// refreshLeaveBalances() — run via monthly time-based trigger
// Credits each employee's balance for the new month
// ------------------------------------------------------------
function refreshLeaveBalances() {
  initLeaveBalances();
  var ss           = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var balanceSheet = ss.getSheetByName('Leave_Balances');
  var data         = balanceSheet.getDataRange().getValues();
  var currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM-yyyy');

  for (var i = 1; i < data.length; i++) {
    var casual  = data[i][2];
    var sick    = data[i][3];
    var earned  = data[i][4];

    // Add monthly credit
    var newCasual = LEAVE_POLICY.Casual.carryForward  ? casual  + LEAVE_POLICY.Casual.perMonth
                                                      : LEAVE_POLICY.Casual.perMonth;
    var newSick   = LEAVE_POLICY.Sick.carryForward    ? sick    + LEAVE_POLICY.Sick.perMonth
                                                      : LEAVE_POLICY.Sick.perMonth;
    var newEarned = LEAVE_POLICY.Earned.carryForward  ? earned  + LEAVE_POLICY.Earned.perMonth
                                                      : LEAVE_POLICY.Earned.perMonth;

    balanceSheet.getRange(i + 1, 3).setValue(newCasual);
    balanceSheet.getRange(i + 1, 4).setValue(newSick);
    balanceSheet.getRange(i + 1, 5).setValue(newEarned);
    balanceSheet.getRange(i + 1, 6).setValue(currentMonth);
  }

  Logger.log('Leave balances refreshed for ' + currentMonth);
}
function initLeaveBalances() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var lookupSheet  = ss.getSheetByName('Lookups');
  var balanceSheet = ss.getSheetByName('Leave_Balances');

  var lookupData   = lookupSheet.getDataRange().getValues();
  var balanceData  = balanceSheet.getDataRange().getValues();
  var currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM-yyyy');

  // Build set of emails already in Leave_Balances
  var existingEmails = {};
  for (var i = 1; i < balanceData.length; i++) {
    if (balanceData[i][1]) existingEmails[balanceData[i][1]] = true;
  }

  // Loop through Lookups and add missing employees
  for (var i = 1; i < lookupData.length; i++) {
    var name   = lookupData[i][0]; // col A
    var email  = lookupData[i][1]; // col B
    var role   = lookupData[i][3]; // col D
    var empId  = lookupData[i][8]; // col I

    if (!email || !name) continue;
    if (role === 'Admin') continue;
    if (existingEmails[email]) continue;

    balanceSheet.appendRow([
      empId,                          // EmployeeId
      email,                          // EmployeeEmail
      LEAVE_POLICY.Casual.perMonth,   // CasualLeft — 1
      LEAVE_POLICY.Sick.perMonth,     // SickLeft — 0.5
      LEAVE_POLICY.Earned.perMonth,   // EarnedLeft — 1.25
      currentMonth                    // Month
    ]);

    existingEmails[email] = true;
    Logger.log('Added: ' + name + ' (' + empId + ')');
  }

  Logger.log('initLeaveBalances complete.');
}
function debugLeaveBalance() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Leave_Balances');
  var data = sheet.getDataRange().getValues();
  var testEmail = 'angelsaniya2016@gmail.com'; // paste exact email here
  
  for (var i = 1; i < data.length; i++) {
    Logger.log('Row ' + i + ': [' + data[i][1] + '] match=' + (data[i][1] === testEmail));
  }
}
function testLeaveSheets() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var r = ss.getSheetByName('Leave_Requests');
  var b = ss.getSheetByName('Leave_Balances');
  Logger.log('Leave_Requests: ' + (r ? 'FOUND' : 'MISSING'));
  Logger.log('Leave_Balances: ' + (b ? 'FOUND' : 'MISSING'));
  Logger.log('All sheets: ' + ss.getSheets().map(function(s){ return s.getName(); }).join(', '));
}
function testGetMyLeaves() {
  var userEmail = 'angelsaniya2016@gmail.com';
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  
  Logger.log('Step 1: Opening sheets...');
  var requestSheet = ss.getSheetByName('Leave_Requests');
  var balanceSheet = ss.getSheetByName('Leave_Balances');
  Logger.log('Step 2: Sheets opened OK');
  
  Logger.log('Step 3: Reading Leave_Requests...');
  var requests = requestSheet.getDataRange().getValues();
  Logger.log('Step 4: Rows found: ' + requests.length);
  Logger.log('Step 5: Headers: ' + requests[0].join(' | '));
  
  Logger.log('Step 6: Reading Leave_Balances...');
  var balances = balanceSheet.getDataRange().getValues();
  Logger.log('Step 7: Balance rows: ' + balances.length);
  Logger.log('Step 8: Balance headers: ' + balances[0].join(' | '));
  
  Logger.log('Step 9: Looking for email: ' + userEmail);
  for (var j = 1; j < balances.length; j++) {
    Logger.log('Row ' + j + ' email: [' + balances[j][1] + ']');
    if (balances[j][1] === userEmail) {
      Logger.log('FOUND balance row!');
    }
  }
  
  Logger.log('Done!');
}
function testGetMyLeavesCall() {
  var result = getMyLeaves('angelsaniya2016@gmail.com');
  Logger.log(JSON.stringify(result));
}
function testGetPendingLeaves() {
  var result = getPendingLeaves('developernotifytechai@gmail.com');
  Logger.log(JSON.stringify(result));
}

// ============================================================
// CALL SYNC WEBHOOK — NotifyTechAI Android App (FIXED- 27-05-26 v2)
// ============================================================

const CONFIG = {
  SHEET_ID: '1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8',
  CALLS_SHEET_NAME: 'Calls',
  LEADS_SHEET_NAME: 'Leads_Master',
  CALLS_HEADERS: [
    'Server Time',
    'Device Time',
    'Agent',
    'Phone Number',
    'Phone Normalized',
    'Duration (s)',
    'Duration',
    'Call Type',
    'Lead ID',
    'Lead Name',
    'BDA Name',
    'Match Status'
  ],
  LEAD_COLUMN_ALIASES: {
    phone: [
      'phone', 'phone number', 'mobile', 'mobile number', 'contact',
      'contact number', 'customer phone', 'lead phone', 'whatsapp',
      'whatsapp number', 'primary phone', 'primary mobile', 'number'
    ],
    leadId: [
      'lead id', 'leadid', 'id', 'customer id', 'crm id', 'lead code'
    ],
    leadName: [
      'lead name', 'name', 'customer name', 'prospect name', 'client name'
    ],
    bdaName: [
      'bda name', 'bda', 'owner', 'assigned to', 'sales person',
      'salesperson', 'agent name', 'rm name', 'executive name'
    ]
  }
};

function doPost(e) {
  try {
    const payload = parseJsonBody_(e);
    const originalPhone = String(payload.number || payload.phone || payload.mobile || '').trim();
    const normalizedPhone = normalizePhone_(originalPhone);

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const callsSheet = getOrCreateCallsSheet_(ss);
    const leadsSheet = ss.getSheetByName(CONFIG.LEADS_SHEET_NAME);

    const secs = parseInt(payload.duration, 10) || 0;
    const formattedDuration = formatDuration_(secs);
    const match = findLeadMatch_(leadsSheet, originalPhone);

    callsSheet.appendRow([
      new Date(),
      payload.timestamp || '',
      payload.agent || 'Unknown',
      originalPhone,
      normalizedPhone,
      secs,
      formattedDuration,
      payload.type || 'UNKNOWN',
      match.leadId,
      match.leadName,
      (match.bdaName || '').trim(),
      match.status
    ]);

    // ── AUTO-UPDATE CALL STATUS TO CONNECTED IF DURATION > 30s ──
    // Only runs when: lead was matched + call lasted more than 30 seconds
    // Only upgrades status from Not Called or Called - No Answer
    // Never overwrites Interested, Unqualified Lead, Called - Callback etc.
    // to change the duration to 35 or 40 just change this line.
    if (match.found && match.leadId && secs > 30) {
      try {
        var autoLeadsData = leadsSheet.getDataRange().getValues();
        var autoHeaders   = autoLeadsData[0];
        var aLeadIDCol    = autoHeaders.indexOf('LeadID');
        var aCallStatusCol = autoHeaders.indexOf('Call Status');
        var aLastDateCol  = autoHeaders.indexOf('Last Date');

        for (var ai = 1; ai < autoLeadsData.length; ai++) {
          if (String(autoLeadsData[ai][aLeadIDCol]).trim() !== String(match.leadId).trim()) continue;

          var currentStatus = String(autoLeadsData[ai][aCallStatusCol] || '').trim();

          // Only auto-upgrade these two statuses — everything else is left untouched
          if (currentStatus === 'Not Called' || currentStatus === 'Called - No Answer' || currentStatus === '') {
            leadsSheet.getRange(ai + 1, aCallStatusCol + 1).setValue('Connected');
            leadsSheet.getRange(ai + 1, aLastDateCol + 1).setValue(new Date());
            Logger.log('Auto-Connected: ' + match.leadId + ' | Duration: ' + secs + 's');
          }
          break;
        }
      } catch (autoErr) {
        // Log error but do NOT let it break the webhook response
        Logger.log('Auto call status update error: ' + autoErr.message);
      }
    }

    return jsonResponse_({
      success: true,
      matched: match.found,
      inputPhone: originalPhone,
      normalizedPhone: normalizedPhone,
      leadId: match.leadId,
      leadName: match.leadName,
      bdaName: match.bdaName,
      matchStatus: match.status
    });
  } catch (err) {
    return jsonResponse_({
      success: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function doGet() {
  return jsonResponse_({
    status: 'NotifyTechAI webhook is live ✅',
    version: 'fixed-lead-matching'
  });
}

function setupSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const callsSheet = getOrCreateCallsSheet_(ss);
  callsSheet.getRange(1, 1, 1, CONFIG.CALLS_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#1A1A2E')
    .setFontColor('#FFFFFF');
  Logger.log('Calls sheet setup complete.');
}

function diagnoseLeadMatching() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const leadsSheet = ss.getSheetByName(CONFIG.LEADS_SHEET_NAME);
  if (!leadsSheet) throw new Error('Leads_Master sheet not found');

  const values = leadsSheet.getDataRange().getDisplayValues();
  if (!values.length) throw new Error('Leads_Master sheet is empty');

  const headerMap = buildHeaderMap_(values[0]);
  const phoneCol = findColumnIndex_(headerMap, CONFIG.LEAD_COLUMN_ALIASES.phone);
  const idCol = findColumnIndex_(headerMap, CONFIG.LEAD_COLUMN_ALIASES.leadId);
  const nameCol = findColumnIndex_(headerMap, CONFIG.LEAD_COLUMN_ALIASES.leadName);
  const bdaCol = findColumnIndex_(headerMap, CONFIG.LEAD_COLUMN_ALIASES.bdaName);

  Logger.log(JSON.stringify({
    detectedColumns: {
      phoneCol: phoneCol,
      idCol: idCol,
      nameCol: nameCol,
      bdaCol: bdaCol
    },
    headers: values[0]
  }, null, 2));
}

function parseJsonBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Empty request body');
  }
  return JSON.parse(e.postData.contents);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDuration_(seconds) {
  const mins = Math.floor(seconds / 60);
  const remSecs = seconds % 60;
  return `${mins}m ${remSecs}s`;
}

function getOrCreateCallsSheet_(ss) {
  let sheet = ss.getSheetByName(CONFIG.CALLS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.CALLS_SHEET_NAME);
  }

  const hasHeaders = sheet.getLastRow() > 0;
  if (!hasHeaders) {
    sheet.appendRow(CONFIG.CALLS_HEADERS);
  } else {
    const firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), CONFIG.CALLS_HEADERS.length)).getDisplayValues()[0];
    const missingHeaders = CONFIG.CALLS_HEADERS.filter((header, index) => firstRow[index] !== header);
    if (missingHeaders.length) {
      sheet.getRange(1, 1, 1, CONFIG.CALLS_HEADERS.length).setValues([CONFIG.CALLS_HEADERS]);
    }
  }

  // Keep phone columns as plain text to avoid scientific notation.
  sheet.getRange('D:E').setNumberFormat('@');
  return sheet;
}

function findLeadMatch_(leadsSheet, inputPhone) {
  if (!leadsSheet) {
    return emptyMatch_('LEADS_SHEET_MISSING');
  }

  const range = leadsSheet.getDataRange();
  const values = range.getValues();
  const displayValues = range.getDisplayValues();

  if (values.length < 2) {
    return emptyMatch_('LEADS_SHEET_EMPTY');
  }

  const headers = displayValues[0];
  const headerMap = buildHeaderMap_(headers);

  const phoneCol = findColumnIndex_(headerMap, CONFIG.LEAD_COLUMN_ALIASES.phone);
  const leadIdCol = findColumnIndex_(headerMap, CONFIG.LEAD_COLUMN_ALIASES.leadId);
  const leadNameCol = findColumnIndex_(headerMap, CONFIG.LEAD_COLUMN_ALIASES.leadName);
  const bdaNameCol = findColumnIndex_(headerMap, CONFIG.LEAD_COLUMN_ALIASES.bdaName);

  if (phoneCol === -1) {
    return emptyMatch_('PHONE_COLUMN_NOT_FOUND');
  }

  const lookupVariants = buildPhoneVariants_(inputPhone);

  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const rawValue = displayValues[rowIndex][phoneCol] || values[rowIndex][phoneCol];
    const rowVariants = buildPhoneVariants_(rawValue);

    if (hasIntersection_(lookupVariants, rowVariants)) {
      return {
        found: true,
        status: 'MATCHED',
        leadId: leadIdCol > -1 ? String(displayValues[rowIndex][leadIdCol] || values[rowIndex][leadIdCol] || '') : '',
        leadName: leadNameCol > -1 ? String(displayValues[rowIndex][leadNameCol] || values[rowIndex][leadNameCol] || '') : '',
        bdaName: bdaNameCol > -1 ? String(displayValues[rowIndex][bdaNameCol] || values[rowIndex][bdaNameCol] || '') : ''
      };
    }
  }

  return emptyMatch_('NOT_FOUND');
}

function emptyMatch_(status) {
  return {
    found: false,
    status: status,
    leadId: '',
    leadName: '',
    bdaName: ''
  };
}

function buildHeaderMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[normalizeHeader_(header)] = index;
  });
  return map;
}

function findColumnIndex_(headerMap, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const alias = normalizeHeader_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(headerMap, alias)) {
      return headerMap[alias];
    }
  }

  // Fallback: partial match when headers contain extra words.
  const keys = Object.keys(headerMap);
  for (let i = 0; i < aliases.length; i++) {
    const alias = normalizeHeader_(aliases[i]);
    const match = keys.find(key => key.indexOf(alias) !== -1 || alias.indexOf(key) !== -1);
    if (match) return headerMap[match];
  }

  return -1;
}

function normalizeHeader_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizePhone_(value) {
  let digits = String(value || '').trim();
  digits = digits.replace(/[^0-9]+/g, '');

  if (!digits) return '';
  if (digits.startsWith('00')) digits = digits.slice(2);

  // India-specific normalization for +91 / 91 / local 10-digit numbers.
  if (digits.length === 13 && digits.startsWith('091')) {
    digits = digits.slice(3);
  }
  if (digits.length === 12 && digits.startsWith('91')) {
    digits = digits.slice(2);
  }
  if (digits.length === 11 && digits.startsWith('0')) {
    digits = digits.slice(1);
  }
  if (digits.length > 10) {
    digits = digits.slice(-10);
  }

  return digits;
}

function buildPhoneVariants_(value) {
  const original = String(value || '').trim();
  let digits = original.replace(/[^0-9]+/g, '');
  if (digits.startsWith('00')) digits = digits.slice(2);

  const normalized = normalizePhone_(original);
  const variants = new Set();

  if (original) variants.add(original);
  if (digits) variants.add(digits);
  if (normalized) variants.add(normalized);
  if (normalized) variants.add(`91${normalized}`);
  if (normalized) variants.add(`+91${normalized}`);
  if (normalized) variants.add(`0${normalized}`);

  return variants;
}

function hasIntersection_(setA, setB) {
  for (const value of setA) {
    if (setB.has(value)) return true;
  }
  return false;
}
function testGetFinancePending() {
  var result = getFinancePendingPayments();
  Logger.log(JSON.stringify(result));
}

