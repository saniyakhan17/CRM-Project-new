// ── WEB APP ENTRY POINT ── WebApp.gs

function doGet(e) {
  var userEmail = e.parameter.user || '';
  var token = e.parameter.token || '';

  if (!userEmail || !token) {
    var loginTemplate = HtmlService.createTemplateFromFile('Login');
    return loginTemplate.evaluate()
      .setTitle('CRM Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (!validateToken(userEmail, token)) {
    var loginTemplate = HtmlService.createTemplateFromFile('Login');
    return loginTemplate.evaluate()
      .setTitle('CRM Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var role = getUserRole(userEmail);
  if (!role) {
    return HtmlService.createHtmlOutput('<h2> Access Denied. Your account is not registered.</h2>');
  }

  var leads = getMyLeads(userEmail, role);
  var userProfile = getUserProfile(userEmail);
  var bdaList = getBDAList(userEmail, role);
  var token = e.parameter.token; // use the validated token directly

  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = userEmail;
  template.role = role;
  template.token = token;
  template.attRow = e.parameter.attRow || '0';
  template.fresh = e.parameter.fresh || '0';  
  template.userProfile = JSON.stringify(userProfile);
  template.leadsJSON = JSON.stringify(leads);
  template.bdaListJSON = JSON.stringify(bdaList);
  template.tlListJSON = JSON.stringify(getTLList());
 var newLeadsData = [];
try {
  if (role === 'BDA') newLeadsData = getNewLeadsForBDA(userEmail) || [];
  else if (role === 'Team Leader') newLeadsData = getNewLeadsForTL(userEmail) || [];
} catch(err) {
  Logger.log('getNewLeads error: ' + err.message);
}
template.newLeadsJSON = JSON.stringify(newLeadsData);


  return template.evaluate()
    .setTitle('CRM Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserProfile(email) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Lookups');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      return {
        name: data[i][0] || '',
        email: data[i][1] || '',
        role: data[i][3] || '',
        reportsTo: data[i][4] || '',
        employeeID: data[i][8] || '',
        phone: data[i][10] || '',        // col K = index 10
        department: data[i][11] || '',   // col L = index 11
        profilePicFileId: data[i][12] || '' // col M = index 12
      };
    }
  }
  return { name: 'Admin', email: email, role: 'Admin', employeeID: 'NT-001', phone: '', department: '', profilePicFileId: '' };
}

function getMyLeads(userEmail, role) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Leads_Master');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var assignedEmailCol = headers.indexOf('Assigned Email');
  var lookupSheet = ss.getSheetByName('Lookups');
  var empData = lookupSheet.getDataRange().getValues();
  var myLeads = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var show = false;
    if (role === 'Admin') show = true;
    else if (role === 'BDA' && data[i][assignedEmailCol] === userEmail) show = true;
    else if (role === 'Team Leader') {
      if (data[i][assignedEmailCol] === userEmail) {
        show = true;
      } else {
        var tlName = '';
        for (var t = 1; t < empData.length; t++) {
          if (empData[t][1] === userEmail) { tlName = empData[t][0]; break; }
        }
        for (var j = 1; j < empData.length; j++) {
          var reportsTo = String(empData[j][4] || '').trim();
          if ((reportsTo === userEmail || reportsTo === tlName) && empData[j][1] === data[i][assignedEmailCol]) {
            show = true; break;
          }
        }
      }
    }
    if (show) {
      var lead = {};
      headers.forEach(function(h, idx) { lead[h] = data[i][idx]; });
      myLeads.push(lead);
    }
  }
  return myLeads;
}

function getBDAList(userEmail, role) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var lookupSheet = ss.getSheetByName('Lookups');
  var data = lookupSheet.getDataRange().getValues();
  var bdas = [];

  // Get TL name for matching reportsTo column
  var callerName = '';
  if (role === 'Team Leader') {
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === userEmail) { callerName = data[i][0]; break; }
    }
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][3] === 'Admin') continue;
    if (data[i][9] === 'Inactive') continue;
    if (role === 'Admin') {
      bdas.push(data[i][0]);
    } else if (role === 'Team Leader' && data[i][3] === 'BDA') {
      var reportsTo = String(data[i][4] || '').trim();
      if (reportsTo === userEmail || reportsTo === callerName) {
        bdas.push(data[i][0]);
      }
    }
  }
  return bdas;
}

function saveLead(leadID, field, value, callerEmail) {
  Logger.log('saveLead called: leadID=[' + leadID + '] field=[' + field + ']');
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Leads_Master');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var leadIDCol = headers.indexOf('LeadID');
  Logger.log('leadIDCol=' + leadIDCol + ' first3IDs=' + JSON.stringify(data.slice(1,4).map(function(r){return r[leadIDCol];})));
  var fieldCol        = headers.indexOf(field);
  var lastDateCol     = headers.indexOf('Last Date');
  var unqualReasonCol = headers.indexOf('Unqualified Reason');
  var leadLockCol     = headers.indexOf('Lead Lock');

  if (fieldCol === -1) return { success: false, message: 'Field not found' };

  var role = getUserRole(callerEmail);
  var allowedFields = ['Call Status', 'Unqualified Reason', 'Remarks', 'Pipeline & Stage',
    'Follow-up Date', 'Name', 'Alternate Phone', 'Service Required', 'Verification Status', 'Email'];
  if (role === 'BDA' && allowedFields.indexOf(field) === -1) {
    return { success: false, message: 'Not allowed' };
    
  }

 for (var i = 1; i < data.length; i++) {
  if (String(data[i][leadIDCol]).trim() === String(leadID).trim()) {
      if (data[i][leadLockCol] === 'Y' && role !== 'Admin') {
        return { success: false, message: 'Lead is locked. Only Admin can edit.' };
      }

      sheet.getRange(i + 1, fieldCol + 1).setValue(value);

      // ── ACTIVITY LOG HOOK ──
      try {
        var oldVal = String(data[i][fieldCol] || '');
        var logLeadName = String(data[i][headers.indexOf('Name')] || '');
        appendActivityLog(String(leadID), logLeadName, getActionTypeForField(field),
          field + ' changed', oldVal, String(value || ''), callerEmail);
      } catch(logErr) {
        Logger.log('Activity log hook error: ' + logErr.message);
      }

      sheet.getRange(i + 1, lastDateCol + 1).setValue(new Date());

      if (field === 'Call Status' && value !== 'Unqualified Lead') {
        sheet.getRange(i + 1, unqualReasonCol + 1).setValue('');
      }

      // ── SCORE RECALC ──
      var scoreResult = null;
      var scoringFields = ['Pipeline & Stage', 'Call Status', 'Remarks', 'Follow-up Date', 'Service Required'];
      if (scoringFields.indexOf(field) !== -1) {
        try {
          var remarksCol   = headers.indexOf('Remarks');
          var pipelineCol2 = headers.indexOf('Pipeline & Stage');
          var csCol2       = headers.indexOf('Call Status');
          var followUpCol2 = headers.indexOf('Follow-up Date');
          var serviceCol2  = headers.indexOf('Service Required');
          var leadScoreCol = headers.indexOf('Lead Score');
          var leadLabelCol = headers.indexOf('Lead Label');

          var freshRow = sheet.getRange(i + 1, 1, 1, headers.length).getValues()[0];
          freshRow[fieldCol] = value;

          scoreResult = calculateLeadScore(
            freshRow[remarksCol] || '',
            freshRow[pipelineCol2] || '',
            freshRow[csCol2] || '',
            freshRow[followUpCol2] || '',
            freshRow[lastDateCol] || '',
            freshRow[serviceCol2] || ''
          );

          if (leadScoreCol !== -1) sheet.getRange(i + 1, leadScoreCol + 1).setValue(scoreResult.score);
          if (leadLabelCol !== -1) sheet.getRange(i + 1, leadLabelCol + 1).setValue(scoreResult.label);
        } catch(scoreErr) {
          Logger.log('Score calculation error: ' + scoreErr.message);
          scoreResult = null;
        }
      }

      // ── RE-ENGAGE FLAGS ──
      var requiresReEngageNote = false;
      var requiresRevivalNote  = false;

      if (field === 'Pipeline & Stage' && value === 'Re-engage Later') {
        requiresReEngageNote = true;
      }

      if (field === 'Unqualified Reason' && (value === 'Budget' || value === 'Too Early')) {
        var csCol3   = headers.indexOf('Call Status');
        var freshCs  = sheet.getRange(i + 1, csCol3 + 1).getValue();
        if (freshCs === 'Unqualified Lead') {
          requiresRevivalNote = true;
        }
      }

      return {
        success: true,
        newScore: scoreResult ? scoreResult.score : null,
        newLabel: scoreResult ? scoreResult.label : null,
        requiresReEngageNote: requiresReEngageNote,
        requiresRevivalNote:  requiresRevivalNote,
        leadID: leadID
      };
    }
  }

  return { success: false, message: 'Lead not found' };
}
function assignLeadToBDA(leadID, bdaName, callerEmail) {
  var role = getUserRole(callerEmail);
  if (role !== 'Admin' && role !== 'Team Leader') {
    return { success: false, message: 'Not allowed' };
  }
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Leads_Master');
  var lookupSheet = ss.getSheetByName('Lookups');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var empData = lookupSheet.getDataRange().getValues();
  var leadIDCol = headers.indexOf('LeadID');
  var assignedToCol = headers.indexOf('Assigned To');
  var assignedEmailCol = headers.indexOf('Assigned Email');
  var bdaNameCol = headers.indexOf('BDA Name');
  var teamLeaderCol = headers.indexOf('Team Leader');
  var bdaEmail = '', managerEmail = '';
  for (var j = 1; j < empData.length; j++) {
    if (empData[j][0] === bdaName) {
      bdaEmail = empData[j][1];
      managerEmail = empData[j][4];
      break;
    }
  }
  var assignDateCol = headers.indexOf('Assign Date');


 for (var i = 1; i < data.length; i++) {
    if (String(data[i][leadIDCol]).trim() === String(leadID).trim()) {
      sheet.getRange(i + 1, assignedToCol + 1).setValue(bdaName);
      // ── ACTIVITY LOG ──
try {
  var oldBDA = String(data[i][assignedToCol] || '');
  var aLeadName = String(data[i][headers.indexOf('Name')] || '');
  var aType = oldBDA ? 'Lead Reassigned' : 'Lead Assigned';
  appendActivityLog(leadID, aLeadName, aType,
    aType + ' to ' + bdaName + (oldBDA ? ' (previously: ' + oldBDA + ')' : ''),
    oldBDA, bdaName, callerEmail);
} catch(logErr) {}
      sheet.getRange(i + 1, assignedEmailCol + 1).setValue(bdaEmail);
      sheet.getRange(i + 1, bdaNameCol + 1).setValue(bdaName);
      for (var k = 1; k < empData.length; k++) {
  if (empData[k][1] === managerEmail || empData[k][0] === managerEmail) {
    sheet.getRange(i + 1, teamLeaderCol + 1).setValue(empData[k][0]);
    break;
  }
}
      // Write Assign Date only when assigning to a BDA
      if (assignDateCol !== -1) {
        sheet.getRange(i + 1, assignDateCol + 1).setValue(new Date());
      }
      queueLeadNotification(bdaEmail, leadID);
      return { success: true };
    }
  }
  // ── QUEUE NEW LEAD NOTIFICATION ─
  return { success: false, message: 'Lead not found' };
}

function generateToken(email, password) {
  // Legacy — kept for compatibility. Login now uses random tokens.
  return '';
}

function validateLogin(email, password) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      if (data[i][9] === 'Inactive') continue; // skip inactive employees
      if (data[i][1] === email && data[i][5] === password) {
        var token = Utilities.getUuid().replace(/-/g, '').substring(0, 24);
        var expiry = new Date();
        expiry.setHours(expiry.getHours() + 8);
        sheet.getRange(i + 1, 7).setValue(token);
        sheet.getRange(i + 1, 8).setValue(expiry);
        
        var attResult = { rowNumber: 0 };
        try { attResult = logLoginTime(email); } catch(e) {}
        
        return { 
          success: true, 
          token: token, 
          role: data[i][3], 
          attRow: attResult.rowNumber || 0 
        };
      }
    }
    return { success: false, message: 'Invalid email or password' };
    
  } catch(e) {
    Logger.log('validateLogin error: ' + e.message);
    return { success: false, message: 'Server busy, please try again.' };
  }
}

function validateToken(email, token) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Lookups');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      var savedToken = data[i][6]; // col G
      var expiry = data[i][7];     // col H
      if (!savedToken || savedToken !== token) return false;
      if (!expiry) return false;
      var expiryDate = new Date(expiry);
      if (new Date() > expiryDate) return false; // expired
      return true;
    }
  }
  return false;
}

function logoutUser(email, rowNumber) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Lookups');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === email) {
      sheet.getRange(i + 1, 7).setValue(''); // clear token
      sheet.getRange(i + 1, 8).setValue(''); // clear expiry
      // Log logout time in same call
      if (rowNumber && rowNumber > 1) {
        logLogoutTime(email, rowNumber);
      }
      return { success: true };
    }
  }
  return { success: false };
}

function getUserRole(email) {
  if (email === 'developernotifytechai@gmail.com') return 'Admin';
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

function getUserPassword(email) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Lookups');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === email) return data[i][5];
  }
  return '';
}

// ── TARGET VS ACHIEVEMENT (BDA) ──
function getTargetData(userEmail) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var leadsSheet = ss.getSheetByName('Leads_Master');
  var targetsSheet = ss.getSheetByName('Targets');

  var leadsData = leadsSheet.getDataRange().getValues();
  var targetsData = targetsSheet.getDataRange().getValues();
  var leadsHeaders = leadsData[0];

  var assignedEmailCol = leadsHeaders.indexOf('Assigned Email');
  var pipelineCol = leadsHeaders.indexOf('Pipeline & Stage');
  var dealAmountCol = leadsHeaders.indexOf('Deal Amount');
  var dateCol = leadsHeaders.indexOf('Last Date');

  var now = new Date();
  var currentMonth = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');

  var leadTarget = 0;
  var revenueTarget = 0;
  for (var i = 1; i < targetsData.length; i++) {
    var targetMonth = '';
    if (targetsData[i][4]) {
      var tm = new Date(targetsData[i][4]);
      targetMonth = Utilities.formatDate(tm, Session.getScriptTimeZone(), 'yyyy-MM');
    }
    if (String(targetsData[i][4]).length === 7) targetMonth = String(targetsData[i][4]);
    if (targetsData[i][1] === userEmail && targetMonth === currentMonth) {
      leadTarget = targetsData[i][2] || 0;
      revenueTarget = targetsData[i][3] || 0;
      break;
    }
  }

  var achievedLeads = 0;
  var achievedRevenue = 0;
  for (var i = 1; i < leadsData.length; i++) {
    if (!leadsData[i][0]) continue;
    if (leadsData[i][assignedEmailCol] !== userEmail) continue;
    if (leadsData[i][pipelineCol] !== 'Won') continue;
    var leadDate = new Date(leadsData[i][dateCol]);
    var leadMonth = Utilities.formatDate(leadDate, Session.getScriptTimeZone(), 'yyyy-MM');
    if (leadMonth !== currentMonth) continue;
    achievedLeads++;
    achievedRevenue += parseFloat(leadsData[i][dealAmountCol]) || 0;
  }

  var percentLeads = leadTarget > 0 ? parseFloat((achievedLeads / leadTarget * 100).toFixed(1)) : 0;
  var percentRevenue = revenueTarget > 0 ? parseFloat((achievedRevenue / revenueTarget * 100).toFixed(1)) : 0;

  return {
    leadTarget: leadTarget,
    achievedLeads: achievedLeads,
    percentLeads: percentLeads,
    revenueTarget: revenueTarget,
    achievedRevenue: achievedRevenue,
    percentRevenue: percentRevenue,
    currentMonth: currentMonth
  };
}

// ── TEAM TARGET DATA (Team Leader + Admin) ──
function getTeamTargetData(userEmail, userRole) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var leadsSheet = ss.getSheetByName('Leads_Master');
  var targetsSheet = ss.getSheetByName('Targets');
  var lookupsSheet = ss.getSheetByName('Lookups');

  var leadsData = leadsSheet.getDataRange().getValues();
  var targetsData = targetsSheet.getDataRange().getValues();
  var lookupsData = lookupsSheet.getDataRange().getValues();
  var leadsHeaders = leadsData[0];

  var assignedEmailCol = leadsHeaders.indexOf('Assigned Email');
  var pipelineCol = leadsHeaders.indexOf('Pipeline & Stage');
  var dealAmountCol = leadsHeaders.indexOf('Deal Amount');
  var dateCol = leadsHeaders.indexOf('Last Date');

  var now = new Date();
  var currentMonth = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');

  // Build people list
  var peopleList = [];
  for (var i = 1; i < lookupsData.length; i++) {
    if (!lookupsData[i][0]) continue;
    var personRole = lookupsData[i][3];
    if (personRole !== 'BDA' && personRole !== 'Team Leader') continue;

    if (userRole === 'Admin') {
      peopleList.push({
        name: lookupsData[i][0],
        email: lookupsData[i][1],
        role: personRole,
        teamLeader: lookupsData[i][4] || ''
      });
    } else if (userRole === 'Team Leader') {
      // TL sees their own row + their BDAs
      if (personRole === 'Team Leader' && lookupsData[i][1] === userEmail) {
        // Include the TL themselves
        peopleList.push({
          name: lookupsData[i][0],
          email: lookupsData[i][1],
          role: personRole,
          teamLeader: ''
        });
        continue;
      }
      if (personRole !== 'BDA') continue;
      var tlName = '';
      for (var t = 1; t < lookupsData.length; t++) {
        if (lookupsData[t][1] === userEmail) { tlName = lookupsData[t][0]; break; }
      }
      if (lookupsData[i][4] === tlName || lookupsData[i][4] === userEmail) {
        peopleList.push({
          name: lookupsData[i][0],
          email: lookupsData[i][1],
          role: personRole,
          teamLeader: lookupsData[i][4] || ''
        });
      }
    }
  }

  // Get targets
  var targetMap = {};
  for (var i = 1; i < targetsData.length; i++) {
    if (!targetsData[i][1]) continue;
    var existingMonth = String(targetsData[i][4]).length === 7 ? String(targetsData[i][4]) : '';
    if (!existingMonth && targetsData[i][4]) {
      existingMonth = Utilities.formatDate(new Date(targetsData[i][4]), Session.getScriptTimeZone(), 'yyyy-MM');
    }
    if (existingMonth === currentMonth) {
      targetMap[targetsData[i][1]] = {
        leadTarget: targetsData[i][2] || 0,
        revenueTarget: targetsData[i][3] || 0
      };
    }
  }

  // Get achievements
  var achievementMap = {};
  for (var i = 1; i < leadsData.length; i++) {
    if (!leadsData[i][0]) continue;
    if (leadsData[i][pipelineCol] !== 'Won') continue;
    var leadDate = new Date(leadsData[i][dateCol]);
    var leadMonth = Utilities.formatDate(leadDate, Session.getScriptTimeZone(), 'yyyy-MM');
    if (leadMonth !== currentMonth) continue;
    var bEmail = leadsData[i][assignedEmailCol];
    if (!achievementMap[bEmail]) achievementMap[bEmail] = { leads: 0, revenue: 0 };
    achievementMap[bEmail].leads++;
    achievementMap[bEmail].revenue += parseFloat(leadsData[i][dealAmountCol]) || 0;
  }
  // ── AGGREGATE TL ACHIEVEMENT FROM THEIR BDAS ──
  for (var i = 1; i < lookupsData.length; i++) {
    if (!lookupsData[i][0]) continue;
    if (lookupsData[i][3] !== 'Team Leader') continue;
    var tlEmail = lookupsData[i][1];
    var tlName = lookupsData[i][0];
    if (!achievementMap[tlEmail]) achievementMap[tlEmail] = { leads: 0, revenue: 0 };

    // Find all BDAs under this TL
    for (var j = 1; j < lookupsData.length; j++) {
      if (!lookupsData[j][0]) continue;
      if (lookupsData[j][3] !== 'BDA') continue;
      var reportsTo = String(lookupsData[j][4] || '').trim();
      if (reportsTo !== tlEmail && reportsTo !== tlName) continue;
      var bdaEmail = lookupsData[j][1];
      if (achievementMap[bdaEmail]) {
        achievementMap[tlEmail].leads += achievementMap[bdaEmail].leads;
        achievementMap[tlEmail].revenue += achievementMap[bdaEmail].revenue;
      }
    }
  }

  // Build rows
  var rows = [];
  for (var b = 0; b < peopleList.length; b++) {
    var pEmail = peopleList[b].email;
    var target = targetMap[pEmail] || { leadTarget: 0, revenueTarget: 0 };
    var achieved = achievementMap[pEmail] || { leads: 0, revenue: 0 };

    // ── IF THIS IS A TL — ADD ALL BDA TARGETS UNDER THEM ──

    var pLeads = target.leadTarget > 0 ? parseFloat((achieved.leads / target.leadTarget * 100).toFixed(1)) : 0;
    var pRev = target.revenueTarget > 0 ? parseFloat((achieved.revenue / target.revenueTarget * 100).toFixed(1)) : 0;

   // Get budget for TL rows
    var budget = null;
    if (peopleList[b].role === 'Team Leader') {
      budget = getTLTargetBudget(pEmail, currentMonth);
      Logger.log('Budget for ' + pEmail + ': ' + JSON.stringify(budget));
    }

    rows.push({
      name: peopleList[b].name,
      email: pEmail,
      role: peopleList[b].role,
      teamLeader: peopleList[b].teamLeader,
      leadTarget: target.leadTarget,
      achievedLeads: achieved.leads,
      percentLeads: pLeads,
      revenueTarget: target.revenueTarget,
      achievedRevenue: achieved.revenue,
      percentRevenue: pRev,
      budget: budget
    });
  }

  // Team summary (Admin only)
  
 var teamSummary = {};
if (userRole === 'Admin') {
  // Build directly from TL rows — no dependency on BDAs existing
  for (var r = 0; r < rows.length; r++) {
    if (rows[r].role !== 'Team Leader') continue;
    teamSummary[rows[r].name] = {
      leadTarget: rows[r].leadTarget,
      achievedLeads: rows[r].achievedLeads,
      revenueTarget: rows[r].revenueTarget,
      achievedRevenue: rows[r].achievedRevenue
    };
  }
}

  return {
    rows: rows,
    teamSummary: teamSummary,
    currentMonth: currentMonth
  };
}

// ── SAVE TARGET ──
function saveTarget(callerEmail, callerRole, bdaEmail, leadTarget, revenueTarget, month) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var lookupsSheet = ss.getSheetByName('Lookups');
  var lookupsData = lookupsSheet.getDataRange().getValues();

  var targetRole = '';
  var targetName = '';
  for (var i = 1; i < lookupsData.length; i++) {
    if (lookupsData[i][1] === bdaEmail) {
      targetRole = lookupsData[i][3];
      targetName = lookupsData[i][0];
      break;
    }
  }

  if (callerRole === 'Team Leader') {
    var callerName = '';
    for (var i = 1; i < lookupsData.length; i++) {
      if (lookupsData[i][1] === callerEmail) { callerName = lookupsData[i][0]; break; }
    }
    if (targetRole !== 'BDA') {
      return { success: false, message: 'Team Leaders can only assign targets to BDAs.' };
    }
    var isMyBDA = false;
    for (var i = 1; i < lookupsData.length; i++) {
      if (lookupsData[i][1] === bdaEmail) {
        if (lookupsData[i][4] === callerEmail || lookupsData[i][4] === callerName) {
          isMyBDA = true;
        }
        break;
      }
    }
    if (!isMyBDA) return { success: false, message: 'Unauthorized — this BDA is not in your team.' };

    var budget = getTLTargetBudget(callerEmail, month);
    if (budget.tlTarget === 0) {
      return { success: false, message: 'You have no target assigned for ' + month + '. Ask Admin to assign your target first.' };
    }

    // Use budget.assignedLeads (reliable) and subtract this BDA's existing allocation
    // so that editing an existing target doesn't double-count
    var alreadyAssigned = budget.assignedLeads;
    var targetsSheetCheck = ss.getSheetByName('Targets');
    var targetsDataCheck = targetsSheetCheck.getDataRange().getValues();
    for (var i = 1; i < targetsDataCheck.length; i++) {
      if (targetsDataCheck[i][1] !== bdaEmail) continue;
      var em = String(targetsDataCheck[i][4]).length === 7 ? String(targetsDataCheck[i][4]) : '';
      if (!em && targetsDataCheck[i][4]) {
        em = Utilities.formatDate(new Date(targetsDataCheck[i][4]), Session.getScriptTimeZone(), 'yyyy-MM');
      }
      if (em === month) {
        alreadyAssigned -= (targetsDataCheck[i][2] || 0);
        break;
      }
    }

    if (alreadyAssigned + leadTarget > budget.tlTarget) {
      return {
        success: false,
        message: 'Cannot assign ' + leadTarget + ' leads. Already assigned ' + alreadyAssigned + ' to other BDAs. Your total target: ' + budget.tlTarget + '. Max here: ' + (budget.tlTarget - alreadyAssigned) + '.'
      };
    }
  }

  if (callerRole !== 'Admin' && callerRole !== 'Team Leader') {
    return { success: false, message: 'Unauthorized.' };
  }

  var targetsSheet = ss.getSheetByName('Targets');
  var targetsData = targetsSheet.getDataRange().getValues();

  for (var i = 1; i < targetsData.length; i++) {
    if (!targetsData[i][1]) continue;
    var existingMonth = String(targetsData[i][4]).length === 7 ? String(targetsData[i][4]) : '';
    if (!existingMonth && targetsData[i][4]) {
      existingMonth = Utilities.formatDate(new Date(targetsData[i][4]), Session.getScriptTimeZone(), 'yyyy-MM');
    }
    if (targetsData[i][1] === bdaEmail && existingMonth === month) {
      targetsSheet.getRange(i + 1, 3).setValue(leadTarget);
      targetsSheet.getRange(i + 1, 4).setValue(revenueTarget);
      return { success: true, message: 'Target updated successfully!' };
    }
  }

  targetsSheet.appendRow([targetName, bdaEmail, leadTarget, revenueTarget, month, targetRole]);
  return { success: true, message: 'Target assigned successfully!' };
}

// ── GET TL BUDGET ──
function getTLTargetBudget(tlEmail, month) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var targetsSheet = ss.getSheetByName('Targets');
  var lookupsSheet = ss.getSheetByName('Lookups');
  var targetsData = targetsSheet.getDataRange().getValues();
  var lookupsData = lookupsSheet.getDataRange().getValues();

  // Get TL name
  var tlName = '';
  for (var i = 1; i < lookupsData.length; i++) {
    if (lookupsData[i][1] === tlEmail) { tlName = lookupsData[i][0]; break; }
  }

  // Get TL's own target
  var tlTarget = 0;
  var tlRevenueTarget = 0;
  for (var i = 1; i < targetsData.length; i++) {
    if (!targetsData[i][1]) continue;
    if (targetsData[i][1] !== tlEmail) continue;
    var tm = String(targetsData[i][4]);
    var existingMonth = tm.length === 7 ? tm : (targetsData[i][4] ? Utilities.formatDate(new Date(targetsData[i][4]), Session.getScriptTimeZone(), 'yyyy-MM') : '');
    if (existingMonth === month) {
      tlTarget = targetsData[i][2] || 0;
      tlRevenueTarget = targetsData[i][3] || 0;
      break;
    }
  }

  // Get total already assigned to BDAs under this TL
  var assignedLeads = 0;
  var assignedRevenue = 0;
  for (var i = 1; i < targetsData.length; i++) {
    if (!targetsData[i][1]) continue;
    // Check role — column F (index 5)
    var rowRole = String(targetsData[i][5] || '').trim();
    if (rowRole !== 'BDA') continue;
    var tm2 = String(targetsData[i][4]);
    var existingMonth2 = tm2.length === 7 ? tm2 : (targetsData[i][4] ? Utilities.formatDate(new Date(targetsData[i][4]), Session.getScriptTimeZone(), 'yyyy-MM') : '');
    if (existingMonth2 !== month) continue;
    // Check if this BDA reports to this TL
    var bdaEmail = targetsData[i][1];
    for (var j = 1; j < lookupsData.length; j++) {
      if (lookupsData[j][1] === bdaEmail) {
        var reportsTo = String(lookupsData[j][4] || '').trim();
        if (reportsTo === tlEmail || reportsTo === tlName) {
          assignedLeads += (parseFloat(targetsData[i][2]) || 0);
          assignedRevenue += (parseFloat(targetsData[i][3]) || 0);
        }
        break;
      }
    }
  }

  Logger.log('TL: ' + tlEmail + ' | tlTarget: ' + tlTarget + ' | assignedLeads: ' + assignedLeads + ' | month: ' + month);

  return {
    tlTarget: tlTarget,
    tlRevenueTarget: tlRevenueTarget,
    assignedLeads: assignedLeads,
    assignedRevenue: assignedRevenue,
    remainingLeads: tlTarget - assignedLeads,
    remainingRevenue: tlRevenueTarget - assignedRevenue,
    month: month
  };
}

// ── SAVE DEAL AMOUNT ──
function saveDealAmount(leadID, amount, transactionID, callerEmail) {
  var role = getUserRole(callerEmail);
  if (role !== 'Admin' && role !== 'Team Leader' && role !== 'BDA') {
    return { success: false, message: 'Unauthorized.' };
  }

  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Leads_Master');
  var paymentSheet = ss.getSheetByName('Payments');

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var leadIDCol         = headers.indexOf('LeadID');
  var dealAmountCol     = headers.indexOf('Deal Amount');
  var assignedEmailCol  = headers.indexOf('Assigned Email');
  var leadLockCol       = headers.indexOf('Lead Lock');
  var transactionIDCol  = headers.indexOf('Transaction ID');
  var verificationCol   = headers.indexOf('Verification Status');

  if (dealAmountCol === -1) return { success: false, message: 'Deal Amount column not found.' };
  if (transactionIDCol === -1) return { success: false, message: 'Transaction ID column not found in Leads_Master. Please add it.' };
  if (verificationCol === -1) return { success: false, message: 'Verification Status column not found in Leads_Master. Please add it.' };

  // ── CROSS CHECK AGAINST PAYMENT SHEET ──
  var verificationStatus = 'Pending';
  var verificationNote = '';

  if (transactionID && paymentSheet) {
    var payData = paymentSheet.getDataRange().getValues();
    var found = false;
    for (var p = 1; p < payData.length; p++) {
      var payCrmID  = String(payData[p][6] || '').trim(); // Column G — CRM/User ID
      var payEmail  = String(payData[p][2] || '').trim(); // Column C — Executive Email
      var payAmount = parseFloat(payData[p][12]) || 0;    // Column M — Final Amount

      if (payCrmID === String(transactionID).trim()) {
        found = true;
        if (payEmail !== callerEmail) {
          verificationStatus = 'Pending';
          verificationNote = 'Email mismatch';
        } else if (Math.abs(payAmount - parseFloat(amount)) > 1) {
          verificationStatus = 'Pending';
          verificationNote = 'Amount mismatch (Payment sheet: ₹' + payAmount + ')';
        } else {
          verificationStatus = 'Verified';
        }
        break;
      }
    }
    if (!found) {
      verificationStatus = 'Pending';
      verificationNote = 'Transaction ID not found in Payment sheet';
    }
  }

for (var i = 1; i < data.length; i++) {
    var cellVal = String(data[i][leadIDCol] || '').trim().replace(/\s+/g,'');
    var searchVal = String(leadID || '').trim().replace(/\s+/g,'');
    if (cellVal === searchVal) {
      if (role === 'BDA' && data[i][assignedEmailCol] !== callerEmail) {
        return { success: false, message: 'You can only edit your own leads.' };
      }
      if (data[i][leadLockCol] === 'Y' && role !== 'Admin') {
        return { success: false, message: '🔒 Lead is locked. Only Admin can edit.' };
      }
      sheet.getRange(i + 1, dealAmountCol + 1).setValue(amount);
      sheet.getRange(i + 1, transactionIDCol + 1).setValue(transactionID || '');
      sheet.getRange(i + 1, verificationCol + 1).setValue(verificationStatus);
      return {
        success: true,
        verificationStatus: verificationStatus,
        verificationNote: verificationNote
      };
    }
  }
  return { success: false, message: 'Lead not found.' };
}

// ── SAVE LEAD PAYMENT SCREENSHOT (BDA uploads proof of payment) ──
function saveLeadPaymentScreenshot(leadID, fileBase64, mimeType, callerEmail) {
  try {
    var role = getUserRole(callerEmail);
    if (role !== 'Admin' && role !== 'Team Leader' && role !== 'BDA') {
      return { success: false, message: 'Unauthorized.' };
    }

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Leads_Master');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var leadIDCol        = headers.indexOf('LeadID');
    var assignedEmailCol = headers.indexOf('Assigned Email');
    var screenshotCol    = headers.indexOf('Payment_Screenshot_FileId');

    if (screenshotCol === -1) return { success: false, message: 'Payment_Screenshot_FileId column not found in Leads_Master.' };

    var rowIndex = -1;
    var searchID = String(leadID || '').trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][leadIDCol] || '').trim() === searchID) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { success: false, message: 'Lead not found.' };

    if (role === 'BDA' && data[rowIndex][assignedEmailCol] !== callerEmail) {
      return { success: false, message: 'You can only upload screenshots for your own leads.' };
    }

    var rootFolder = DriveApp.getRootFolder();
    var folders = rootFolder.getFoldersByName('CRM_Payment_Screenshots');
    var folder = folders.hasNext() ? folders.next() : rootFolder.createFolder('CRM_Payment_Screenshots');

    var existingFileId = data[rowIndex][screenshotCol];
    if (existingFileId) {
      try { DriveApp.getFileById(existingFileId).setTrashed(true); } catch(e) {}
    }

    var ext = mimeType === 'image/png' ? '.png' : mimeType === 'image/webp' ? '.webp' : mimeType === 'application/pdf' ? '.pdf' : '.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, searchID + '_payment' + ext);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();

    sheet.getRange(rowIndex + 1, screenshotCol + 1).setValue(fileId);

    appendActivityLog(searchID, String(data[rowIndex][headers.indexOf('Name')] || ''),
      'Payment Screenshot Uploaded', 'BDA attached proof of payment', '', 'Uploaded', callerEmail);

    return { success: true, fileId: fileId };
  } catch(e) {
    Logger.log('saveLeadPaymentScreenshot error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function queueLeadNotification(bdaEmail, leadID) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var qSheet = ss.getSheetByName('Notification_Queue');
    if (!qSheet) return;
    qSheet.appendRow([bdaEmail, leadID, new Date(), 'N']);
  } catch(e) {
    Logger.log('queueLeadNotification error: ' + e.message);
  }
}

function getNewLeadsForBDA(userEmail) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Leads_Master');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var assignedEmailCol = headers.indexOf('Assigned Email');
  var lastDateCol      = headers.indexOf('Last Date');
  var createdDateCol   = headers.indexOf('Date');
  var nameCol          = headers.indexOf('Name');
  var mobileCol        = headers.indexOf('Mobile');
  var serviceCol       = headers.indexOf('Service Required');
  var leadIDCol        = headers.indexOf('LeadID');
  var callStatusCol    = headers.indexOf('Call Status');

  var cutoff = new Date();
  cutoff.setHours(cutoff.getHours() - 24);

  var newLeads = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][assignedEmailCol] !== userEmail) continue;
    var cs = data[i][callStatusCol] || '';
    if (cs !== 'Not Called' && cs !== '') continue;
    var checkDate = new Date(data[i][createdDateCol] || data[i][lastDateCol]);
    if (checkDate < cutoff) continue;
    newLeads.push({
      leadID:  data[i][leadIDCol],
      name:    data[i][nameCol],
      mobile:  data[i][mobileCol],
      service: data[i][serviceCol]
    });
  }
  return newLeads;
}
function getNewLeadsForTL(userEmail) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Leads_Master');
  var lookupsSheet = ss.getSheetByName('Lookups');
  var data = sheet.getDataRange().getValues();
  var lookupsData = lookupsSheet.getDataRange().getValues();
  var headers = data[0];

  var assignedEmailCol = headers.indexOf('Assigned Email');
  var lastDateCol      = headers.indexOf('Last Date');
  var nameCol          = headers.indexOf('Name');
  var mobileCol        = headers.indexOf('Mobile');
  var serviceCol       = headers.indexOf('Service Required');
  var leadIDCol        = headers.indexOf('LeadID');
  var callStatusCol    = headers.indexOf('Call Status');

  // Get all BDA emails under this TL
  // Get TL name first
var tlName = '';
for (var i = 1; i < lookupsData.length; i++) {
  if (lookupsData[i][1] === userEmail) { tlName = lookupsData[i][0]; break; }
}
var myBDAs = [];
for (var i = 1; i < lookupsData.length; i++) {
  var reportsTo = String(lookupsData[i][4] || '').trim();
  if (lookupsData[i][3] === 'BDA' && (reportsTo === userEmail || reportsTo === tlName)) {
    myBDAs.push(lookupsData[i][1]);
  }
}
Logger.log('TL: ' + userEmail + ' | tlName: ' + tlName + ' | myBDAs: ' + JSON.stringify(myBDAs));
  var cutoff = new Date();
  cutoff.setHours(cutoff.getHours() - 24);

  var newLeads = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (myBDAs.indexOf(data[i][assignedEmailCol]) === -1) continue;
    var cs = data[i][callStatusCol] || '';
if (cs !== 'Not Called' && cs !== '') continue;
    var lastDate = new Date(data[i][lastDateCol]);
    if (lastDate < cutoff) continue;
    newLeads.push({
      leadID:  data[i][leadIDCol],
      name:    data[i][nameCol],
      mobile:  data[i][mobileCol],
      service: data[i][serviceCol]
    });
  }
  return newLeads;
}

function testTLNewLeads() {
  var result = getNewLeadsForTL('zoyakhantesting@gmail.com');
  Logger.log('Count: ' + result.length);
  Logger.log('Result: ' + JSON.stringify(result));
}

function testTLBDAMatch() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var lookupsSheet = ss.getSheetByName('Lookups');
  var lookupsData = lookupsSheet.getDataRange().getValues();
  
  var userEmail = 'zoyakhantesting@gmail.com';
  
  // Get TL name
  var tlName = '';
  for (var i = 1; i < lookupsData.length; i++) {
    if (lookupsData[i][1] === userEmail) { 
      tlName = lookupsData[i][0]; 
      Logger.log('Found TL: ' + tlName + ' at row ' + (i+1));
      break; 
    }
  }
  
  // Log all BDA rows and their ReportsTo values
  Logger.log('--- All BDA rows ---');
  for (var i = 1; i < lookupsData.length; i++) {
    if (lookupsData[i][3] === 'BDA') {
      Logger.log('BDA: ' + lookupsData[i][0] + 
        ' | email: ' + lookupsData[i][1] + 
        ' | reportsTo: [' + lookupsData[i][4] + ']' +
        ' | matches tlEmail: ' + (lookupsData[i][4] === userEmail) +
        ' | matches tlName: ' + (lookupsData[i][4] === tlName));
    }
  }
}
function getTLList() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Lookups');
  var data = sheet.getDataRange().getValues();
  var tls = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][3] === 'Team Leader') {
      tls.push(data[i][0]); // col A = Name
    }
  }
  return tls;
}


// ── LOG LOGIN TIME ──
function logLoginTime(email) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Attendance');

    // Set headers if sheet is empty
   if (sheet.getLastRow() === 0) {
  sheet.appendRow(['Date', 'Employee Name', 'Email', 'Role', 'Login Time', 'Logout Time', 'Duration (mins)', 'Status']);
}
// Always ensure the formula exists in H2
var h2 = sheet.getRange('H2');
if (!h2.getFormula()) {
  h2.setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(F2:F="","⚠️ Incomplete",IF((G2:G)<(1/24),"🔴 Very Short (<1min)","✅ Complete"))))');
}

    var profile = getUserProfile(email);
    var now = new Date();
    var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    sheet.appendRow([dateStr, profile.name, email, profile.role, now, '', '']);

    // Return the row number so logout can update it
    var rowNumber = sheet.getLastRow();
    return { success: true, rowNumber: rowNumber };
  } catch (e) {
    Logger.log('logLoginTime error: ' + e.message);
    return { success: false, rowNumber: null };
  }
}

// ── LOG LOGOUT TIME ──
function logLogoutTime(email, rowNumber) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Attendance');

    if (!rowNumber || rowNumber < 2) return { success: false };

    var now = new Date();
    var loginTimeCell = sheet.getRange(rowNumber, 5).getValue(); // col E = Login Time
    var loginTime = new Date(loginTimeCell);

    // Calculate duration in hours (rounded to 2 decimal places)
    var durationMins = Math.round((now - loginTime) / (1000 * 60));
    var logoutTimeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
sheet.getRange(rowNumber, 6).setValue(logoutTimeStr);
sheet.getRange(rowNumber, 7).setValue(durationMins);

    return { success: true };
  } catch (e) {
    Logger.log('logLogoutTime error: ' + e.message);
    return { success: false };
  }
}

function testLogLoginTime() {
  var result = logLoginTime('zoyakhantesting@gmail.com');
  Logger.log(JSON.stringify(result));
}

function testValidateLogin() {
  var result = validateLogin('angelsaniya2016@gmail.com', 'Angel@123');
  Logger.log(JSON.stringify(result));
}

// ── UPDATE LAST SEEN (Heartbeat) ──
function updateLastSeen(email) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Attendance');
    var data = sheet.getDataRange().getValues();
    var tz = 'Asia/Kolkata';
    var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][2] !== email) continue;  // not this employee
      if (data[i][5]) continue;            // already logged out

      // Only update if the login row is from TODAY
      var rowDate = '';
      if (data[i][0] instanceof Date) {
        rowDate = Utilities.formatDate(new Date(data[i][0]), tz, 'yyyy-MM-dd');
      } else {
        rowDate = String(data[i][0] || '').substring(0, 10);
      }

      if (rowDate !== todayStr) return { success: false }; // login was yesterday or older — skip

      var lastSeenStr = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
      sheet.getRange(i + 1, 9).setValue(lastSeenStr);
      return { success: true };
    }
    return { success: false };
  } catch(e) {
    Logger.log('updateLastSeen error: ' + e.message);
    return { success: false };
  }
}

// ── FILL INCOMPLETE ATTENDANCE (Nightly) ──
function fillIncompleteAttendance() {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var attSheet = ss.getSheetByName('Attendance');
    var attData = attSheet.getDataRange().getValues();
    var tz = 'Asia/Kolkata';
    var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    for (var r = 1; r < attData.length; r++) {
      if (!attData[r][0]) continue;   // skip empty rows
      if (attData[r][5]) continue;    // skip if Logout already filled

      // Normalize row date
      var rowDate = '';
      if (attData[r][0] instanceof Date) {
        rowDate = Utilities.formatDate(new Date(attData[r][0]), tz, 'yyyy-MM-dd');
      } else {
        rowDate = String(attData[r][0] || '').substring(0, 10);
      }
      if (rowDate !== today) continue; // only process today's open rows

      var lastSeen = attData[r][8]; // col I = Last Seen
      if (!lastSeen) continue;      // no heartbeat — skip

      // KEY FIX: verify Last Seen is also from today
      var lastSeenDate = '';
      try {
        var lsDt = new Date(lastSeen);
        lastSeenDate = Utilities.formatDate(lsDt, tz, 'yyyy-MM-dd');
      } catch(e) {}
      if (lastSeenDate !== today) continue; // Last Seen is from a different day — skip

      var loginTime = new Date(attData[r][4]);
      var logoutTime = new Date(lastSeen);
      var durationMins = Math.round((logoutTime - loginTime) / (1000 * 60));
      if (durationMins < 0) continue; // bad data — skip

      attSheet.getRange(r + 1, 6).setValue(lastSeen);
      attSheet.getRange(r + 1, 7).setValue(durationMins);
      attSheet.getRange(r + 1, 8).setValue('Estimated');
    }

    Logger.log('fillIncompleteAttendance complete');
  } catch(e) {
    Logger.log('fillIncompleteAttendance error: ' + e.message);
  }
}

function getPortalHTML(email, token, attRow) {
  var role = getUserRole(email);
  var leads = getMyLeads(email, role);
  var userProfile = getUserProfile(email);
  var bdaList = getBDAList(email, role);

  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = email;
  template.role = role;
  template.token = token;
  template.attRow = attRow;
  template.userProfile = JSON.stringify(userProfile);
  template.leadsJSON = JSON.stringify(leads);
  template.bdaListJSON = JSON.stringify(bdaList);
  template.tlListJSON = JSON.stringify(getTLList());

  var newLeadsData = [];
  try {
    if (role === 'BDA') newLeadsData = getNewLeadsForBDA(email) || [];
    else if (role === 'Team Leader') newLeadsData = getNewLeadsForTL(email) || [];
  } catch(err) {}
  template.newLeadsJSON = JSON.stringify(newLeadsData);

  return template.evaluate().getContent();
}

function addEmployee(name, email, password, role, reportsTo, callerEmail) {
  try {
    var callerRole = getUserRole(callerEmail);
    if (callerRole !== 'Admin') {
      return { success: false, message: 'Only Admin can add employees.' };
    }

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();

    // Check if email already exists
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === email) {
        return { success: false, message: 'An employee with this email already exists.' };
      }
    }

    // Auto generate Employee ID
    var maxID = 0;
    for (var i = 1; i < data.length; i++) {
      var empID = String(data[i][8] || '');
      if (empID.startsWith('NT-')) {
        var num = parseInt(empID.replace('NT-', '')) || 0;
        if (num > maxID) maxID = num;
      }
    }
    var newEmpID = 'NT-' + String(maxID + 1).padStart(3, '0');

    // Append new row
    // A: Name, B: Email, C: Last Assigned, D: Role, E: ReportsTo, F: Password, G: Token, H: Token Expiry, I: Employee ID
    sheet.appendRow([name, email, '', role, reportsTo, password, '', '', newEmpID]);

  // ── SEND WELCOME EMAIL ──
    try {
      var portalURL = 'https://script.google.com/macros/s/AKfycbwU-nk5mdbbpru0GEAoZXa16CLqw2wwik-HxUiZj1eoResk6jc6vID5wSAWpc4hSCZ9hA/exec';
      
      // Get manager name + email
      var managerName = 'Admin';
      var managerEmail = 'developernotifytechai@gmail.com';
      if (reportsTo) {
        for (var m = 1; m < data.length; m++) {
          if (data[m][0] === reportsTo || data[m][1] === reportsTo) {
            managerName = data[m][0];
            managerEmail = data[m][1];
            break;
          }
        }
      }

      var subject = 'Welcome to NotifyTechAI — Your CRM Portal Access';
      var body =
        'Hi ' + name + ',\n\n' +
        'Welcome to the team at NotifyTechAI! 🎉\n\n' +
        'Your CRM portal account has been created. Here are your login details:\n\n' +
        '📧 Email: ' + email + '\n' +
        '🔑 Password: ' + password + '\n' +
        '🌐 Portal URL: ' + portalURL + '\n\n' +
        'Your Role: ' + role + '\n' +
        'You report to: ' + managerName + ' (' + managerEmail + ')\n\n' +
        'Please log in, and if you would like to change your password or encounter any issues, kindly reach out to your team leader or manager for assistance.\n\n' +
        'Looking forward to working with you!\n\n' +
        'Best regards,\n' +
        'Admin — NotifyTechAI\n' +
        'developernotifytechai@gmail.com';

      MailApp.sendEmail(email, subject, body);
    } catch(mailErr) {
      Logger.log('Welcome email failed: ' + mailErr.message);
      // Don't fail the whole function if email fails
    }

    // ── RE-COLOR LOOKUPS AFTER NEW EMPLOYEE ADDED ──
    try { colorLookups(); } catch(e) { Logger.log('colorLookups error: ' + e.message); }

    return { success: true, message: 'Employee added successfully! ID: ' + newEmpID };

  } catch(e) {
    Logger.log('addEmployee error: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  }
}
function testUpdateLastSeen() {
  var result = updateLastSeen('ravenmistiek2002@gmail.com');
  Logger.log(JSON.stringify(result));
}

// ── EDIT EMPLOYEE ──
function editEmployee(targetEmail, newName, newPassword, newRole, newReportsTo, callerEmail) {
  try {
    var callerRole = getUserRole(callerEmail);
    if (callerRole !== 'Admin') {
      return { success: false, message: 'Only Admin can edit employees.' };
    }

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === targetEmail) {
        if (newName) sheet.getRange(i + 1, 1).setValue(newName);
        if (newPassword) sheet.getRange(i + 1, 6).setValue(newPassword);
        if (newRole) sheet.getRange(i + 1, 4).setValue(newRole);
        if (newReportsTo !== undefined) sheet.getRange(i + 1, 5).setValue(newReportsTo);
        return { success: true, message: 'Employee updated successfully!' };
      }
    }
    return { success: false, message: 'Employee not found.' };
  } catch(e) {
    Logger.log('editEmployee error: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ── DELETE EMPLOYEE (soft delete) ──
function deleteEmployee(targetEmail, callerEmail) {
  try {
    var callerRole = getUserRole(callerEmail);
    if (callerRole !== 'Admin') {
      return { success: false, message: 'Only Admin can delete employees.' };
    }
    if (targetEmail === callerEmail) {
      return { success: false, message: 'You cannot delete your own account.' };
    }
    if (targetEmail === 'developernotifytechai@gmail.com' || targetEmail === 'saniyakhan1709@gmail.com') {
      return { success: false, message: 'Cannot delete Admin accounts.' };
    }

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === targetEmail) {
        sheet.getRange(i + 1, 10).setValue('Inactive'); // col J = Status
        sheet.getRange(i + 1, 7).setValue('');          // clear token
        sheet.getRange(i + 1, 8).setValue('');          // clear expiry
        return { success: true, message: 'Employee deactivated successfully.' };
      }
    }
    return { success: false, message: 'Employee not found.' };
  } catch(e) {
    Logger.log('deleteEmployee error: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ── GET EMPLOYEE LIST (Admin only) ──
function getEmployeeList(callerEmail) {
  try {
    var callerRole = getUserRole(callerEmail);
    if (callerRole !== 'Admin') {
      return { success: false, message: 'Only Admin can view employee list.' };
    }

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();
    var employees = [];

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || !data[i][1]) continue;
      if (data[i][3] === 'Admin') continue;
      if (data[i][9] === 'Inactive') continue;
      employees.push({
        name: data[i][0],
        email: data[i][1],
        role: data[i][3],
        reportsTo: data[i][4] || '',
        employeeID: data[i][8] || ''
      });
    }
    return { success: true, employees: employees };
  } catch(e) {
    Logger.log('getEmployeeList error: ' + e.message);
    return { success: false, message: 'Error: ' + e.message };
  }
}
// ── SAVE MONTHLY SNAPSHOT (runs 1st of each month) ──
function saveMonthlySnapshot() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var leadsSheet    = ss.getSheetByName('Leads_Master');
  var lookupsSheet  = ss.getSheetByName('Lookups');
  var targetsSheet  = ss.getSheetByName('Targets');
  var historySheet  = ss.getSheetByName('Performance_Sheet');

  if (!historySheet) { Logger.log('Performance_Sheet sheet not found!'); return; }

  if (historySheet.getLastRow() === 0) {
    historySheet.appendRow([
      'Month', 'BDA Name', 'Email', 'Role', 'Team Leader',
      'Leads Assigned', 'Won', 'Revenue', 'Calls Made', 'Conversion %',
      'Leads Target Hit', 'Revenue Target Hit'
    ]);
  }

  var leadsData    = leadsSheet.getDataRange().getValues();
  var leadsHeaders = leadsData[0];
  var lookupsData  = lookupsSheet.getDataRange().getValues();
  var targetsData  = targetsSheet.getDataRange().getValues();

  var assignedEmailCol = leadsHeaders.indexOf('Assigned Email');
  var pipelineCol      = leadsHeaders.indexOf('Pipeline & Stage');
  var dealAmountCol    = leadsHeaders.indexOf('Deal Amount');
  var lastDateCol      = leadsHeaders.indexOf('Last Date');
  var callStatusCol    = leadsHeaders.indexOf('Call Status');
  var assignDateCol    = leadsHeaders.indexOf('Assign Date');

  var now = new Date();
  var lastMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var lastMonth = Utilities.formatDate(lastMonthDate, Session.getScriptTimeZone(), 'yyyy-MM');

  // Check duplicate
  var existingData = historySheet.getDataRange().getValues();
  for (var i = 1; i < existingData.length; i++) {
    var existingVal = existingData[i][0];
    var existingStr = '';
    var es = String(existingVal);
    if (/^\d{4}-\d{2}$/.test(es)) {
      existingStr = es;
    } else if (existingVal) {
      try {
        var ed = new Date(existingVal);
        if (!isNaN(ed)) existingStr = Utilities.formatDate(ed, Session.getScriptTimeZone(), 'yyyy-MM');
      } catch(e) {}
    }
    if (existingStr === lastMonth) {
      Logger.log('Snapshot for ' + lastMonth + ' already exists. Skipping.');
      return;
    }
  }

  // ── BUILD STATS — based on Last Date (when the action happened) ──
  var stats = {};
  for (var i = 1; i < leadsData.length; i++) {
    if (!leadsData[i][0]) continue;
    var bEmail = leadsData[i][assignedEmailCol];
    if (!bEmail) continue;
    if (!stats[bEmail]) stats[bEmail] = { assigned: 0, won: 0, revenue: 0, calls: 0 };

    var lastDate = leadsData[i][lastDateCol] ? new Date(leadsData[i][lastDateCol]) : null;
    if (!lastDate) continue;
    var leadMonth = Utilities.formatDate(lastDate, Session.getScriptTimeZone(), 'yyyy-MM');
    if (leadMonth !== lastMonth) continue;

    // Count every lead touched this month
    stats[bEmail].assigned++;

    if (leadsData[i][pipelineCol] === 'Won') {
      stats[bEmail].won++;
      stats[bEmail].revenue += parseFloat(leadsData[i][dealAmountCol]) || 0;
    }
    var cs = leadsData[i][callStatusCol] || '';
    if (cs !== 'Not Called' && cs !== '') stats[bEmail].calls++;
  }

  // ── BUILD LOOKUP MAPS ──
  var nameMap = {}, roleMap = {}, tlMap = {};
  for (var i = 1; i < lookupsData.length; i++) {
    if (!lookupsData[i][1]) continue;
    nameMap[lookupsData[i][1]] = lookupsData[i][0];
    roleMap[lookupsData[i][1]] = lookupsData[i][3];
    tlMap[lookupsData[i][1]]   = lookupsData[i][4] || '';
  }

  // ── ENSURE ALL ACTIVE EMPLOYEES ARE INCLUDED even if no activity ──
  for (var i = 1; i < lookupsData.length; i++) {
    if (!lookupsData[i][1]) continue;
    if (lookupsData[i][9] === 'Inactive') continue;
    var empRole = lookupsData[i][3];
    if (empRole !== 'BDA' && empRole !== 'Team Leader') continue;
    var empEmail = lookupsData[i][1];
    if (!stats[empEmail]) stats[empEmail] = { assigned: 0, won: 0, revenue: 0, calls: 0 };
  }

  // ── TARGET MAP ──
  var targetMap = {};
  for (var i = 1; i < targetsData.length; i++) {
    if (!targetsData[i][1]) continue;
    var tm = String(targetsData[i][4]);
    var rowMonth = tm.length === 7 ? tm : '';
    if (!rowMonth && targetsData[i][4]) {
      try { rowMonth = Utilities.formatDate(new Date(targetsData[i][4]), Session.getScriptTimeZone(), 'yyyy-MM'); } catch(e) {}
    }
    if (rowMonth === lastMonth) {
      targetMap[targetsData[i][1]] = {
        leadTarget:    targetsData[i][2] || 0,
        revenueTarget: targetsData[i][3] || 0
      };
    }
  }

 
 // ── WRITE ROWS ──
  for (var email in stats) {
    var s = stats[email];
    var conversion = s.assigned > 0 ? parseFloat((s.won / s.assigned * 100).toFixed(1)) : 0;
    var tgt = targetMap[email] || { leadTarget: 0, revenueTarget: 0 };
    var leadsHit   = tgt.leadTarget > 0 ? (s.won >= tgt.leadTarget ? '✅ Yes' : '❌ No') : '—';
    var revenueHit = tgt.revenueTarget > 0 ? (s.revenue >= tgt.revenueTarget ? '✅ Yes' : '❌ No') : '—';

    var newRow = historySheet.getLastRow() + 1;
    historySheet.appendRow([
      '',
      nameMap[email] || email,
      email,
      roleMap[email] || 'BDA',
      tlMap[email] || '',
      s.assigned,
      s.won,
      s.revenue,
      s.calls,
      conversion,
      leadsHit,
      revenueHit,
      tgt.leadTarget,
      tgt.revenueTarget
    ]);
    historySheet.getRange(newRow, 1).setNumberFormat('@').setValue(lastMonth);
  }

  Logger.log('✅ Monthly snapshot saved for ' + lastMonth);
}

// ── GET HISTORICAL PERFORMANCE ──
function getHistoricalPerformance(callerEmail, callerRole) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var historySheet = ss.getSheetByName('Performance_Sheet');
  if (!historySheet || historySheet.getLastRow() < 2) return [];

  var data = historySheet.getDataRange().getValues();
  var lookupsSheet = ss.getSheetByName('Lookups');
  var lookupsData  = lookupsSheet.getDataRange().getValues();

  var tlName = '';
  if (callerRole === 'Team Leader') {
    for (var i = 1; i < lookupsData.length; i++) {
      if (lookupsData[i][1] === callerEmail) { tlName = lookupsData[i][0]; break; }
    }
  }

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var rowEmail = data[i][2];
    var rowTL    = String(data[i][4] || '').trim();
    var include  = false;

    if (callerRole === 'Admin') {
      include = true;
    } else if (callerRole === 'Team Leader') {
      if (rowEmail === callerEmail) include = true;
      else if (rowTL === callerEmail || rowTL === tlName) include = true;
    } else if (callerRole === 'BDA') {
      if (rowEmail === callerEmail) include = true;
    }

    if (include) {
      rows.push({
        month:         String(data[i][0]),
        name:          data[i][1],
        email:         rowEmail,
        role:          data[i][3],
        teamLeader:    data[i][4],
        assigned:      data[i][5],
        won:           data[i][6],
        revenue:       data[i][7],
        calls:         data[i][8],
        conversion:    data[i][9],
        leadsHit:      data[i][10],
        revenueHit:    data[i][11],
        leadTarget:    data[i][12] || 0,
        revenueTarget: data[i][13] || 0
      });
    }
  }
  return rows;
}


// ── CREATE SNAPSHOT TRIGGER (run once manually) ──
function createSnapshotTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'saveMonthlySnapshot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('saveMonthlySnapshot')
    .timeBased()
    .onMonthDay(1)
    .atHour(7)
    .create();
  Logger.log('✅ Snapshot trigger created!');
}

// ── CHAT: SEND MESSAGE ──
function sendChatMessage(senderEmail, message, channel) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('ChatLog');
    if (!sheet) return { success: false, message: 'ChatLog sheet not found' };

    // ── ACCESS CONTROL ──
    var role = getUserRole(senderEmail);
    if (channel === 'tl-admin' && role === 'BDA') {
      return { success: false, message: 'Access denied' };
    }
    if (channel === 'general' && role === 'BDA') {
      return { success: false, message: 'BDAs cannot send in general' };
    }
    if (channel.indexOf('team-') === 0) {
      // BDA can only send in their own team channel
      if (role === 'BDA') {
        var lookupsSheet = ss.getSheetByName('Lookups');
        var lookupsData = lookupsSheet.getDataRange().getValues();
        for (var i = 1; i < lookupsData.length; i++) {
          if (lookupsData[i][1] === senderEmail) {
            var reportsTo = lookupsData[i][4].toString().trim().toLowerCase().replace(/\s+/g, '');
            var allowedChannel = 'team-' + reportsTo;
            if (channel !== allowedChannel) return { success: false, message: 'Access denied' };
            break;
          }
        }
      }
    }

   var profile = getUserProfile(senderEmail);
    var senderName = profile.name || senderEmail;
    var senderRole = role || 'BDA';

    // Extract @mentions from message
    var mentionMatches = message.match(/@([A-Za-z0-9 ]+?)(?=\s|$|[^A-Za-z0-9 ])/g) || [];
    var mentions = mentionMatches.map(function(m) { return m.replace('@', '').trim(); }).join(',');

    sheet.appendRow([new Date(), senderEmail, senderName, message, channel || 'general', senderRole, '', mentions]);
    return { success: true };
  } catch(e) {
    Logger.log('sendChatMessage error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── CHAT: GET MESSAGES ──
function getChatMessages(channel, callerEmail) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('ChatLog');
    if (!sheet || sheet.getLastRow() < 2) return JSON.stringify([]);

    // ── ACCESS CONTROL ──
    var role = callerEmail ? getUserRole(callerEmail) : 'BDA';
    if (channel === 'tl-admin' && role === 'BDA') return JSON.stringify([]);

    var data = sheet.getDataRange().getValues();
    var messages = [];
    for (var i = 1; i < data.length; i++) {
      var msgChannel = data[i][4] || 'general';
      if (msgChannel !== channel) continue;
      var reactionsRaw = data[i][6] || '';
      var reactions = {};
      try { if (reactionsRaw) reactions = JSON.parse(reactionsRaw); } catch(e) {}
      var mentions = data[i][7] ? String(data[i][7]).split(',').map(function(m){ return m.trim(); }) : [];
      messages.push({
        rowIndex: i + 1,
        timestamp: data[i][0] ? new Date(data[i][0]).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }) : '',
        senderEmail: data[i][1],
        senderName: data[i][2] || data[i][1],
        message: data[i][3],
        senderRole: data[i][5] || 'BDA',
        reactions: reactions,
        mentions: mentions
      });
    }
    return JSON.stringify(messages.slice(-50));
  } catch(e) {
    Logger.log('getChatMessages error: ' + e.message);
    return JSON.stringify([]);
  }
}

// ── GET MY CHANNELS ──
function getMyChannels(callerEmail) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Lookups');
  var data = sheet.getDataRange().getValues();

  var role = getUserRole(callerEmail);
  var channels = [];

  // Everyone gets general
  channels.push({ id: 'general', label: '# general' });

  if (role === 'Admin') {
    // Admin only gets tl-admin — no team channels
    channels.push({ id: 'tl-admin', label: '# tl-admin' });

  } else if (role === 'Team Leader') {
    // TL gets tl-admin + their own team channel
    channels.push({ id: 'tl-admin', label: '# tl-admin' });
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === callerEmail) {
        var tlName = data[i][0].toString().trim().toLowerCase().replace(/\s+/g, '');
        channels.push({ id: 'team-' + tlName, label: '# team-' + tlName });
        break;
      }
    }

  } else if (role === 'BDA') {
    // BDA gets their team channel — resolve reportsTo to a name
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === callerEmail) {
        var reportsToRaw = data[i][4].toString().trim();
        var tlName = '';

        // Check if reportsTo is an email — if so resolve to name
        if (reportsToRaw.indexOf('@') !== -1) {
          for (var j = 1; j < data.length; j++) {
            if (data[j][1] === reportsToRaw) {
              tlName = data[j][0].toString().trim().toLowerCase().replace(/\s+/g, '');
              break;
            }
          }
        } else {
          tlName = reportsToRaw.toLowerCase().replace(/\s+/g, '');
        }

        if (tlName) channels.push({ id: 'team-' + tlName, label: '# team-' + tlName });
        break;
      }
    }
  }

  return JSON.stringify(channels);
}

// ── GET ROLE FOR CHAT COLOR ──
function getChatSenderRole(email) {
  return getUserRole(email) || 'BDA';
}


function addReaction(rowIndex, emoji, callerEmail, channel) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('ChatLog');
    if (!sheet) return { success: false };

    var role = getUserRole(callerEmail);
    var profile = getUserProfile(callerEmail);
    var callerName = profile.name || callerEmail;

    // BDAs in general can only react if they're mentioned
    if (channel === 'general' && role === 'BDA') {
      var mentionsCell = sheet.getRange(rowIndex, 8).getValue();
      var mentions = mentionsCell ? String(mentionsCell).split(',').map(function(m){ return m.trim(); }) : [];
      if (mentions.indexOf(callerName) === -1) {
        return { success: false, message: 'You can only react to messages where you are mentioned.' };
      }
    }

    var cell = sheet.getRange(rowIndex, 7);
    var raw = cell.getValue() || '';
    var reactions = {};
    try { if (raw) reactions = JSON.parse(raw); } catch(e) {}

    if (!reactions[emoji]) reactions[emoji] = [];
    var idx = reactions[emoji].indexOf(callerName);
    if (idx === -1) {
      reactions[emoji].push(callerName); // add reaction
    } else {
      reactions[emoji].splice(idx, 1);   // toggle off
      if (reactions[emoji].length === 0) delete reactions[emoji];
    }

    cell.setValue(JSON.stringify(reactions));
    return { success: true, reactions: reactions };
  } catch(e) {
    Logger.log('addReaction error: ' + e.message);
    return { success: false };
  }
}

function getChatMembers(channel, callerEmail) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();
    var members = [];

    if (channel === 'general') {
      // Everyone — Admin, TLs, BDAs
      for (var i = 1; i < data.length; i++) {
        if (!data[i][0] || !data[i][1]) continue;
        if (data[i][9] === 'Inactive') continue;
        if (data[i][3] === 'Admin') continue; // skip admin from tag list
        members.push(data[i][0]);
      }
      members.unshift('Admin');

    } else if (channel === 'tl-admin') {
      // TLs + Admin only
      members.push('Admin');
      for (var i = 1; i < data.length; i++) {
        if (!data[i][0] || !data[i][1]) continue;
        if (data[i][9] === 'Inactive') continue;
        if (data[i][3] === 'Team Leader') members.push(data[i][0]);
      }

    } else if (channel.indexOf('team-') === 0) {
      // TL of this team + their BDAs
      var channelSlug = channel.replace('team-', '');
      for (var i = 1; i < data.length; i++) {
        if (!data[i][0] || !data[i][1]) continue;
        if (data[i][9] === 'Inactive') continue;
        var nameSlug = data[i][0].toString().trim().toLowerCase().replace(/\s+/g, '');
        if (data[i][3] === 'Team Leader' && nameSlug === channelSlug) {
          members.unshift(data[i][0]); // TL first
        } else if (data[i][3] === 'BDA') {
          var reportsTo = String(data[i][4] || '').trim().toLowerCase().replace(/\s+/g, '');
          if (reportsTo === channelSlug) members.push(data[i][0]);
        }
      }
    }

    return JSON.stringify(members);
  } catch(e) {
    Logger.log('getChatMembers error: ' + e.message);
    return JSON.stringify([]);
  }
}

// ── DM: GET CONTACTS ──
function getDMContacts(callerEmail) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();

    var role = getUserRole(callerEmail);
    var contacts = [];

    // Find caller's own row
    var callerName = '';
    var callerReportsTo = '';
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === callerEmail) {
        callerName = data[i][0];
        callerReportsTo = String(data[i][4] || '').trim();
        break;
      }
    }

   if (role === 'Admin') {
      // Admin can DM: TLs only
      for (var i = 1; i < data.length; i++) {
        if (!data[i][0] || !data[i][1]) continue;
        if (data[i][9] === 'Inactive') continue;
        if (data[i][1] === callerEmail) continue;
        if (data[i][3] !== 'Team Leader') continue;
        contacts.push({ name: data[i][0], email: data[i][1], role: data[i][3] });
      }

    } else if (role === 'Team Leader') {
      // TL can DM: their BDAs + Admin
      contacts.push({ name: 'Developer', email: 'developernotifytechai@gmail.com', role: 'Admin' });
      for (var i = 1; i < data.length; i++) {
        if (!data[i][0] || !data[i][1]) continue;
        if (data[i][9] === 'Inactive') continue;
        if (data[i][3] !== 'BDA') continue;
        var reportsTo = String(data[i][4] || '').trim();
        if (reportsTo === callerEmail || reportsTo === callerName) {
          contacts.push({ name: data[i][0], email: data[i][1], role: 'BDA' });
        }
      }

    } else if (role === 'BDA') {
      // BDA can DM: their TL only
      for (var i = 1; i < data.length; i++) {
        if (!data[i][0] || !data[i][1]) continue;
        if (data[i][9] === 'Inactive') continue;
        // Match by name or email
        if (data[i][0] === callerReportsTo || data[i][1] === callerReportsTo) {
          contacts.push({ name: data[i][0], email: data[i][1], role: data[i][3] });
          break;
        }
      }
    }

    return JSON.stringify(contacts);
  } catch(e) {
    Logger.log('getDMContacts error: ' + e.message);
    return JSON.stringify([]);
  }
}

// ── DM: GET MESSAGES ──
function getDMMessages(callerEmail, recipientEmail) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('ChatLog');
    if (!sheet || sheet.getLastRow() < 2) return JSON.stringify([]);

    var data = sheet.getDataRange().getValues();
    var messages = [];

    for (var i = 1; i < data.length; i++) {
      var isDM = data[i][8]; // col I = IsDM
      if (String(isDM).toUpperCase() !== 'TRUE') continue;

      var sender = data[i][1];
      var recipient = data[i][9]; // col J = RecipientEmail

      // Only show messages between these two people
      var isMatch = (sender === callerEmail && recipient === recipientEmail) ||
                    (sender === recipientEmail && recipient === callerEmail);
      if (!isMatch) continue;

      var reactionsRaw = data[i][6] || '';
      var reactions = {};
      try { if (reactionsRaw) reactions = JSON.parse(reactionsRaw); } catch(e) {}

      messages.push({
        rowIndex: i + 1,
        timestamp: data[i][0] ? new Date(data[i][0]).toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }) : '',
        senderEmail: sender,
        senderName: data[i][2] || sender,
        message: data[i][3],
        senderRole: data[i][5] || 'BDA',
        reactions: reactions,
        mentions: []
      });
    }
    return JSON.stringify(messages.slice(-50));
  } catch(e) {
    Logger.log('getDMMessages error: ' + e.message);
    return JSON.stringify([]);
  }
}

// ── DM: SEND MESSAGE ──
function sendDM(senderEmail, recipientEmail, message) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('ChatLog');
    if (!sheet) return { success: false, message: 'ChatLog sheet not found' };

    var profile = getUserProfile(senderEmail);
    var senderName = profile.name || senderEmail;
    var role = getUserRole(senderEmail) || 'BDA';

    // cols: A=Timestamp, B=SenderEmail, C=SenderName, D=Message, E=Channel,
    //       F=SenderRole, G=Reactions, H=Mentions, I=IsDM, J=RecipientEmail
    sheet.appendRow([new Date(), senderEmail, senderName, message, 'DM', role, '', '', true, recipientEmail]);
    return { success: true };
  } catch(e) {
    Logger.log('sendDM error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── DASHBOARD TARGET CARDS + RANKINGS ──
function getDashboardTargetCards(userEmail, userRole) {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var leadsSheet   = ss.getSheetByName('Leads_Master');
  var targetsSheet = ss.getSheetByName('Targets');
  var lookupsSheet = ss.getSheetByName('Lookups');

  var leadsData    = leadsSheet.getDataRange().getValues();
  var targetsData  = targetsSheet.getDataRange().getValues();
  var lookupsData  = lookupsSheet.getDataRange().getValues();
  var leadsHeaders = leadsData[0];

  var assignedEmailCol = leadsHeaders.indexOf('Assigned Email');
  var pipelineCol      = leadsHeaders.indexOf('Pipeline & Stage');
  var dealAmountCol    = leadsHeaders.indexOf('Deal Amount');
  var dateCol          = leadsHeaders.indexOf('Last Date');

  var now = new Date();
  var currentMonth = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');

  // ── BUILD LOOKUP MAPS ──
  var nameMap = {}, roleMap = {}, reportsToMap = {};
  // tlNameMap: TL name -> TL email
  var tlNameToEmail = {};
  // bdaEmailToTLName: BDA email -> TL name
  var bdaEmailToTLName = {};

  for (var i = 1; i < lookupsData.length; i++) {
    if (!lookupsData[i][1]) continue;
    var lEmail = lookupsData[i][1];
    nameMap[lEmail]      = lookupsData[i][0];
    roleMap[lEmail]      = lookupsData[i][3];
    reportsToMap[lEmail] = String(lookupsData[i][4] || '').trim();
    if (lookupsData[i][3] === 'Team Leader') {
      tlNameToEmail[lookupsData[i][0]] = lEmail;
    }
  }

  // Build BDA -> TL name map
  for (var i = 1; i < lookupsData.length; i++) {
    if (!lookupsData[i][1]) continue;
    if (lookupsData[i][3] !== 'BDA') continue;
    var bdaEmail  = lookupsData[i][1];
    var reportsTo = String(lookupsData[i][4] || '').trim();
    // reportsTo could be TL name or TL email
    var tlName = '';
    if (reportsTo.indexOf('@') !== -1) {
      tlName = nameMap[reportsTo] || reportsTo;
    } else {
      tlName = reportsTo; // it's already a name
    }
    bdaEmailToTLName[bdaEmail] = tlName;
  }

  var callerName = nameMap[userEmail] || '';

  // ── BUILD ACHIEVEMENT MAP from Leads_Master ──
  var achieveMap = {};
  for (var i = 1; i < leadsData.length; i++) {
    if (!leadsData[i][0]) continue;
    if (leadsData[i][pipelineCol] !== 'Won') continue;
    var ld = leadsData[i][dateCol];
    if (!ld) continue;
    var ldMonth = Utilities.formatDate(new Date(ld), Session.getScriptTimeZone(), 'yyyy-MM');
    if (ldMonth !== currentMonth) continue;
    var em = leadsData[i][assignedEmailCol];
    if (!achieveMap[em]) achieveMap[em] = { leads: 0, revenue: 0 };
    achieveMap[em].leads++;
    achieveMap[em].revenue += parseFloat(leadsData[i][dealAmountCol]) || 0;
  }

  // ── AGGREGATE FROM TARGETS SHEET DIRECTLY ──
  var totalLeadTarget = 0, totalLeadAchieved = 0;
  var totalRevTarget  = 0, totalRevAchieved  = 0;
  var teamBreakdown   = {}; // tlName -> { leadTarget, leadAchieved, revTarget, revAchieved }

  for (var i = 1; i < targetsData.length; i++) {
    if (!targetsData[i][1]) continue;

    var tEmail = targetsData[i][1];

    // ── MONTH CHECK ──
    var tRawVal = targetsData[i][4];
    var tRowMonth = '';
    var tRawStr = String(tRawVal).trim();
    if (/^\d{4}-\d{2}$/.test(tRawStr)) {
      tRowMonth = tRawStr;
    } else if (tRawVal) {
      try { tRowMonth = Utilities.formatDate(new Date(tRawVal), Session.getScriptTimeZone(), 'yyyy-MM'); } catch(e) {}
    }
    if (tRowMonth !== currentMonth) continue;

    // ── ROLE CHECK: skip TLs and Admins using Lookups ──
    var tLookupRole = roleMap[tEmail] || '';
    if (tLookupRole === 'Admin') continue;
    // Include both BDA and Team Leader targets in total
    // BDAs (active or deleted) pass through

    // ── SCOPE FILTER for TL and BDA roles ──
    if (userRole === 'Team Leader') {
  // Include TL's own row + BDAs that report to this TL
  if (tEmail !== userEmail) {
    var rt = reportsToMap[tEmail] || '';
    var rtResolved = rt.indexOf('@') !== -1 ? nameMap[rt] || rt : rt;
    if (rtResolved !== callerName && rt !== userEmail) continue;
  }
} else if (userRole === 'BDA') {
      if (tEmail !== userEmail) continue;
    }
    // Admin: include all

    var tLeadTarget = parseFloat(targetsData[i][2]) || 0;
    var tRevTarget  = parseFloat(targetsData[i][3]) || 0;
    var ach = achieveMap[tEmail] || { leads: 0, revenue: 0 };

    totalLeadTarget   += tLeadTarget;
    totalLeadAchieved += ach.leads;
    totalRevTarget    += tRevTarget;
    totalRevAchieved  += ach.revenue;

    // ── TEAM BREAKDOWN (Admin only) ──
    if (userRole === 'Admin') {
      // Resolve TL name for this BDA
      var bdaTLName = bdaEmailToTLName[tEmail] || '';
      if (!bdaTLName) {
        // BDA deleted from Lookups — try to find TL from Targets sheet name column
        // We can't know their TL, put under 'Unassigned'
        bdaTLName = 'Unassigned';
      }
      if (!teamBreakdown[bdaTLName]) {
        teamBreakdown[bdaTLName] = { leadTarget: 0, leadAchieved: 0, revTarget: 0, revAchieved: 0 };
      }
      teamBreakdown[bdaTLName].leadTarget   += tLeadTarget;
      teamBreakdown[bdaTLName].leadAchieved += ach.leads;
      teamBreakdown[bdaTLName].revTarget    += tRevTarget;
      teamBreakdown[bdaTLName].revAchieved  += ach.revenue;
    }
  }

  // ── RANKINGS: Top 5 active BDAs by revenue ──
  var rankings = [];
  for (var i = 1; i < lookupsData.length; i++) {
    if (!lookupsData[i][1] || lookupsData[i][3] !== 'BDA') continue;
    if (lookupsData[i][9] === 'Inactive') continue;
    var bdaEmail = lookupsData[i][1];
    var bdaName  = lookupsData[i][0];
    var bdaTL    = bdaEmailToTLName[bdaEmail] || '';

    if (userRole === 'Team Leader') {
      var rt = reportsToMap[bdaEmail] || '';
      var rtR = rt.indexOf('@') !== -1 ? nameMap[rt] || rt : rt;
      if (rtR !== callerName && rt !== userEmail) continue;
    } else if (userRole === 'BDA') {
      continue; // no rankings for BDA
    }

    var ach = achieveMap[bdaEmail] || { leads: 0, revenue: 0 };
    rankings.push({ name: bdaName, revenue: ach.revenue, leads: ach.leads, tlName: bdaTL });
  }
  rankings.sort(function(a, b) { return b.revenue - a.revenue; });
  rankings = rankings.slice(0, 5);

  return {
    currentMonth:      currentMonth,
    totalLeadTarget:   totalLeadTarget,
    totalLeadAchieved: totalLeadAchieved,
    totalRevTarget:    totalRevTarget,
    totalRevAchieved:  totalRevAchieved,
    teamBreakdown:     teamBreakdown,
    rankings:          rankings
  };
}

function testTargetCards2() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var targetsSheet = ss.getSheetByName('Targets');
  var lookupsSheet = ss.getSheetByName('Lookups');
  var targetsData  = targetsSheet.getDataRange().getValues();
  var lookupsData  = lookupsSheet.getDataRange().getValues();

  var currentMonth = '2026-04';
  
  // Build roleMap
  var roleMap = {};
  for (var i = 1; i < lookupsData.length; i++) {
    if (!lookupsData[i][1]) continue;
    roleMap[lookupsData[i][1]] = lookupsData[i][3];
  }

  Logger.log('=== CHECKING EACH TARGETS ROW ===');
  var total = 0;
  for (var i = 1; i < targetsData.length; i++) {
    if (!targetsData[i][1]) continue;
    var tEmail = targetsData[i][1];
    var tRawVal = targetsData[i][4];
    var tRawStr = String(tRawVal).trim();
    var tRowMonth = '';
    if (/^\d{4}-\d{2}$/.test(tRawStr)) {
      tRowMonth = tRawStr;
    } else if (tRawVal) {
      try { tRowMonth = Utilities.formatDate(new Date(tRawVal), Session.getScriptTimeZone(), 'yyyy-MM'); } catch(e) {}
    }
    
    var tLookupRole = roleMap[tEmail] || 'NOT_IN_LOOKUPS';
    var leads = parseFloat(targetsData[i][2]) || 0;
    var included = tRowMonth === currentMonth && tLookupRole !== 'Team Leader' && tLookupRole !== 'Admin';
    if (included) total += leads;
    
    Logger.log((included ? '✅' : '❌') + ' ' + targetsData[i][0] + 
      ' | role=' + tLookupRole + 
      ' | month=' + tRowMonth + 
      ' | leads=' + leads);
  }
 Logger.log('=== TOTAL COUNTED: ' + total + ' ===');
}

// ── GET ORG CHART DATA ──
function getOrgChartData(callerEmail) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();
    var employees = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || !data[i][1]) continue;
      if (data[i][9] === 'Inactive') continue;
     employees.push({
        name:             String(data[i][0] || ''),
        email:            String(data[i][1] || ''),
        role:             String(data[i][3] || ''),
        reportsTo:        String(data[i][4] || ''),
        employeeID:       String(data[i][8] || ''),
        phone:            String(data[i][10] || ''),
        department:       String(data[i][11] || ''),
        profilePicFileId: String(data[i][12] || '')
      });
    }
    return { success: true, employees: employees };
  } catch(e) {
    Logger.log('getOrgChartData error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── SAVE PROFILE PIC (Admin only) ──
function saveProfilePic(fileBase64, mimeType, targetEmail, callerEmail) {
  try {
    var callerRole = getUserRole(callerEmail);
    if (callerRole !== 'Admin' && callerEmail !== targetEmail) {
      return { success: false, message: 'You can only update your own profile picture.' };
    }

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();

    // Get or create CRM_Profile_Pictures folder
    var rootFolder = DriveApp.getRootFolder();
    var folders = rootFolder.getFoldersByName('CRM_Profile_Pictures');
    var folder = folders.hasNext() ? folders.next() : rootFolder.createFolder('CRM_Profile_Pictures');

    // Find employee name for filename
    var empName = targetEmail;
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === targetEmail) { empName = data[i][0]; break; }
    }

    // Delete old file if exists
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === targetEmail && data[i][12]) {
        try { DriveApp.getFileById(data[i][12]).setTrashed(true); } catch(e) {}
        break;
      }
    }

    // Save new file
    var ext = mimeType === 'image/png' ? '.png' : mimeType === 'image/webp' ? '.webp' : '.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, empName + '_pic' + ext);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();

    // Save file ID to col M (index 12)
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === targetEmail) {
        sheet.getRange(i + 1, 13).setValue(fileId);
        return { success: true, fileId: fileId };
      }
    }
    return { success: false, message: 'Employee not found.' };
  } catch(e) {
    Logger.log('saveProfilePic error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── UPDATE EMPLOYEE PROFILE (Admin only — phone + department) ──
function updateEmployeeProfile(targetEmail, phone, department, callerEmail) {
  try {
    var callerRole = getUserRole(callerEmail);
    if (callerRole !== 'Admin') return { success: false, message: 'Only Admin can update profiles.' };

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Lookups');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === targetEmail) {
        sheet.getRange(i + 1, 11).setValue(phone || '');       // col K
        sheet.getRange(i + 1, 12).setValue(department || '');  // col L
        return { success: true, message: 'Profile updated successfully.' };
      }
    }
    return { success: false, message: 'Employee not found.' };
  } catch(e) {
    Logger.log('updateEmployeeProfile error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── TEAM MONITOR DATA ──
function getTeamMonitorData(callerEmail, callerRole) {
  try {
    Logger.log('getTeamMonitorData called: ' + callerEmail + ' | ' + callerRole);
    
    if (!callerEmail || !callerRole) {
      return { success: false, message: 'Missing email or role' };
    }
    if (callerRole !== 'Admin' && callerRole !== 'Team Leader') {
      return { success: false, message: 'Access denied for role: ' + callerRole };
    }

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var lookupsSheet = ss.getSheetByName('Lookups');
    var attSheet     = ss.getSheetByName('Attendance');

    if (!lookupsSheet) return { success: false, message: 'Lookups sheet not found' };
    if (!attSheet)     return { success: false, message: 'Attendance sheet not found' };

    var lookupsData = lookupsSheet.getDataRange().getValues();

    var callerName = '';
    for (var i = 1; i < lookupsData.length; i++) {
      if (lookupsData[i][1] === callerEmail) { callerName = lookupsData[i][0]; break; }
    }

    // ── BUILD EMPLOYEE LIST FIRST (fast) ──
    var employees = [];
    var emailSet = {};
    for (var i = 1; i < lookupsData.length; i++) {
      if (!lookupsData[i][0] || !lookupsData[i][1]) continue;
      if (lookupsData[i][9] === 'Inactive') continue;
      var empRole = lookupsData[i][3];
      if (empRole === 'Admin') continue;

      if (callerRole === 'Team Leader') {
        if (empRole !== 'BDA') continue;
        var reportsTo = String(lookupsData[i][4] || '').trim();
        if (reportsTo !== callerEmail && reportsTo !== callerName) continue;
      }

      var empEmail = lookupsData[i][1];
      emailSet[empEmail] = true;
      employees.push({
        name:        lookupsData[i][0],
        email:       empEmail,
        role:        empRole,
        reportsTo:   lookupsData[i][4] || '',
        employeeID:  lookupsData[i][8] || '',
        isOnline:    false,
        attendanceDays: 0,
        lastSeen:    ''
      });
    }

    // ── READ ATTENDANCE (only if employees exist) ──
    if (employees.length > 0) {
      var tz = 'Asia/Kolkata';
      var now = new Date();
      var todayStr  = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
      var thisMonth = Utilities.formatDate(now, tz, 'yyyy-MM');

      var attData = attSheet.getDataRange().getValues();

     var onlineEmails    = {};
      var lastSeenMap     = {};
      var monthDurationMap = {}; // email -> { 'yyyy-MM-dd': totalMins }

      for (var r = 1; r < attData.length; r++) {
        var rowEmail = String(attData[r][2] || '').trim();
        if (!rowEmail || !emailSet[rowEmail]) continue;

        var rawDate = attData[r][0];
        var normalizedDate = '';
        if (rawDate instanceof Date) {
          normalizedDate = Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd');
        } else {
          normalizedDate = String(rawDate || '').substring(0, 10);
        }

        var logout   = attData[r][5];
        var lastSeen = attData[r][8];

        // ── ONLINE CHECK: today + no logout ──
        if (normalizedDate === todayStr && !logout) {
          onlineEmails[rowEmail] = true;
        }

        // ── LAST SEEN MAP ──
        if (lastSeen) lastSeenMap[rowEmail] = String(lastSeen);

        // ── DURATION ACCUMULATION (only for current month) ──
        if (normalizedDate.indexOf(thisMonth) === 0) {
          var durationMins = parseFloat(attData[r][6]) || 0; // col G = Duration (mins)

          // If session has no logout yet (still active), estimate from Last Seen
          if (!logout && lastSeen && normalizedDate === todayStr) {
            var loginTime  = new Date(attData[r][4]);
            var lastSeenDt = new Date(lastSeen);
            if (!isNaN(loginTime.getTime()) && !isNaN(lastSeenDt.getTime())) {
              var estimatedMins = (lastSeenDt - loginTime) / 60000;
              if (estimatedMins > durationMins) durationMins = estimatedMins;
            }
          }

          if (!monthDurationMap[rowEmail]) monthDurationMap[rowEmail] = {};
          monthDurationMap[rowEmail][normalizedDate] =
            (monthDurationMap[rowEmail][normalizedDate] || 0) + durationMins;
        }
      }

      // ── ENRICH EMPLOYEES ──
      for (var e = 0; e < employees.length; e++) {
        var em = employees[e].email;
        employees[e].isOnline = !!onlineEmails[em];

        // Count days where total session time >= 480 mins (8 hours)
        var empDates = monthDurationMap[em] || {};
        var fullDays = 0;
        for (var d in empDates) {
          if (empDates[d] >= 480) fullDays++;
        }
        employees[e].attendanceDays = fullDays;

        var ls = lastSeenMap[em];
        if (ls) {
          try {
            employees[e].lastSeen = Utilities.formatDate(
              new Date(ls), 'Asia/Kolkata', 'dd MMM yyyy HH:mm'
            );
          } catch(fe) {
            employees[e].lastSeen = String(ls);
          }
        } else {
          employees[e].lastSeen = '';
        }
      }
    }

    employees.sort(function(a, b) {
      if (a.isOnline && !b.isOnline) return -1;
      if (!a.isOnline && b.isOnline) return 1;
      return a.name.localeCompare(b.name);
    });

    Logger.log('getTeamMonitorData success: ' + employees.length + ' employees');
    return { 
  success: true, 
  employees: employees, 
  month: String(Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM'))
};

  } catch(e) {
    Logger.log('getTeamMonitorData ERROR: ' + e.message + '\n' + e.stack);
    return { success: false, message: e.message };
  }
}
function testTeamMonitor() {
  var result = getTeamMonitorData('developernotifytechai@gmail.com', 'Admin');
  Logger.log(JSON.stringify(result));
}

function getEmployeeLeadStats(targetEmail, callerEmail, callerRole) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');

    // BDAs cannot access anyone's stats
    if (callerRole === 'BDA') {
      return { success: false, error: 'Access denied.' };
    }

    // --- Read Leads_Master ---
    var leadsSheet = ss.getSheetByName('Leads_Master');
    var leadsData  = leadsSheet.getDataRange().getValues();

    var colAssignedEmail = 13; // col N (0-based)
    var colStage         = 18; // col S
    var colFollowupDate  = 19; // col T

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var stats = { assigned: 0, contacted: 0, won: 0, followupDue: 0, unqualified: 0, lost: 0 };

    for (var r = 1; r < leadsData.length; r++) {
      var row = leadsData[r];
      if (!row[colAssignedEmail]) continue;
      if (row[colAssignedEmail].toString().trim().toLowerCase() !== targetEmail.trim().toLowerCase()) continue;

      stats.assigned++;
      var stage = (row[colStage] || '').toString().trim().toLowerCase();

      if (stage.indexOf('contact') !== -1)  stats.contacted++;
      if (stage.indexOf('won')     !== -1)  stats.won++;
      if (stage.indexOf('unqualif')!== -1)  stats.unqualified++;
      if (stage.indexOf('lost')    !== -1)  stats.lost++;

      if (row[colFollowupDate] && stage.indexOf('won') === -1 && stage.indexOf('lost') === -1) {
        var fDate = new Date(row[colFollowupDate]);
        fDate.setHours(0, 0, 0, 0);
        if (fDate <= today) stats.followupDue++;
      }
    }

    return { success: true, stats: stats };

  } catch(e) {
    return { success: false, error: e.message };
  }
}
// ── IMPORT LEADS FROM CSV ──
function importLeadsFromCSV(rows, assignToName, callerEmail) {
  var role = getUserRole(callerEmail);
  if (role !== 'Admin' && role !== 'Team Leader') {
    return { success: false, message: 'Access denied.' };
  }

  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet    = ss.getSheetByName('Leads_Master');
  var lookups  = ss.getSheetByName('Lookups');
  var data     = sheet.getDataRange().getValues();
  var headers  = data[0];
  var empData  = lookups.getDataRange().getValues();

  // ── COLUMN INDICES ──
  var colLeadID       = headers.indexOf('LeadID');
  var colDate         = headers.indexOf('Date');
  var colLastDate     = headers.indexOf('Last Date');
  var colName         = headers.indexOf('Name');
  var colMobile       = headers.indexOf('Mobile');
  var colAltPhone     = headers.indexOf('Alternate Phone');
  var colEmail        = headers.indexOf('Email');
  var colWhatsApp     = headers.indexOf('WhatsApp');
  var colCompany      = headers.indexOf('Company');
  var colSource       = headers.indexOf('Source');
  var colService      = headers.indexOf('Service Required');
  var colAssignDate   = headers.indexOf('Assign Date');
  var colAssignedTo   = headers.indexOf('Assigned To');
  var colAssignEmail  = headers.indexOf('Assigned Email');
  var colBDAName      = headers.indexOf('BDA Name');
  var colTL           = headers.indexOf('Team Leader');
  var colCallStatus   = headers.indexOf('Call Status');
  var colUnqualReason = headers.indexOf('Unqualified Reason');
  var colPipeline     = headers.indexOf('Pipeline & Stage');
  var colRemarks      = headers.indexOf('Remarks');
  var colFollowUp     = headers.indexOf('Follow-up Date');
  var colVerify       = headers.indexOf('Verification Status');
  var colLeadLock     = headers.indexOf('Lead Lock');
  var colLeadScore    = headers.indexOf('Lead Score');
  var colLeadLabel    = headers.indexOf('Lead Label');

  // ── BUILD EXISTING MOBILE SET ──
  var existingMobiles = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][colMobile]) {
      existingMobiles[data[i][colMobile].toString().trim()] = true;
    }
  }

  // ── RESOLVE ASSIGNEE if bulk-assigning ──
  var bulkBdaName  = '';
  var bulkBdaEmail = '';
  var bulkTLName   = '';

  if (assignToName) {
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][0] === assignToName) {
        bulkBdaName  = empData[i][0];
        bulkBdaEmail = empData[i][1];
        var managerRef = empData[i][4];
        for (var j = 1; j < empData.length; j++) {
          if (empData[j][1] === managerRef || empData[j][0] === managerRef) {
            bulkTLName = empData[j][0]; break;
          }
        }
        break;
      }
    }
  }

  // ── TL SCOPE CHECK ──
  // If caller is TL, the assignToName (if given) must be one of their BDAs
  if (role === 'Team Leader' && assignToName) {
    var callerName = '';
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][1] === callerEmail) { callerName = empData[i][0]; break; }
    }
    var isMyBDA = false;
    for (var i = 1; i < empData.length; i++) {
      if (empData[i][0] === assignToName) {
        var rt = String(empData[i][4] || '').trim();
        if (rt === callerEmail || rt === callerName) { isMyBDA = true; break; }
      }
    }
    if (!isMyBDA) return { success: false, message: 'You can only assign to your own BDAs.' };
  }

  var imported = 0, skipped = [], now = new Date();

  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var mobile = (row['Mobile'] || '').toString().trim().replace(/[\s\-\(\)\+\.]/g, '');
    if (!mobile) { skipped.push({ row: r + 1, name: row['Name'] || '—', reason: 'Missing mobile' }); continue; }
    if (existingMobiles[mobile]) { skipped.push({ row: r + 1, name: row['Name'] || '—', reason: 'Duplicate mobile: ' + mobile }); continue; }

    // ── RESOLVE PER-ROW ASSIGNMENT (if CSV has Assigned To column) ──
    var rowBdaName  = bulkBdaName;
    var rowBdaEmail = bulkBdaEmail;
    var rowTLName   = bulkTLName;

    var csvAssignTo = (row['Assigned To'] || '').toString().trim();
    if (csvAssignTo && !assignToName) {
      // CSV has assignment — validate scope for TL
      for (var i = 1; i < empData.length; i++) {
        if (empData[i][0] === csvAssignTo) {
          if (role === 'Team Leader') {
            var callerNameCheck = '';
            for (var k = 1; k < empData.length; k++) {
              if (empData[k][1] === callerEmail) { callerNameCheck = empData[k][0]; break; }
            }
            var rt2 = String(empData[i][4] || '').trim();
            if (rt2 !== callerEmail && rt2 !== callerNameCheck) {
              skipped.push({ row: r + 1, name: row['Name'] || '—', reason: csvAssignTo + ' is not in your team' });
              csvAssignTo = ''; break;
            }
          }
          rowBdaName  = empData[i][0];
          rowBdaEmail = empData[i][1];
          var manRef  = empData[i][4];
          for (var j = 1; j < empData.length; j++) {
            if (empData[j][1] === manRef || empData[j][0] === manRef) { rowTLName = empData[j][0]; break; }
          }
          break;
        }
      }
      if (!rowBdaEmail && csvAssignTo) {
        skipped.push({ row: r + 1, name: row['Name'] || '—', reason: 'BDA "' + csvAssignTo + '" not found' });
        continue;
      }
    }

    var leadID = 'LD-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + (imported + 1);

    // ── SCORE ──
    var scoreResult = calculateLeadScore('', row['Pipeline Stage'] || '', row['Call Status'] || 'Not Called', row['Follow-up Date'] || '', now, row['Service Required'] || '');

    // ── BUILD ROW matching Leads_Master column order ──
    var newRow = new Array(headers.length).fill('');
    newRow[colLeadID]       = leadID;
    newRow[colDate]         = now;
    newRow[colLastDate]     = now;
    newRow[colName]         = (row['Name'] || '').toString().trim();
    newRow[colMobile]       = mobile;
    newRow[colAltPhone]     = (row['Alt Phone'] || '').toString().trim();
    newRow[colEmail]        = (row['Email'] || '').toString().trim();
    newRow[colWhatsApp]     = (row['WhatsApp'] || '').toString().trim();
    newRow[colCompany]      = (row['Company'] || '').toString().trim();
    newRow[colSource]       = (row['Source'] || '').toString().trim();
    newRow[colService]      = (row['Service Required'] || '').toString().trim();
    newRow[colCallStatus]   = (row['Call Status'] || 'Not Called').toString().trim();
    newRow[colUnqualReason] = (row['Unqualified Reason'] || '').toString().trim();
    newRow[colPipeline]     = (row['Pipeline Stage'] || '').toString().trim();
    newRow[colRemarks]      = (row['Remarks'] || '').toString().trim();
    newRow[colFollowUp]     = row['Follow-up Date'] ? new Date(row['Follow-up Date']) : '';
    newRow[colVerify]       = (row['Verification'] || '').toString().trim();
    newRow[colLeadLock]     = 'N';
    newRow[colLeadScore]    = scoreResult.score;
    newRow[colLeadLabel]    = scoreResult.label;

    if (rowBdaName) {
      newRow[colAssignedTo]  = rowBdaName;
      newRow[colAssignEmail] = rowBdaEmail;
      newRow[colBDAName]     = rowBdaName;
      newRow[colTL]          = rowTLName;
      newRow[colAssignDate]  = now;
    }

    sheet.appendRow(newRow);
    existingMobiles[mobile] = true;

    if (rowBdaEmail) queueLeadNotification(rowBdaEmail, leadID);
    imported++;
  }

  return { success: true, imported: imported, skipped: skipped };
}
// ============================================================
// ARCHIVE SYSTEM — Paste these functions into WebApp.gs
// ============================================================

// ── AUTO ARCHIVE handler (Admin button — older than 6 months) ──
function runArchiveLeads(callerEmail) {
  var role = getUserRole(callerEmail);
  if (role !== 'Admin') {
    return { success: false, message: 'Only Admin can archive leads.' };
  }
  return archiveLeads(); // defined in Code.gs
}

// ── MONTH ARCHIVE handler (month picker button) ──
function runArchiveLeadsByMonth(month, callerEmail) {
  var role = getUserRole(callerEmail);
  if (role !== 'Admin') {
    return { success: false, message: 'Only Admin can archive leads.' };
  }
  return archiveLeadsByMonth(month); // defined in Code.gs
}

// ── PRE-CHECK COUNT handler (called before confirmation prompt) ──
function runCountArchivableByMonth(month, callerEmail) {
  var role = getUserRole(callerEmail);
  if (role !== 'Admin') {
    return { success: false, message: 'Access denied.' };
  }
  return countArchivableByMonth(month); // defined in Code.gs
}

// ── GET ARCHIVED LEADS (paginated, Admin only) ──
function getArchivedLeads(callerEmail, page, pageSize) {
  var role = getUserRole(callerEmail);
  if (role !== 'Admin') {
    return { success: false, message: 'Only Admin can view archived leads.' };
  }

  page     = page     || 1;
  pageSize = pageSize || 50;

  var ss    = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var sheet = ss.getSheetByName('Leads_Archive');
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, leads: [], total: 0, page: 1, pageSize: pageSize };
  }

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows    = [];

 for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var lead = {};
    headers.forEach(function(h, idx) {
      var val = data[i][idx];
      if (val instanceof Date) {
        lead[h] = val.toISOString();
      } else {
        lead[h] = val;
      }
    });
    rows.push(lead);
}

  // Most recently archived first
  rows.reverse();

  var total      = rows.length;
  var startIndex = (page - 1) * pageSize;
  var paged      = rows.slice(startIndex, startIndex + pageSize);

  return { success: true, leads: paged, total: total, page: page, pageSize: pageSize };
}

// ── RESTORE SINGLE LEAD from archive back to Leads_Master ──
function restoreArchivedLead(leadID, callerEmail) {
  var role = getUserRole(callerEmail);
  if (role !== 'Admin') {
    return { success: false, message: 'Only Admin can restore leads.' };
  }

  var ss      = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var archive = ss.getSheetByName('Leads_Archive');
  var master  = ss.getSheetByName('Leads_Master');

  if (!archive) return { success: false, message: 'Leads_Archive sheet not found.' };

  var archiveData = archive.getDataRange().getValues();
  var headers     = archiveData[0];
  var leadIDCol   = headers.indexOf('LeadID');

  for (var i = 1; i < archiveData.length; i++) {
    if (archiveData[i][leadIDCol] === leadID) {
      master.appendRow(archiveData[i]);
      archive.deleteRow(i + 1);
      try { applyLeadColors(); } catch(e) {}
      return { success: true, message: 'Lead ' + leadID + ' restored to Leads_Master.' };
    }
  }

  return { success: false, message: 'Lead not found in archive.' };
}
// ============================================================
// LEAD TIMELINE — Phase 1
// ============================================================

function appendActivityLog(leadId, leadName, actionType, details, oldValue, newValue, callerEmail) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var logSheet = ss.getSheetByName('Lead_Activity_Log');
    if (!logSheet) return;
    var userEmail = callerEmail || '';
    var timestamp = new Date().toISOString();
    var newRow = logSheet.getLastRow() + 1;
    logSheet.appendRow([
      timestamp,
      String(leadId || ''),
      String(leadName || ''),
      userEmail,
      String(actionType || ''),
      String(details || ''),
      String(oldValue || ''),
      String(newValue || '')
    ]);
    // Force column B to plain text so Sheets never auto-converts the LeadID
    logSheet.getRange(newRow, 2).setNumberFormat('@');
  } catch(e) {
    Logger.log('appendActivityLog error: ' + e.message);
  }
}

function logCall(leadId, outcome, notes, nextFollowUp, callerEmail) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Leads_Master');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var leadIDCol     = headers.indexOf('LeadID');
    var nameCol       = headers.indexOf('Name');
    var callStatusCol = headers.indexOf('Call Status');
    var remarksCol    = headers.indexOf('Remarks');
    var followUpCol   = headers.indexOf('Follow-up Date');
    var lastDateCol   = headers.indexOf('Last Date');

    var role = getUserRole(callerEmail);
    var allowedOutcomes = ['Connected', 'Called - No Answer', 'Called - Callback', 'Unqualified Lead', 'Interested'];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][leadIDCol]) !== String(leadId)) continue;

      var leadName = data[i][nameCol] || '';
      var oldCallStatus = data[i][callStatusCol] || '';
      var oldRemark = data[i][remarksCol] || '';
      var oldFollowUp = data[i][followUpCol] ? new Date(data[i][followUpCol]).toISOString().split('T')[0] : '';

      // Update Call Status
      if (outcome && allowedOutcomes.indexOf(outcome) !== -1) {
        sheet.getRange(i + 1, callStatusCol + 1).setValue(outcome);
      }

      // Update Remarks
      if (notes) {
        sheet.getRange(i + 1, remarksCol + 1).setValue(notes);
      }

      // Update Follow-up Date
      if (nextFollowUp) {
        sheet.getRange(i + 1, followUpCol + 1).setValue(new Date(nextFollowUp));
      }

      // Update Last Date
      sheet.getRange(i + 1, lastDateCol + 1).setValue(new Date());

      // Log the call
      var details = 'Call outcome: ' + outcome + (notes ? '. Notes: ' + notes : '');
      appendActivityLog(leadId, leadName, 'Call Logged', details, oldCallStatus, outcome, callerEmail);

      // Log follow-up change if provided
      if (nextFollowUp && nextFollowUp !== oldFollowUp) {
        appendActivityLog(leadId, leadName, 'Follow-up Date Changed',
          'Follow-up set after call to ' + nextFollowUp,
          oldFollowUp, nextFollowUp, callerEmail);
      }

      // Log remark change if notes provided
      if (notes && notes !== oldRemark) {
        appendActivityLog(leadId, leadName, 'Remark Updated',
          'Remark updated after call',
          oldRemark, notes, callerEmail);
      }

      return { success: true };
    }

    return { success: false, message: 'Lead not found' };
  } catch(e) {
    Logger.log('logCall error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function getLeadTimeline(leadId) {
  try {
    Logger.log('getLeadTimeline called with: [' + leadId + '] type: ' + typeof leadId);
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var logSheet = ss.getSheetByName('Lead_Activity_Log');
    if (!logSheet || logSheet.getLastRow() < 2) return [];

    var data = logSheet.getDataRange().getValues();
    var timeline = [];
    var searchId = String(leadId || '').trim();

    for (var i = 1; i < data.length; i++) {
      var storedId = String(data[i][1] || '').trim();
      // Log first few to help debug
      if (i <= 3) Logger.log('Row ' + i + ' storedId: [' + storedId + '] match: ' + (storedId === searchId));
      if (storedId !== searchId) continue;
      timeline.push({
        timestamp:  data[i][0] ? new Date(data[i][0]).toISOString() : '',
        leadId:     storedId,
        leadName:   String(data[i][2] || ''),
        userEmail:  String(data[i][3] || ''),
        actionType: String(data[i][4] || ''),
        details:    String(data[i][5] || ''),
        oldValue:   String(data[i][6] || ''),
        newValue:   String(data[i][7] || '')
      });
    }

   Logger.log('getLeadTimeline returning ' + timeline.length + ' entries for [' + searchId + ']');
    return JSON.stringify(timeline.reverse());
  } catch(e) {
    Logger.log('getLeadTimeline error: ' + e.message);
    return [];
  }
}

function getAllActivityLog(filters) {
  try {
    Logger.log('getAllActivityLog filters: ' + JSON.stringify(filters));
    
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var logSheet = ss.getSheetByName('Lead_Activity_Log');
    if (!logSheet || logSheet.getLastRow() < 2) return JSON.stringify([]);

    var data = logSheet.getDataRange().getValues();
    var userEmail = filters && filters.callerEmail ? filters.callerEmail : '';
    var userRole  = getUserRole(userEmail);
    
    Logger.log('userRole: ' + userRole + ' | total rows: ' + data.length);

    // ── BUILD SCOPE SETS ──
    var myLeadIDs = {};

    if (userRole === 'Team Leader') {
      var lookupsSheet = ss.getSheetByName('Lookups');
      var lookupsData  = lookupsSheet.getDataRange().getValues();
      
      var callerName = '';
      for (var i = 1; i < lookupsData.length; i++) {
        if (lookupsData[i][1] === userEmail) { callerName = lookupsData[i][0]; break; }
      }

      // BDA emails under this TL
      var myBDAEmails = {};
      myBDAEmails[userEmail] = true; // TL can see their own actions
      for (var i = 1; i < lookupsData.length; i++) {
        if (lookupsData[i][3] !== 'BDA') continue;
        var rt = String(lookupsData[i][4] || '').trim();
        if (rt.toLowerCase() === userEmail.toLowerCase() || rt.toLowerCase() === callerName.toLowerCase()) {
          myBDAEmails[lookupsData[i][1]] = true;
        }
      }

      // Lead IDs assigned to those BDAs
      var leadsSheet2 = ss.getSheetByName('Leads_Master');
      var leadsData2  = leadsSheet2.getDataRange().getValues();
      var lHeaders    = leadsData2[0];
      var lAssignedEmailCol = lHeaders.indexOf('Assigned Email');
      var lLeadIDCol        = lHeaders.indexOf('LeadID');
      for (var r = 1; r < leadsData2.length; r++) {
        if (myBDAEmails[leadsData2[r][lAssignedEmailCol]]) {
          myLeadIDs[String(leadsData2[r][lLeadIDCol])] = true;
        }
      }
    }

    var results = [];
    for (var i = 1; i < data.length; i++) {
      var row = {
        timestamp:  data[i][0] ? new Date(data[i][0]).toISOString() : '',
        leadId:     String(data[i][1] || ''),
        leadName:   data[i][2] || '',
        userEmail:  data[i][3] || '',
        actionType: data[i][4] || '',
        details:    data[i][5] || '',
        oldValue:   data[i][6] || '',
        newValue:   data[i][7] || ''
      };

      // Role scoping
      if (userRole === 'BDA' && row.userEmail !== userEmail) continue;
      if (userRole === 'Team Leader' && !myLeadIDs[row.leadId]) continue;

      // Optional filters
      if (filters) {
        if (filters.bdaEmail && row.userEmail !== filters.bdaEmail) continue;
        if (filters.actionType && filters.actionType !== 'all' && row.actionType !== filters.actionType) continue;
        if (filters.dateFrom && row.timestamp && row.timestamp < filters.dateFrom) continue;
        if (filters.dateTo && row.timestamp && row.timestamp > filters.dateTo + 'T23:59:59') continue;
      }

      results.push(row);
    }

    Logger.log('returning ' + results.length + ' results');
    return JSON.stringify(results.reverse());

  } catch(e) {
    Logger.log('getAllActivityLog error: ' + e.message);
    return JSON.stringify([]);
  }
}
function getActionTypeForField(field) {
  if (field === 'Remarks')          return 'Remark Updated';
  if (field === 'Follow-up Date')   return 'Follow-up Date Changed';
  if (field === 'Pipeline & Stage') return 'Pipeline Stage Changed';
  if (field === 'Call Status')      return 'Call Logged';
  if (field === 'Assigned To')      return 'Lead Assigned';
  return 'Field Updated';
}
function saveWinDebrief(leadID, debriefData, callerEmail) {
  try {
    var role = getUserRole(callerEmail);
    if (!role) return { success: false, message: 'Unauthorized.' };

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var leadsSheet    = ss.getSheetByName('Leads_Master');
    var learningsSheet = ss.getSheetByName('Won_Learnings');

    if (!learningsSheet) return { success: false, message: 'Won_Learnings sheet not found.' };

    var leadsData    = leadsSheet.getDataRange().getValues();
    var leadsHeaders = leadsData[0];

    var leadIDCol       = leadsHeaders.indexOf('LeadID');
    var nameCol         = leadsHeaders.indexOf('Name');
    var assignedEmailCol = leadsHeaders.indexOf('Assigned Email');
    var assignedToCol   = leadsHeaders.indexOf('Assigned To');
    var teamLeaderCol   = leadsHeaders.indexOf('Team Leader');
    var serviceCol      = leadsHeaders.indexOf('Service Required');
    var dealAmountCol   = leadsHeaders.indexOf('Deal Amount');

    // Find the lead
    var leadRow = null;
    for (var i = 1; i < leadsData.length; i++) {
      if (String(leadsData[i][leadIDCol]) === String(leadID)) {
        leadRow = leadsData[i];
        break;
      }
    }
    if (!leadRow) return { success: false, message: 'Lead not found.' };

    // Validate caller is assigned BDA or Admin/TL
    var assignedEmail = leadRow[assignedEmailCol];
    if (role === 'BDA' && assignedEmail !== callerEmail) {
      return { success: false, message: 'You can only submit debriefs for your own leads.' };
    }

    // Resolve BDA name + email
    var bdaEmail = assignedEmail;
    var bdaName  = leadRow[assignedToCol] || '';
    var tlName   = leadRow[teamLeaderCol] || '';
    if (!tlName) {
  var lookupData2 = ss.getSheetByName('Lookups').getDataRange().getValues();
  for (var lu = 1; lu < lookupData2.length; lu++) {
    if (lookupData2[lu][1] === assignedEmail) {
      var rt = String(lookupData2[lu][4] || '').trim();
      // reportsTo may be a name or email — resolve to name
      for (var lv = 1; lv < lookupData2.length; lv++) {
        if (lookupData2[lv][1] === rt || lookupData2[lv][0] === rt) {
          tlName = lookupData2[lv][0];
          break;
        }
      }
      break;
    }
  }
}

    learningsSheet.appendRow([
      new Date(),                              // Timestamp
      leadID,                                  // LeadID
      leadRow[nameCol] || '',                  // LeadName
      bdaEmail,                                // BDAEmail
      bdaName,                                 // BDAName
      tlName,                                  // TeamLeader
      leadRow[serviceCol] || '',               // ServiceWon
      leadRow[dealAmountCol] || 0,             // DealValue
      debriefData.keyFactor || '',             // Win_KeyFactor
      debriefData.objectionsOvercome || '',    // Win_ObjectionsOvercome
      debriefData.pitchUsed || '',             // Win_ServicePitchUsed
      debriefData.salesNotes || '',            // Win_SalesNotes
      debriefData.skipped ? 'TRUE' : 'FALSE'  // Skipped
    ]);

    // Force LeadID column to plain text
    var newRow = learningsSheet.getLastRow();
    learningsSheet.getRange(newRow, 2).setNumberFormat('@');

    // Activity log
    appendActivityLog(
      leadID,
      leadRow[nameCol] || '',
      'Win Debrief Saved',
      debriefData.skipped ? 'BDA skipped win debrief' : 'Win debrief submitted. Key factor: ' + (debriefData.keyFactor || '—'),
      '',
      debriefData.skipped ? 'Skipped' : 'Completed',
      callerEmail
    );

    return { success: true, message: 'Win debrief saved.' };
  } catch(e) {
    Logger.log('saveWinDebrief error: ' + e.message);
    return { success: false, message: e.message };
  }
}
function getWinLibrary(callerEmail, callerRole, filters) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Won_Learnings');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, rows: [] };

    var data    = sheet.getDataRange().getValues();
    var headers = data[0];

    var lookupsSheet = ss.getSheetByName('Lookups');
    var lookupsData  = lookupsSheet.getDataRange().getValues();

    // Resolve caller's TL name (for TL scope)
    var callerName = '';
    for (var i = 1; i < lookupsData.length; i++) {
      if (lookupsData[i][1] === callerEmail) { callerName = lookupsData[i][0]; break; }
    }

    // Get BDA emails under this TL
    var myBDAEmails = {};
    if (callerRole === 'Team Leader') {
      myBDAEmails[callerEmail] = true;
      for (var i = 1; i < lookupsData.length; i++) {
        if (lookupsData[i][3] !== 'BDA') continue;
        var rt = String(lookupsData[i][4] || '').trim();
        if (rt === callerEmail || rt === callerName) {
          myBDAEmails[lookupsData[i][1]] = true;
        }
      }
    }

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;

      var row = {};
      headers.forEach(function(h, idx) {
        var val = data[i][idx];
        row[h] = val instanceof Date ? val.toISOString() : val;
      });

      // Scope filter
      if (callerRole === 'BDA' && row['BDAEmail'] !== callerEmail) continue;
      if (callerRole === 'Team Leader' && !myBDAEmails[row['BDAEmail']]) continue;

      // Optional filters
      if (filters) {
        if (filters.service && row['ServiceWon'] !== filters.service) continue;
        if (filters.bdaEmail && row['BDAEmail'] !== filters.bdaEmail) continue;
        if (filters.keyFactor && row['Win_KeyFactor'] !== filters.keyFactor) continue;
        if (filters.dateFrom && row['Timestamp'] && row['Timestamp'] < filters.dateFrom) continue;
        if (filters.dateTo && row['Timestamp'] && row['Timestamp'] > filters.dateTo + 'T23:59:59') continue;
        if (filters.skipSkipped && row['Skipped'] === 'TRUE') continue;
      }

      rows.push(row);
    }

    // Most recent first
    rows.sort(function(a, b) { return (b['Timestamp'] || '').localeCompare(a['Timestamp'] || ''); });

    return { success: true, rows: rows };
  } catch(e) {
    Logger.log('getWinLibrary error: ' + e.message);
    return { success: false, message: e.message };
  }
}
function testTeamTarget() {
  var result = getTeamTargetData('zoyakhantesting@gmail.com', 'Team Leader');
  Logger.log(JSON.stringify(result.rows));
}

// ── RE-ENGAGE: SAVE NOTE (BDA submits modal) ──
function saveReEngageNote(leadID, reason, note, type, callerEmail) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var leadsSheet = ss.getSheetByName('Leads_Master');
    var poolSheet  = ss.getSheetByName('ReEngage_Pool');

    if (!poolSheet) return { success: false, message: 'ReEngage_Pool sheet not found.' };

    var data    = leadsSheet.getDataRange().getValues();
    var headers = data[0];

    var leadIDCol       = headers.indexOf('LeadID');
    var nameCol         = headers.indexOf('Name');
    var mobileCol       = headers.indexOf('Mobile');
    var serviceCol      = headers.indexOf('Service Required');
    var bdaNameCol      = headers.indexOf('BDA Name');
    var assignedEmailCol = headers.indexOf('Assigned Email');
    var reEngStatusCol  = headers.indexOf('ReEngage_Status');
    var reEngTypeCol    = headers.indexOf('ReEngage_Type');

    if (reEngStatusCol === -1 || reEngTypeCol === -1) {
      return { success: false, message: 'ReEngage_Status / ReEngage_Type columns not found in Leads_Master.' };
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][leadIDCol]) !== String(leadID)) continue;

      var leadName = String(data[i][nameCol] || '');
      var now      = new Date();

      // ── Write 2 marker cols on Leads_Master ──
      leadsSheet.getRange(i + 1, reEngStatusCol + 1).setValue('Pending Review');
      leadsSheet.getRange(i + 1, reEngTypeCol  + 1).setValue(type);

      // ── Write full record to ReEngage_Pool ──
      // Check if a row for this leadID already exists (upsert)
      var poolData    = poolSheet.getDataRange().getValues();
      var poolHeaders = poolData[0];
      var pLeadIDCol  = poolHeaders.indexOf('LeadID');
      var existingRow = -1;
      for (var p = 1; p < poolData.length; p++) {
        if (String(poolData[p][pLeadIDCol]) === String(leadID)) { existingRow = p + 1; break; }
      }

      var poolRow = [
        leadID,
        leadName,
        String(data[i][assignedEmailCol] || callerEmail),
        type,
        reason,
        note,
        '',   // ReEngage_After_Date — blank until TL sets it
        '',   // ReEngage_Assigned_To
        '',   // ReEngage_Assigned_Email
        now,  // Submitted_At
        now   // Last_Updated
      ];

      if (existingRow > 1) {
        poolSheet.getRange(existingRow, 1, 1, poolRow.length).setValues([poolRow]);
      } else {
        poolSheet.appendRow(poolRow);
        // Force LeadID col to plain text
        poolSheet.getRange(poolSheet.getLastRow(), 1).setNumberFormat('@');
      }

      // ── Activity log ──
      var actionType = type === 'Unqualified - Revival' ? 'Revival Requested' : 'Re-engage Requested';
      appendActivityLog(leadID, leadName, actionType,
        actionType + '. Reason: ' + reason + '. Note: ' + note,
        '', type, callerEmail);

      return { success: true };
    }
    return { success: false, message: 'Lead not found' };
  } catch(e) {
    Logger.log('saveReEngageNote error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── RE-ENGAGE: GET POOL (for the tab) ──
function getReEngagePool(callerEmail, callerRole) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var poolSheet    = ss.getSheetByName('ReEngage_Pool');
    var lookupsSheet = ss.getSheetByName('Lookups');

    if (!poolSheet) return { success: false, message: 'ReEngage_Pool sheet not found.' };

    var poolData     = poolSheet.getDataRange().getValues();
    var lookupsData  = lookupsSheet.getDataRange().getValues();

    if (poolData.length <= 1) return { success: true, pool: [] };

    var headers = poolData[0];
    var pLeadIDCol        = headers.indexOf('LeadID');
    var pLeadNameCol      = headers.indexOf('LeadName');
    var pBDAEmailCol      = headers.indexOf('BDA_Email');
    var pTypeCol          = headers.indexOf('ReEngage_Type');
    var pReasonCol        = headers.indexOf('ReEngage_Reason');
    var pNoteCol          = headers.indexOf('ReEngage_BDA_Note');
    var pAfterDateCol     = headers.indexOf('ReEngage_After_Date');
    var pAssignedToCol    = headers.indexOf('ReEngage_Assigned_To');
    var pAssignedEmailCol = headers.indexOf('ReEngage_Assigned_Email');
    var pSubmittedCol     = headers.indexOf('Submitted_At');
    var pTLNoteCol = headers.indexOf('TL_Note');

    // Also read ReEngage_Status from Leads_Master for filtering
    var leadsSheet   = ss.getSheetByName('Leads_Master');
    var leadsData    = leadsSheet.getDataRange().getValues();
    var leadsHeaders = leadsData[0];
    var lLeadIDCol   = leadsHeaders.indexOf('LeadID');
    var lStatusCol   = leadsHeaders.indexOf('ReEngage_Status');
    var lMobileCol   = leadsHeaders.indexOf('Mobile');
    var lServiceCol  = leadsHeaders.indexOf('Service Required');
    var lBDANameCol  = leadsHeaders.indexOf('BDA Name');
    var lTLCol       = leadsHeaders.indexOf('Team Leader');

    // Build quick lookup from Leads_Master
    var leadsMap = {};
    for (var r = 1; r < leadsData.length; r++) {
      var lid = String(leadsData[r][lLeadIDCol] || '');
      if (!lid) continue;
      leadsMap[lid] = {
        reEngageStatus: String(leadsData[r][lStatusCol] || ''),
        mobile:         String(leadsData[r][lMobileCol] || ''),
        service:        String(leadsData[r][lServiceCol] || ''),
        bdaName:        String(leadsData[r][lBDANameCol] || ''),
        teamLeader:     String(leadsData[r][lTLCol] || '')
      };
    }

    // Resolve caller context for scoping
    var callerName = '';
    var myBDAEmails = {};
    for (var l = 1; l < lookupsData.length; l++) {
      if (lookupsData[l][1] === callerEmail) { callerName = lookupsData[l][0]; break; }
    }
    if (callerRole === 'Team Leader') {
      myBDAEmails[callerEmail] = true;
      for (var l = 1; l < lookupsData.length; l++) {
        if (lookupsData[l][3] !== 'BDA') continue;
        var rt = String(lookupsData[l][4] || '').trim();
        if (rt === callerEmail || rt === callerName) myBDAEmails[lookupsData[l][1]] = true;
      }
    }

    var closedStatuses = ['Won', 'Re-unqualified', 'Re-contacted'];
    var pool = [];

    for (var i = 1; i < poolData.length; i++) {
      var leadID   = String(poolData[i][pLeadIDCol] || '');
      if (!leadID) continue;

      var lmData   = leadsMap[leadID] || {};
      var status   = lmData.reEngageStatus || '';

      // Skip closed
      if (closedStatuses.indexOf(status) !== -1) continue;

      var bdaEmail = String(poolData[i][pBDAEmailCol] || '');

      // Scope filter
      if (callerRole === 'BDA' && bdaEmail !== callerEmail) continue;
      if (callerRole === 'Team Leader' && !myBDAEmails[bdaEmail]) continue;

      var afterDate = poolData[i][pAfterDateCol];
      var afterDateStr = '';
      if (afterDate) {
        try { afterDateStr = Utilities.formatDate(new Date(afterDate), Session.getScriptTimeZone(), 'yyyy-MM-dd'); } catch(e) {}
      }

      var submittedAt = poolData[i][pSubmittedCol];
      var submittedStr = '';
      if (submittedAt) {
        try { submittedStr = Utilities.formatDate(new Date(submittedAt), Session.getScriptTimeZone(), 'dd MMM yyyy'); } catch(e) {}
      }

      var today = new Date(); today.setHours(0,0,0,0);
      var isReady = afterDateStr && new Date(afterDateStr) <= today && status === 'Date Set';

      pool.push({
        leadID:           leadID,
        leadName:         String(poolData[i][pLeadNameCol] || ''),
        bdaEmail:         bdaEmail,
        bdaName:          lmData.bdaName || '',
        teamLeader:       lmData.teamLeader || '',
        mobile:           lmData.mobile || '',
        service:          lmData.service || '',
        tlNote: String(poolData[i][pTLNoteCol] || ''),
        reEngageType:     String(poolData[i][pTypeCol] || ''),
        reEngageReason:   String(poolData[i][pReasonCol] || ''),
        reEngageBDANote:  String(poolData[i][pNoteCol] || ''),
        reEngageAfterDate: afterDateStr,
        reEngageAssignedTo: String(poolData[i][pAssignedToCol] || ''),
        reEngageStatus:   status,
        submittedAt:      submittedStr,
        isReady:          isReady
      });
    }

    return { success: true, pool: pool };
  } catch(e) {
    Logger.log('getReEngagePool error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── RE-ENGAGE: SET DATE (TL/Admin only) ──
function setReEngageDate(leadID, afterDate, assignToBDAName, callerEmail, tlNote) {
  try {
    var role = getUserRole(callerEmail);
    if (role !== 'Admin' && role !== 'Team Leader') {
      return { success: false, message: 'Only TL or Admin can set re-engage dates.' };
    }

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var leadsSheet   = ss.getSheetByName('Leads_Master');
    var poolSheet    = ss.getSheetByName('ReEngage_Pool');
    var lookupsSheet = ss.getSheetByName('Lookups');

    var leadsData    = leadsSheet.getDataRange().getValues();
    var leadsHeaders = leadsData[0];
    var lLeadIDCol   = leadsHeaders.indexOf('LeadID');
    var lNameCol     = leadsHeaders.indexOf('Name');
    var lStatusCol   = leadsHeaders.indexOf('ReEngage_Status');

    // Update Leads_Master status
    var leadName = '';
    for (var i = 1; i < leadsData.length; i++) {
      if (String(leadsData[i][lLeadIDCol]) !== String(leadID)) continue;
      leadName = String(leadsData[i][lNameCol] || '');
      leadsSheet.getRange(i + 1, lStatusCol + 1).setValue('Date Set');
      break;
    }

    // Update ReEngage_Pool row
    var poolData    = poolSheet.getDataRange().getValues();
    var poolHeaders = poolData[0];
    var pLeadIDCol        = poolHeaders.indexOf('LeadID');
    var pAfterDateCol     = poolHeaders.indexOf('ReEngage_After_Date');
    var pAssignedToCol    = poolHeaders.indexOf('ReEngage_Assigned_To');
    var pAssignedEmailCol = poolHeaders.indexOf('ReEngage_Assigned_Email');
    var pLastUpdCol       = poolHeaders.indexOf('Last_Updated');
    var pTLNoteCol = poolHeaders.indexOf('TL_Note');

    var assignedEmail = '';
    if (assignToBDAName) {
      var lookupsData = lookupsSheet.getDataRange().getValues();
      for (var l = 1; l < lookupsData.length; l++) {
        if (lookupsData[l][0] === assignToBDAName) { assignedEmail = lookupsData[l][1]; break; }
      }
    }

   var pTLNoteCol = poolHeaders.indexOf('TL_Note');

for (var p = 1; p < poolData.length; p++) {
  if (String(poolData[p][pLeadIDCol]) !== String(leadID)) continue;
  poolSheet.getRange(p + 1, pAfterDateCol + 1).setValue(new Date(afterDate));
  if (assignToBDAName) {
    poolSheet.getRange(p + 1, pAssignedToCol    + 1).setValue(assignToBDAName);
    poolSheet.getRange(p + 1, pAssignedEmailCol + 1).setValue(assignedEmail);
  }
  if (pTLNoteCol !== -1 && tlNote) {
    poolSheet.getRange(p + 1, pTLNoteCol + 1).setValue(tlNote);
  }
  poolSheet.getRange(p + 1, pLastUpdCol + 1).setValue(new Date());
  break;
}

    appendActivityLog(leadID, leadName, 'Re-engage Approved',
      'Date set to ' + afterDate + (assignToBDAName ? '. Assigned to: ' + assignToBDAName : ''),
      '', afterDate, callerEmail);

    if (assignToBDAName) {
      appendActivityLog(leadID, leadName, 'Re-engage Reassigned',
        'Reassigned to ' + assignToBDAName, '', assignToBDAName, callerEmail);
    }

    return { success: true };
  } catch(e) {
    Logger.log('setReEngageDate error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── RE-ENGAGE: MARK RE-CONTACTED ──
function markReContacted(leadID, callerEmail) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var leadsSheet = ss.getSheetByName('Leads_Master');
    var poolSheet  = ss.getSheetByName('ReEngage_Pool');

    var leadsData    = leadsSheet.getDataRange().getValues();
    var leadsHeaders = leadsData[0];
    var lLeadIDCol   = leadsHeaders.indexOf('LeadID');
    var lNameCol     = leadsHeaders.indexOf('Name');
    var lCSCol       = leadsHeaders.indexOf('Call Status');
    var lStatusCol   = leadsHeaders.indexOf('ReEngage_Status');
    var lTypeCol     = leadsHeaders.indexOf('ReEngage_Type');

    var leadName = '';
    for (var i = 1; i < leadsData.length; i++) {
      if (String(leadsData[i][lLeadIDCol]) !== String(leadID)) continue;
      leadName = String(leadsData[i][lNameCol] || '');
      // Reset Call Status, clear both marker cols
      leadsSheet.getRange(i + 1, lCSCol    + 1).setValue('Not Called');
      leadsSheet.getRange(i + 1, lStatusCol + 1).setValue('Re-contacted');
      leadsSheet.getRange(i + 1, lTypeCol   + 1).setValue('');
      break;
    }

    // Mark pool row as Re-contacted (keep for history)
    var poolData    = poolSheet.getDataRange().getValues();
    var poolHeaders = poolData[0];
    var pLeadIDCol  = poolHeaders.indexOf('LeadID');
    var pLastUpdCol = poolHeaders.indexOf('Last_Updated');

    for (var p = 1; p < poolData.length; p++) {
      if (String(poolData[p][pLeadIDCol]) !== String(leadID)) continue;
      poolSheet.getRange(p + 1, pLastUpdCol + 1).setValue(new Date());
      break;
    }

    appendActivityLog(leadID, leadName, 'Re-contact Initiated',
      'Lead returned to active queue. Call Status reset to Not Called.',
      '', 'Not Called', callerEmail);

    return { success: true };
  } catch(e) {
    Logger.log('markReContacted error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── RE-ENGAGE: GET BADGE COUNT (for sidebar) ──
function getReEngageBadgeCount(callerEmail, callerRole) {
  try {
    var result = getReEngagePool(callerEmail, callerRole);
    if (!result.success) return { count: 0 };
    var pendingCount = result.pool.filter(function(item) {
      return item.reEngageStatus === 'Pending Review';
    }).length;
    var readyCount = result.pool.filter(function(item) {
      return item.isReady;
    }).length;
    return { count: pendingCount + readyCount, pending: pendingCount, ready: readyCount };
  } catch(e) {
    return { count: 0 };
  }
}

function saveLeadBatch(leadID, changes, callerEmail) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Leads_Master');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
var leadIDCol   = headers.indexOf('LeadID');
var lastDateCol = headers.indexOf('Last Date');

// // Force LeadID column to text to prevent number conversion
// sheet.getRange(2, leadIDCol + 1, sheet.getLastRow() - 1, 1).setNumberFormat('@');

  var rowIndex = -1;
var searchID = String(leadID || '').trim();
for (var i = 1; i < data.length; i++) {
  var cellID = String(data[i][leadIDCol] || '').trim();
  if (cellID === searchID || cellID.replace(/\s/g,'') === searchID.replace(/\s/g,'')) { 
    rowIndex = i; 
    break; 
  }
}
if (rowIndex === -1) {
  Logger.log('Lead not found. Searching for: [' + searchID + ']');
  Logger.log('First 5 IDs in sheet: ' + data.slice(1,6).map(function(r){ return '['+r[leadIDCol]+']'; }).join(', '));
  return { success: false, message: 'Lead not found' };
}

    var newScore = null, newLabel = null;

    for (var c = 0; c < changes.length; c++) {
      var field  = changes[c].field;
      var value  = changes[c].value;
      var colIdx = headers.indexOf(field);
      if (colIdx === -1) continue;
      sheet.getRange(rowIndex + 1, colIdx + 1).setValue(value);
      data[rowIndex][colIdx] = value;

      // Activity log each change
      try {
        appendActivityLog(
          String(leadID),
          String(data[rowIndex][headers.indexOf('Name')] || ''),
          getActionTypeForField(field),
          field + ' updated',
          '',
          String(value || ''),
          callerEmail
        );
      } catch(logErr) {
        Logger.log('Activity log error: ' + logErr.message);
      }
    }

    // Update Last Date
    sheet.getRange(rowIndex + 1, lastDateCol + 1).setValue(new Date());

    // Recalculate lead score
    var scoringFields = ['Pipeline & Stage','Call Status','Remarks','Follow-up Date','Service Required'];
    var needsRescore  = changes.some(function(c) { return scoringFields.indexOf(c.field) !== -1; });
    if (needsRescore) {
      try {
        var remarksCol   = headers.indexOf('Remarks');
        var pipelineCol  = headers.indexOf('Pipeline & Stage');
        var csCol        = headers.indexOf('Call Status');
        var followUpCol  = headers.indexOf('Follow-up Date');
        var serviceCol   = headers.indexOf('Service Required');
        var leadScoreCol = headers.indexOf('Lead Score');
        var leadLabelCol = headers.indexOf('Lead Label');

        var scoreResult = calculateLeadScore(
          data[rowIndex][remarksCol]  || '',
          data[rowIndex][pipelineCol] || '',
          data[rowIndex][csCol]       || '',
          data[rowIndex][followUpCol] || '',
          new Date(),
          data[rowIndex][serviceCol]  || ''
        );

        newScore = scoreResult.score;
        newLabel = scoreResult.label;
        if (leadScoreCol !== -1) sheet.getRange(rowIndex + 1, leadScoreCol + 1).setValue(newScore);
        if (leadLabelCol !== -1) sheet.getRange(rowIndex + 1, leadLabelCol + 1).setValue(newLabel);
      } catch(scoreErr) {
        Logger.log('Score error: ' + scoreErr.message);
      }
    }

    return { success: true, newScore: newScore, newLabel: newLabel };
  } catch(e) {
    Logger.log('saveLeadBatch error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function getNotifications(userEmail, userRole) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var notifications = [];
    var now = new Date();
    var tz = 'Asia/Kolkata';

    var leadsSheet = ss.getSheetByName('Leads_Master');
    var leadsData  = leadsSheet.getDataRange().getValues();
    var lHeaders   = leadsData[0];

    var lAssignedEmailCol = lHeaders.indexOf('Assigned Email');
    var lCallStatusCol    = lHeaders.indexOf('Call Status');
    var lLastDateCol      = lHeaders.indexOf('Last Date');
    var lDateCol          = lHeaders.indexOf('Date');
    var lLeadIDCol        = lHeaders.indexOf('LeadID');
    var lNameCol          = lHeaders.indexOf('Name');
    var lFollowUpCol      = lHeaders.indexOf('Follow-up Date');
    var lPipelineCol      = lHeaders.indexOf('Pipeline & Stage');
    var lVerifyCol        = lHeaders.indexOf('Verification Status');
    var lAssignedToCol    = lHeaders.indexOf('Assigned To');
    var lReEngStatusCol   = lHeaders.indexOf('ReEngage_Status');

    // Build TL scope
    var lookupSheet = ss.getSheetByName('Lookups');
    var lookupData  = lookupSheet.getDataRange().getValues();
    var callerName  = '';
    var myBDAEmails = {};
    for (var i = 1; i < lookupData.length; i++) {
      if (lookupData[i][1] === userEmail) { callerName = lookupData[i][0]; break; }
    }
    if (userRole === 'Team Leader') {
      for (var i = 1; i < lookupData.length; i++) {
        if (lookupData[i][3] !== 'BDA') continue;
        var rt = String(lookupData[i][4] || '').trim();
        if (rt === userEmail || rt === callerName) myBDAEmails[lookupData[i][1]] = true;
      }
    }

    // Today boundaries
    var todayStart = new Date();
    todayStart.setHours(0, 0, 0, 0);
    var todayEnd = new Date(todayStart);
    todayEnd.setDate(todayEnd.getDate() + 1);
    var cutoff24h = new Date(now.getTime() - 24 * 60 * 60 * 1000);

    var newLeadCount      = 0;
    var followupTodayCount = 0;
    var overdueCount      = 0;

    for (var i = 1; i < leadsData.length; i++) {
      if (!leadsData[i][0]) continue;

      var assignedEmail = String(leadsData[i][lAssignedEmailCol] || '');
      var isMyLead = false;
      if (userRole === 'Admin') {
        isMyLead = true;
      } else if (userRole === 'BDA') {
        isMyLead = (assignedEmail === userEmail);
      } else if (userRole === 'Team Leader') {
        isMyLead = (assignedEmail === userEmail || myBDAEmails[assignedEmail]);
      }
      if (!isMyLead) continue;

      var pipeline = String(leadsData[i][lPipelineCol] || '');
      var cs       = String(leadsData[i][lCallStatusCol] || '');

      // ── NEW LEADS for BDA (Not Called, assigned in last 24h) ──
      if (userRole === 'BDA') {
        if (cs === 'Not Called' || cs === '') {
          var checkDateRaw = leadsData[i][lDateCol] || leadsData[i][lLastDateCol];
          if (checkDateRaw) {
            var checkDate = new Date(checkDateRaw);
            if (!isNaN(checkDate.getTime()) && checkDate >= cutoff24h) {
              newLeadCount++;
            }
          }
        }
      }

      // ── FOLLOW-UPS ──
      var fuRaw = leadsData[i][lFollowUpCol];
      if (fuRaw) {
        var fDate = new Date(fuRaw);
        if (!isNaN(fDate.getTime())) {
          fDate.setHours(0, 0, 0, 0);
          if (fDate >= todayStart && fDate < todayEnd) {
            followupTodayCount++;
          } else if (fDate < todayStart) {
            overdueCount++;
          }
        }
      }

      // ── PENDING VERIFICATION (Admin only) ──
      if (userRole === 'Admin' && pipeline === 'Won') {
        var vs = String(leadsData[i][lVerifyCol] || '');
        if (vs === 'Pending' || vs === '') {
          var leadName = String(leadsData[i][lNameCol] || '');
          var lastDateRaw = leadsData[i][lLastDateCol];
          var lastDateStr = lastDateRaw ? new Date(lastDateRaw).toISOString() : now.toISOString();
          notifications.push({
            icon: 'deal_verify',
            title: 'Deal pending verification — ' + leadName,
            sub: 'Won lead needs Admin review. Check Leads tab.',
            time: lastDateStr,
            unread: true,
            type: 'verify'
          });
        }
      }
    }

    // ── AGGREGATE COUNT NOTIFICATIONS ──
    if (newLeadCount > 0 && userRole === 'BDA') {
      notifications.push({
        icon: 'new_lead',
        title: newLeadCount + ' new lead' + (newLeadCount > 1 ? 's' : '') + ' assigned to you',
        sub: 'Check your leads — assigned in the last 24 hours',
        time: now.toISOString(),
        unread: true,
        type: 'lead'
      });
    }
    if (followupTodayCount > 0) {
      notifications.push({
        icon: 'followup',
        title: followupTodayCount + ' follow-up' + (followupTodayCount > 1 ? 's' : '') + ' due today',
        sub: 'Leads waiting for your call today',
        time: now.toISOString(),
        unread: true,
        type: 'followup'
      });
    }
    if (overdueCount > 0) {
      notifications.push({
        icon: 'overdue',
        title: overdueCount + ' overdue follow-up' + (overdueCount > 1 ? 's' : ''),
        sub: 'Past their follow-up date — action needed',
        time: now.toISOString(),
        unread: true,
        type: 'overdue'
      });
    }

    // ── RE-ENGAGE: ready leads notification (daily, for BDA) ──
if (userRole === 'BDA') {
  var rePoolSheet = ss.getSheetByName('ReEngage_Pool');
  if (rePoolSheet && rePoolSheet.getLastRow() > 1) {
    var rePoolData = rePoolSheet.getDataRange().getValues();
    var rePoolH = rePoolData[0];
    var rpBDAEmailCol      = rePoolH.indexOf('BDA_Email');
    var rpAssignedEmailCol = rePoolH.indexOf('ReEngage_Assigned_Email');
    var rpAfterDateCol     = rePoolH.indexOf('ReEngage_After_Date');
    var rpLeadNameCol      = rePoolH.indexOf('LeadName');
    var rpLeadIDCol        = rePoolH.indexOf('LeadID');

    var leadsSheetN    = ss.getSheetByName('Leads_Master');
    var leadsDataN     = leadsSheetN.getDataRange().getValues();
    var leadsHeadersN  = leadsDataN[0];
    var lStatusColN    = leadsHeadersN.indexOf('ReEngage_Status');
    var lLeadIDColN    = leadsHeadersN.indexOf('LeadID');

    // Build status map
    var statusMapN = {};
    for (var rn = 1; rn < leadsDataN.length; rn++) {
      statusMapN[String(leadsDataN[rn][lLeadIDColN])] = String(leadsDataN[rn][lStatusColN] || '');
    }

    var todayN = new Date(); todayN.setHours(0,0,0,0);

    for (var rp = 1; rp < rePoolData.length; rp++) {
      var rpAssigned = String(rePoolData[rp][rpAssignedEmailCol] || '');
      var rpOrigBDA  = String(rePoolData[rp][rpBDAEmailCol] || '');
      var isMyLead   = (rpAssigned === userEmail || (!rpAssigned && rpOrigBDA === userEmail));
      if (!isMyLead) continue;

      var rpAfterRaw = rePoolData[rp][rpAfterDateCol];
      if (!rpAfterRaw) continue;
      var rpAfterDate = new Date(rpAfterRaw); rpAfterDate.setHours(0,0,0,0);
      if (rpAfterDate > todayN) continue; // not ready yet

      var rpLeadID = String(rePoolData[rp][rpLeadIDCol] || '');
      var rpStatus = statusMapN[rpLeadID] || '';
      if (rpStatus === 'Re-contacted' || rpStatus === 'Won') continue; // already handled

      var rpLeadName = String(rePoolData[rp][rpLeadNameCol] || 'A lead');
      notifications.push({
        icon: 'reengage',
        title: 'Re-engage lead ready — ' + rpLeadName,
        sub: 'This lead is due for re-contact today. Check your leads.',
        time: now.toISOString(),
        unread: true,
        type: 'reengage'
      });
    }
  }
}

    // ── RE-ENGAGE: assigned to BDA ──
    if (userRole === 'BDA') {
      var poolSheet = ss.getSheetByName('ReEngage_Pool');
      if (poolSheet && poolSheet.getLastRow() > 1) {
        var poolData = poolSheet.getDataRange().getValues();
        var pH = poolData[0];
        var pAssignedEmailCol = pH.indexOf('ReEngage_Assigned_Email');
        var pBDAEmailCol      = pH.indexOf('BDA_Email');
        var pLeadNameCol      = pH.indexOf('LeadName');
        var pLastUpdCol       = pH.indexOf('Last_Updated');
        for (var i = 1; i < poolData.length; i++) {
          var reAssigned = String(poolData[i][pAssignedEmailCol] || '');
          var origBDA    = String(poolData[i][pBDAEmailCol] || '');
          if (reAssigned === userEmail && origBDA !== userEmail) {
            notifications.push({
              icon: 'reengage',
              title: 'Re-engage lead assigned to you',
              sub: (String(poolData[i][pLeadNameCol] || 'A lead')) + ' from re-engage pool is in your queue',
              time: poolData[i][pLastUpdCol] ? new Date(poolData[i][pLastUpdCol]).toISOString() : now.toISOString(),
              unread: true,
              type: 'reengage'
            });
          }
        }
      }
    }

    // ── RE-ENGAGE: pending review (TL / Admin) ──
    if (userRole === 'Admin' || userRole === 'Team Leader') {
      var poolSheet2 = ss.getSheetByName('ReEngage_Pool');
      if (poolSheet2 && poolSheet2.getLastRow() > 1) {
        var poolData2 = poolSheet2.getDataRange().getValues();
        var pH2 = poolData2[0];
        var pBDAEmailCol2   = pH2.indexOf('BDA_Email');
        var pAfterDateCol2  = pH2.indexOf('ReEngage_After_Date');
        var pendingReEngCount = 0;
        for (var i = 1; i < poolData2.length; i++) {
          var bdaEmail = String(poolData2[i][pBDAEmailCol2] || '');
          var isVisible = userRole === 'Admin' ? true : myBDAEmails[bdaEmail];
          if (isVisible && !poolData2[i][pAfterDateCol2]) pendingReEngCount++;
        }
        if (pendingReEngCount > 0) {
          notifications.push({
            icon: 'reengage_pending',
            title: pendingReEngCount + ' re-engage lead' + (pendingReEngCount > 1 ? 's' : '') + ' awaiting date',
            sub: 'Leads in Pending Review need a date set',
            time: now.toISOString(),
            unread: true,
            type: 'reengage_pending'
          });
        }
      }
    }

    // ── LEAVE APPROVALS (TL / Admin) ──
    if (userRole === 'Admin' || userRole === 'Team Leader') {
      var leavSheet = ss.getSheetByName('Leave_Requests');
      if (leavSheet && leavSheet.getLastRow() > 1) {
        var leavData = leavSheet.getDataRange().getValues();
        var leavH = leavData[0];
        var lApproverCol = leavH.indexOf('ApproverEmail');
        var lStatusColL  = leavH.indexOf('Status');
        var lEmpNameCol  = leavH.indexOf('EmployeeName');
        if (lApproverCol === -1) lApproverCol = 10;
        if (lStatusColL  === -1) lStatusColL  = 11;
        if (lEmpNameCol  === -1) lEmpNameCol  = 2;
        var pendingLeaves = 0;
        var leavNames = [];
        for (var i = 1; i < leavData.length; i++) {
          if (String(leavData[i][lStatusColL] || '') !== 'Pending') continue;
          if (userRole !== 'Admin' && String(leavData[i][lApproverCol] || '') !== userEmail) continue;
          pendingLeaves++;
          if (leavNames.length < 2) leavNames.push(String(leavData[i][lEmpNameCol] || ''));
        }
        if (pendingLeaves > 0) {
          notifications.push({
            icon: 'leave_pending',
            title: pendingLeaves + ' pending leave approval' + (pendingLeaves > 1 ? 's' : ''),
            sub: leavNames.join(', ') + (pendingLeaves > 2 ? ' and more' : '') + ' awaiting your decision',
            time: now.toISOString(),
            unread: true,
            type: 'leave_pending'
          });
        }
      }
    }

    // ── MY LEAVE STATUS CHANGED (within last 7 days) ──
    var leavSheet2 = ss.getSheetByName('Leave_Requests');
    if (leavSheet2 && leavSheet2.getLastRow() > 1) {
      var leavData2 = leavSheet2.getDataRange().getValues();
      var leavH2 = leavData2[0];
      var lEmpEmailCol2 = leavH2.indexOf('EmployeeEmail');
      var lStatusCol2   = leavH2.indexOf('Status');
      var lLeaveTypeCol = leavH2.indexOf('LeaveType');
      var lActionTsCol  = leavH2.indexOf('ActionTimestamp');
      if (lEmpEmailCol2 === -1) lEmpEmailCol2 = 3;
      if (lStatusCol2   === -1) lStatusCol2   = 11;
      if (lLeaveTypeCol === -1) lLeaveTypeCol = 4;
      var weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
      for (var i = 1; i < leavData2.length; i++) {
        if (String(leavData2[i][lEmpEmailCol2] || '') !== userEmail) continue;
        var status2 = String(leavData2[i][lStatusCol2] || '');
        if (status2 !== 'Approved' && status2 !== 'Rejected') continue;
        var actionTs = leavData2[i][lActionTsCol] ? new Date(leavData2[i][lActionTsCol]) : null;
        if (!actionTs || isNaN(actionTs.getTime()) || actionTs < weekAgo) continue;
        var icon2 = status2 === 'Approved' ? 'leave_approved' : 'leave_rejected';
        notifications.push({
          icon: icon2,
          title: 'Leave ' + status2.toLowerCase(),
          sub: String(leavData2[i][lLeaveTypeCol] || '') + ' leave request has been ' + status2.toLowerCase(),
          time: actionTs.toISOString(),
          unread: true,
          type: icon2
        });
      }
    }

    // ── WIN DEBRIEF SUBMITTED (TL / Admin, last 2 days) ──
    if (userRole === 'Admin' || userRole === 'Team Leader') {
      var wlSheet = ss.getSheetByName('Won_Learnings');
      if (wlSheet && wlSheet.getLastRow() > 1) {
        var wlData = wlSheet.getDataRange().getValues();
        var wlH = wlData[0];
        var wlBDAEmailCol  = wlH.indexOf('BDAEmail');
        var wlBDANameCol   = wlH.indexOf('BDAName');
        var wlTimestampCol = wlH.indexOf('Timestamp');
        var wlSkippedCol   = wlH.indexOf('Skipped');
        if (wlTimestampCol === -1) wlTimestampCol = 0;
        var twoDaysAgo = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000);
        for (var i = wlData.length - 1; i >= 1; i--) {
          var debriefEmail = String(wlData[i][wlBDAEmailCol] || '');
          var isVisible2 = userRole === 'Admin' ? true : myBDAEmails[debriefEmail];
          if (!isVisible2) continue;
          var debriefTs = wlData[i][wlTimestampCol] ? new Date(wlData[i][wlTimestampCol]) : null;
          if (!debriefTs || isNaN(debriefTs.getTime()) || debriefTs < twoDaysAgo) continue;
          if (String(wlData[i][wlSkippedCol] || '') === 'TRUE') continue;
          notifications.push({
            icon: 'debrief',
            title: 'Win debrief submitted',
            sub: String(wlData[i][wlBDANameCol] || 'A BDA') + ' submitted a win debrief — check Win Library',
            time: debriefTs.toISOString(),
            unread: true,
            type: 'debrief'
          });
          break;
        }
      }
    }

    // ── CHAT MENTIONS (last 6 hours) ──
    var chatSheet = ss.getSheetByName('ChatLog');
    if (chatSheet && chatSheet.getLastRow() > 1) {
      var chatData = chatSheet.getDataRange().getValues();
      var mentionName = callerName || userEmail;
      var sixHrsAgo = new Date(now.getTime() - 6 * 60 * 60 * 1000);
      var mentionCount = 0;
      for (var i = chatData.length - 1; i >= 1; i--) {
        var ts = chatData[i][0] ? new Date(chatData[i][0]) : null;
        if (!ts || isNaN(ts.getTime()) || ts < sixHrsAgo) break;
        var mentions = chatData[i][7] ? String(chatData[i][7]) : '';
        if (mentions.indexOf(mentionName) !== -1) mentionCount++;
      }
      if (mentionCount > 0) {
        notifications.push({
          icon: 'chat',
          title: 'You were mentioned in chat ' + mentionCount + ' time' + (mentionCount > 1 ? 's' : ''),
          sub: 'Check team chat for your @mentions',
          time: now.toISOString(),
          unread: true,
          type: 'chat'
        });
      }
    }

    // Sort by time desc, limit 20
    notifications.sort(function(a, b) {
      return new Date(b.time || 0) - new Date(a.time || 0);
    });

    return { success: true, notifications: notifications.slice(0, 20) };

  } catch(e) {
    Logger.log('getNotifications error: ' + e.message + '\n' + e.stack);
    return { success: false, notifications: [], message: e.message };
  }
}

function getReEngageDataForLead(leadID) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var poolSheet = ss.getSheetByName('ReEngage_Pool');
    if (!poolSheet || poolSheet.getLastRow() < 2) return {};

    var data    = poolSheet.getDataRange().getValues();
    var headers = data[0];

    var pLeadIDCol        = headers.indexOf('LeadID');
    var pTypeCol          = headers.indexOf('ReEngage_Type');
    var pReasonCol        = headers.indexOf('ReEngage_Reason');
    var pNoteCol          = headers.indexOf('ReEngage_BDA_Note');
    var pAfterDateCol     = headers.indexOf('ReEngage_After_Date');
    var pSubmittedCol     = headers.indexOf('Submitted_At');
    var pTLNoteCol        = headers.indexOf('TL_Note');
    var pAssignedToCol    = headers.indexOf('ReEngage_Assigned_To');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][pLeadIDCol]) !== String(leadID)) continue;

      var afterDate = data[i][pAfterDateCol];
      var afterDateStr = '';
      if (afterDate) {
        try { afterDateStr = Utilities.formatDate(new Date(afterDate), Session.getScriptTimeZone(), 'yyyy-MM-dd'); } catch(e) {}
      }
      var submittedAt = data[i][pSubmittedCol];
      var submittedStr = '';
      if (submittedAt) {
        try { submittedStr = Utilities.formatDate(new Date(submittedAt), Session.getScriptTimeZone(), 'dd MMM yyyy'); } catch(e) {}
      }

      return {
        reEngageType:     String(data[i][pTypeCol]       || ''),
        reEngageReason:   String(data[i][pReasonCol]     || ''),
        reEngageBDANote:  String(data[i][pNoteCol]       || ''),
        reEngageAfterDate: afterDateStr,
        submittedAt:      submittedStr,
        tlNote:           String(pTLNoteCol !== -1 ? data[i][pTLNoteCol] || '' : ''),
        assignedTo:       String(data[i][pAssignedToCol] || '')
      };
    }
    return {};
  } catch(e) {
    Logger.log('getReEngageDataForLead error: ' + e.message);
    return {};
  }
}

// ── CALLS PAGE ──
function getCallsData(userEmail, userRole) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var callsSheet = ss.getSheetByName('Calls');
    if (!callsSheet || callsSheet.getLastRow() < 2) return { success: true, calls: [] };

    var data = callsSheet.getDataRange().getValues();
    var headers = data[0];

    var COL_SERVER_TIME  = headers.indexOf('Server Time');
    var COL_DEVICE_TIME  = headers.indexOf('Device Time');
    var COL_AGENT        = headers.indexOf('Agent');
    var COL_PHONE        = headers.indexOf('Phone Number');
    var COL_PHONE_NORM   = headers.indexOf('Phone Normalized');
    var COL_DUR_S        = headers.indexOf('Duration (s)');
    var COL_DUR          = headers.indexOf('Duration');
    var COL_TYPE         = headers.indexOf('Call Type');
    var COL_LEAD_ID      = headers.indexOf('Lead ID');
    var COL_LEAD_NAME    = headers.indexOf('Lead Name');
    var COL_BDA_NAME     = headers.indexOf('BDA Name');
    var COL_MATCH        = headers.indexOf('Match Status');

    // Build TL -> BDA names map for TL scoping
    var lookupsSheet = ss.getSheetByName('Lookups');
    var lookupsData  = lookupsSheet.getDataRange().getValues();

    var callerName = '';
    for (var i = 1; i < lookupsData.length; i++) {
      if (lookupsData[i][1] === userEmail) { callerName = lookupsData[i][0]; break; }
    }

    var myBDANames = {};
    if (userRole === 'Team Leader') {
      myBDANames[callerName] = true;
      for (var i = 1; i < lookupsData.length; i++) {
        if (lookupsData[i][3] !== 'BDA') continue;
        var rt = String(lookupsData[i][4] || '').trim();
       if (rt.toLowerCase() === userEmail.toLowerCase() || rt.toLowerCase() === callerName.toLowerCase()){
          myBDANames[lookupsData[i][0]] = true;
        }
      }
    }

  var myBDAName = '';
if (userRole === 'BDA') {
  for (var i = 1; i < lookupsData.length; i++) {
    if (lookupsData[i][1] === userEmail) { myBDAName = String(lookupsData[i][0] || '').trim(); break; }
  }
}

    var calls = [];
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row[COL_SERVER_TIME]) continue;

      var bdaName = String(row[COL_BDA_NAME] || '').trim();

      // Scope filter
      if (userRole === 'BDA' && bdaName !== myBDAName) continue;
      if (userRole === 'Team Leader' && !myBDANames[bdaName] && !myBDANames[bdaName.toLowerCase()]) continue;
      // Admin sees all

      var serverTime = row[COL_SERVER_TIME];
      var serverTimeStr = '';
      try {
        serverTimeStr = Utilities.formatDate(
          new Date(serverTime), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'
        );
      } catch(e) { serverTimeStr = String(serverTime); }

      calls.push({
        serverTime : serverTimeStr,
        deviceTime : String(row[COL_DEVICE_TIME] || ''),
        agent      : String(row[COL_AGENT]       || ''),
        phone      : String(row[COL_PHONE]        || ''),
        phoneNorm  : String(row[COL_PHONE_NORM]   || ''),
        durationS  : parseInt(row[COL_DUR_S])     || 0,
        duration   : String(row[COL_DUR]          || ''),
        callType   : String(row[COL_TYPE]         || ''),
        leadId     : String(row[COL_LEAD_ID]      || ''),
        leadName   : String(row[COL_LEAD_NAME]    || ''),
        bdaName    : bdaName,
        matchStatus: String(row[COL_MATCH]        || '')
      });
    }

    // Most recent first
    calls.reverse();

    return { success: true, calls: calls };
  } catch(e) {
    Logger.log('getCallsData error: ' + e.message);
    return { success: false, message: e.message };
  }
}

// ── LEAD SUMMARY FOR CALLS MODAL ──
function getLeadSummaryForCall(leadId) {
  try {
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Leads_Master');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var leadIDCol      = headers.indexOf('LeadID');
    var nameCol        = headers.indexOf('Name');
    var mobileCol      = headers.indexOf('Mobile');
    var serviceCol     = headers.indexOf('Service Required');
    var callStatusCol  = headers.indexOf('Call Status');
    var pipelineCol    = headers.indexOf('Pipeline & Stage');
    var remarksCol     = headers.indexOf('Remarks');
    var followUpCol    = headers.indexOf('Follow-up Date');
    var assignedToCol  = headers.indexOf('Assigned To');
    var teamLeaderCol  = headers.indexOf('Team Leader');
    var leadScoreCol   = headers.indexOf('Lead Score');
    var leadLabelCol   = headers.indexOf('Lead Label');
    var sourceCol      = headers.indexOf('Source');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][leadIDCol]).trim() !== String(leadId).trim()) continue;
      return {
        success    : true,
        leadId     : String(data[i][leadIDCol]     || ''),
        name       : String(data[i][nameCol]        || ''),
        mobile     : String(data[i][mobileCol]      || ''),
        service    : String(data[i][serviceCol]     || ''),
        callStatus : String(data[i][callStatusCol]  || ''),
        pipeline   : String(data[i][pipelineCol]    || ''),
        remarks    : String(data[i][remarksCol]      || ''),
        followUp   : data[i][followUpCol] ? Utilities.formatDate(new Date(data[i][followUpCol]), 'Asia/Kolkata', 'dd MMM yyyy') : '',
        assignedTo : String(data[i][assignedToCol]  || ''),
        teamLeader : String(data[i][teamLeaderCol]  || ''),
        leadScore  : data[i][leadScoreCol] || 0,
        leadLabel  : String(data[i][leadLabelCol]   || ''),
        source     : String(data[i][sourceCol]      || '')
      };
    }
    return { success: false, message: 'Lead not found' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ── CONTACT MANAGEMENT ──

function createPotentialContactsSheet() {
  var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
  var existing = ss.getSheetByName('Potential_Contacts');
  if (existing) {
    return { success: false, message: 'Sheet already exists.' };
  }
  var sheet = ss.insertSheet('Potential_Contacts');
  var headers = [
    'ContactID', 'Name', 'Mobile', 'Email', 'Company', 'Location',
    'Service Interest', 'Notes', 'Status', 'Added By Email',
    'Added By Name', 'Added At', 'Last Updated By Email',
    'Last Updated By Name', 'Last Updated At', 'IsDeleted'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#0b1120')
    .setFontColor('#3ddc84')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('@');
  return { success: true, message: 'Potential_Contacts sheet created successfully.' };
}

function getContacts(callerEmail) {
  try {
    var role = getUserRole(callerEmail);
    if (role !== 'Admin' && role !== 'Team Leader') {
      return { success: false, message: 'Access denied.' };
    }
    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Potential_Contacts');
    if (!sheet || sheet.getLastRow() < 2) return { success: true, contacts: [] };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var contacts = [];

    for (var i = 1; i < data.length; i++) {
      var row = {};
      headers.forEach(function(h, idx) {
        var val = data[i][idx];
        row[h] = val instanceof Date ? val.toISOString() : String(val || '');
      });
      if (row['IsDeleted'] === 'TRUE') continue;
      contacts.push(row);
    }

    // TL sees only contacts they added
    if (role === 'Team Leader') {
      contacts = contacts.filter(function(c) {
        return c['Added By Email'] === callerEmail;
      });
    }

    contacts.sort(function(a, b) {
      return (b['Added At'] || '').localeCompare(a['Added At'] || '');
    });

    return { success: true, contacts: contacts };
  } catch(e) {
    Logger.log('getContacts error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function saveContact(data, callerEmail) {
  try {
    var role = getUserRole(callerEmail);
    if (role !== 'Admin' && role !== 'Team Leader') {
      return { success: false, message: 'Access denied.' };
    }

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Potential_Contacts');
    if (!sheet) return { success: false, message: 'Potential_Contacts sheet not found. Ask Admin to create it.' };

    var sheetData = sheet.getDataRange().getValues();
    var headers = sheetData[0];
    var callerName = getUserNameFromEmail(callerEmail);
    var now = new Date();

    // ── EDIT existing contact ──
    if (data.contactID) {
      var contactIDCol = headers.indexOf('ContactID');
      for (var i = 1; i < sheetData.length; i++) {
        if (String(sheetData[i][contactIDCol]) !== String(data.contactID)) continue;

        var oldRow = {};
        headers.forEach(function(h, idx) { oldRow[h] = sheetData[i][idx]; });

        var fieldsToUpdate = {
          'Name':             data.name || '',
          'Mobile':           data.mobile || '',
          'Email':            data.email || '',
          'Company':          data.company || '',
          'Location':         data.location || '',
          'Service Interest': data.serviceInterest || '',
          'Notes':            data.notes || '',
          'Status':           data.status || 'New',
          'Last Updated By Email': callerEmail,
          'Last Updated By Name':  callerName,
          'Last Updated At':       now
        };

        Object.keys(fieldsToUpdate).forEach(function(field) {
          var col = headers.indexOf(field);
          if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(fieldsToUpdate[field]);
        });

        // Log changes to activity log
        Object.keys(fieldsToUpdate).forEach(function(field) {
          if (field.indexOf('Last Updated') !== -1) return;
          var oldVal = String(oldRow[field] || '');
          var newVal = String(fieldsToUpdate[field] || '');
          if (oldVal !== newVal) {
            appendActivityLog(
              'Contact_' + data.contactID,
              data.name || oldRow['Name'] || '',
              'Contact Field Updated',
              field + ' changed',
              oldVal, newVal, callerEmail
            );
          }
        });

        return { success: true, message: 'Contact updated.' };
      }
      return { success: false, message: 'Contact not found.' };
    }

    // ── ADD new contact ──
    var newID = 'CT-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    var newRow = new Array(headers.length).fill('');

    var fieldMap = {
      'ContactID':            newID,
      'Name':                 data.name || '',
      'Mobile':               data.mobile || '',
      'Email':                data.email || '',
      'Company':              data.company || '',
      'Location':             data.location || '',
      'Service Interest':     data.serviceInterest || '',
      'Notes':                data.notes || '',
      'Status':               data.status || 'New',
      'Added By Email':       callerEmail,
      'Added By Name':        callerName,
      'Added At':             now,
      'Last Updated By Email': callerEmail,
      'Last Updated By Name':  callerName,
      'Last Updated At':       now,
      'IsDeleted':            'FALSE'
    };

    Object.keys(fieldMap).forEach(function(field) {
      var col = headers.indexOf(field);
      if (col !== -1) newRow[col] = fieldMap[field];
    });

    sheet.appendRow(newRow);
    var newRowNum = sheet.getLastRow();
    sheet.getRange(newRowNum, 1).setNumberFormat('@');

    appendActivityLog(
      'Contact_' + newID,
      data.name || '',
      'Contact Created',
      'New potential contact added',
      '', data.status || 'New', callerEmail
    );

    return { success: true, message: 'Contact added.', contactID: newID };
  } catch(e) {
    Logger.log('saveContact error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function deleteContact(contactID, callerEmail) {
  try {
    var role = getUserRole(callerEmail);
    if (role !== 'Admin' && role !== 'Team Leader') {
      return { success: false, message: 'Access denied.' };
    }

    var ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    var sheet = ss.getSheetByName('Potential_Contacts');
    if (!sheet) return { success: false, message: 'Sheet not found.' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var contactIDCol  = headers.indexOf('ContactID');
    var isDeletedCol  = headers.indexOf('IsDeleted');
    var nameCol       = headers.indexOf('Name');

    // TL can only delete their own contacts
    var addedByCol = headers.indexOf('Added By Email');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][contactIDCol]) !== String(contactID)) continue;

      if (role === 'Team Leader' && data[i][addedByCol] !== callerEmail) {
        return { success: false, message: 'You can only delete contacts you added.' };
      }

      sheet.getRange(i + 1, isDeletedCol + 1).setValue('TRUE');

      appendActivityLog(
        'Contact_' + contactID,
        String(data[i][nameCol] || ''),
        'Contact Deleted',
        'Contact soft-deleted',
        '', 'Deleted', callerEmail
      );

      return { success: true, message: 'Contact deleted.' };
    }
    return { success: false, message: 'Contact not found.' };
  } catch(e) {
    Logger.log('deleteContact error: ' + e.message);
    return { success: false, message: e.message };
  }
}

function getContactHistory(contactID, callerEmail) {
  try {
    var role = getUserRole(callerEmail);
    if (role !== 'Admin' && role !== 'Team Leader') {
      return { success: false, message: 'Access denied.' };
    }
    var raw = getLeadTimeline('Contact_' + contactID);
    var parsed = [];
    try { parsed = JSON.parse(raw || '[]'); } catch(e) {}
    return { success: true, history: parsed };
  } catch(e) {
    return { success: false, message: e.message };
  }
}
function getFinancePendingPayments() {
  try {
    const ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    const sheet = ss.getSheetByName('Payments');
    if (!sheet) return { success: false, error: 'Payments sheet not found' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    // ── Build CRM screenshot lookup by Transaction ID ──
    const leadsSheet = ss.getSheetByName('Leads_Master');
    const crmScreenshotMap = {};
    if (leadsSheet) {
      const leadsData = leadsSheet.getDataRange().getValues();
      const lHeaders = leadsData[0];
      const txnCol = lHeaders.indexOf('Transaction ID');
      const shotCol = lHeaders.indexOf('Payment_Screenshot_FileId');
      if (txnCol !== -1 && shotCol !== -1) {
        for (let i = 1; i < leadsData.length; i++) {
          const txnId = String(leadsData[i][txnCol] || '').trim();
          const fileId = leadsData[i][shotCol];
          if (txnId && fileId) {
            crmScreenshotMap[txnId] = 'https://drive.google.com/file/d/' + fileId + '/view';
          }
        }
      }
    }

    const pending = [];

    rows.forEach((row, i) => {
      const verificationStatus = row[17]; // Col R
      if (verificationStatus === '' || verificationStatus === null || verificationStatus === undefined) {
        const crmId = String(row[6] || '').trim();
        pending.push({
          rowIndex: i + 2, // actual sheet row
          date: row[0] ? row[0].toString() : '',
          executiveName: row[1],
          executiveEmail: row[2],
          companyName: row[3],
          contactNumber: row[4],
          userType: row[5],
          crmUserId: row[6],
          products: row[7],
          productDescription: row[8],
          creditsToAdd: row[9],
          paymentMode: row[10],
          originalAmount: row[11],
          finalAmount: row[12],
          taxPaid: row[13],
          gstin: row[14],
          tdsAmount: row[15],
          submittedAt: row[16] ? row[16].toString() : '',
          screenshot: row[21],
          crmScreenshot: crmScreenshotMap[crmId] || ''
        });
      }
    });

    return { success: true, data: pending };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
function verifyPayment(crmUserId) {
  try {
    const ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    const paymentsSheet = ss.getSheetByName('Payments');
    const reviewsSheet = ss.getSheetByName('Payment_Reviews');
    const leadsSheet = ss.getSheetByName('Leads_Master');

    if (!paymentsSheet || !reviewsSheet || !leadsSheet) 
      return { success: false, error: 'Sheet not found' };

    const payData = paymentsSheet.getDataRange().getValues();
    let targetRow = -1;
    let companyName = '';
    let executiveEmail = '';
    let executiveName = '';  // ← ADDED

    for (let i = 1; i < payData.length; i++) {
      if (String(payData[i][6]) === String(crmUserId) && (payData[i][17] === '' || payData[i][17] === null)) {
        targetRow = i + 1;
        executiveName = payData[i][1];  // ← ADDED (col B)
        executiveEmail = payData[i][2];
        companyName = payData[i][3]; 
        break;
      }
    }

    if (targetRow === -1) return { success: false, error: 'Payment record not found or already reviewed' };

    const now = new Date();
    const reviewer = Session.getActiveUser().getEmail();

    // Update Payments sheet
    paymentsSheet.getRange(targetRow, 18).setValue('VERIFIED');     // Col R
    paymentsSheet.getRange(targetRow, 19).setValue(reviewer);       // Col S
    paymentsSheet.getRange(targetRow, 20).setValue(now);            // Col T

    // Update Leads_Master — find by Transaction ID (col W) ← FIXED
    const leadsData = leadsSheet.getDataRange().getValues();
    const lHeaders = leadsData[0];
    const txnCol = lHeaders.indexOf('Transaction ID');
    for (let i = 1; i < leadsData.length; i++) {
      if (String(leadsData[i][txnCol]) === String(crmUserId)) {
        leadsSheet.getRange(i + 1, 19).setValue('Won');        // Col S - Pipeline & Stage
        leadsSheet.getRange(i + 1, 24).setValue('VERIFIED');   // Col X - Verification Status
        break;
      }
    }

    // Write to Payment_Reviews
    const reviewsData = reviewsSheet.getDataRange().getValues();
    const lastRow = reviewsData.length;
    const reviewId = 'REV-' + String(lastRow).padStart(3, '0');

    reviewsSheet.appendRow([
      reviewId,        // A - Review_ID
      crmUserId,       // B - CRM_ID
      'VERIFIED',      // C - Action
      '',              // D - Remark
      reviewer,        // E - Reviewed_By
      now,             // F - Reviewed_At
      '',              // G - Previous_Status
      'VERIFIED',      // H - New_Status
      '',              // I - Notified_To
      'NO'             // J - Notification_Sent
    ]);

    // ── NOTIFY BDA ON PAYMENT APPROVAL ── ← ADDED
    try {
      if (executiveEmail) {
        GmailApp.sendEmail(
          executiveEmail,
          ' Payment Verified — ' + companyName + ' (' + crmUserId + ')',
          'Hi ' + (executiveName || 'there') + ',\n\n' +
          'Great news! The payment for the following deal has been verified by Finance.\n\n' +
          '━━━━━━━━━━━━━━━━━━━━\n' +
          'CRM / Transaction ID : ' + crmUserId + '\n' +
          'Company              : ' + companyName + '\n' +
          'Verified By          : ' + reviewer + '\n' +
          'Verified At          : ' + now + '\n' +
          '━━━━━━━━━━━━━━━━━━━━\n\n' +
          'The deal has been marked as WON in the CRM.\n\n' +
          'Next steps:\n' +
          '1. Initiate client onboarding\n' +
          '2. Send WhatsApp + Email to client with onboarding form\n' +
          '3. Collect required documents from client\n\n' +
          'Keep up the great work! 🎉\n\n' +
          'Regards,\n' +
          'NotifyTechAI Finance Team'
        );
      }
    } catch (mailErr) {
      Logger.log('BDA approval email failed: ' + mailErr.message);
    }

    return { success: true, message: 'Payment verified successfully' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function rejectPayment(crmUserId, remark) {
  try {
    if (!remark || remark.trim() === '') {
      return { success: false, error: 'Remark is required for rejection' };
    }

    const ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    const paymentsSheet = ss.getSheetByName('Payments');
    const reviewsSheet = ss.getSheetByName('Payment_Reviews');
    const leadsSheet = ss.getSheetByName('Leads_Master');

    if (!paymentsSheet || !reviewsSheet || !leadsSheet)
      return { success: false, error: 'Sheet not found' };

    const payData = paymentsSheet.getDataRange().getValues();
    let targetRow = -1;
    let companyName = '';
    let executiveEmail = '';
    let executiveName = '';
    let finalAmount = '';
    let paymentMode = '';

    for (let i = 1; i < payData.length; i++) {
      if (String(payData[i][6]) === String(crmUserId) && (payData[i][17] === '' || payData[i][17] === null)) {
        targetRow = i + 1;
        executiveName = payData[i][1];
        executiveEmail = payData[i][2];
        companyName = payData[i][3];
        paymentMode = payData[i][10];
        finalAmount = payData[i][12];
        break;
      }
    }

    if (targetRow === -1) return { success: false, error: 'Payment record not found or already reviewed' };

    const now = new Date();
    const reviewer = Session.getActiveUser().getEmail();

    // Update Payments sheet
    paymentsSheet.getRange(targetRow, 18).setValue('REJECTED');      // Col R
    paymentsSheet.getRange(targetRow, 19).setValue(reviewer);        // Col S
    paymentsSheet.getRange(targetRow, 20).setValue(now);             // Col T
    paymentsSheet.getRange(targetRow, 21).setValue(remark);          // Col U
// Update Leads_Master — On Hold
const leadsData = leadsSheet.getDataRange().getValues();
const lHeaders = leadsData[0];
const txnCol = lHeaders.indexOf('Transaction ID');
for (let i = 1; i < leadsData.length; i++) {
  if (String(leadsData[i][txnCol]) === String(crmUserId)) {
    leadsSheet.getRange(i + 1, 19).setValue('On Hold');          // Col S - Pipeline & Stage
    leadsSheet.getRange(i + 1, 24).setValue('REJECTED');         // Col X - Verification Status
    break;
  }
}

    // Write to Payment_Reviews
    const reviewsData = reviewsSheet.getDataRange().getValues();
    const lastRow = reviewsData.length;
    const reviewId = 'REV-' + String(lastRow).padStart(3, '0');

    reviewsSheet.appendRow([
      reviewId,
      crmUserId,
      'REJECTED',
      remark,
      reviewer,
      now,
      'PAYMENT_SUBMITTED',
      'ON_HOLD',
      'ksidhant880@gmail.com',
      'NO'
    ]);

    // Email to Sidhant
    GmailApp.sendEmail(
      'ksidhant880@gmail.com',
      'Payment Rejected — ' + companyName + ' (' + crmUserId + ')',
      'A payment has been rejected.\n\n' +
      'CRM ID: ' + crmUserId + '\n' +
      'Company: ' + companyName + '\n' +
      'BDA: ' + executiveName + ' (' + executiveEmail + ')\n' +
      'Amount: ₹' + finalAmount + '\n' +
      'Payment Mode: ' + paymentMode + '\n' +
      'Rejected By: ' + reviewer + '\n' +
      'Rejected At: ' + now + '\n' +
      'Remark: ' + remark + '\n\n' +
      'Please take necessary action.'
    );

    // Update Payment_Reviews notification status
    const updatedRow = reviewsSheet.getLastRow();
    reviewsSheet.getRange(updatedRow, 10).setValue('YES');           // Col J

    return { success: true, message: 'Payment rejected and Sidhant notified' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function requestInfo(crmUserId, message) {
  try {
    const ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    const paymentsSheet = ss.getSheetByName('Payments');
    const reviewsSheet = ss.getSheetByName('Payment_Reviews');

    if (!paymentsSheet || !reviewsSheet)
      return { success: false, error: 'Sheet not found' };

    const payData = paymentsSheet.getDataRange().getValues();
    let targetRow = -1;
    let executiveEmail = '';
    let executiveName = '';
    let companyName = '';

    for (let i = 1; i < payData.length; i++) {
      if (String(payData[i][6]) === String(crmUserId) && (payData[i][17] === '' || payData[i][17] === null)) {
        targetRow = i + 1;
        executiveName = payData[i][1];
        executiveEmail = payData[i][2];
        companyName = payData[i][3];
        break;
      }
    }

    if (targetRow === -1) return { success: false, error: 'Payment record not found or already reviewed' };

    const now = new Date();
    const reviewer = Session.getActiveUser().getEmail();

    // Update Payments sheet status
    paymentsSheet.getRange(targetRow, 18).setValue('INFO_REQUESTED');  // Col R

    // Write to Payment_Reviews
    const reviewsData = reviewsSheet.getDataRange().getValues();
    const lastRow = reviewsData.length;
    const reviewId = 'REV-' + String(lastRow).padStart(3, '0');

    reviewsSheet.appendRow([
      reviewId,
      crmUserId,
      'INFO_REQUESTED',
      message || '',
      reviewer,
      now,
      'PAYMENT_SUBMITTED',
      'INFO_REQUESTED',
      executiveEmail,
      'NO'
    ]);

    // Email to BDA
    if (executiveEmail) {
      const emailBody = 'Hi ' + executiveName + ',\n\n' +
        'Additional information has been requested for the payment submitted for ' + companyName + '.\n\n' +
        'CRM ID: ' + crmUserId + '\n' +
        (message ? 'Message from Finance:\n' + message + '\n\n' : '') +
        'Please update the payment details or resubmit with the required information.\n\n' +
        'Regards,\nNotifyTechAI Finance Team';

      GmailApp.sendEmail(
        executiveEmail,
        'Action Required — Payment Info Needed (' + crmUserId + ')',
        emailBody
      );

      // Update notification status
      const updatedRow = reviewsSheet.getLastRow();
      reviewsSheet.getRange(updatedRow, 10).setValue('YES');  // Col J
    }

    return { success: true, message: 'Info requested and BDA notified' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getReviewHistory() {
  try {
    const ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    const sheet = ss.getSheetByName('Payment_Reviews');
    if (!sheet) return { success: false, error: 'Payment_Reviews sheet not found' };
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };
    return { success: true, data: data.slice(1) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
// ── AUTO REMINDER: Unverified Payments > 48 hours ──
function sendUnverifiedPaymentReminders() {
  try {
    const ss = SpreadsheetApp.openById('1ugzWsQDRiXiWAfQYUPGxMmznYGawbw05shR3lgj_0r8');
    const paymentsSheet = ss.getSheetByName('Payments');
    if (!paymentsSheet) return;

    const data = paymentsSheet.getDataRange().getValues();
    const headers = data[0];
    const now = new Date();
    const cutoff = new Date(now.getTime() - 48 * 60 * 60 * 1000); // 48 hours ago

    const FINANCE_EMAIL = 'developernotifytechai@gmail.com'; // ← change to actual finance email if different

    let remindersSent = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip if already verified/rejected
      const verificationStatus = row[17]; // Col R
      if (verificationStatus !== '' && verificationStatus !== null && verificationStatus !== undefined) continue;

      // Check submitted date — col Q (index 16)
      const submittedAt = row[16];
      if (!submittedAt) continue;

      const submittedDate = new Date(submittedAt);
      if (isNaN(submittedDate.getTime())) continue;

      // Only remind if older than 48 hours
      if (submittedDate > cutoff) continue;

      const executiveName  = row[1]  || 'BDA';
      const executiveEmail = row[2]  || '';
      const companyName    = row[3]  || '';
      const crmUserId      = row[6]  || '';
      const finalAmount    = row[12] || '';
      const paymentMode    = row[10] || '';

      const hoursOld = Math.floor((now - submittedDate) / (1000 * 60 * 60));

      // ── Email Finance ──
      try {
        GmailApp.sendEmail(
          FINANCE_EMAIL,
          '⏰ Payment Pending Review — ' + hoursOld + 'h overdue (' + crmUserId + ')',
          'Hi Finance Team,\n\n' +
          'The following payment has been pending verification for ' + hoursOld + ' hours.\n\n' +
          '━━━━━━━━━━━━━━━━━━━━\n' +
          'CRM / Transaction ID : ' + crmUserId + '\n' +
          'Company              : ' + companyName + '\n' +
          'BDA                  : ' + executiveName + ' (' + executiveEmail + ')\n' +
          'Amount               : ₹' + finalAmount + '\n' +
          'Payment Mode         : ' + paymentMode + '\n' +
          'Submitted At         : ' + submittedDate + '\n' +
          'Hours Pending        : ' + hoursOld + 'h\n' +
          '━━━━━━━━━━━━━━━━━━━━\n\n' +
          'Please review and take action at the earliest.\n\n' +
          'Regards,\n' +
          'NotifyTechAI CRM System'
        );
      } catch (e) {
        Logger.log('Finance reminder email failed: ' + e.message);
      }

      // ── Email BDA ──
      try {
        if (executiveEmail) {
          GmailApp.sendEmail(
            executiveEmail,
            '⏰ Your Payment is Still Under Review — ' + companyName,
            'Hi ' + executiveName + ',\n\n' +
            'Your payment submission for ' + companyName + ' has been pending Finance review for ' + hoursOld + ' hours.\n\n' +
            '━━━━━━━━━━━━━━━━━━━━\n' +
            'CRM / Transaction ID : ' + crmUserId + '\n' +
            'Company              : ' + companyName + '\n' +
            'Amount               : ₹' + finalAmount + '\n' +
            'Submitted At         : ' + submittedDate + '\n' +
            '━━━━━━━━━━━━━━━━━━━━\n\n' +
            'Finance has been reminded to review your payment.\n' +
            'You will be notified once it is approved or if any action is needed.\n\n' +
            'Regards,\n' +
            'NotifyTechAI Finance Team'
          );
        }
      } catch (e) {
        Logger.log('BDA reminder email failed: ' + e.message);
      }

      remindersSent++;
      Logger.log('Reminder sent for: ' + crmUserId + ' (' + hoursOld + 'h old)');
    }

    Logger.log('Total reminders sent: ' + remindersSent);
  } catch (e) {
    Logger.log('sendUnverifiedPaymentReminders error: ' + e.message);
  }
}
// ── CREATE TRIGGER (run this ONCE manually) ──
function createPaymentReminderTrigger() {
  // Delete existing trigger if any to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendUnverifiedPaymentReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Run every 6 hours
  ScriptApp.newTrigger('sendUnverifiedPaymentReminders')
    .timeBased()
    .everyHours(6)
    .create();

  Logger.log('✅ Payment reminder trigger created — runs every 6 hours');
}