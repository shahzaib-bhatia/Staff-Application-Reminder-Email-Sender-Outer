function myFunction() {
  
}

function scheduledReminder() {
  var contacts = new Map;
  contacts = getContacts();

  var spreadsheetParent = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(spreadsheetParent.getSheetByName('Staff Review'));
  var applicationSheet = SpreadsheetApp.getActiveSheet();
  var applications = applicationSheet.getRange("A2:K").getValues();

  var acceptedApps = new Map;
  var pendingApps = new Map;
  var rejectedApps = new Array;
  var unknownApps = new Array;

  for (x=0; x<applications.length; x++) {
    let application = applications[x];
    let appFirstName = application[0];
    let appLastName = application[1];
    let appEmail = application[3];
    let appSC1Name = application[6];
    let appSC1Status = application[7];
    let appSC2Name = application[8];
    let appSC2Status = application[9];

    // Probably an empty row. Skip.
    if (!appFirstName){ 
      continue;
    }

    // what it says on the tin
    if (!appEmail) { 
      appEmail = "unknown email - check with HR to create account";
    }

    let appSummary = appFirstName + " " + appLastName + " <" + appEmail + ">";
    Logger.log("Looking at application: " + appSummary);

    // If they have been denied by the first department just shuffle the vars around so we don't need to think too hard.
    if (appSC1Status == 'Denied' || appSC1Status == 'Withdrawn') { 
      appSC1Status = appSC2Status;
      appSC1Name = appSC2Name;
    }

    // They were rejected by both. Technically we don't need to check both (see above) but it's easier to read this way.
    if (((appSC1Status == 'Denied' || appSC1Status == 'Withdrawn') && (appSC2Status == 'Denied' || appSC2Status == 'Withdrawn'))) {
      rejectedApps.push(appSummary);
      continue;
    }

    // They were accepted to both choices. We'll handle the second one here and let the regular handler do the first one.
    if (appSC1Status == 'Confirmed' && appSC2Status == 'Confirmed') { 
      appSummary = appSummary + " (Shared: " + appSC1Name + "/" + appSC2Name + ")";
      // I hate that we have to do this but we can't pretend that it's an array until it is an array
      if (!acceptedApps.has(appSC2Name)) {
        acceptedApps.set(appSC2Name,[]);
      }
      acceptedApps.set(appSC2Name,[...acceptedApps.get(appSC2Name), appSummary]);
      // Don't 'continue' since we still have to handle the first one (later)
    }

    // They were accepted to their primary choice
    if (appSC1Status == 'Confirmed') {
      if (!acceptedApps.has(appSC1Name)) {
        acceptedApps.set(appSC1Name,[]);
      }
      acceptedApps.set(appSC1Name,[...acceptedApps.get(appSC1Name), appSummary]);
      continue;
    }

    // Application is still pending
    if (appSC1Status == '' || appSC1Status == 'Contacted') {
      if (!pendingApps.has(appSC1Name)) {
        pendingApps.set(appSC1Name,[]);
      }
      appSummary = appSummary + " (Next choice: " + appSC2Name + ")";     
      pendingApps.set(appSC1Name,[...pendingApps.get(appSC1Name), appSummary]);
      continue;
    }

    // We shouldn't be here
    unknownApps.push(appSummary);
  }

  for (let [dept, addr] of contacts.entries()) {
    if (!addr.endsWith("@example.org")){
      Logger.log("Skipping " + dept + " due to bad email address.")
      continue;
    }
    if (!pendingApps.has(dept)){
      pendingApps.set(dept,[]);
    }
    if (!acceptedApps.has(dept)){
      acceptedApps.set(dept,[]);
    }
    sendEmail(dept, addr, acceptedApps.get(dept), pendingApps.get(dept));
    pendingApps.delete(dept);
    acceptedApps.delete(dept);
  }

  let uncontacted = [...new Set([ ...pendingApps.keys(), ...acceptedApps.keys() ])];
  Logger.log(uncontacted);
  let errorReport = new String;

  if (!uncontacted.length == 0) {
    errorReport += "Departments without contact addresses:\n";
      for (let uncontactedAddr of uncontacted){
        errorReport += " - " + uncontactedAddr + "\n";
      }
  } else {
    errorReport += "All departments had contact addresses! (gold star)\n"
  }
  errorReport += divider;
  if (!rejectedApps.length == 0) {
    errorReport += "Applicants rejected from all requested departments:\n";
    for (let rejected of rejectedApps){
        errorReport += " - " + rejected + "\n";
      }
  } else {
    errorReport += "No rejected applicants found! (gold star)\n"
  }
  errorReport += divider;
  if (!unknownApps.length == 0) {
    errorReport += "Applicants with unknown state:\n";
    for (let unknown of unknownApps){
        errorReport += " - " + unknown + "\n";
      }
  } else {
    errorReport += "No unknown state applicants found! (gold star)\n"
  }
  errorReport += divider;
  Logger.log(errorReport);
  MailApp.sendEmail({
    to: "nullroute@example.org, blackhole@example.org",
    subject: "Staff App Email Reminder Error Report",
    body: errorReport,
    replyTo: "blackhole@example.org"
  })
}

function getContacts() {
  var spreadsheetParent = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(spreadsheetParent.getSheetByName('Sector Contacts (for reminder script)'));
  var contactSheet = SpreadsheetApp.getActiveSheet();

  var contacts = contactSheet.getRange('A2:B').getValues();
  var contactsMap = new Map;
    for (x=0; x<contacts.length; x++) {
      if (!contacts[x][0]){
        continue;
      }
      contactsMap.set(contacts[x][0],contacts[x][1]);
    }
  return contactsMap;
}

function sendEmail (dept, addr, acceptedApps, pendingApps) {
  var template = HtmlService.createTemplateFromFile('EmailTemplate');
  template.dept = dept
  template.addr = addr
  template.pending = pendingApps;
  template.accepted = acceptedApps;
  Logger.log("Mail Quota Remaining: " + MailApp.getRemainingDailyQuota());
  MailApp.sendEmail({
    to: addr,
    subject: "Staff Application Status for " + dept,
    htmlBody: template.evaluate().getContent(),
    replyTo: "nullroute@example.org"
  })
}
