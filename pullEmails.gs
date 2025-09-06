/*
How to use the stop feature
Open the Apps Script editor.

Go to View > Show project properties > Script properties (or use the Apps Script API).

Add a property:

Key: stopNow

Value: true

The next time main() runs, it will detect the flag, write the report, clear progress, and stop.

After stopping, the flag is automatically cleared so you can restart fresh later.

https://developers.google.com/apps-script/guides/services/quotas

Find email sent to targetEmail in the last 15 years
deliveredto:myemail@gmail.com after:2009/01/01 before:2025/01/01

*/

function Triggered_Search() {
  //const TARGET_EMAIL = 'adevine@gmail.com';  // Change if needed
  //const BATCH_SIZE = 1000;                    // Number of threads to process per run
  //const query = 'deliveredto:' + TARGET_EMAIL;
  const props = PropertiesService.getScriptProperties();

  // Setup variables
  const batchSize = parseInt(props.getProperty('batchSize'));
  const targetEmail = props.getProperty('targetEmail');
  const query = 'deliveredto:' + targetEmail + ' ' + 'after:2009/01/01 before:2025/01/01';
  let start = parseInt(props.getProperty('start')) || 0;
  let iteration = parseInt(props.getProperty('iteration')) || 0;

  Logger.log('targetEmail: ' + targetEmail);
  Logger.log('batchSize: ' + batchSize);
  Logger.log('stopNow: ' + props.getProperty('stopNow'));
  
  // Check for stop flag
  //if(true) {
  if (props.getProperty('stopNow') === 'true') {
    const savedCounts = props.getProperty('domainCounts');
    const domainCounts = savedCounts ? JSON.parse(savedCounts) : {};
    writeResultsAndClear(domainCounts, targetEmail, props);
    props.deleteProperty('stopNow'); // reset flag
    Logger.log('Stop flag detected: finishing and exiting.');
    return 'Stopped early and wrote results.';
  }
  
  Logger.log('Starting iteration:' + iteration);
  Logger.log('Starting index: ' + start);

  // Load progress
  let domainCounts = {};
  const savedCounts = props.getProperty('domainCounts');
  if (savedCounts) {
    domainCounts = JSON.parse(savedCounts);
  }

  // Fetch batch of threads that match query
  const threads = GmailApp.search(query, start, batchSize);
  if (!threads || threads.length === 0) {
    // No more threads - finalize
    writeResultsAndClear(domainCounts, targetEmail, props);
    return;
  }

  const messagesPerThread = GmailApp.getMessagesForThreads(threads);

  for (let t = 0; t < messagesPerThread.length; t++) {
    const msgs = messagesPerThread[t];
    for (let m = 0; m < msgs.length; m++) {
      const msg = msgs[m];
      const from = msg.getFrom();

      // Extract sender email
      let emailMatch = from.match(/<([^>]+)>/);
      let senderEmail;
      if (emailMatch && emailMatch[1]) {
        senderEmail = emailMatch[1];
      } else {
        emailMatch = from.match(/([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/);
        senderEmail = emailMatch ? emailMatch[1] : null;
      }

      if (senderEmail) {
        const domain = senderEmail.split('@').pop().toLowerCase();
        if (domain) domainCounts[domain] = (domainCounts[domain] || 0) + 1;
      }
    }
  }
  iteration++;
  
  // Save progress
  start += threads.length;
  props.setProperty('start', start.toString());
  props.setProperty('iteration', iteration.toString());
  props.setProperty('domainCounts', JSON.stringify(domainCounts));

  Logger.log(`Processed ${start} threads so far. Run the script again to continue.`);
  return `Processed ${start} threads so far. Run the script again to continue.`;
}

function writeResultsAndClear(domainCounts, targetEmail, props) {
  // Convert counts to sorted array
  const rows = Object.keys(domainCounts).map(d => [d, domainCounts[d]]);
  rows.sort((a, b) => b[1] - a[1]);

  // Create spreadsheet and write results
  const ssName = `Sender domains for ${targetEmail} - ${new Date().toISOString().slice(0,10)}`;
  const ss = SpreadsheetApp.create(ssName);
  const sheet = ss.getActiveSheet();
  sheet.appendRow(['domain', 'message_count']);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 2).setValues(rows);

  Logger.log(`Finished processing! Report written to: ${ss.getUrl()}`);
  Logger.log('start: ' + start + ' iteration: ' + iteration);

  // Clear stored progress
  props.deleteProperty('start');
  props.deleteProperty('domainCounts');
  props.deleteProperty('iteration');
}

/**
 * Utility: Set the stop flag (optional helper you can run manually)
 */
function setStopFlag() {
  PropertiesService.getScriptProperties().setProperty('stopNow', 'true');
  Logger.log('Set stopNow = true. Next run will finalize and write results.');
}

/**
 * Utility: Clear saved progress (use with care)
 */
function clearProgress() {
  PropertiesService.getScriptProperties().deleteProperty('start');
  PropertiesService.getScriptProperties().deleteProperty('domainCounts');
  PropertiesService.getScriptProperties().deleteProperty('stopNow');
  PropertiesService.getScriptProperties().deleteProperty('batchSize');
  PropertiesService.getScriptProperties().deleteProperty('targetEmail');

  Logger.log('Cleared progress, stop flag, targetEmail, and batchSize.');
}

/**
 * Utility: Defines properties
 */
function setupScript() {
  PropertiesService.getScriptProperties().setProperty('start','0');
  PropertiesService.getScriptProperties().setProperty('stopNow', 'false');
  PropertiesService.getScriptProperties().setProperty('batchSize', '500');
  PropertiesService.getScriptProperties().setProperty('targetEmail', 'adevine@gmail.com');
  
  Logger.log('Setup parameters.');
}

function setupSpreadsheet() {
  PropertiesService.getScriptProperties().setProperty('spreadsht', 'https://docs.google.com/spreadsheets/d/1BIgaUspNot3d9TWD3oxadNjAULzhxu311TrU_7kKi0g/');
}

function categorizeDomainsFromSheet() {
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = PropertiesService.getScriptProperties().getProperty('spreadsht');
  var ss = SpreadsheetApp.openByUrl(f);
  var sheet = ss.getSheetByName("Sheet1");
  if (!sheet) {
    Logger.log("Sheet " + f + " not found.");
    return;
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues(); // assumes domains start in A2
  var categorizedData = [];

  // Lists of known domains
  var personalDomains = ["gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "icloud.com", "protonmail.com"];
  var corporateDomains = ["openai.com", "microsoft.com", "apple.com"]; // Add your trusted work domains here
  var financialDomains = ["paypal.com", "chase.com", "bankofamerica.com", "wellsfargo.com", "bmo.com"];
  var retailDomains = ["amazon.com", "ebay.com", "etsy.com", "michaels.com", "vs.com", "vspink.com", "victoriassecret.com", "starbucks.com", "ulta.com", "kohls.com"];
  var socialDomains = ["facebookmail.com", "linkedin.com", "twitter.com", "instagram.com", "google.com"];
  var marketingDomains = ["mailchimp.com", "sendgrid.net", "constantcontact.com", "hubspot.com"];
  var securityDomains = ["securityweek.com", "sans.org","paloaltonetworks.com", "fortinet.com", "checkpoint.com", "cisco.com", "ciscosecure.com",
  "crowdstrike.com", "rapid7.com", "qualys.com", "tenable.com", "splunk.com",
  "arcticwolf.com", "kaspersky.com", "trendmicro.com", "sophos.com", "mcafee.com",
  "bitdefender.com", "eset.com", "avast.com", "sentinelone.com", "fireeye.com", "trellix.com",
  "mitre.org", "cisa.gov", "us-cert.gov", "sans.org", "first.org", "abuse.ch", "shadowserver.org",
  "threatpost.com", "darkreading.com", "thehackernews.com", "bleepingcomputer.com",
  "isc2.org", "isaca.org", "nist.gov", "owasp.org", "enisa.europa.eu", "antisyphontraining.com", "blackhillsinfosec.com", "comptia.org", "tryhackme.com", "activecountermeasures.com"];
  var jobDomains = ["clearancejobs.com", "careerbuilder.com", "ziprecruiter.com", "monster.com"];
  var allowList = ["bluehalo.com", "depaul.edu", "tacobell.com", "saloninteractive.com"];


  // Function to categorize domain
  function categorizeDomain(domain) {
    var d = domain.toLowerCase().trim();
    // Match Subdomains
    if (personalDomains.some(base => domain.endsWith(base))) return ["Personal/Free", "Low"];
    if (corporateDomains.some(base => domain.endsWith(base))) return ["Corporate", "Low"];
    if (financialDomains.some(base => domain.endsWith(base))) return ["Financial/Banking", "Medium"];
    if (retailDomains.some(base => domain.endsWith(base))) return ["Retail/E-commerce", "Medium"];
    if (socialDomains.some(base => domain.endsWith(base))) return ["Social Media", "Medium"];
    
    if (marketingDomains.some(base => domain.endsWith(base))) return ["Marketing/Newsletter", "Medium"];

    if (securityDomains.some(base => domain.endsWith(base))) {
      return ["Cybersecurity / Trusted","Low"];
    }

    if (jobDomains.some(base => domain.endsWith(base))) return ["Job Search", "Low"];

    if (allowList.some(base => domain.endsWith(base))) return ["Allow List", "Low"];
    if (d.match(/\.(biz|click|top|xyz)$/) || d.match(/(prize|lotto|cheap|offer)/)) return ["Known Spam/Junk", "High"];
    return ["Unknown", "Medium"];
  }

  // Process all rows
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0]) {
      categorizedData.push(["", ""]);
      continue;
    }
    var categoryRisk = categorizeDomain(data[i][0]);
    categorizedData.push(categoryRisk);
  }

  // Write results to columns C and D
  sheet.getRange(2, 3, categorizedData.length, 2).setValues(categorizedData);

  Logger.log("Categorization complete. Added Category and Risk Level.");
}

function sumDuplicatesInSheet() {
  //const ss = SpreadsheetApp.getActiveSpreadsheet();
    var f = PropertiesService.getScriptProperties().getProperty('spreadsht');
  var ss = SpreadsheetApp.openByUrl(f);
  var sheet = ss.getSheetByName("Sheet1");
  if (!sheet) {
    Logger.log("Sheet " + f + " not found.");
    return;
  }
  //const sheet = ss.getActiveSheet();

  // Get all data
  const data = sheet.getDataRange().getValues();

  // Object to store sums by unique key in column A
  const sums = {};

  for (let i = 0; i < data.length; i++) {
    const key = data[i][0]; // Column A
    const value = parseFloat(data[i][1]) || 0; // Column B, force number

    if (sums[key] === undefined) {
      sums[key] = value;
    } else {
      sums[key] += value;
    }
  }

  // Clear existing sheet
  sheet.clear();

  // Write back unique entries with sums
  const output = Object.entries(sums);
  sheet.getRange(1, 1, output.length, 2).setValues(output);
}



