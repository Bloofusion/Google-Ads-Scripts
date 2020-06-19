/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
* ETA Migration Checker v1.0
* Written by Martin Roettgerding.
* © 2017 Martin Roettgerding, Bloofusion Germany GmbH.
* www.bloofusion.de
*/

// How many days of ad stats should be taken into account? Default is 7 for the last 7 days.
// Also possible are strings like: "THIS_MONTH", "LAST_MONTH", etc. (check AWQL documentation for more).
// Also possible are concrete timeframes, like "20161201, 20161231".
var statsTime = 7;

// In case you only want to check some of the accounts in an MCC, label them and specify the label's name here:
var accountLabelName = "";

// Use "de" for German. Everything else defaults to English.
var language = "de";

function main() {
  if(typeof statsTime == "number") {
    var date = new Date();
    date.setDate(date.getDate() - 1);
    var endDateString = Utilities.formatDate(date, AdWordsApp.currentAccount().getTimeZone(), "yyyyMMdd");
    date.setDate(date.getDate() - statsTime + 1);
    var startDateString = Utilities.formatDate(date, AdWordsApp.currentAccount().getTimeZone(), "yyyyMMdd");
    var reportDateString = startDateString + ", " + endDateString;
  }else var reportDateString = statsTime;
  
  var params = JSON.stringify({ "reportDateString" : reportDateString});

  // Check if this is running on MCC level or account level.
  if(typeof MccApp !== 'undefined'){
    if(accountLabelName) var accountSelector = MccApp.accounts().withCondition("LabelNames CONTAINS '" + accountLabelName + "'").withLimit(50);
    else var accountSelector = MccApp.accounts().withLimit(50);
    accountSelector.executeInParallel("checkEtaStatus", "callback", params);
  }else{
    var result = checkEtaStatus(params);
    var passOn = [{"status" : "OK", "result" : JSON.parse(result)}];
    saveToSpreadsheet(passOn);
  }
}

/*
* Collects all the data for an account.
*/
function checkEtaStatus(params){
  params = JSON.parse(params);
  
  var searchCampaignIds = getCampaignIds("SEARCH");
  
  var awql = "SELECT AdGroupId, CampaignName, AdGroupName, AdType, Impressions, Clicks, Conversions, ConversionValue FROM AD_PERFORMANCE_REPORT WHERE CampaignId IN " + JSON.stringify(searchCampaignIds) + " AND CampaignStatus = 'ENABLED' AND AdGroupStatus = 'ENABLED' AND Status = 'ENABLED' AND AdType IN ['TEXT_AD', 'EXPANDED_TEXT_AD'] AND AdNetworkType2 = 'SEARCH'";
  awql += " DURING " + params['reportDateString'];
  
  var reportRows = AdWordsApp.report(awql).rows();
  var results = {};
  var highest = {};
  var a = 0;
  while(reportRows.hasNext()){
    var row = reportRows.next();
    if(!results.hasOwnProperty(row['AdGroupId'])){
      results[row['AdGroupId']] = [ row['CampaignName'], row['AdGroupName'], 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
      highest[row['AdGroupId']] = [0, 0, 0, 0, 0, 0, 0, 0];
    }
    
    if(row['AdType'] == "Expanded text ad"){      
      a = 0;
    }else{
      a = 1;
    }
    
    results[row['AdGroupId']][2 + a] ++;
    results[row['AdGroupId']][4 + a] += parseInt(row['Impressions']);
    results[row['AdGroupId']][6 + a] += parseInt(row['Clicks']);
    results[row['AdGroupId']][8 + a] += parseInt(row['Conversions']);
    results[row['AdGroupId']][10 + a] += parseInt(row['ConversionValue']);
    
    // Store the ads CTR/CR/CpI/VpI if it's the highest in the group.
    if(row['Impressions'] > 0){        
      if(row['Clicks'] / row['Impressions'] > highest[row['AdGroupId']][0 + a]) highest[row['AdGroupId']][0 + a] = row['Clicks'] / row['Impressions'];
      if(row['Conversions'] / row['Impressions'] > highest[row['AdGroupId']][4 + a]) highest[row['AdGroupId']][4 + a] = row['Conversions'] / row['Impressions'];
      if(row['ConversionValue'] / row['Impressions'] > highest[row['AdGroupId']][6 + a]) highest[row['AdGroupId']][6 + a] = row['ConversionValue'] / row['Impressions'];
    }
    if(row['Impressions'] > 0){
      if(row['Conversions'] / row['Clicks'] > highest[row['AdGroupId']][2 + a]) highest[row['AdGroupId']][2 + a] = row['Conversions'] / row['Clicks'];
    }
  }
  
  var rows = [];

  var sums = {
    "Account Name" : AdWordsApp.currentAccount().getName(),
    "AdWords Client ID" : AdWordsApp.currentAccount().getCustomerId(),
    "ETAs" : 0,
    "STAs" : 0,
    "ETA Impressions" : 0,
    "STA Impressions" : 0,
    "ETA Clicks" : 0,
    "STA Clicks" : 0,
    "ETA Conversions" : 0,
    "STA Conversions" : 0,
    "ETA ConversionValue" : 0,
    "STA ConversionValue" : 0,
    "Ad Groups" : 0,
    "Ad Groups with ETAs" : 0
  };
  
  for(var adGroupId in results){
    row = results[adGroupId];
    
    for(var i = 0; i < 8; i += 2){ if(highest[adGroupId][i] > highest[adGroupId][i + 1]) row.push("ETA");
      else if(highest[adGroupId][i] < highest[adGroupId][i + 1]) row.push("STA");
      else row.push("--");
    }
    
    rows.push(row);
    
    var i = 0;
    for(var index in sums){
      
      if(i < 4){ i++; continue; } sums[index] += parseInt(row[i]); i++; if(i == 12) break; } sums['Ad Groups']++; if(row[2] > 0) sums["Ad Groups with ETAs"]++;
  }
  
  var sheetName = AdWordsApp.currentAccount().getName() + " (" + AdWordsApp.currentAccount().getCustomerId() + ")";

  return JSON.stringify({"sums" : sums, "table" : rows, "sheetName" : sheetName});
}

/*
* Translates the results into an object to be used by the writeBack function.
*/
function callback(results){
  var passOn = [];
  
  for (var i = 0; i < results.length; i++){
    passOn[i] = {
      "status" : results[i].getStatus(),
      "result" : JSON.parse(results[i].getReturnValue())
    };
  }
  // Sort the results sheet name (which comes down to client name). This way the sheets will be in order later.
  passOn.sort(sortBySheetName);
  
  saveToSpreadsheet(passOn);
}

/*
* Creates a spreadsheet and saves the results.
*/
function saveToSpreadsheet(results){
  // All templates are stored in a master spreadsheet.
  var masterSpreadsheetId = "1xM-OOPzaGzfuo6eZvFfMY5yp0yfAg-KxBrek2cEpkN8";
  var masterSpreadsheet = SpreadsheetApp.openById(masterSpreadsheetId);
  
  var spreadsheet = SpreadsheetApp.create("ETA Migration Status");  
  var originalSheet = spreadsheet.getActiveSheet();
  
  if(language == "de"){
    var masterSheet = masterSpreadsheet.getSheetByName("Account DE");
    var masterSummary = masterSpreadsheet.getSheetByName("Summary DE");  
    var summarySheet = masterSummary.copyTo(spreadsheet).setName("Zusammenfassung");  
  }else{
    var masterSheet = masterSpreadsheet.getSheetByName("Account EN"); 
    var masterSummary = masterSpreadsheet.getSheetByName("Summary EN");  
    var summarySheet = masterSummary.copyTo(spreadsheet).setName("Summary");  
  }
  
  spreadsheet.deleteSheet(originalSheet);
  
  // Log the spreadsheet's URL so that it can be found by the user.
  Logger.log("A spreadsheet with the results is here:");
  Logger.log(spreadsheet.getUrl());
  
  var rowCounter = 0;
  var row;
  for (var i = 0; i < results.length; i++) { if(results[i]['status'] == "OK"){ var res = results[i]['result']; var sums = res['sums']; row = [sums['Account Name'], sums['AdWords Client ID'], sums['Ad Groups']]; // Share of groups with ETAs if(sums['Ad Groups'] > 0) row.push(sums['Ad Groups with ETAs'] / (sums['Ad Groups'])); else row.push(0);
      
      // Calculate ETA shares.
      if(sums['STA Impressions'] + sums['ETA Impressions'] > 0) row.push(sums['ETA Impressions'] / (sums['STA Impressions'] + sums['ETA Impressions'])); else row.push(0);
      if(sums['STA Clicks'] + sums['ETA Clicks'] > 0) row.push(sums['ETA Clicks'] / (sums['STA Clicks'] + sums['ETA Clicks'])); else row.push(0);
      if(sums['STA Conversions'] + sums['ETA Conversions'] > 0) row.push(sums['ETA Conversions'] / (sums['STA Conversions'] + sums['ETA Conversions'])); else row.push(0);
      if(sums['STA ConversionValue'] + sums['ETA ConversionValue'] > 0) row.push(sums['ETA ConversionValue'] / (sums['STA ConversionValue'] + sums['ETA ConversionValue'])); else row.push(0);
      
      // Write the summary data right away.
      summarySheet.getRange(2 + rowCounter, 1, 1, row.length).setValues([ row ]);
      rowCounter++;
      
      var table = res['table'];      
      
      var sheetName = res['sheetName'];
      var sheet = masterSheet.copyTo(spreadsheet).setName(sheetName);
      
      if(table.length > 0) sheet.getRange(20, 1, table.length, 16).setValues(table).sort([1, 2]);
    }
  }

  summarySheet.getRange(2, 1, rowCounter, row.length).sort(1);
  
  if(language == "de"){
    var masterAbout = masterSpreadsheet.getSheetByName("About DE");
    masterAbout.copyTo(spreadsheet).setName("Über");
  }else{
    var masterAbout = masterSpreadsheet.getSheetByName("About EN");
    masterAbout.copyTo(spreadsheet).setName("About");
  }
}

function getCampaignIds(campaignType){
  var awql = "SELECT CampaignId FROM CAMPAIGN_PERFORMANCE_REPORT WHERE AdvertisingChannelType = '" + campaignType + "'";
  var reportRows = AdWordsApp.report(awql).rows();
  var campaignIds = [];
  
  while(reportRows.hasNext()){
    var row = reportRows.next();
    campaignIds.push(row['CampaignId']);
  }
  return campaignIds;
}

function sortBySheetName(a, b){
  if (a['result']['sheetName'].toLowerCase() > b['result']['sheetName'].toLowerCase()) return 1;
  if (a['result']['sheetName'].toLowerCase() < b['result']['sheetName'].toLowerCase()) return -1;
  return 0; 
}
