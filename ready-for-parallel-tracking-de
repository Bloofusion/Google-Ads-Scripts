/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
* Ready for parallel tracking? (Deutsche Version)
* © 2018 Martin Roettgerding, Bloofusion Germany GmbH.
* www.bloofusion.de
*
* Hintergrund
* Am 30. Oktober müssen alle Google-Ads-Kampagnen paralleles Tracking unterstützen. Wer ohne Tracking-URLs auskommt, hat damit kein Problem.
* Wer URLs allerdings per Tracking-Vorlage über eine Domain weiterleitet, muss sicherstellen, dass dabei paralleles Tracking unterstützt wird. Dieses Skript prüft, ob sich solch eine Einstellung möglicherweise irgendwo eingeschlichen hat.
* Geprüft werden aktive Kampagnen, Anzeigengruppen, Anzeigen und Ausrichtungskriterien (z. B. Keywords, Zielgruppen). Nicht geprüft werden Kontoeinstellungen, Sitelinks oder Produktdatenfeeds.
* 
* Die Kontoeinstellung sollte in jedem Fall manuell geprüft werden.
* Das geht so: Alle Kampagnen (keine auswählen!) > Einstellungen (graue Navigation) > Kontoeinstellungen (zweite obere Navigation) > Tracking
* Dort prüfen, ob eine Tracking-Vorlage gesetzt ist. Falls diese leer ist oder mit {lpurl} oder {unescapedlpurl} beginnt, ist alles gut.
*/

// Um nur bestimmte Konten in einem MCC zu prüfen, müssen diese mit einem Label versehen werden. Der Name des Labels ist dann hier einzutragen:
var accountLabelName = "";

function main() {
  Logger.log("Achtung: Per Script können Kontoeinstellung, Sitelinks oder Produktdatenfeeds nicht geprüft werden.");
  Logger.log("Es wird empfohlen, die Kontoeinstellung sowie ggf. Produktdatenfeeds selbst zu prüfen.");
  Logger.log("Weitere Infos: https://blog.bloofusion.de/paralleles-tracking-fuer-google-ads/");
  Logger.log("---------------------------------------------------------------------------------------------------");
    
  // Check if this is running on MCC level or account level.
  if(typeof MccApp !== 'undefined'){
    // Based on https://outshine.com/blog/run-your-adwords-scripts-across-a-lot-of-accounts
    var accountSelector = MccApp.accounts();
    if(accountLabelName) accountSelector = accountSelector.withCondition("LabelNames CONTAINS '" + accountLabelName + "'");
    
    var accountIterator = accountSelector.get();
    var accountIds = [];
    while (accountIterator.hasNext()) {
      var account = accountIterator.next();
      accountIds.push(account.getCustomerId());
    }
    var parallelIds = accountIds.slice(0, 50);
    var sequentialIds = accountIds.slice(50);
    	
    Logger.log(accountIds.length + " Konten zu verarbeiten.");
    var params = { "sequentialIds" : sequentialIds };
    MccApp.accounts().withIds(parallelIds).executeInParallel("checkParallelTrackingStatus", "allDone", JSON.stringify(params));

  }else{
    var params = JSON.stringify({ });
    checkParallelTrackingStatus(params);
  }
}

function checkParallelTrackingStatus(params){
  var currentAccountName = AdWordsApp.currentAccount().getName();
  if(currentAccountName == "") currentAccountName = AdWordsApp.currentAccount().getCustomerId();
  
  var trackingUrlTemplates = {};
  
  trackingUrlTemplates = checkForTrackingUrlTemplates(trackingUrlTemplates, "campaigns");
  trackingUrlTemplates = checkForTrackingUrlTemplates(trackingUrlTemplates, "adgroups");
  trackingUrlTemplates = checkForTrackingUrlTemplates(trackingUrlTemplates, "ads");
  trackingUrlTemplates = checkForTrackingUrlTemplates(trackingUrlTemplates, "criteria");
    
  var foundSomething = false;
  for(var trackingUrlTemplate in trackingUrlTemplates){
    if(!isTrackingUrlTemplateReadyForParallelTracking(trackingUrlTemplate)){
      foundSomething = true;
      var levels = "";
      for(var level in trackingUrlTemplates[trackingUrlTemplate]){
        if(levels == "") levels = level;
        else levels += ", " + level;
      }
        
      Logger.log("-- Möglicherweise problematische Tracking-Vorlage im Konto '" + currentAccountName + "': " + trackingUrlTemplate + " (gefunden auf folgenden Ebenen: " + levels + ").");
    }
  }
  if(!foundSomething){
    Logger.log("+ Kein Problem gefunden im Konto '" + currentAccountName + "'.");
  }

  return JSON.stringify({ "params" :  JSON.parse(params) });
}

function isTrackingUrlTemplateReadyForParallelTracking(trackingUrlTemplate){
  if(trackingUrlTemplate.match(/^\{(lpurl|unescapedlpurl)\}/)) return true;
  return false;
}

function checkForTrackingUrlTemplates(trackingUrlTemplates, level){
  var field = "TrackingUrlTemplate";
  var awql;
  switch(level){
    case "campaigns":
      awql = "SELECT " + field + " FROM CAMPAIGN_PERFORMANCE_REPORT WHERE CampaignStatus = 'ENABLED' AND " + field + " != '{lpurl}' AND " + field + " != ''";
      break;
    case "adgroups":
      awql = "SELECT " + field + " FROM ADGROUP_PERFORMANCE_REPORT WHERE CampaignStatus = 'ENABLED' AND AdGroupStatus = 'ENABLED' AND " + field + " != '{lpurl}' AND " + field + " != ''";
      break;
    case "ads":
      field = "CreativeTrackingUrlTemplate";
      awql = "SELECT " + field + " FROM AD_PERFORMANCE_REPORT WHERE CampaignStatus = 'ENABLED' AND AdGroupStatus = 'ENABLED' AND Status = 'ENABLED' AND " + field + " != '{lpurl}' AND " + field + " != ''";
      break;
    case "criteria":
      awql = "SELECT " + field + ", CriteriaType FROM CRITERIA_PERFORMANCE_REPORT WHERE CampaignStatus = 'ENABLED' AND AdGroupStatus = 'ENABLED' AND Status = 'ENABLED' AND " + field + " != '{lpurl}' AND " + field + " != ''";
      break;
  }
  
  var reportRows = AdWordsApp.report(awql).rows();
  
  // Translate level name into something easier to understand.
  var levelTranslations = {
    'campaigns' : 'Kampagne',
    'adgroups' : 'Anzeigengruppe',
    'ads' : 'Anzeige'
  };  
  if(levelTranslations.hasOwnProperty(level)) var translatedLevel = levelTranslations[level];
  else var translatedLevel = level;

  while(reportRows.hasNext()){
    var row = reportRows.next();
    if(row[field] == '{lpurl}' || row[field] == '') continue;
    
    if(level == "criteria"){
      translatedLevel = row['CriteriaType'];
    }
    
    if(!trackingUrlTemplates.hasOwnProperty(row[field])) trackingUrlTemplates[row[field]] = {};
    if(!trackingUrlTemplates.hasOwnProperty(row[field][translatedLevel])) trackingUrlTemplates[row[field]][translatedLevel] = 0;
    trackingUrlTemplates[row[field]][translatedLevel]++;
  }
  return trackingUrlTemplates;
}
  
function allDone(results){
  var params;
  for (var i = 0; i < results.length; i++) { if(results[i].getStatus() == "OK"){ params = JSON.parse(results[i].getReturnValue())['params']; break; } } if(!params){ Logger.log("Keine erfolgreich durchlaufenen Konten - breche ab ..."); return; } if(params['sequentialIds'].length > 0){
    Logger.log("Noch " + params['sequentialIds'].length + " Konten sequenziell zu verarbeiten.");
    
    var accountIterator = MccApp.accounts().withIds(params['sequentialIds']).get();
    while(accountIterator.hasNext()){
      var account = accountIterator.next();
      MccApp.select(account);
      results = checkParallelTrackingStatus(JSON.stringify(params));
      // Logger.log(account.getName() + ": erfolgreich verarbeitet");
    }    
  }
  Logger.log("Verarbeitung abgeschlossen.");
}
