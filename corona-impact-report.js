/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Corona Impact Report v1.0
 * 
 * This script charts search behavior regarding devices, weekdays, and hour of day for several periods.
 * By default, it compares the weeks before and at the beginning of the corona crisis.
 * 
 * © 2020 Martin Roettgerding, Bloofusion Germany GmbH.
 * https://github.com/Bloofusion/Google-Ads-Scripts
 */

/*
 * Change this configuration to customize the script.
 */
var config = {
    /*
     * Include all periods you want to compare. The name will show up in your charts.
     * You can include as many periods as you want (or just one). There has to be a comma between entries, but not after the last one.
     * Note: The periods don't need to have the same length.
     * Note: Make sure periods are a multiple of seven days long. Otherwise the distribution over weekdays won't make sense (e.g. if there are more Mondays than Tuesdays, there will be more clicks on Mondays).
     */
    periods: {
        "February": {start: "2020-02-01", end: "2020-02-28"},
        "Corona crisis": {start: "2020-03-19", end: "2020-04-15"}
    },
    /*
     * Lets you focus the analysis on one or several ad networks. Leave empty to analyze all networks.
     * Available options are: SEARCH, SEARCH_PARTNERS, CONTENT, YOUTUBE_SEARCH, YOUTUBE_WATCH, and MIXED.
     * (CONTENT refers to the display netzwork, MIXED refers to cross-network)
     * Use commas to separate multiple values. Example: "YOUTUBE_SEARCH, YOUTUBE_WATCH"
     */
    adNetworkTypes: "",
    /* 
     * These are the dimensions that will be displayed in charts. There's usually no reason to change this.
     * Available options are: Device, DayOfWeek, and HourOfDay.
     */
    rowDimensions: ['Device', 'DayOfWeek', 'HourOfDay'],
    /* 
     * These are the metrics that will be displayed in charts. 
     * Available options are: Impressions, Clicks, Cost, Conversions, and ConversionValue.
     */
    metricNames: ['Clicks', 'Conversions'],
    /*
     * If you only want to analyze some accounts in your MCC, you can label them.
     * Put the label's name into this setting to have the script only look at accounts with the label.
     * This does not affect account level runs of the script.
     */
    accountLabelName: "",
    /*
     * How many accounts the script should look at. The default is 50, which is the number of accounts that Google Ads scripts can process at the same time.
     * This script can go beyond this limit by processing the rest afterwards, one by one. Depending on the number, this can result in the script going over the time limit.
     * This does not affect account level runs of the script.
     */
    maxAccounts: 50,
    /*
     * What to name the spreadsheet.
     */
    spreadsheetName: "Corona Impact Report",
    /*
     * A list of emails to be added as viewers to the spreadsheet.
     * Example: "alice@example.com, bob@example.com"
     */
    viewerEmails: "",
    /*
     * A list of emails to be added as editors to the spreadsheet.
     * Example: "alice@example.com, bob@example.com"
     * There's no need to add someone as both viewer and editor. If you do this, they'll become editors.
     */
    editorEmails: "",
    /*
     * Use this to not include account names in the spreadsheet.
     * If set to true, names will be replaced by Client 1, Client 2, and so on.
     */
    anonymizeAccountNames: false
}

function main() {
    var params = {};
    // Check if this script is running on MCC level.
    if (typeof AdsManagerApp !== 'undefined') {
        // The script is currently running in an MCC.

        // The selection of accounts and how to go beyond the 50 account limit is based on https://outshine.com/blog/run-your-adwords-scripts-across-a-lot-of-accounts
        var accountSelector = AdsManagerApp.accounts();
        if (config['accountLabelName']) {
            accountSelector = accountSelector.withCondition("LabelNames CONTAINS '" + config['accountLabelName'] + "'");
        }
        if (config['maxAccounts']) {
            accountSelector = accountSelector.withLimit(config['maxAccounts']);
        }

        // Get accounts for parallel run sorted by cost. That way, bigger accounts can be handled in parallel and smaller ones can later be handled sequentially.
        var accountIterator = accountSelector.forDateRange("LAST_30_DAYS").orderBy("Cost DESC").get();
        var accountIds = [];
        while (accountIterator.hasNext()) {
            var account = accountIterator.next();
            accountIds.push(account.getCustomerId());
        }
        var parallelIds = accountIds.slice(0, 50);
        var sequentialIds = accountIds.slice(50);

        params['sequentialIds'] = sequentialIds;
        Logger.log("Starting parallel run ...");
        AdsManagerApp.accounts().withIds(parallelIds).executeInParallel("processAccount", "parallelDone", JSON.stringify(params));
    } else {
        // The script is currently running in a single account.
        var accountResults = processAccount(JSON.stringify(params));
        var passOn = [
            {
                "status": "OK",
                "accountResults": JSON.parse(accountResults)
            }
        ];
        saveToSpreadsheet(passOn);
    }
}

/*
 * Processes a single account.
 * @param {String} params The params passed to all accounts as JSON.
 * @returns {String} Results are returned as JSON.
 */
function processAccount(params) {
    // Get the account's name. If the account doesn't have a name, use it's customer ID instead.
    var currentAccountName = AdsApp.currentAccount().getName();
    if (currentAccountName == "") {
        currentAccountName = AdsApp.currentAccount().getCustomerId();
    }

    Logger.log("Now processing account " + currentAccountName);

    var data = {};
    var colDimensions = [{values: [], current: ""}];
    for (var periodName in config['periods']) {
        colDimensions[0]['values'].push(periodName);
    }

    for (var periodName in config['periods']) {
        var period = config['periods'][periodName];
        colDimensions[0]['current'] = periodName;
        var timeZone = AdWordsApp.currentAccount().getTimeZone();
        var format = 'yyyyMMdd';
        var start = Utilities.formatDate(new Date(period['start']), timeZone, format);
        var end = Utilities.formatDate(new Date(period['end']), timeZone, format);
        var awql = "SELECT " + config['rowDimensions'].concat(config['metricNames']).toString() + " FROM ACCOUNT_PERFORMANCE_REPORT";
        if (config['adNetworkType']) {
            var adNetworkTypes = config['adNetworkTypes'].split(/[,; ]+/);
            awql += " WHERE AdNetworkType2 IN " + JSON.stringify(adNetworkTypes)
        }
        awql += " DURING " + start + ", " + end;

        var reportRows = AdsApp.report(awql, {includeZeroImpressions: true}).rows();
        while (reportRows.hasNext()) {
            var row = reportRows.next();
            row['ConversionValue'] = parseConversionValueAsFloat(row['ConversionValue']);
            collectData(data, row, config['rowDimensions'], config['metricNames'], colDimensions);
        }
    }
    
    trackInAnalytics("1.0");

    return JSON.stringify({accountName: currentAccountName, data: data, originalParams: JSON.parse(params)});
}

/*
 * Arranges the data from the a report row and adds it to the data collected so far.
 * @param {Object} data The data collected so far.
 * @param {Object} reportRow A row from a Google Ads report.
 * @param {Array} rowDimensions All column names from the reportRow that are to be used as dimensions in the first columns in every row.
 * @param {Array} metricNames All column names from the reportRow that are to be used as metrics.
 * @param {Array} colDimensions Dimensions that become separate columns for each metric (like: Impressions A, Impressions B) [ { values: [A, B], current: A }, { ... } ]
 * @returns {Object} The updated data.
 */
function collectData(data, reportRow, rowDimensions, metricNames, colDimensions) {
    var key = "";
    var dataRow = [];
    var weekdayNumbers = {Monday: 1, Tuesday: 2, Wednesday: 3, Thursday: 4, Friday: 5, Saturday: 6, Sunday: 7};
    var rowDimensionsCount = 0;
    for (var i = 0; i < rowDimensions.length; i++) {
        var rowDimension = rowDimensions[i];
        var rowDimensionValue = reportRow[rowDimension];
        key += rowDimensionValue + "##";
        dataRow.push(rowDimensionValue);
        rowDimensionsCount++;
        // For the day of the week, its number is stored in an additional column. This will be used to sort the table later.
        if (rowDimension == "DayOfWeek") {
            var weekdayNumber = weekdayNumbers[rowDimensionValue];
            dataRow.push(weekdayNumber);
            rowDimensionsCount++;
        }
    }

    valuesOnly = [];
    for (var i = 0; i < metricNames.length; i++) {
        var metricName = metricNames[i];
        for (var j = 0; j < colDimensions.length; j++) {
            var colDimension = colDimensions[j];
            for (var k = 0; k < colDimension['values'].length; k++) {
                var colDimensionValue = colDimension['values'][k];
                if (colDimensionValue == colDimension['current']) {
                    if (metricName == "ConversionValue" || metricName == "Cost") {
                        var metricValue = parseConversionValueAsFloat(reportRow[metricName]);
                    } else {
                        var metricValue = parseInt(reportRow[metricName]);
                    }
                    valuesOnly.push(metricValue);
                } else {
                    valuesOnly.push(0);
                }
            }
        }
    }

    if (!data.hasOwnProperty(key)) {
        data[key] = dataRow.concat(valuesOnly);
    } else {
        for (var i = 0; i < valuesOnly.length; i++) {
            data[key][(rowDimensionsCount + i)] += valuesOnly[i];
        }
    }

    return data;
}

/*
 * Prepares the results from the parallel runs and saves everything to a spreadsheet.
 * @param {Array} results The list of results, one set per account.
 * @returns {undefined} Returns nothing.
 */
function parallelDone(results) {
    Logger.log("Parallel run completed.");
    var originalParams;
    for (var i = 0; i < results.length; i++) {
        if (results[i].getStatus() == "OK") {
            originalParams = JSON.parse(results[i].getReturnValue())['originalParams'];
            break;
        }
    }

    if (!originalParams) {
        Logger.log("No successfully processed accounts.");
        return;
    }

    var passOn = [];
    for (var i = 0; i < results.length; i++) {
        passOn[i] = {
            "status": results[i].getStatus(),
            "accountResults": JSON.parse(results[i].getReturnValue())
        };
    }

    if (originalParams['sequentialIds'].length > 0) {
        Logger.log(originalParams['sequentialIds'].length + " more accounts to process sequentially.");
        var accountIterator = MccApp.accounts().withIds(originalParams['sequentialIds']).get();
        while (accountIterator.hasNext()) {
            var account = accountIterator.next();
            MccApp.select(account);
            result = processAccount(JSON.stringify(originalParams));
            passOn.push({
                "status": "OK",
                "accountResults": JSON.parse(result)
            });
            Logger.log(account.getName() + ": processed.");
        }
    }

    saveToSpreadsheet(passOn);
}

/*
 * Prepares a spreadsheet with the results based on a template and the current configuration.
 * @param {Array} results The results from one or more successful run at the account level.
 * @returns {undefined} Returns nothing.
 */
function saveToSpreadsheet(results) {
    // Create an empty spreadsheet.
    var spreadsheet = SpreadsheetApp.create(config['spreadsheetName']);

    Logger.log("Saving to speadsheet:\r\n" + spreadsheet.getUrl());

    // Add viewers and editors to the spreadsheet.
    if (config['viewerEmails']) {
        var viewerEmails = config['viewerEmails'].split(/[,; ]+/);
        spreadsheet.addViewers(viewerEmails);
    }
    if (config['editorEmails']) {
        var editorEmails = config['editorEmails'].split(/[,; ]+/);
        spreadsheet.addEditors(editorEmails);
    }

    // All templates are stored in a master spreadsheet.
    var templateSpreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1QVVmonxVr3hjVBBFW8R53DpZ4sfnILs3USqEwUFIvL4/edit#gid=0");
    var accountTemplateSheet = templateSpreadsheet.getSheetByName("Account Template");
    // Copy the 'About' sheet and rename it from 'Copy of About'.
    templateSpreadsheet.getSheetByName("About").copyTo(spreadsheet).setName("About");

    // The header row contains the regular headers, starting with the row dimensions.
    var headerRow = [];
    // The metric headers row is above the regular headers. It contains the metric names whereas the regular headers for metrics will just be period names.
    var metricHeaders = [];
    for (i = 0; i < config['rowDimensions'].length; i++) {
        var rowDimension = config['rowDimensions'][i]
        headerRow.push(rowDimension);
        metricHeaders.push("");
        // After DayOfWeek, add an additional column by which the days can be sorted later.
        if (rowDimension == "DayOfWeek") {
            headerRow.push("WeekdayNumber");
            metricHeaders.push("");
        }
    }

    // Prepare the headers for metric columns. Each metric gets two columns for each period: one for the absolute values and one for the percentages.
    var headerRowPercentages = [];
    var metricColumnsCount = 0;
    var metricPercentageHeaders = [];
    for (i = 0; i < config['metricNames'].length; i++) {
        for (var periodName in config['periods']) {
            var metricName = config['metricNames'][i];
            // The following tow cells will be right above/below each other.
            metricHeaders.push(metricName);
            headerRow.push(periodName);
            // The following two cells will be right above/below each other.
            metricPercentageHeaders.push("% " + metricName);
            headerRowPercentages.push(periodName);

            metricColumnsCount++;
        }
    }

    // Prepare the formulas for every row. Since these relative formulas are always the same for each row, this only needs to be done once.
    var formulasRow = [];
    for (i = 0; i < config['metricNames'].length; i++) {
        for (var periodName in config['periods']) {
            formulasRow.push("=R[0]C[-" + metricColumnsCount + "] / SUBTOTAL(109, C[-" + metricColumnsCount + "]:C[-" + metricColumnsCount + "])");
        }
    }
    var metricHeaderRowPosition = 92;
    var headerRowPosition = 93;
    var firstDataRowPosition = headerRowPosition + 1;

    var periodCount = 0;
    for (var periodName in config['periods']) {
        periodCount++;
    }

    // Find out the column numbers for each row dimension and metric.
    var rowDimensionColumnNumbers = {};
    for (d = 0; d < config['rowDimensions'].length; d++) {
        rowDimensionColumnNumbers[config['rowDimensions'][d]] = d + 1;
    }

    // Metric column numbers point to the first percentage column for a metric.
    var metricColumnNumbers = {};
    for (m = 0; m < config['metricNames'].length; m++) {
        metricColumnNumbers[config['metricNames'][m]] = config['rowDimensions'].length + config['metricNames'].length * periodCount + m * periodCount + 1;
    }

    // Determine the order in which the columns will be sorted later: Device, DayOfWeek, and HourOfDay.
    var sortOrder = [];
    if (rowDimensionColumnNumbers['Device']) {
        sortOrder.push(rowDimensionColumnNumbers['Device']);
    }
    if (rowDimensionColumnNumbers['DayOfWeek']) {
        sortOrder.push(rowDimensionColumnNumbers['DayOfWeek']);
    }
    if (rowDimensionColumnNumbers['HourOfDay']) {
        sortOrder.push(rowDimensionColumnNumbers['HourOfDay']);
    }

    // There is an additional column for WeekdayNumber right after the DayOfWeek column to sort by DayOfWeek. The columns to sort by move one to the right from there.
    if (rowDimensionColumnNumbers['DayOfWeek']) {
        var move = false;
        for (var s = 0; s < sortOrder.length; s++) {
            if (sortOrder[s] == rowDimensionColumnNumbers['DayOfWeek']) {
                move = true;
            }
            if (move) {
                sortOrder[s]++;
            }
        }
    }

    // Go through every account's results.
    for (var accountIndex = 0; accountIndex < results.length; accountIndex++) {
        var rows = [];
        var rowCounter = 1;
        var formulas = [];
        if (results[accountIndex]['status'] == "OK") {
            if (config['anonymizeAccountNames']) {
                accountName = "Client " + (accountIndex < 9 ? "0" + (accountIndex + 1) : (accountIndex + 1));
            } else {
                var accountName = results[accountIndex]['accountResults']['accountName'];
            }

            // Prepare data and formulas for the sheet.
            var data = results[accountIndex]['accountResults']['data'];
            for (var key in data) {
                rows.push(data[key]);
                formulas.push(formulasRow);
                rowCounter++;
            }

            // Add a new sheet for the account by copying the account template sheet.
            var sheet = addAccountSheet(accountName, spreadsheet, accountTemplateSheet);
            // Get the charts in the account template.
            var charts = sheet.getCharts();

            // Insert the metric headers row, containing the actual metric names.
            sheet.getRange(metricHeaderRowPosition, 1, 1, metricHeaders.concat(metricPercentageHeaders).length).setValues([metricHeaders.concat(metricPercentageHeaders)]);

            // Insert header row and format as text.
            sheet.getRange(headerRowPosition, 1, 1, headerRow.concat(headerRowPercentages).length).setValues([headerRow.concat(headerRowPercentages)]).setNumberFormat("@");
            // Insert values (dimensions and metrics) and sort by device, weekday, and hour of day.
            sheet.getRange(firstDataRowPosition, 1, rows.length, headerRow.length).setValues(rows).sort(sortOrder);

            // Delete the 'WeekdayNumber' column - it was only needed to sort weekdays.
            if (rowDimensionColumnNumbers['DayOfWeek']) {
                sheet.getRange(metricHeaderRowPosition, rowDimensionColumnNumbers['DayOfWeek'] + 1, rows.length + 2, 1).deleteCells(SpreadsheetApp.Dimension.COLUMNS);
                // Now there's an empty column left that can be deleted.
                sheet.deleteColumn(sheet.getLastColumn() + 1);
            }
            // Insert formulas to calculate the percentages.
            sheet.getRange(firstDataRowPosition, headerRow.length, formulas.length, headerRowPercentages.length).setFormulasR1C1(formulas).setNumberFormat("0.00%");

            // The account sheet contains some charts from the template. These need to be removed or updated with the correct data.
            var rowsToDelete = [];
            for (var c in charts) {
                var chart = charts[c];
                // Get the chart's container info in order to get it's position.
                var containerInfo = chart.getContainerInfo();

                // What metric is the chart about? Find out from its row.
                switch (containerInfo.getAnchorRow()) {
                    case 2:
                        var metricName = "Impressions";
                        break;
                    case 20:
                        var metricName = "Clicks";
                        break;
                    case 38:
                        var metricName = "Cost";
                        break;
                    case 56:
                        var metricName = "Conversions";
                        break;
                    case 74:
                        var metricName = "ConversionValue";
                        break;
                    default:
                        var metricName = false;
                }

                // What dimension is the chart about? Find out from its column.
                switch (containerInfo.getAnchorColumn()) {
                    case 1:
                        var dimensionName = "Device";
                        break;
                    case 3:
                        var dimensionName = "DayOfWeek";
                        break;
                    case 5:
                        var dimensionName = "HourOfDay";
                        break;
                    default:
                        var dimensionName = false;
                }

                // If this metric or dimension is not covered, remove this chart.
                if (!metricName || config['metricNames'].indexOf(metricName) == -1) {
                    sheet.removeChart(chart);
                    // Since this metric is not covered, mark its rows for removal.
                    rowsToDelete.push(containerInfo.getAnchorRow());
                    continue;
                }
                if (!dimensionName || config['rowDimensions'].indexOf(dimensionName) == -1) {
                    sheet.removeChart(chart);
                    continue;
                }

                var chartBuilder = chart.modify();

                // Remove the old data ranges. Their formatting will be relayed to the new ones.
                var chartRanges = chartBuilder.getRanges();
                for (var r = 0; r < chartRanges.length; r++) {
                    chartBuilder.removeRange(chartRanges[r]);
                }

                var metricColIndex = metricColumnNumbers[metricName];
                var rowDimensionColIndex = rowDimensionColumnNumbers[dimensionName];

                // Add the row dimension column. Its values will be shown on the horizontal axis.
                chartBuilder.addRange(sheet.getRange(headerRowPosition, rowDimensionColIndex, rows.length + 1, 1));
                // Add the percentage columns. Each period gets its own bar or line.
                chartBuilder.addRange(sheet.getRange(headerRowPosition, metricColIndex, rows.length + 1, periodCount));

                // Execute the changes and update the sheet.
                chart = chartBuilder.build();
                sheet.updateChart(chart);
            }

            // If some charts from the templates were not needed, there are now some empty rows where they would've been. These rows need to be deleted now.
            // Sort the rows to delete in reverse order, then delete them from the bottom up.
            rowsToDelete.sort().reverse().forEach(function (value, index, self) {
                // if this is the first value or if it is not the same as the last one: delete 18 rows, starting with this one.
                if (!index || value != self[index - 1]) {
                    sheet.deleteRows(value, 18);
                }
            });
        } else {
            // This has never happened ...
            Logger.log("Account not successfully processed. Skipping ...");
        }

    }

    // Delete the empty default sheet from the spreadsheet.
    spreadsheet.deleteActiveSheet();
    // Sort the account sheets by name.
    sortGoogleSheets(spreadsheet);
    // Log the spreadsheet's URL so that it can be found by the user.
    Logger.log("Done.");
}

/* 
* Tracks the execution of the script as an event in Google Analytics.
* Sends the version number and a random UUID (basically just a random number as required by Analytics).
* Basically tells that somewhere someone ran the script with a certain version.
* Credit for the idea goes to Russel Savage, who posted his version at http://www.freeadwordsscripts.com/2013/11/track-adwords-script-runs-with-google.html.
* @param {String} version The current version of the script. 
*/
function trackInAnalytics(version){
  // Create the random UUID from 30 random hex numbers gets them into the format xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx (with y being 8, 9, a, or b).
  var uuid = "";
  for(var i = 0; i < 30; i++){
    uuid += parseInt(Math.random()*16).toString(16);
  }
  uuid = uuid.substr(0, 8) + "-" + uuid.substr(8, 4) + "-4" + uuid.substr(12, 3) + "-" + parseInt(Math.random() * 4 + 8).toString(16) + uuid.substr(15, 3) + "-" + uuid.substr(18, 12);
  
  var url = "http://www.google-analytics.com/collect?v=1&t=event&tid=UA-74705456-1&cid=" + uuid + "&ds=adwordsscript&an=coronaimpactreport&av="
  + version
  + "&ec=AdWords%20Scripts&ea=Script%20Execution&el=Corona%20Impact%20Report%20v" + version;
  UrlFetchApp.fetch(url);
}

/*
 * Copy the account template into a spreadsheet.
 * @param {String} accountName Used as a name for the new sheet. Also used as a headline in the first cell of the sheet.
 * @param {Spreadsheet} spreadsheet The new sheet is added to this spreadsheet.
 * @param {Sheet} accountTemplateSheet This sheet gets copied to the spreadsheet.
 * @returns {Sheet} The newly added sheet.
 */
function addAccountSheet(accountName, spreadsheet, accountTemplateSheet) {
    var accountSheet = accountTemplateSheet.copyTo(spreadsheet).setName(accountName);
    accountSheet.getRange(1, 1).setValue(accountName);
    return accountSheet;
}

/* 
 * Sorts the sheets in the spreadsheet alphabetically and then puts the About sheet at the end.
 * Based on https://gist.github.com/chipoglesby/26fa70a35f0b420ffc23
 * @param {Spreadsheet} spreadsheet The spreadsheet to sort.
 */
function sortGoogleSheets(spreadsheet) {
    var sheetNameArray = [];
    var sheets = spreadsheet.getSheets();
    for (var i = 0; i < sheets.length; i++) {
        sheetNameArray.push(sheets[i].getName().toLowerCase());
    }

    sheetNameArray.sort();
    for (var j = 0; j < sheets.length; j++) {
        spreadsheet.setActiveSheet(spreadsheet.getSheetByName(sheetNameArray[j]));
        spreadsheet.moveActiveSheet(j + 1);
    }

    // Move the about sheet to the last position.
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName("About"));
    spreadsheet.moveActiveSheet(sheets.length);
}

/*
 * Parses a localized number string into a float value. Example: "123,456.78" and "123.456,78" both become 123456.78
 * @param {mixed} value The value to parse.
 * @returns {float} The parsed value.
 */
function parseConversionValueAsFloat(value) {
    string = String(value);
    return parseFloat(string.replace(/(\.|,)([0-9]{3})/g, "$2").replace(/,([0-9]{2})($| )/, ".$1").replace(/[^0-9.]/g, ""));
}
