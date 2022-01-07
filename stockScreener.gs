function populateStockScreenerData() {
  // Check the cache if the same script is being executed, if so, end this script
  if(checkRunning() == true) {
    Logger.log("Another process is still running!")
    return;
  }

  // Mark the process status as running, to avoid the same script for being executed simultaneously by the time-driven trigger
  setIsRunning(true);

  // Stock screener sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock Screener');

  var startingColumn = 3;
  var startingRow = 5;

  var currentDate = Date();
  var currentTime = currentDate.toString();

  // Fetch last update log and compare it
  var date = new Date();
  var lastUpdateLog = sheet.getRange(3 ,25).getValue();
  var currentDateLog = Math.floor((date.getTime()/1000)).toString();

  Logger.log("Current date: " + currentDateLog);
  Logger.log("Last update: " + lastUpdateLog);

  // If the sheet is updated in the last 24 hours, then skip updating the sheet.
  if (lastUpdateLog >= currentDateLog-86400) {
    Logger.log("Skipping update. The sheet is updated already!");
    return;
  }

  // Continue from the last updated row
  if (sheet.getRange(3, 24).getValue() != 0) {
    startingRow = sheet.getRange(3, 24).getValue();
    Logger.log("Continue from the last updated row @ " + sheet.getRange(3, 24).getValue());
  }

  for (var i = startingRow; i <= 9999; i++) {
    symbol = sheet.getRange(i, startingColumn).getValue();
    Logger.log("Detected symbol: " + symbol);

    if (symbol.length > 0) {
      // Fetch stock info
      Logger.log("Fetching Y! Finance data for " + symbol);
      var url = "https://query2.finance.yahoo.com/v10/finance/quoteSummary/" + symbol + "?modules=assetProfile%2CsummaryDetail%2Cprice%2CdefaultKeyStatistics%2CrecommendationTrend";
      var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
      var json = response.getContentText();
      var data = JSON.parse(json).quoteSummary.result;

      // Set the last updated row
      sheet.getRange(3, 24).setValue(i);

      if (data == undefined) {
        continue;
      }

      // Set the stokc data
      sheet.getRange(i, 5).setValue(data[0].price.longName);
      if ("assetProfile" in data[0]) {
        if ("sector" in data[0].assetProfile) {
          sheet.getRange(i, 6).setValue(data[0].assetProfile.sector);
        }
      }
      sheet.getRange(i, 7).setValue((data[0].price.marketCap.raw/1000000000).toFixed(2));
      sheet.getRange(i, 9).setValue(data[0].price.regularMarketPrice.raw);
      sheet.getRange(i, 10).setValue(data[0].price.regularMarketChangePercent.raw);
      if ("summaryDetail" in data[0]) {
        if ("fiftyTwoWeekLow" in data[0].summaryDetail) {
          sheet.getRange(i, 11).setValue(data[0].summaryDetail.fiftyTwoWeekLow.raw);
        }
        if ("fiftyTwoWeekHigh" in data[0].summaryDetail) {
          sheet.getRange(i, 12).setValue(data[0].summaryDetail.fiftyTwoWeekHigh.raw);
        }      
      }
      if ("defaultKeyStatistics" in data[0]) {
        if ("trailingEps" in data[0].defaultKeyStatistics) {
          sheet.getRange(i, 15).setValue(data[0].defaultKeyStatistics.trailingEps.raw);
        }  
      }
      if ("summaryDetail" in data[0]) {
        if ("trailingPE" in data[0].summaryDetail) {
          sheet.getRange(i, 16).setValue(data[0].summaryDetail.trailingPE.raw);
        }
        if ("trailingAnnualDividendRate" in data[0].summaryDetail) {
          sheet.getRange(i, 17).setValue(data[0].summaryDetail.trailingAnnualDividendRate.raw);
        }
        if ("trailingAnnualDividendYield" in data[0].summaryDetail) {
          sheet.getRange(i, 18).setValue(data[0].summaryDetail.trailingAnnualDividendYield.raw);
        }
      }
      if ("recommendationTrend" in data[0]) {
        if ("trend" in data[0].recommendationTrend) {
          if (0 in data[0].recommendationTrend.trend) {
            sheet.getRange(i, 19).setValue(data[0].recommendationTrend.trend[0].buy)
            sheet.getRange(i, 20).setValue(data[0].recommendationTrend.trend[0].hold)
            sheet.getRange(i, 21).setValue(data[0].recommendationTrend.trend[0].sell)
          }
        }
      }
      
      Utilities.sleep(100);
    } else {
      sheet.getRange(3, 23).setValue(currentTime);
      sheet.getRange(3, 24).setValue(0);
      sheet.getRange(3, 25).setValue(currentDateLog)
      return;
    }
  }

  // End the script
  sheet.getRange(3, 23).setValue(currentTime);
  sheet.getRange(3, 24).setValue(0);
  sheet.getRange(3, 25).setValue(currentDateLog)

  // Clear the script's running status
  CacheService.getScriptCache().remove("isRunning");

  return;
}

function setIsRunning(value){
  CacheService.getScriptCache().put("isRunning", value.toString(), 360);
}

function checkRunning() {
  var currentState = CacheService.getScriptCache().get("isRunning");
  return (currentState == 'true');
}
