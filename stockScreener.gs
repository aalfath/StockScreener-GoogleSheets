function populateStockScreenerData() {
  // Skip the script if its weekend
  var today = new Date();
  if(today.getDay() == 6 || today.getDay() == 0) {
    Logger.log("It's weekend, no need to fetch the update from the market!")
    return;
  }
  
  // Check the cache if the same script is being executed, if so, end this script
  if(checkRunning() == true) {
    Logger.log("Another process is still running!")
    return;
  }

  // Build indexed column to perform an update
  columnVals =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock Screener')
    .getRange('D:D')
    .getValues();

  var indexedRow = [];
  var rowCounter = 1;
  columnVals.forEach(function(ticker) {
    indexedRow[ticker] = rowCounter;
    rowCounter++;
  });

  // Mark the process status as running, to avoid the same script for being executed simultaneously by the time-driven trigger
  setIsRunning(true);

  // Stock screener sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock Screener');

  // Set the starting column and row
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
  if (lastUpdateLog >= currentDateLog-43200) {
    Logger.log("Skipping update. The sheet is updated already!");
    return;
  }

  // Continue from the last updated row
  if (sheet.getRange(3, 24).getValue() != 0) {
    startingRow = sheet.getRange(3, 24).getValue();
    Logger.log("Continue from the last updated row @ " + sheet.getRange(3, 24).getValue());
  }

  var requests = []
  for (var i = startingRow; i <= 9999; i++) {
    symbol = sheet.getRange(i, startingColumn).getValue();
    Logger.log("Detected symbol: " + symbol);

    if (symbol.length > 0) {
      // Fetch stock info
      Logger.log("Fetching Y! Finance data for " + symbol);
      var url = ("https://query2.finance.yahoo.com/v10/finance/quoteSummary/" + symbol + "?modules=assetProfile%2CsummaryDetail%2Cprice%2CdefaultKeyStatistics%2CrecommendationTrend");
      var request1 = {
        'url': url,
        'method' : 'get',
        'muteHttpExceptions': true
      };
      
      // Push the request to the requests array. Calls will be made asynchronously.
      requests.push(url);

      // Process the batch request
      if (i % 25 == 0 && requests.length > 0) {
        // Call the requests and process the response
        var response = UrlFetchApp.fetchAll(requests);
        
        response.forEach(function(response) {
          var json = response.getContentText();
          var data = JSON.parse(json).quoteSummary.result;

          Logger.log(data);

          if (data == undefined) {
            return;
          }

          // Set the stock data
          Logger.log("Updating row for the ticker: " + data[0].price.symbol);

          sheet.getRange(indexedRow[data[0].price.symbol], 5).setValue(data[0].price.longName);
          if ("assetProfile" in data[0]) {
            if ("sector" in data[0].assetProfile) {
              sheet.getRange(indexedRow[data[0].price.symbol], 6).setValue(data[0].assetProfile.sector);
            }
          }

          var marketCap = 0;
          if ("price" in data[0]) {
            if ("marketCap" in data[0].price) {
              marketCap = (data[0].price.marketCap.raw/1000000000).toFixed(2);

              // Market cap data for some tickers aren't available
              if (isNaN(marketCap)) {
                marketCap = 0;
              }  
            }

            sheet.getRange(indexedRow[data[0].price.symbol], 7).setValue(marketCap);
            sheet.getRange(indexedRow[data[0].price.symbol], 9).setValue(data[0].price.regularMarketPrice.raw);
            sheet.getRange(indexedRow[data[0].price.symbol], 10).setValue(data[0].price.regularMarketChangePercent.raw);
          }
          
          if ("summaryDetail" in data[0]) {
            if ("fiftyTwoWeekLow" in data[0].summaryDetail) {
              sheet.getRange(indexedRow[data[0].price.symbol], 11).setValue(data[0].summaryDetail.fiftyTwoWeekLow.raw);
            }
            if ("fiftyTwoWeekHigh" in data[0].summaryDetail) {
              sheet.getRange(indexedRow[data[0].price.symbol], 12).setValue(data[0].summaryDetail.fiftyTwoWeekHigh.raw);
            }      
          }
          if ("defaultKeyStatistics" in data[0]) {
            if ("trailingEps" in data[0].defaultKeyStatistics) {
              sheet.getRange(indexedRow[data[0].price.symbol], 15).setValue(data[0].defaultKeyStatistics.trailingEps.raw);
            }  
          }
          if ("summaryDetail" in data[0]) {
            if ("trailingPE" in data[0].summaryDetail) {
              sheet.getRange(indexedRow[data[0].price.symbol], 16).setValue(data[0].summaryDetail.trailingPE.raw);
            }
            if ("trailingAnnualDividendRate" in data[0].summaryDetail) {
              sheet.getRange(indexedRow[data[0].price.symbol], 17).setValue(data[0].summaryDetail.trailingAnnualDividendRate.raw);
            }
            if ("trailingAnnualDividendYield" in data[0].summaryDetail) {
              sheet.getRange(indexedRow[data[0].price.symbol], 18).setValue(data[0].summaryDetail.trailingAnnualDividendYield.raw);
            }
          }
          if ("recommendationTrend" in data[0]) {
            if ("trend" in data[0].recommendationTrend) {
              if (0 in data[0].recommendationTrend.trend) {
                sheet.getRange(indexedRow[data[0].price.symbol], 19).setValue(data[0].recommendationTrend.trend[0].buy)
                sheet.getRange(indexedRow[data[0].price.symbol], 20).setValue(data[0].recommendationTrend.trend[0].hold)
                sheet.getRange(indexedRow[data[0].price.symbol], 21).setValue(data[0].recommendationTrend.trend[0].sell)
              }
            }
          }
        });
        
        // Set the last updated row
        sheet.getRange(3, 24).setValue(i);

        // Reset the batch request
        requests = [];
        Utilities.sleep(100);
      }
    } else {
      sheet.getRange(3, 23).setValue(currentTime);
      sheet.getRange(3, 24).setValue(0);
      sheet.getRange(3, 25).setValue(currentDateLog)
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
