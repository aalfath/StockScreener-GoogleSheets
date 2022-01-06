![image](https://user-images.githubusercontent.com/7084258/148468038-2954e79c-b713-4197-8849-f62df051f53a.png)

# StockScreener-GoogleSheets
A Google script to fetch the data from Yahoo! Finance. Contains around 7826 tickers from NASDAQ, NYSE and AMEX.

# Installation
1. Make a copy of the initial spreadsheet @ https://docs.google.com/spreadsheets/d/1ra4FRXjUIVxda1Rz0RlualMgJGGoKKqHqtdncWigpwM/edit?usp=sharing
2. Create a Google Script for that sheet, and paste the content of the stockScreener.gs
3. If Google asks for a permission to run the script, you might have to create new GCP project and include it in that script.
4. Due to the limitation of the maximum execution time of Google Script, it is advised to set time-driven trigger for the script for every 10 minutes.
5. It takes around 2-3 hours for the script to fetch the data for all tickers.
