![image](https://user-images.githubusercontent.com/7084258/148468038-2954e79c-b713-4197-8849-f62df051f53a.png)

# StockScreener-GoogleSheets
A Google script to fetch the data from Yahoo! Finance. Contains around 7826 tickers from NASDAQ, NYSE and AMEX.

# Installation
1. Make a copy of the initial spreadsheet @ https://docs.google.com/spreadsheets/d/1ra4FRXjUIVxda1Rz0RlualMgJGGoKKqHqtdncWigpwM/edit?usp=sharing. This is required because the public sheet that I have provided doesn't have write access which you will need.
2. Create a new Google Script for that sheet, and paste the content of the stockScreener.gs
3. If Google asks for a permission to run the script, you might have to create a new GCP project and include it in that script.
4. Due to the limitation of the maximum execution time of Google Script, it is advised to set a time-driven trigger for the script for every 1 minute (don't worry, the script has caching to prevent it from being executed simultaneously).
5. It takes around 2-3 hours for the script to fetch the data for all tickers.

# Disclaimer
Sometimes the API doesn't return proper data, so be sure to double check with the other sources!

Also, obligatory I am not a financial advisor. Do your own research.

Consult a professional investment advisor before making any investment decisions! This sheet is shared for the purpose of knowledge sharing only!

---

Copyright 2021 https://github.com/aalfath/

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
