Description
1.	The Portfolio Tracker monitors the real-time market value of the holdings and calculates the unrealized gain/loss. The market prices are obtained from Coingecko exchange. The tool also includes a graphical representation of portfolio allocation. 
2.	The Deal Calculator allows users to input individual buy and sell transactions, automatically calculating the VWAP and realized gains from these deals. 
Deployment environment
•	Excel Version: Compatible with Excel 2016 or newer.
•	System Requirements: Excel must support VBA macros, and Internet access is required to fetch real-time data from CoinGecko.
•	Macro Settings: Ensure that macros are enabled in Excel.
3.3 How to use the application:
1.	Entering Transactions:
o	Open the "My Deals" section of the program. 
 
o	Click button “Enter Deal”
 
o	UserForm appears with fields for:
	Buy/Sell: Choose the transaction type.
	Coin: Select the cryptocurrency from the dropdown.
	Deal Date: Enter the transaction date.
	Deal Time: Enter the transaction time.
	Quantity of Coins: Input how much of the cryptocurrency you bought or sold.
	Price per Coin: Enter the price at which the deal occurred.
	Deal Amount: The system calculates this as Quantity x Price per Coin.
o	If you change the coin, the recent market price will be put 
o	Once all details are entered, click "Submit Data" to record the deal.
2.	Tracking Portfolio (Portfolio):
o	Navigate to the "Portfolio".
 
o	Click the "Update" button to fetch the latest cryptocurrency prices from CoinGecko and calculate portfolio details.
o	The program automatically updates:
	Market Value of Portfolio.
	Unrealized Gain/Loss.
	Portfolio Allocation (displayed in a pie chart).
3.	Error Handling and Warnings:
o	If the API request limit is reached (Rate Limit Exceeded), the program will display a warning and ask you to wait before making further requests.
o	Data validation errors (e.g., entering an invalid quantity or price) will be indicated with a message box.
o	For instance, if you sell more coins that you own, the error message appears, and corresponding rows will be highlighted by red.
 
4.	Assumptions:
o	The program assumes that initial prices and holdings are inputted correctly, and calculations are based on these values.
3.4 Important Notes:
•	Rate Limit for API Requests: The program fetches data from CoinGecko, which has a request limit. To avoid errors, we recommend take breaks between updates or wait at least 1-2 minutes before updating again.
3.5 Troubleshooting:
1.	If the program does not update prices:

o	Ensure you have an active internet connection.
o	Wait for a few moments if the rate limit is exceeded.
3.	If incorrect results are displayed:
o	Check if the initial prices and holdings are correctly entered.
Ensure that the quantity and price values are valid and correctly formatted.
