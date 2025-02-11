# Live Cryptocurrency Data Tracker

This project fetches live cryptocurrency data for the top 50 cryptocurrencies, analyzes it, and presents the data in a continuously updating Excel sheet.  A PDF report summarizing the analysis is also generated.

## Objective

The goal of this project is to provide a real-time view of the cryptocurrency market, focusing on the top 50 cryptocurrencies by market capitalization.
## Running the Script
Clone the Repository
## Install Dependencies
```bash
pip install requests pandas openpyxl fpdf msal
```

## Run the Script
```bash
python crypto_tracker.py
```

The script will run continuously, fetching data every 5 minutes and updating the Excel sheet and PDF report.

## Notes
1. The Excel sheet (crypto_data.xlsx) will be updated locally. For live updates on OneDrive, the script will upload the local file to your OneDrive folder every 5 minutes.  
2. The PDF report (crypto_report.pdf) will be regenerated with each data fetch.  
3. The script requires an active internet connection to fetch data from the CoinGecko API.  
4. Configure OneDrive/Google Sheets (For Live Updates):  
**OneDrive:**  
To enable live updates directly to OneDrive, you'll need to configure an Azure AD application and grant it the necessary permissions.
Follow these steps:
1. Register an application in the Azure portal (portal.azure.com).  
2. Grant the application "Files.ReadWrite.All" (or more restrictive) API permissions for Microsoft Graph.  
3. Create a client secret for your application.  
4. Replace the placeholder values in the crypto_tracker.py script for CLIENT_ID, CLIENT_SECRET, SCOPES, and AUTHORITY with your Azure AD app registration details. Also update ONE_DRIVE_FOLDER with your desired folder name.  
**Google Sheets:**  
Live updates to Google Sheets require a different approach using the Google Sheets API.  You'll need to set up a Google Cloud project, enable the Sheets API, and obtain credentials.  The crypto_tracker.py script would need to be modified to use the Google Sheets API instead of directly writing to an Excel file.  This project currently uses OneDrive..  

## Assessment Steps

1. **Fetch Live Data:** The `crypto_tracker.py` script fetches data from the CoinGecko API, retrieving the top 50 cryptocurrencies by market cap. The data includes:
    * Cryptocurrency Name
    * Symbol
    * Current Price (USD)
    * Market Capitalization
    * 24-hour Trading Volume
    * 24-hour Price Change (percentage)

2. **Data Analysis:** The script performs the following analysis:
    * Identifies the top 5 cryptocurrencies by market cap.
    * Calculates the average price of the top 50 cryptocurrencies.
    * Determines the highest and lowest 24-hour percentage price changes.

3. **Live-Running Excel Sheet:** The `crypto_tracker.py` script updates the `crypto_data.xlsx` file with the latest data every 5 minutes.  The Excel sheet displays the live prices and other key metrics.

4. **Analysis Report:** The `crypto_report.pdf` file contains a summary of the key insights and analysis derived from the fetched data.

## Thank You  
Thank you for reviewing this project! I appreciate your time and consideration.  I hope this project demonstrates my ability to fetch, analyze, and present live cryptocurrency data effectively.  If you have any questions or feedback, please don't hesitate to reach ou

