# QBOSalesReceiptExtract
by EccentricWorkshop (ecc.ws or github.com/EccentricWkshp)
Initial on 06/05/2024

Saves QuickBooks Online sales receipt information to xlsx.

## Dependencies
- json
- argparse
- requests
- pandas
- openpyxl
- pycountry

## Configuring
You will need your QBO API credentials which involves setting up an app through your QBO developer account.

Should be pretty self explanatory, but create a config.json file with:
{
    "client_id": "your_client_id",
    "client_secret": "your_client_secret",
    "refresh_token": "your_refresh_token",
    "realm_id": "your_realm_id",
    "redirect_uri": "http://localhost",
    "sandbox": false,
	"default_days": 30,
	"address_debug": false,
	"receipt_debug": false
}


## Current Features
- Connects to QBO API.
- Requires QBO developer account, creation of an app, and generation of OAuth credentials.
- Gets sales receipts for specified number of days.
- Extracts date, name, state or country, total amount, shipping amount, and SKUs along with quantity if greater than 1.
- Saves everything to sales_receipts.xlsx in the current directory.
- Has debug feature for BillAddr, ShipAddr, and sales_receipts.

QBO API Docs: https://developer.intuit.com/app/developer/qbo/docs/api/accounting/all-entities/salesreceipt
