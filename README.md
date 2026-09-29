# FinancialHelper

FinanceControl frontend and Vercel API backed by the private Google Sheet database.

## Required Vercel environment variables

- GOOGLE_SHEET_ID
- GOOGLE_SERVICE_ACCOUNT_JSON

The Google service-account email must have Editor access to the FinanceControl Database spreadsheet.

## Health check

After deployment, open:

/api/health

Expected result:

{"ok":true,"database":"FinanceControl Database",...}

The frontend reads and writes finance data through /api/state. Do not expose the Google service-account JSON in client-side code.
