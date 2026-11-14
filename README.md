# Ledger Loan Manager

A lightweight loan and repayment tracker that runs in your browser without a server or account.

## Run

Run a local web server from this folder, then open `http://localhost:8000` in a modern browser:

```sh
python3 -m http.server 8000
```

Use the browser's **Install app** option (or the install prompt when available) to add Ledger to your home screen. After the first online visit, the app shell is available offline. Your data and Excel file link are retained in the browser profile for this exact origin.

To keep an Excel backup updated automatically, choose **Connect Excel** once and select a dedicated `.xlsx` file. New loans, edits, deletions, and repayments update its **Loans** and **Repayments** sheets automatically. If the browser needs to renew file permission after a refresh, tap **Reconnect Excel**; the saved workbook handle is reused. When setting up the PWA on a new origin, choosing a Ledger backup restores its loans and repayments automatically. Without a linked workbook, records still save locally in the browser.

## Features

- Track personal, home, vehicle, education, and business loans.
- Calculate monthly EMIs and reducing-balance repayment schedules.
- Record full EMI payments, with a payment date and optional reference.
- Review outstanding principal, collections, upcoming installments, and overdue loans.
- Search and filter loans. No loan or repayment details are sent to an external service.

EMIs are calculated from the principal, annual rate, and term using the standard monthly reducing-balance formula. The first scheduled installment is one month after the disbursement date. This is a tracking aid, not a lender statement; actual lender schedules, fees, payment timing, and rounding may differ.
