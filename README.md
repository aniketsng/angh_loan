# Ledger Loan Manager

A lightweight loan and repayment tracker that runs in your browser without a server or account.

## Run

Open `index.html` in a modern Chromium-based browser. Your loans and repayment history are saved in that browser's local storage on this device.

To keep an Excel backup updated automatically, choose **Connect Excel** once and select a dedicated `.xlsx` file. New loans, edits, deletions, and repayments update its **Loans** and **Repayments** sheets automatically. The browser asks for file access; if that permission needs renewing, reconnect the workbook. Without a linked workbook, records still save locally in the browser.

## Features

- Track personal, home, vehicle, education, and business loans.
- Calculate monthly EMIs and reducing-balance repayment schedules.
- Record full EMI payments, with a payment date and optional reference.
- Review outstanding principal, collections, upcoming installments, and overdue loans.
- Search and filter loans. No loan or repayment details are sent to an external service.

EMIs are calculated from the principal, annual rate, and term using the standard monthly reducing-balance formula. The first scheduled installment is one month after the disbursement date. This is a tracking aid, not a lender statement; actual lender schedules, fees, payment timing, and rounding may differ.
