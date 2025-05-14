# Automated Invoice Tracker & Reminder System (Google Apps Script)

**Version: V.0.7.7 (Automated Dashboard Setup - Stable Core)**

A Google Apps Script to automate the tracking of sent invoices from Gmail, process payment replies, update invoice statuses, send overdue reminders, and visualize key metrics on a dynamic Google Sheets dashboard.

## Overview

Manually tracking invoices, following up on payments, and understanding your financial overview can be time-consuming and prone to errors, especially for freelancers and small businesses. This script aims to solve that by leveraging the power of Google Workspace (Gmail and Google Sheets) to create an automated and "plug and play" invoice management system.

The system automatically:
*   Scans your "Sent" Gmail items for new invoices based on subject keywords.
*   Parses key details from invoice emails (invoice number, client name, amounts, dates).
*   Logs all invoice information into a structured Google Sheet.
*   Creates a dynamic dashboard in Google Sheets to visualize outstanding amounts, overdue invoices, aging analysis, and more.
*   Checks for payment replies in Gmail threads and updates invoice statuses.
*   Automatically sends tiered, customizable overdue reminders for unpaid invoices.
*   Sets up its own time-driven triggers for continuous, automated operation.

## Key Features

*   **One-Step Initial Setup:** Run the `initialSetup()` function once to create and configure a new Google Spreadsheet with all necessary sheets ("Invoice Log", "Error Log", "Dashboard"), formatting, formulas, placeholder charts, and triggers.
*   **Automated Invoice Logging:** Processes new invoices sent from your primary invoicing email.
*   **Smart Email Parsing:** Extracts details like invoice number, client name, amounts, invoice date, and due date from email bodies and subjects (with ongoing refinements).
*   **Dynamic Dashboard:** Provides an at-a-glance overview of:
    *   Key Financial Metrics (Total Overdue, Total Outstanding)
    *   Action & Review Center (Needs Review, Parsing Discrepancies, Script Errors)
    *   Overdue Aging Analysis (by Amount and Count)
    *   Top 5 Clients with Overdue Balances
*   **Payment Reply Processing:** Detects payment-related keywords in client replies and updates invoice status accordingly (e.g., "Paid", "Partially Paid", "Pending Confirmation"). Includes cross-referencing of invoice numbers in replies.
*   **Automated Overdue Reminders:** Sends configurable, multi-stage reminder emails for overdue invoices.
*   **Error Logging:** Captures script errors and parsing issues in a dedicated "Error Log" sheet for review.
*   **Gmail Label Management:** Automatically creates and uses Gmail labels to manage the processing workflow.

## Visuals

*(Consider adding a screenshot of your Dashboard here with dummy data)*
`![Dashboard Screenshot](assets/dashboard_screenshot.png)`

*(Optional: A GIF showing the `initialSetup()` process or the dashboard populating)*
`![Setup Flow GIF](assets/setup_flow.gif)`

## Technologies Used

*   **Google Apps Script** (JavaScript environment for Google Workspace)
*   **Gmail API** (via `GmailApp` service)
*   **Google Sheets API** (via `SpreadsheetApp` service)

## Setup & Installation Guide

1.  **Access Google Apps Script:**
    *   Open Google Sheets (you can create a new blank sheet, or the script will create one if no ID is provided).
    *   Go to **Extensions > Apps Script**.
2.  **Copy the Script:**
    *   Delete any boilerplate `Code.gs` content.
    *   Copy the entire content of the `InvoiceAutomationScript.gs` file from this repository and paste it into the Apps Script editor.
    *   Save the project (File > Save). Give it a name like "Invoice Automation System".
3.  **Configure Critical Constants:**
    *   In the script, locate the "--- CONFIGURATION - USER TO UPDATE THESE VALUES ---" section.
    *   **IMPORTANT:** Update the following constants:
        *   `SPREADSHEET_ID`:
            *   **If you want the script to create a brand new spreadsheet for you:** Leave `SPREADSHEET_ID` blank (e.g., `const SPREADSHEET_ID = '';`).
            *   **If you want to use an existing spreadsheet (e.g., a blank one you just created):** Get its ID from the URL (e.g., `https://docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit`) and paste it here.
        *   `MY_PRIMARY_INVOICING_EMAIL`: Your primary email address used for sending invoices (e.g., `libutti123@gmail.com`).
        *   `YOUR_COMPANY_NAME`: Your name or your company's name for email templates (e.g., `"Francis"`).
        *   `YOUR_SUPPORT_EMAIL_OR_CONTACT_INFO`: Contact information for reminder email footers (e.g., `"libutti123@gmail.com"`).
    *   Review other constants like `INVOICE_SUBJECT_KEYWORD` and `REMINDER_TIERS` and customize if needed.
4.  **Run Initial Setup:**
    *   In the Apps Script editor, select the `initialSetup` function from the function dropdown (next to the "Debug" and "Run" buttons).
    *   Click the **Run** button.
    *   **Authorization:** You will be prompted to authorize the script. Review the permissions and allow them. Google will show a warning screen because the script isn't verified by them (since it's your own); click "Advanced" and "Go to [Your Script Name] (unsafe)."
    *   The script will then:
        *   Create Gmail labels.
        *   If `SPREADSHEET_ID` was blank, it will create a new spreadsheet and show you an alert with the **NEW SPREADSHEET ID**. **You MUST copy this new ID back into the `SPREADSHEET_ID` constant in the script and save the script for future runs.**
        *   Set up the "Invoice Log", "Error Log", and "Dashboard" sheets with headers, formatting, formulas, and placeholder charts.
        *   Set up time-driven triggers.
        *   You'll see an alert confirming setup completion.
5.  **Verify Triggers:**
    *   In the Apps Script editor, go to the "Triggers" section (clock icon on the left).
    *   You should see triggers for `processSentInvoices`, `processPaymentReplies`, `checkAndUpdateOverdueStatuses`, and `sendOverdueReminders`.

## How to Use

*   **Initial Setup:** As described above (only needs to be done once correctly).
*   **Sending Invoices:** Simply send your invoices from the email address specified in `MY_PRIMARY_INVOICING_EMAIL` containing the `INVOICE_SUBJECT_KEYWORD`.
*   **Automated Processing:** The script will run automatically on the configured triggers:
    *   `processSentInvoices`: Periodically scans for new invoices.
    *   `checkAndUpdateOverdueStatuses`: Daily updates statuses to "Overdue".
    *   `sendOverdueReminders`: Daily sends reminders based on `REMINDER_TIERS`.
    *   `processPaymentReplies`: Periodically checks for replies on logged invoice threads.
*   **Monitoring:**
    *   Check the **Dashboard** sheet for an overview of your invoice status.
    *   Review the **Invoice Log** for detailed entries.
    *   Check the **Error Log** for any script errors or significant parsing issues.
    *   Monitor the "Executions" log in the Apps Script editor for any failed trigger runs.

## Current Limitations (V.0.7.7)

*   **Dashboard Currency Aggregation:** KPIs like "Total Amount Outstanding" sum values directly. If invoices use multiple currencies, this total will be a mixed-currency sum.
*   **Invoice Number Parsing in Replies:** `findInvoiceNumber` is tuned for original invoices. In replies, it might occasionally pick up other numbers. Mitigation is in place (`NeedsKeywordReview`).
*   **Chart Sophistication:** Charts are basic; advanced styling requires manual edits or migration to Looker Studio.
*   **Diverse Invoice Format Parsing:** Robustness depends on patterns. "Potential Parsing Discrepancies" highlights opportunities for refinement.
*   **I18n & Advanced Date/Term Parsing:** Limited support for non-US/European date formats, non-English keywords, or terms like "Net 30."

## Future Development Ideas

*   **Maintainability & Usability:**
    *   Move script constants to a "Configuration" sheet tab.
    *   Enhanced admin email notifications for critical failures.
*   **Dashboard V2 & Deeper Insights:**
    *   Consider Looker Studio migration.
    *   Currency-specific totals or conversion.
    *   Metrics like "Reminder Effectiveness."
*   **Parsing & I18n Expansion:**
    *   Address common parsing discrepancies.
    *   Expand date parsing.
    *   Multi-language support for payment keywords.
    *   Logic for terms like "Net 30".

## Contributing

While this is a personal project, suggestions and feedback are welcome via GitHub Issues. If you'd like to contribute code:
1.  Fork the repository.
2.  Create a new branch (`git checkout -b feature/your-feature-name`).
3.  Make your changes.
4.  Commit your changes (`git commit -am 'Add some feature'`).
5.  Push to the branch (`git push origin feature/your-feature-name`).
6.  Create a new Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgements

*   AIStudio assistance

---
