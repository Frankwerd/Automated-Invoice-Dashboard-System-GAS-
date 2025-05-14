// =============== CONFIGURATION & SCRIPT METADATA ===============
// Last Updated: 2025-05-13 (YYYY-MM-DD)
// Version: V.0.7.7 (Automated Dashboard Setup - Stable Core)// Author: Francis John LiButti (with AI assistance)
// Description: This script automates the tracking of sent invoices from Gmail, processes replies,
//              updates statuses, sends overdue reminder emails, sets up its own triggers, and creates a dashboard.


// --- CONFIGURATION - USER TO UPDATE THESE VALUES ---
const SPREADSHEET_ID = '19OYK5w-zC3vn4L2tjuuzpvnr5JeDSgt8CKWlu1oPBuE'; // <<<< PLEASE VERIFY/UPDATE THIS ID
const MY_PRIMARY_INVOICING_EMAIL = 'libutti123@gmail.com';
const YOUR_COMPANY_NAME = "Francis"; // <<< PLEASE UPDATE (Used in email templates)
const YOUR_SUPPORT_EMAIL_OR_CONTACT_INFO = "libutti123@gmail.com"; // <<< PLEASE UPDATE (Used in reminder footer)


const INVOICE_LOG_SHEET_NAME = 'Invoice Log';
const ERROR_LOG_SHEET_NAME = 'Error Log';
const DASHBOARD_SHEET_NAME = 'Dashboard'; // For the new dashboard
const INVOICE_SUBJECT_KEYWORD = 'invoice';
const PROCESS_INVOICES_AFTER_DATE = '2023-01-01'; // Example: '2024-01-01' - Adjust as needed


// --- LABEL CONFIGURATION (NESTED) ---
const PARENT_LABEL_NAME = 'Assisted Invoice Tracker';
const LABEL_PROCESSING = PARENT_LABEL_NAME + '/Processing';
const LABEL_SUCCESS = PARENT_LABEL_NAME + '/Logged';
const LABEL_ERROR = PARENT_LABEL_NAME + '/Error';
const LABEL_PAYMENT_REPLY_PROCESSED = PARENT_LABEL_NAME + '/PaymentReplyProcessed';
const REVIEW_LABEL_NAME = PARENT_LABEL_NAME + '/NeedsKeywordReview';


const LABELS_TO_SKIP_IF_PRESENT = [LABEL_SUCCESS, LABEL_ERROR];
const DEFAULT_CURRENCY = 'USD';
const SCRIPT_TIMEZONE = Session.getScriptTimeZone();


// --- DEBUGGING CONSTANT ---
const DEBUG_THREAD_ID_FOR_AMOUNTS = ""; // Example: "196ad3ef3673406e"


// --- REMINDER CONFIGURATION ---
const REMINDER_SUBJECT_EXCLUSION_TAG = "(DO NOT REPLY)"; // For exclusion and display
const DO_NOT_REPLY_FOOTER_TEXT = "\n\n--- Please do not reply directly to this automated email. If you have questions, please contact us at [Your Support Email or Preferred Contact Method]. ---";


const REMINDER_TIERS = [
  {
    stage: 'Level 1',
    daysOverdue: 7,
    subject: `Gentle Reminder: Invoice [Invoice #] Due ${REMINDER_SUBJECT_EXCLUSION_TAG}`,
    bodyTemplate: "Hi [Client Name],\n\nJust a friendly reminder that invoice [Invoice #] for [Currency] [Amount Due] was due on [Due Date].\n\nPlease let us know if payment has been arranged or if you have any questions.\n\nBest regards,\n[Your Company Name]" + DO_NOT_REPLY_FOOTER_TEXT
  },
  {
    stage: 'Level 2',
    daysOverdue: 30,
    subject: `Reminder: Invoice [Invoice #] Overdue ${REMINDER_SUBJECT_EXCLUSION_TAG}`,
    bodyTemplate: "Hi [Client Name],\n\nThis is a follow-up reminder that invoice [Invoice #] for [Currency] [Amount Due] was due on [Due Date] and is now significantly overdue.\n\nYour prompt attention to this matter would be greatly appreciated. Please let us know the status of this payment.\n\nThanks,\n[Your Company Name]" + DO_NOT_REPLY_FOOTER_TEXT
  },
  {
    stage: 'Final',
    daysOverdue: 60,
    subject: `URGENT: Invoice [Invoice #] Seriously Overdue ${REMINDER_SUBJECT_EXCLUSION_TAG}`,
    bodyTemplate: "Hi [Client Name],\n\nWe need to bring your urgent attention to invoice [Invoice #] for [Currency] [Amount Due], due on [Due Date], which is now 60+ days overdue.\n\nFailure to settle this outstanding amount may lead to further action. Please contact us immediately to arrange payment or discuss this critical matter.\n\nSincerely,\n[Your Company Name]" + DO_NOT_REPLY_FOOTER_TEXT
  },
];




// --- SPREADSHEET COLUMN HEADERS ---
const COL_PROCESSED_TS = 'Processed Timestamp';
const COL_EMAIL_SUBJECT = 'Email Subject';
const COL_SENT_DATE = 'Sent Date';
const COL_INVOICE_SENDER_EMAIL = 'Invoice Sender Email';
const COL_CLIENT_EMAIL = 'Client Email(s)';
const COL_INVOICE_NUM = 'Invoice #';
const COL_INVOICE_DATE = 'Invoice Date';
const COL_PARSED_GROSS = 'Parsed Gross Total';
const COL_PARSED_DISCOUNT = 'Parsed Discount';
const COL_AMOUNT_DUE = 'Amount Due';
const COL_CURRENCY = 'Currency';
const COL_CLIENT_NAME = 'Client Name';
const COL_STREET_ADDRESS = 'Street Address';
const COL_DUE_DATE = 'Due Date';
const COL_DATA_SOURCE = 'Source of Data';
const COL_ATTACH_PRESENT = 'Attachment Present';
const COL_ATTACH_NAMES = 'Attachment Name(s)';
const COL_STATUS = 'Status';
const COL_REPLY_PAID_DATE = 'Reply Paid Date';
const COL_REPLY_MENTIONED_AMOUNT = 'Reply Mentioned Amount';
const COL_REPLY_MENTIONED_CURRENCY = 'Reply Mentioned Currency';
const COL_PAYMENT_METHOD = 'Payment Method';
const COL_PAYMENT_NOTES = 'Payment Notes';
const COL_LINK_GMAIL = 'Link to Gmail Thread';
const COL_PARSING_NOTES = 'Notes/Error Details';
const COL_REMINDER_STAGE_SENT = 'Reminder Stage Sent';


const INVOICE_LOG_COLUMN_HEADERS = [
  COL_PROCESSED_TS, COL_EMAIL_SUBJECT, COL_SENT_DATE,
  COL_INVOICE_SENDER_EMAIL, COL_CLIENT_EMAIL,
  COL_INVOICE_NUM, COL_INVOICE_DATE,
  COL_PARSED_GROSS, COL_PARSED_DISCOUNT, COL_AMOUNT_DUE, COL_CURRENCY,
  COL_CLIENT_NAME, COL_STREET_ADDRESS, COL_DUE_DATE, COL_DATA_SOURCE,
  COL_ATTACH_PRESENT, COL_ATTACH_NAMES,
  COL_STATUS, COL_REPLY_PAID_DATE, COL_REPLY_MENTIONED_AMOUNT, COL_REPLY_MENTIONED_CURRENCY,
  COL_PAYMENT_METHOD, COL_PAYMENT_NOTES,
  COL_LINK_GMAIL, COL_PARSING_NOTES,
  COL_REMINDER_STAGE_SENT
];
const ERROR_LOG_COLUMN_HEADERS = ['Error Timestamp', 'Email Subject', 'Email Date', 'Recipient(s)','Function Name', 'Error Message', 'Link to Gmail Thread', 'Raw Email Body (Snippet)'];

// =============== INITIAL SETUP FUNCTION ===============
function initialSetup() {
  try { // <<<< MAIN TRY FOR THE ENTIRE FUNCTION 
    const SCRIPT_VERSION_FOR_LOG = "V.0.7.6q (Dashboard Full Setup + Unconditional Charts)"; 
    Logger.log(`Starting Initial Setup ${SCRIPT_VERSION_FOR_LOG}...`);
    let ui = null;
    try { ui = SpreadsheetApp.getUi(); } catch (e) { /* no UI */ }

    // --- Configuration Checks ---
    if (!MY_PRIMARY_INVOICING_EMAIL || !MY_PRIMARY_INVOICING_EMAIL.includes('@') ||
        !YOUR_COMPANY_NAME || YOUR_COMPANY_NAME === "Your Company Name Here" ||
        !YOUR_SUPPORT_EMAIL_OR_CONTACT_INFO || YOUR_SUPPORT_EMAIL_OR_CONTACT_INFO === "billing@example.com") {
        const errorMsg = "Configuration Error: Update MY_PRIMARY_INVOICING_EMAIL, YOUR_COMPANY_NAME, or YOUR_SUPPORT_EMAIL_OR_CONTACT_INFO.";
        Logger.log("E: " + errorMsg);
        if (ui) ui.alert("Configuration Error", errorMsg, ui.ButtonSet.OK);
        return;
    }

    Logger.log('Step 1: Gmail labels...');
    getOrCreateLabel(PARENT_LABEL_NAME);
    [LABEL_PROCESSING, LABEL_SUCCESS, LABEL_ERROR, LABEL_PAYMENT_REPLY_PROCESSED, REVIEW_LABEL_NAME].forEach(l => { if (l) getOrCreateLabel(l); });
    Logger.log('Gmail labels OK.');

    Logger.log('Step 2: Spreadsheet...');
    let spreadsheet = null; 
    let newSpreadsheetFileCreated = false; 
    let ssId = SPREADSHEET_ID ? SPREADSHEET_ID.trim() : '';

    if (ssId && ssId !== 'YOUR_SPREADSHEET_ID_HERE' && ssId !== 'PASTE YOUR SPREADSHEET ID HERE' && ssId !== '') {
      try { spreadsheet = SpreadsheetApp.openById(ssId); Logger.log(`Opened existing sheet: "${spreadsheet.getName()}" (ID: ${ssId})`); }
      catch (e) { Logger.log(`W: Cannot open SPREADSHEET_ID "${ssId}". Will create new. E: ${e.message}`); ssId = ''; spreadsheet = null; }
    } else { Logger.log('SPREADSHEET_ID not configured. Will create new.'); ssId = ''; }

    if (!spreadsheet) { 
      const dateStr = Utilities.formatDate(new Date(), SCRIPT_TIMEZONE, 'yyyy-MM-dd');
      const newSheetName = `Automated Invoice Log - ${dateStr}`;
      spreadsheet = SpreadsheetApp.create(newSheetName); 
      newSpreadsheetFileCreated = true; 
      ssId = spreadsheet.getId();
      Logger.log(`!!! NEW SPREADSHEET FILE CREATED: "${newSheetName}" ID: ${ssId}. ACTION: Please copy this ID into the SPREADSHEET_ID constant for future runs. !!!`);
      if (ui) { 
          ui.alert('New Spreadsheet File Created!', `Name: "${newSheetName}".\n\nIMPORTANT: For the script to work correctly in the future, copy this ID into the SPREADSHEET_ID constant in the script:\n\n${ssId}\n\nThe script will now proceed to set up sheets.`, ui.ButtonSet.OK);
      }
    }
    
    if (!spreadsheet) { 
        Logger.log("FATAL: Spreadsheet object is null after create/open attempt."); 
        if(ui) ui.alert('Fatal Error', 'Spreadsheet initialization failed critically.', ui.ButtonSet.OK); 
        return; 
    }

    // Call setupSheet for Invoice Log and Error Log
    // The 'newSpreadsheetFileCreated' flag ensures they get full formatting if the Spreadsheet file itself is new.
    // Assumes setupSheet V.0.7.6p is in your script.
    setupSheet(spreadsheet, INVOICE_LOG_SHEET_NAME, INVOICE_LOG_COLUMN_HEADERS, true, newSpreadsheetFileCreated);
    setupSheet(spreadsheet, ERROR_LOG_SHEET_NAME, ERROR_LOG_COLUMN_HEADERS, false, newSpreadsheetFileCreated);

    // Dashboard Sheet Setup 
    Logger.log(`Setting up sheet: "${DASHBOARD_SHEET_NAME}"...`);
    let dashboardSheet = spreadsheet.getSheetByName(DASHBOARD_SHEET_NAME);
    if (!dashboardSheet) {
      dashboardSheet = spreadsheet.insertSheet(DASHBOARD_SHEET_NAME); Logger.log(`Sheet "${DASHBOARD_SHEET_NAME}" created.`);
    } else {
      Logger.log(`Sheet "${DASHBOARD_SHEET_NAME}" exists. Clearing for fresh setup.`);
      dashboardSheet.clearContents(); dashboardSheet.clearFormats(); dashboardSheet.clearNotes();
      let existingCharts = dashboardSheet.getCharts(); 
      existingCharts.forEach(function(chart){ dashboardSheet.removeChart(chart); });
      if (existingCharts.length > 0) Logger.log(`Removed ${existingCharts.length} old charts from Dashboard.`);
    }
    
    const dashboardLayoutSource = [ 
        ["Key Financial Overview", "Value", "", "Notes & Guidance"], 
        ["Total Amount Currently Overdue:", ""], 
        ["Total Amount Outstanding (Unpaid/Overdue/Pending):", ""], 
        ["Number of Overdue Invoices:", ""], 
        ["Average Days Overdue (for Overdue Invoices):", ""], 
        [], 
        ["Action & Review Center", "Count"], 
        ["Invoices: Needs Payment Keyword/Reply Review", ""], 
        ["Invoices: Potential Parsing Discrepancies", ""],   
        ["Invoices: Critical Script Errors (in Invoice Log)", ""], 
        ["Entries in Error Log Sheet:", ""], 
        [],
        ["Overdue Aging Analysis", "Total Amount ($)", "Number of Invoices (#)"], 
        ["1-30 Days Overdue", "", ""],   
        ["31-60 Days Overdue", "", ""],  
        ["61-90 Days Overdue", "", ""],  
        ["91+ Days Overdue", "", ""],    
        [],
        ["Top 5 Clients with Overdue Balances", "Total Overdue ($)"], 
        ["", ""], ["", ""], ["", ""], ["", ""], ["", ""], 
    ];
    const maxCols = Math.max(...dashboardLayoutSource.map(r => r.length ? r.length : 0)); 
    const dashboardLayoutFormatted = dashboardLayoutSource.map(row => { const newRow = [...row]; while (newRow.length < maxCols) newRow.push(""); return newRow; });
    dashboardSheet.getRange(1, 1, dashboardLayoutFormatted.length, maxCols).setValues(dashboardLayoutFormatted);
    dashboardSheet.getRange("D1").setValue("Data auto-updates. Charts created by script.");
    Logger.log(`"${DASHBOARD_SHEET_NAME}" text headers applied.`);

    Logger.log(`Setting formulas in "${DASHBOARD_SHEET_NAME}"...`);
    const ilSheetRef = `'${INVOICE_LOG_SHEET_NAME.replace(/'/g, "''")}'!`; 
    const errSheetRef = `'${ERROR_LOG_SHEET_NAME.replace(/'/g, "''")}'!`;

    dashboardSheet.getRange("B2").setFormula(`=IFERROR(SUMIF(${ilSheetRef}R:R, "Overdue", ${ilSheetRef}J:J), 0)`);
    dashboardSheet.getRange("B3").setFormula(`=IFERROR(SUM(SUMIFS(${ilSheetRef}J:J, ${ilSheetRef}R:R, {"Unpaid";"Overdue";"Partially Paid";"Pending Confirmation"})), 0)`);
    dashboardSheet.getRange("B4").setFormula(`=IFERROR(COUNTIF(${ilSheetRef}R:R, "Overdue"), 0)`);
    dashboardSheet.getRange("B5").setFormula(`=IFERROR(AVERAGE(FILTER(TODAY()-${ilSheetRef}N2:N, ${ilSheetRef}R2:R="Overdue", ISNUMBER(${ilSheetRef}N2:N))), 0)`);
    dashboardSheet.getRange("B8").setFormula(`=IFERROR(COUNTIF(${ilSheetRef}Y:Y, "*NEEDS KEYWORD REVIEW*"), 0)`);
    dashboardSheet.getRange("B9").setFormula(`=IFERROR(SUM(ARRAYFORMULA(COUNTIF(${ilSheetRef}Y:Y, {"*Client Name NP.*";"*Invoice # NP.*";"*Gross Total NP.*";"*Inv Date NP.*";"*Due Date NP.*"}))), 0)`);
    dashboardSheet.getRange("B10").setFormula(`=IFERROR(COUNTIF(${ilSheetRef}Y:Y, "*SCRIPT ERROR*"), 0)`);
    dashboardSheet.getRange("B11").setFormula(`=IFERROR(COUNTA(IFERROR(FILTER(${errSheetRef}A2:A, LEN(${errSheetRef}A2:A)>0))), 0)`);
    dashboardSheet.getRange("B14").setFormula(`=IFERROR(SUMIFS(${ilSheetRef}J:J, ${ilSheetRef}R:R, "Overdue", ${ilSheetRef}N:N, ">="&TODAY()-30, ${ilSheetRef}N:N, "<"&TODAY()), 0)`);
    dashboardSheet.getRange("C14").setFormula(`=IFERROR(COUNTIFS(${ilSheetRef}R:R, "Overdue", ${ilSheetRef}N:N, ">="&TODAY()-30, ${ilSheetRef}N:N, "<"&TODAY()), 0)`);
    dashboardSheet.getRange("B15").setFormula(`=IFERROR(SUMIFS(${ilSheetRef}J:J, ${ilSheetRef}R:R, "Overdue", ${ilSheetRef}N:N, ">="&TODAY()-60, ${ilSheetRef}N:N, "<"&TODAY()-30), 0)`);
    dashboardSheet.getRange("C15").setFormula(`=IFERROR(COUNTIFS(${ilSheetRef}R:R, "Overdue", ${ilSheetRef}N:N, ">="&TODAY()-60, ${ilSheetRef}N:N, "<"&TODAY()-30), 0)`);
    dashboardSheet.getRange("B16").setFormula(`=IFERROR(SUMIFS(${ilSheetRef}J:J, ${ilSheetRef}R:R, "Overdue", ${ilSheetRef}N:N, ">="&TODAY()-90, ${ilSheetRef}N:N, "<"&TODAY()-60), 0)`);
    dashboardSheet.getRange("C16").setFormula(`=IFERROR(COUNTIFS(${ilSheetRef}R:R, "Overdue", ${ilSheetRef}N:N, ">="&TODAY()-90, ${ilSheetRef}N:N, "<"&TODAY()-60), 0)`);
    dashboardSheet.getRange("B17").setFormula(`=IFERROR(SUMIFS(${ilSheetRef}J:J, ${ilSheetRef}R:R, "Overdue", ${ilSheetRef}N:N, "<"&TODAY()-90), 0)`);
    dashboardSheet.getRange("C17").setFormula(`=IFERROR(COUNTIFS(${ilSheetRef}R:R, "Overdue", ${ilSheetRef}N:N, "<"&TODAY()-90), 0)`);
    dashboardSheet.getRange("A20").setFormula(`=IFERROR(QUERY(${ilSheetRef}A:Y, "SELECT L, SUM(J) WHERE R = 'Overdue' AND L IS NOT NULL AND L <> '' GROUP BY L ORDER BY SUM(J) DESC LIMIT 5 LABEL L '', SUM(J) ''", 0), IFERROR(SPLIT("No overdue clients| ", "|"), ""))`);
    
    dashboardSheet.getRange("A1:D1").setFontWeight("bold"); dashboardSheet.getRange("A7:B7").setFontWeight("bold"); dashboardSheet.getRange("A13:C13").setFontWeight("bold"); dashboardSheet.getRange("A19:B19").setFontWeight("bold");
    dashboardSheet.getRange("B2:B5").setFontSize(12).setFontWeight("bold").setHorizontalAlignment("right"); 
    dashboardSheet.getRange("B8:B11").setFontSize(11).setHorizontalAlignment("right"); 
    dashboardSheet.getRange("B14:C17").setHorizontalAlignment("right");
    dashboardSheet.getRange("B20:B24").setHorizontalAlignment("right"); 
    dashboardSheet.getRange("B2:B3").setNumberFormat("$#,##0.00;($#,##0.00);\"--\""); 
    dashboardSheet.getRange("B5").setNumberFormat("0.00;0;\"--\""); 
    dashboardSheet.getRange("B4").setNumberFormat("0;\"0\";\"--\"");     
    dashboardSheet.getRange("B8:B11").setNumberFormat("0;\"0\";\"--\""); 
    dashboardSheet.getRange("B14:B17").setNumberFormat("$#,##0.00;($#,##0.00);\"--\"");
    dashboardSheet.getRange("C14:C17").setNumberFormat("0;\"0\";\"--\""); 
    dashboardSheet.getRange("B20:B24").setNumberFormat("$#,##0.00;($#,##0.00);\"--\""); 
    Logger.log('Formulas and cell formatting set for Dashboard.');
    
    // --- Create Basic Charts (Unconditionally) ---
    Logger.log('Creating basic charts on Dashboard (will populate as data appears)...');
    try { 
        let charts = dashboardSheet.getCharts(); 
        charts.forEach(function(chart) { dashboardSheet.removeChart(chart);});
        if (charts.length > 0) Logger.log(`Removed ${charts.length} existing charts for fresh setup.`);

        // 1. Aging Analysis Chart (Column Chart for Amount)
        let agingAmountDataRange = dashboardSheet.getRange("A13:B17"); 
        let agingChartAmount = dashboardSheet.newChart().setChartType(Charts.ChartType.COLUMN).addRange(agingAmountDataRange)
            .setOption('title', 'Overdue Invoice Aging - By Amount ($)').setOption('hAxis', { title: 'Aging Bucket', slantedText: true, slantedTextAngle: 30 })
            .setOption('vAxis', { title: 'Total Amount Overdue', format: '$#,##0.00', viewWindow: { min: 0} })
            .setOption('legend', { position: 'none' }) 
            .setOption('height', 300).setOption('width', 500)
            .setPosition(2, 5, 0, 0) // Anchor: Row 2, Col E (5th col), offset 0,0
            .build();
        dashboardSheet.insertChart(agingChartAmount); 
        Logger.log('Aging Analysis chart (Amount) placeholder created.');
        
        // 2. Aging Analysis Chart (Column Chart for Count)
        let agingCountCategoriesRange = dashboardSheet.getRange("A13:A17"); 
        let agingCountSeriesRange = dashboardSheet.getRange("C13:C17"); 
        let agingChartCount = dashboardSheet.newChart().setChartType(Charts.ChartType.COLUMN)
            .addRange(agingCountCategoriesRange) 
            .addRange(agingCountSeriesRange)   
            .setOption('title', 'Overdue Invoice Aging - By Count (#)').setOption('hAxis', { title: 'Aging Bucket', slantedText: true, slantedTextAngle: 30 })
            .setOption('vAxis', { title: 'Number of Invoices', format: '0', viewWindow: { min: 0 } })
            .setOption('legend', { position: 'none' }) 
            .setOption('height', 300).setOption('width', 500)
            .setPosition(18, 5, 0, 0) // Anchor: Row 18, Col E (5th col)
            .build();
        dashboardSheet.insertChart(agingChartCount); 
        Logger.log('Aging Analysis chart (Count) placeholder created.');

        // 3. Top 5 Overdue Clients (Bar Chart)
        let topClientsDataRange = dashboardSheet.getRange("A19:B24"); 
        let topClientsChart = dashboardSheet.newChart().setChartType(Charts.ChartType.BAR)
            .addRange(topClientsDataRange)
            .setOption('title', 'Top 5 Clients by Overdue Amount')
            .setOption('vAxis', { title: 'Client' }) 
            .setOption('hAxis', { title: 'Total Overdue Amount', format: '$#,##0.00', viewWindow: { min: 0 } })
            .setOption('legend', { position: 'none' })
            .setOption('height', 320) 
            .setOption('width', 500)
            .setPosition(2, 10, 0, 0) // Anchor: Row 2, Col J (10th col)
            .build();
        dashboardSheet.insertChart(topClientsChart); 
        Logger.log('Top Overdue Clients chart placeholder created.');

    } catch(e) { 
        Logger.log(`W: Error during chart creation: ${e.message}\n${e.stack}`); 
    }
    Logger.log('Dashboard chart creation placeholders finished.');
        
    try { 
      if (dashboardSheet) { 
        spreadsheet.setActiveSheet(dashboardSheet); 
        spreadsheet.moveActiveSheet(1); 
        Logger.log(`Sheet "${DASHBOARD_SHEET_NAME}" moved to first position.`); 
      } else {
        Logger.log(`W: Dashboard sheet object not available for reordering (should not happen).`);
      }
    } catch (e) { 
        Logger.log(`W: Could not move dashboard: ${e.message}`); 
    }

    if (newSpreadsheetFileCreated) {
        const defaultSheet = spreadsheet.getSheetByName('Sheet1');
        if (defaultSheet && spreadsheet.getSheets().length > 1) { 
            try { spreadsheet.deleteSheet(defaultSheet); Logger.log('Default "Sheet1" removed.'); } 
            catch(e){ Logger.log(`W: Could not delete "Sheet1": ${e.message}`);}
        }
    }

    Logger.log('Step 3: Triggers...');
    setupTriggers(); // Assumes setupTriggers is defined elsewhere in your script

    let completionMessage = `Initial Setup (${SCRIPT_VERSION_FOR_LOG}) Complete! Using spreadsheet: "${spreadsheet.getName()}" (ID: ${ssId}).\nAll sheets, dashboard formulas, basic charts, and triggers are set up.`;
    if (newSpreadsheetFileCreated) { 
        completionMessage += `\n\nACTION: A NEW SPREADSHEET FILE WAS CREATED. For future script runs, please copy the new Spreadsheet ID into the SPREADSHEET_ID constant in your script and save.\nID: ${ssId}`;
    }
    Logger.log(completionMessage);
    if(ui) { ui.alert(`Setup Complete (${SCRIPT_VERSION_FOR_LOG})`, completionMessage, ui.ButtonSet.OK); }
    else { Logger.log("Setup Complete - UI Alert skipped."); }
    Logger.log('Initial setup function finished.');

  } catch (e) { // <<<< FINAL CATCH FOR THE ENTIRE initialSetup (This is line 296 approx in this block)
    Logger.log(`FATAL ERROR during initialSetup: ${e.toString()}\nStack: ${e.stack}`);
    try { 
        let uiForError = SpreadsheetApp.getUi(); 
        uiForError.alert('Fatal Error During Setup', `Error: ${e.message}. Check logs.`, uiForError.ButtonSet.OK); 
    } catch (uiError) {
        Logger.log(`Could not display fatal error UI alert: ${uiError.message}`);
    }
  } // <<<< FINAL CLOSING BRACE for initialSetup function
}

// =============== MAIN PROCESSING FUNCTION (FOR NEW INVOICES) ===============
function processSentInvoices() {
 Logger.log("Starting processSentInvoices run (V.0.7.5)...");
 if (!SPREADSHEET_ID || SPREADSHEET_ID.startsWith('YOUR_') || SPREADSHEET_ID.startsWith('PASTE_') || SPREADSHEET_ID === '' || !MY_PRIMARY_INVOICING_EMAIL || !MY_PRIMARY_INVOICING_EMAIL.includes('@')) {
   Logger.log("E: SPREADSHEET_ID or MY_PRIMARY_INVOICING_EMAIL not configured. Run initialSetup if a new sheet ID was generated and needs to be updated in the script."); return;
 }
 let spreadsheet, invoiceLogSheet, errorLogSheet;
 try { spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); invoiceLogSheet = spreadsheet.getSheetByName(INVOICE_LOG_SHEET_NAME); errorLogSheet = spreadsheet.getSheetByName(ERROR_LOG_SHEET_NAME); if (!invoiceLogSheet||!errorLogSheet) throw new Error("Log sheet(s) missing. Run initialSetup.");}
 catch (e) { Logger.log(`E opening sheets: ${e.message}. Run initialSetup if sheet ID needs refresh or sheet missing.`); return; }


 const scriptUserEmailLower = Session.getActiveUser().getEmail().toLowerCase();
 const myPrimaryInvoicingEmailLower = MY_PRIMARY_INVOICING_EMAIL.toLowerCase();
 const ownSendingEmails = [scriptUserEmailLower];
 if (myPrimaryInvoicingEmailLower && myPrimaryInvoicingEmailLower !== scriptUserEmailLower) ownSendingEmails.push(myPrimaryInvoicingEmailLower);


 const skipLabelsQueryPart = LABELS_TO_SKIP_IF_PRESENT.map(l => `-label:"${l.replace(/"/g, '\\"')}"`).join(' ');
 const processingLabelStr = LABEL_PROCESSING.replace(/"/g, '\\"');
 const subjectKeywordStr = INVOICE_SUBJECT_KEYWORD.replace(/"/g, '\\"');
  let dateQueryPart = "";
 if (PROCESS_INVOICES_AFTER_DATE && /^\d{4}-\d{2}-\d{2}$/.test(PROCESS_INVOICES_AFTER_DATE)) {
   const [year, month, day] = PROCESS_INVOICES_AFTER_DATE.split('-');
   dateQueryPart = ` after:${year}/${month}/${day}`;
   Logger.log(`Processing invoices sent after: ${PROCESS_INVOICES_AFTER_DATE}`);
 } else {
   Logger.log(`PROCESS_INVOICES_AFTER_DATE is not set or invalid. Processing all applicable invoices.`);
 }


 const searchQuery = `subject:(${subjectKeywordStr}) -subject:"${REMINDER_SUBJECT_EXCLUSION_TAG}" in:sent ${skipLabelsQueryPart} -label:"${processingLabelStr}" -label:"${REVIEW_LABEL_NAME.replace(/"/g, '\\"')}" -label:"${LABEL_PAYMENT_REPLY_PROCESSED.replace(/"/g, '\\"')}"${dateQueryPart}`;
 Logger.log(`Invoice thread search: ${searchQuery}`);
 let threads;
 try { threads = GmailApp.search(searchQuery, 0, 50); if (!threads) threads = []; Logger.log(`Found ${threads.length} candidate threads.`);}
 catch (e) { Logger.log(`E during Gmail search: ${e.message}`); return; }
 if (threads.length === 0) { Logger.log("No new invoice threads found matching criteria."); return; }


 const procLabel=getOrCreateLabel(LABEL_PROCESSING); const succLabel=getOrCreateLabel(LABEL_SUCCESS); const errLabel=getOrCreateLabel(LABEL_ERROR);
 if (!procLabel||!succLabel||!errLabel) { Logger.log("E: Critical labels missing."); return;}


 threads.forEach(thread => {
   let invoiceMessageToProcess = null; const messages = thread.getMessages();
   if (messages.length === 0) { Logger.log(`Thread ${thread.getId()} empty.`); return; }


   let isDebugTargetThread = (DEBUG_THREAD_ID_FOR_AMOUNTS && thread.getId() === DEBUG_THREAD_ID_FOR_AMOUNTS);


   for (let k = 0; k < messages.length; k++) {
       const currentMessage = messages[k]; const msgFrom = currentMessage.getFrom().toLowerCase();
       const msgFromEmailOnlyMatch = msgFrom.match(/<([^>]+)>/);
       const msgFromEmailOnly = msgFromEmailOnlyMatch ? msgFromEmailOnlyMatch[1] : msgFrom;


       if (ownSendingEmails.includes(msgFromEmailOnly)) {
           const subj = currentMessage.getSubject().toLowerCase();
           if (!subj.startsWith("re:") && !subj.startsWith("fw:") && !subj.startsWith("fwd:") && !subj.includes(REMINDER_SUBJECT_EXCLUSION_TAG.toLowerCase())) {
               invoiceMessageToProcess = currentMessage;
               break;
           }
       }
   }


   if (!invoiceMessageToProcess) {
     Logger.log(`Skipping thread ${thread.getId()} (Subj: "${thread.getFirstMessageSubject()}") as no eligible base invoice message found (e.g., it might be a reminder with exclusion tag, or reply only).`);
     return;
   }


   const msgSubj = invoiceMessageToProcess.getSubject(); const msgDate = invoiceMessageToProcess.getDate();
   const msgId = invoiceMessageToProcess.getId();
   const threadIdToStore = thread.getId();
   Logger.log(`Processing invoice: "${msgSubj}" (MsgID: ${msgId}) in thread ${threadIdToStore}`);
   try { thread.addLabel(procLabel); } catch(e) { /* ignore */ }


   let emailBody = null; let recipients = null;
   try {
     recipients = invoiceMessageToProcess.getTo(); const sentDate = formatDate(msgDate);
     emailBody = invoiceMessageToProcess.getPlainBody();


     const attachments = invoiceMessageToProcess.getAttachments(); const attachPresent = attachments.length > 0;
     const attachNames = attachPresent ? attachments.map(att => att.getName()).join(', ') : '';
     const parsedData = parseInvoiceDetailsFromEmail(emailBody, msgSubj, attachments, recipients, isDebugTargetThread);
     const gross = parsedData.grossTotal === null ? 0 : parsedData.grossTotal;
     const disc = parsedData.discount === null ? 0 : parsedData.discount;
     const amountDue = gross - disc;


     const rowData = [
       formatDate(new Date()), msgSubj, sentDate, MY_PRIMARY_INVOICING_EMAIL, recipients,
       parsedData.invoiceNumber||'', parsedData.invoiceDate||sentDate,
       parsedData.grossTotal==null?'':parsedData.grossTotal, parsedData.discount==null||parsedData.discount==0?'':parsedData.discount,
       amountDue==null?'':amountDue.toFixed(2), parsedData.currency||DEFAULT_CURRENCY,
       parsedData.clientName||'', parsedData.streetAddress||'', parsedData.dueDate||'', "Email Body",
       attachPresent?"Yes":"No", attachNames,
       'Unpaid',
       '', '', '',
       '', '',
       threadIdToStore, parsedData.parsingNotes||'',
       ''
     ];
     invoiceLogSheet.appendRow(rowData);
     Logger.log(`Logged Inv#: ${parsedData.invoiceNumber||'N/A'}. AmtDue: ${amountDue.toFixed(2)}`);
     thread.removeLabel(procLabel); thread.addLabel(succLabel);
   } catch (e) {
     Logger.log(`E processing "${msgSubj}" (MsgID: ${msgId}): ${e.message}\n${e.stack}`);
     logErrorToSheetAdvanced(errorLogSheet, {subject:msgSubj,date:msgDate,recipients:recipients,functionName:'processSentInvoicesLoop',errorMessage:e.message,threadLink:threadIdToStore,rawBodySnippet:emailBody?emailBody.substring(0,500):"N/A"});
     try{thread.removeLabel(procLabel); thread.addLabel(errLabel);}catch(e){}
   }
 });
 Logger.log("processSentInvoices finished.");
}

// =============== PROCESS PAYMENT REPLIES FUNCTION ===============
// V.0.7.6d - Added Invoice # cross-check in reply body
function processPaymentReplies() {
 Logger.log("Starting processPaymentReplies run (V.0.7.6d - Inv# Check)...");
 if (!SPREADSHEET_ID || SPREADSHEET_ID.startsWith('YOUR_') || SPREADSHEET_ID.startsWith('PASTE_') || SPREADSHEET_ID === '' || !MY_PRIMARY_INVOICING_EMAIL || !MY_PRIMARY_INVOICING_EMAIL.includes('@')) {
   Logger.log("E: SPREADSHEET_ID or MY_PRIMARY_INVOICING_EMAIL not configured. Run initialSetup."); return;
 }
 let spreadsheet, sheet;
 try { spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); sheet = spreadsheet.getSheetByName(INVOICE_LOG_SHEET_NAME); if (!sheet) throw new Error("Sheet missing.");}
 catch (e) { Logger.log(`E opening sheet: ${e.message}. Run initialSetup if ID needs refresh or sheet missing.`); return; }

 const loggedLabel = getOrCreateLabel(LABEL_SUCCESS);       if (!loggedLabel) { Logger.log(`E: Label "${LABEL_SUCCESS}" missing.`); return; }
 const reviewLabel = getOrCreateLabel(REVIEW_LABEL_NAME); if (!reviewLabel) { Logger.log(`E: Label "${REVIEW_LABEL_NAME}" missing.`); return; }
 const replyProcessedLabel = getOrCreateLabel(LABEL_PAYMENT_REPLY_PROCESSED); if (!replyProcessedLabel) { Logger.log(`E: Label "${LABEL_PAYMENT_REPLY_PROCESSED}" missing.`); return;}

 const threads = GmailApp.search(
   `label:"${loggedLabel.getName()}" -label:"${LABEL_PAYMENT_REPLY_PROCESSED.replace(/"/g, '\\"')}"`,
   0, 50);
 Logger.log(`Found ${threads.length} logged threads for reply check (excluding those already reply-processed).`);
 if (threads.length === 0) { Logger.log("No new threads to check for payment replies."); return; }

  const scriptUserMail = Session.getActiveUser().getEmail().toLowerCase();
 const primaryInvoiceMail = MY_PRIMARY_INVOICING_EMAIL.toLowerCase();
 const ownMails = [scriptUserMail]; if (primaryInvoiceMail && scriptUserMail !== primaryInvoiceMail) ownMails.push(primaryInvoiceMail);
 Logger.log(`"Own" emails for reply check: [${ownMails.join('; ')}]`);

 const basePaymentKeywords = [
   'paid', 'payment made', 'payment sent', 'settled', 'transferred', 'completed payment',
   'payment has been made', 'invoice settled', 'remitted the payment', 'funds sent', 'partial payment',
   'processed the payment', 'payment has been completed', 'been completed',
   'payment was sent', 'made the payment', 'remitted', 'remittance', 'wired', 'wire transfer',
   'ach transfer', 'funds are on their way', 'payment is on its way', 'payment processed',
   'payment completed', 'payment submitted', 'transaction complete', 'transaction successful',
   'payment confirmed', 'payment should reflect', 'payment initiated', 'sent via PayPal',
   'paid with PayPal', 'PayPal payment sent', 'PayPal transfer complete', 'payment through PayPal',
   'sent via Venmo', 'paid with Venmo', 'Venmo payment sent', 'sent you a Venmo', 'Venmo transfer complete',
   'payment through Venmo', 'sent Venmo', 'sent via Zelle', 'paid with Zelle', 'Zelle payment sent',
   'Zelle transfer complete', 'payment through Zelle', 'Zelle payment made', 'card payment successful',
   'payment by card completed', 'credit card payment made', 'bank transfer initiated', 'bank transfer made',
   'direct deposit sent', 'direct deposit made', 'e-transfer sent', 'electronic funds transfer',
   'EFT sent', 'EFT made', 'money sent', 'payment done', 'transfer done', 'funds transferred', 'been sent'
 ];
 const paymentKeywords = [...new Set(basePaymentKeywords)];
 Logger.log("DEBUG: Keywords list (" + paymentKeywords.length + " unique). First 5: " + JSON.stringify(paymentKeywords.slice(0,5)) + "...");

 const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
 const linkIdx           = headers.indexOf(COL_LINK_GMAIL);
 const notesIdx          = headers.indexOf(COL_PAYMENT_NOTES);
 const invNumIdx         = headers.indexOf(COL_INVOICE_NUM);
 const clientMailIdx     = headers.indexOf(COL_CLIENT_EMAIL);
 const parsingNotesSheetIdx = headers.indexOf(COL_PARSING_NOTES);
 const statusIdx         = headers.indexOf(COL_STATUS);
 const replyPaidDateIdx  = headers.indexOf(COL_REPLY_PAID_DATE);
 const replyAmountIdx    = headers.indexOf(COL_REPLY_MENTIONED_AMOUNT);
 const replyCurrencyIdx  = headers.indexOf(COL_REPLY_MENTIONED_CURRENCY);
 const invCurrencyIdx    = headers.indexOf(COL_CURRENCY);
 const amountDueIdx      = headers.indexOf(COL_AMOUNT_DUE);
 const dueDateSheetIdx   = headers.indexOf(COL_DUE_DATE);
 const reminderSentIdx   = headers.indexOf(COL_REMINDER_STAGE_SENT);

 if ([linkIdx, notesIdx, invNumIdx, clientMailIdx, parsingNotesSheetIdx, statusIdx, replyPaidDateIdx, replyAmountIdx, replyCurrencyIdx, invCurrencyIdx, amountDueIdx, dueDateSheetIdx, reminderSentIdx].includes(-1)) {
   Logger.log(`E: One or more required column indices not found. Check COL_ constants and sheet headers.`); return;
 }

 const allDataRange = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow()-1), sheet.getLastColumn());
 const allData = allDataRange.getValues();
 const allFormulas = allDataRange.getFormulas();

 threads.forEach(thread => {
   const currentGmailThreadId = thread.getId();
   let rowData = null; let rowIndex = -1; let actualRow = -1; let foundMatchInSheet = false;

   for (let j=0; j<allData.length; j++) {
       let idFromSheet = '';
       const formulaValue = allFormulas[j][linkIdx];
       if (formulaValue && typeof formulaValue === 'string' && formulaValue.toUpperCase().startsWith('=HYPERLINK("')) {
           const formulaMatch = formulaValue.match(/#inbox\/([^"]+)"/);
           if (formulaMatch && formulaMatch[1]) idFromSheet = formulaMatch[1].trim();
           else idFromSheet = String(allData[j][linkIdx]).trim();
       } else { idFromSheet = String(allData[j][linkIdx]).trim(); }
       if (idFromSheet === currentGmailThreadId) {
           rowData = allData[j]; rowIndex = j; actualRow = j + 2; foundMatchInSheet = true; break;
       }
   }

   if (!foundMatchInSheet) {
       Logger.log(`Thread ${currentGmailThreadId} (Subj: "${thread.getFirstMessageSubject()}"): NO MATCHING ROW. SKIPPED.`); return;
   }

   const invNumFromSheetRaw = String(rowData[invNumIdx] || "N/A_SHEET_INV_NUM");
   const clientEmailsStr = rowData[clientMailIdx];
   const clientMails = clientEmailsStr ? String(clientEmailsStr).toLowerCase().split(',').map(email => {const m=email.match(/<([^>]+)>/); return m?m[1].trim():email.trim();}).filter(e=>e&&e.includes('@')) : [];
   const invoiceCurrency = rowData[invCurrencyIdx] || DEFAULT_CURRENCY;
   const originalAmountDue = parseFloat(rowData[amountDueIdx]);
   const originalStatusOnSheet = rowData[statusIdx];

   Logger.log(`-----\nInv# ${invNumFromSheetRaw} (Thread: ${currentGmailThreadId}, Row: ${actualRow}, Original Inv Due: ${originalAmountDue} ${invoiceCurrency}, Current Sheet Status: ${originalStatusOnSheet}):`);
   Logger.log(`   Raw Client Email(s) from sheet: "${clientEmailsStr}"`);
   Logger.log(`   Parsed Client Emails for this row: [${clientMails.join('; ')}]`);

   let paymentReplyProcessedForThisThread = false; let processedAReplyInThisRun = false;

   try {
     const messages = thread.getMessages();
     Logger.log(`   Thread ${currentGmailThreadId} has ${messages.length} messages.`);

     for (let k = messages.length - 1; k >= 0; k--) {
       if(paymentReplyProcessedForThisThread) break;

       const msg = messages[k];
       const fromFull = msg.getFrom(); const fromDate = msg.getDate(); const fromSubj = msg.getSubject();
       const fromMailMatch = fromFull.match(/<([^>]+)>/);
       const fromMail = fromMailMatch ? fromMailMatch[1].toLowerCase() : fromFull.toLowerCase().trim();
       let isOwn = ownMails.includes(fromMail); let isListedClient = clientMails.includes(fromMail);

       Logger.log(`   -> Inspecting msg [${k+1}/${messages.length}] from "${fromFull}" (Subj: "${fromSubj}"). IsOwn: ${isOwn}. IsListedClient: ${isListedClient}.`);
       if (isOwn || !isListedClient) continue;
       Logger.log(`      PASSED CHECKS: Processing reply from client "${fromFull}" for Inv# ${invNumFromSheetRaw}`);

       const originalBody = msg.getPlainBody();
       let newReplyTextOnly = originalBody;
       const replySeparators = [ /\nOn .* wrote:\s*\n/i, /\n----- Original Message -----\n/i, /\nFrom: .* <.*@.*>\s*\nSent: .*\nTo: .*\nSubject: .*/i, /\n>? ?From: .*\n>? ?Date: .*\n>? ?Subject: .*\n>? ?To: .*/i, /^\s*>+.*\n?/gm ];
       let separatorFoundInLoop = false;
       for (const separator of replySeparators) {
           const match = originalBody.match(separator);
           if (match) { newReplyTextOnly = originalBody.substring(0, match.index).trim(); separatorFoundInLoop = true; break; }
       }
       if (!separatorFoundInLoop) { /* Use full body if no separator */ }
       else if (newReplyTextOnly.length < 10) { newReplyTextOnly = originalBody; }

       let bodyToTestInLoop = newReplyTextOnly.replace(/[\s\u00A0]+/g, ' ').trim().replace(/[.,;:!?*()[\]{}'"`~]/g, " ").replace(/\s+/g, ' ').trim();
       Logger.log('---DEBUG PROCESSED NEW REPLY BODY FOR LOOP (Inv# ' + invNumFromSheetRaw + ', Msg '+(k+1)+')---\n[' + bodyToTestInLoop.substring(0,200) + '...]\n---END PROCESSED BODY---');

       const invNumInReplyRaw = findInvoiceNumber(newReplyTextOnly, fromSubj);
       const cleanInvNumInReply = invNumInReplyRaw ? String(invNumInReplyRaw).trim().replace(/^[.,;:!?()\s]+|[.,;:!?()\s]+$/g, "").toLowerCase() : null;
       const cleanInvNumFromSheet = invNumFromSheetRaw ? String(invNumFromSheetRaw).trim().replace(/^[.,;:!?()\s]+|[.,;:!?()\s]+$/g, "").toLowerCase() : null;

       if (cleanInvNumInReply && cleanInvNumFromSheet && cleanInvNumFromSheet !== "n/a_sheet_inv_num" && cleanInvNumInReply !== cleanInvNumFromSheet) {
           Logger.log(`      WARNING: Cleaned Invoice# in reply ("${cleanInvNumInReply}" from raw "${invNumInReplyRaw}") does NOT match cleaned Invoice# for this thread ("${cleanInvNumFromSheet}" from raw "${invNumFromSheetRaw}"). Skipping payment processing for this message.`);
           const parsingNotesCell = sheet.getRange(actualRow, parsingNotesSheetIdx + 1); let currentParsingNotes = String(parsingNotesCell.getValue() || "").trim();
           const mismatchNote = `Reply (raw inv# "${invNumInReplyRaw}") mentions different Inv# than thread's Inv# ("${invNumFromSheetRaw}"). Payment not processed from this msg.`;
           if (!currentParsingNotes.includes("Reply mentions different Inv#")) parsingNotesCell.setValue((currentParsingNotes ? currentParsingNotes + "; " : "") + mismatchNote);
           if (reviewLabel && !thread.getLabels().some(label => label.getName() === REVIEW_LABEL_NAME)) { try { thread.addLabel(reviewLabel); Logger.log(`      Applied '${REVIEW_LABEL_NAME}' due to Inv# mismatch.`); } catch (rlErr) {}}
           continue;
       } else if (cleanInvNumInReply) { Logger.log(`      Invoice# in reply ("${cleanInvNumInReply}" from raw "${invNumInReplyRaw}") MATCHES sheet Inv# ("${cleanInvNumFromSheet}" from raw "${invNumFromSheetRaw}") or sheet Inv# is N/A. Proceeding.`);
       } else { Logger.log(`      No distinct Invoice# found in reply text. Proceeding with keyword check on thread context.`); }

       let keywordFoundInMessage = false; let matchedKeyword = "";
       for (const keyword of paymentKeywords) {
         if (new RegExp(`\\b${keyword}\\b`, 'i').test(bodyToTestInLoop)) {
           keywordFoundInMessage = true; matchedKeyword = keyword;
           Logger.log(`      Iterative Keyword Matched: '${matchedKeyword}' in NEW reply body for Inv# ${invNumFromSheetRaw}`); break;
         }
       }

       if (keywordFoundInMessage) {
         processedAReplyInThisRun = true; paymentReplyProcessedForThisThread = true;
         let newStatus = "Pending Confirmation";
         Logger.log(`      KEYWORDS FOUND ('${matchedKeyword}') for Inv# ${invNumFromSheetRaw}. Processing sheet updates.`);
         sheet.getRange(actualRow, replyPaidDateIdx + 1).setValue(fromDate).setNumberFormat("yyyy-MM-dd HH:mm:ss");
         const amountDetails = findAmountInPaymentReply(originalBody, invoiceCurrency);
         if (amountDetails.amount !== null) {
           sheet.getRange(actualRow, replyAmountIdx + 1).setValue(amountDetails.amount).setNumberFormat("0.00##");
           sheet.getRange(actualRow, replyCurrencyIdx + 1).setValue(amountDetails.currency);
           Logger.log(`      Logged replied amount: ${amountDetails.currency} ${amountDetails.amount}`);
           if (!isNaN(originalAmountDue) && amountDetails.amount < originalAmountDue && amountDetails.amount > 0) { newStatus = "Partially Paid"; Logger.log(`      Amount (${amountDetails.amount}) < due (${originalAmountDue}). Status "Partially Paid".`); }
           else if (!isNaN(originalAmountDue) && amountDetails.amount > originalAmountDue) { newStatus = "Paid"; Logger.log(`      Amount (${amountDetails.amount}) > due (${originalAmountDue}). Status "Paid" (OVERPAYMENT).`);
               const parsingNotesCell = sheet.getRange(actualRow, parsingNotesSheetIdx + 1); let currentParsingNotes = String(parsingNotesCell.getValue() || "").trim();
               const overpaymentNote = `Potential Overpayment: Replied ${amountDetails.currency} ${amountDetails.amount} vs Due ${rowData[invCurrencyIdx]} ${originalAmountDue}.`;
               if(!currentParsingNotes.includes("Potential Overpayment")) parsingNotesCell.setValue((currentParsingNotes ? currentParsingNotes + "; " : "") + overpaymentNote);
           }
           else if (!isNaN(originalAmountDue) && amountDetails.amount === originalAmountDue){ newStatus = "Paid"; Logger.log(`      Amount (${amountDetails.amount}) matches due. Status "Paid".`);}
           else { newStatus = "Pending Confirmation"; Logger.log(`      Amount ${amountDetails.amount} vs Due (${originalAmountDue}) inconclusive. Status "Pending Confirmation".`); }
         } else {
           Logger.log(`      No amount in reply for Inv# ${invNumFromSheetRaw}, keywords present. Status "Pending Confirmation".`);
           sheet.getRange(actualRow, replyAmountIdx + 1).clearContent(); sheet.getRange(actualRow, replyCurrencyIdx + 1).clearContent();
         }
         sheet.getRange(actualRow, statusIdx + 1).setValue(newStatus);
         if (newStatus === 'Paid' || newStatus === 'Partially Paid' || newStatus === 'Pending Confirmation') {
            const reminderSentCell = sheet.getRange(actualRow, reminderSentIdx + 1);
            if (String(reminderSentCell.getValue() || "").trim() !== "") { reminderSentCell.clearContent(); Logger.log(`      Cleared Reminder Stage for Inv# ${invNumFromSheetRaw} due to status '${newStatus}'.`); }
         }
         const notesCell = sheet.getRange(actualRow, notesIdx + 1); const currentNotes = notesCell.getValue() || "";
         const notePrefix = `Reply from ${fromFull} (${Utilities.formatDate(fromDate, SCRIPT_TIMEZONE, "yyyy-MM-dd HH:mm")}): `;
         const snippet = `"${newReplyTextOnly.substring(0, 150).replace(/\n|\r/g," ").trim()}..."`;
         const newNote = notePrefix + snippet;
         if (!currentNotes.includes(snippet.substring(0,50))) notesCell.setValue((currentNotes ? currentNotes + "\n---\n" : "") + newNote);
         Logger.log(`      Updated sheet for Inv# ${invNumFromSheetRaw} (Row ${actualRow}) to "${newStatus}", Reply Date: ${formatDate(fromDate)}.`);
         if (reviewLabel && thread.getLabels().some(label => label.getName() === REVIEW_LABEL_NAME)) { try { thread.removeLabel(reviewLabel); Logger.log(`      Removed '${REVIEW_LABEL_NAME}'.`); } catch (rlErr) {}}
         if (replyProcessedLabel && !thread.getLabels().some(label => label.getName() === LABEL_PAYMENT_REPLY_PROCESSED)) { try { thread.addLabel(replyProcessedLabel); Logger.log(`      Applied '${LABEL_PAYMENT_REPLY_PROCESSED}'.`); } catch (alErr) {}}
       } else { Logger.log(`      No payment keywords found in NEW reply body for Inv# ${invNumFromSheetRaw}.`); }
     }
     if (!processedAReplyInThisRun && (originalStatusOnSheet === 'Unpaid' || originalStatusOnSheet === 'Partially Paid')) {
       const dueDateString = rowData[dueDateSheetIdx];
       if (dueDateString) {
           const dueDateEval = parseDateString(String(dueDateString));
           if (dueDateEval) {
               const todayEval = new Date(); dueDateEval.setHours(0,0,0,0); todayEval.setHours(0,0,0,0);
               if (dueDateEval < todayEval) { sheet.getRange(actualRow, statusIdx + 1).setValue("Overdue"); Logger.log(`      Inv# ${invNumFromSheetRaw} (Row ${actualRow}) status (was ${originalStatusOnSheet}) updated to "Overdue" (Courtesy Check).`); }
           } else { Logger.log(`      Inv# ${invNumFromSheetRaw} (Row ${actualRow}): Could not parse Due Date '${dueDateString}' for courtesy check.`); }
       }
     }
   } catch (e) {
       Logger.log(`E processing replies for thread ${currentGmailThreadId} (Inv# ${invNumFromSheetRaw}): ${e.message}\nStack: ${e.stack}`);
       if (foundMatchInSheet && parsingNotesSheetIdx > -1 && actualRow <= sheet.getLastRow()) {
            try {
               const parsingNotesCell = sheet.getRange(actualRow, parsingNotesSheetIdx + 1); let currentParsingNotes = String(parsingNotesCell.getValue() || "").trim();
               const errorNote = `SCRIPT ERROR: ${e.message}`;
               if (!currentParsingNotes.includes("SCRIPT ERROR")) parsingNotesCell.setValue((currentParsingNotes ? currentParsingNotes + "; " : "") + errorNote);
           } catch (sheetErr) {}
       }
   }
 });
 Logger.log("processPaymentReplies finished.");
}


// =============== HELPER: PARSE AMOUNT FROM PAYMENT REPLY TEXT ===============
// V.0.7.6c - Attempt to isolate new reply text before parsing amount
function findAmountInPaymentReply(replyBodyText, invoiceCurrencyHint) {
 Logger.log(`findAmountInPaymentReply: Hinted currency: ${invoiceCurrencyHint}`);
 let newReplyText = replyBodyText;
 const replySeparators = [ /\nOn .* wrote:\s*\n/i, /\n----- Original Message -----\n/i, /\nFrom: .* <.*@.*>\s*\nSent: .*\nTo: .*\nSubject: .*/i, /\n>? ?From: .*\n>? ?Date: .*\n>? ?Subject: .*\n>? ?To: .*/i, /^\s*>+.*\n?/gm ];
 let separatorFound = false;
 for (const separator of replySeparators) {
    const match = replyBodyText.match(separator);
    if (match) { newReplyText = replyBodyText.substring(0, match.index).trim(); separatorFound = true; Logger.log(`findAmountInPaymentReply: Separator found (${separator.source.substring(0,30)}...). Using text before it.`); break; }
 }
 if (!separatorFound) { Logger.log(`findAmountInPaymentReply: No common reply separator. Searching full body.`); }
 else if (newReplyText.length < 10) { Logger.log(`findAmountInPaymentReply: Text before separator too short. Reverting to search full body.`); newReplyText = replyBodyText; }
 Logger.log(`findAmountInPaymentReply: Searching within text (first 300): ${newReplyText.substring(0,300).replace(/\n/g," ")}`);

 let amount = null; let currency = null;
 const amountPatterns = [ /(?:(\d{1,3}(?:[,.]\d{3})*[,.]\d{2}|\d+[.,]\d{2}|\d+))\s*(USD|EUR|GBP|CAD|AUD|JPY|CHF|CNY|INR|NZD|ZAR)\b/i, /\b(USD|EUR|GBP|CAD|AUD|JPY|CHF|CNY|INR|NZD|ZAR)\s*(\d{1,3}(?:[,.]\d{3})*[,.]\d{2}|\d+[.,]\d{2}|\d+)/i, /(?:[\$€£¥])\s*(\d{1,3}(?:[,.]\d{3})*[,.]\d{2}|\d+[.,]\d{2}|\d+)/i, ];
 const currencySymbols = {'$':'USD', '€':'EUR', '£':'GBP', '¥':'JPY'};
 for (const pattern of amountPatterns) {
   const match = newReplyText.match(pattern);
   if (match) {
     let rawAmountStr = ""; let foundCurrencySymbolOrCode = "";
     if (pattern.source.includes("\\b(USD|EUR")) { foundCurrencySymbolOrCode = match[1].toUpperCase(); rawAmountStr = match[2]; currency = foundCurrencySymbolOrCode; Logger.log(`findAmountInPaymentReply: Matched (Code First): AmountStr="${rawAmountStr}", Currency="${currency}"`); }
     else if (pattern.source.includes("(USD|EUR|GBP)\\b")) { rawAmountStr = match[1]; foundCurrencySymbolOrCode = match[2].toUpperCase(); currency = foundCurrencySymbolOrCode; Logger.log(`findAmountInPaymentReply: Matched (Amount First): AmountStr="${rawAmountStr}", Currency="${currency}"`); }
     else if (pattern.source.startsWith("(?:[\\$€£¥])")) { rawAmountStr = match[1]; const symbolIndex = match.index;
         if (symbolIndex >= 0) { const charBefore = symbolIndex > 0 ? newReplyText[symbolIndex -1] : ''; const potentialSymbol = charBefore.match(/[\$€£¥]/) ? charBefore : newReplyText[symbolIndex]; foundCurrencySymbolOrCode = potentialSymbol; currency = currencySymbols[foundCurrencySymbolOrCode] || invoiceCurrencyHint || DEFAULT_CURRENCY; Logger.log(`findAmountInPaymentReply: Matched (Symbol): AmountStr="${rawAmountStr}", Symbol="${foundCurrencySymbolOrCode}", Currency="${currency}"`); }
         else { continue; }
     }
     const cleanedAmount = cleanCurrencyValue(rawAmountStr);
     if (cleanedAmount !== null && cleanedAmount > 0) { amount = cleanedAmount; Logger.log(`findAmountInPaymentReply: Successfully parsed: Amount=${amount}, Currency=${currency}`); break; }
     else { Logger.log(`findAmountInPaymentReply: cleanCurrencyValue null/zero/negative for "${rawAmountStr}". Continuing.`); amount = null; currency = null; }
   }
 }
 if (amount === null) { Logger.log("findAmountInPaymentReply: No definitive amount in NEW reply text."); }
 return { amount: amount, currency: currency };
}


// =============== FUNCTION: CHECK & UPDATE OVERDUE STATUSES ===============
function checkAndUpdateOverdueStatuses() {
  Logger.log("Starting checkAndUpdateOverdueStatuses run (V.0.7.6h)...");
  if (!SPREADSHEET_ID || SPREADSHEET_ID.startsWith('YOUR_')) { Logger.log("E: SPREADSHEET_ID not configured."); return; }
  let spreadsheet, sheet;
  try { spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); sheet = spreadsheet.getSheetByName(INVOICE_LOG_SHEET_NAME); if (!sheet) throw new Error("Sheet missing: " + INVOICE_LOG_SHEET_NAME); }
  catch (e) { Logger.log(`E opening sheet: ${e.message}. Run initialSetup.`); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusIdx = headers.indexOf(COL_STATUS); const dueDateIdx = headers.indexOf(COL_DUE_DATE); const invNumIdx = headers.indexOf(COL_INVOICE_NUM);
  if (statusIdx === -1 || dueDateIdx === -1 || invNumIdx === -1) { Logger.log(`E: Required columns missing - Status (${statusIdx}), Due Date (${dueDateIdx}), Invoice # (${invNumIdx}).`); return; }
  const lastRow = sheet.getLastRow(); if (lastRow < 2) { Logger.log("No data rows for overdue check."); return; }
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const allData = dataRange.getValues(); const statusesToUpdate = allData.map(row => [row[statusIdx]]);
  const today = new Date(); today.setHours(0, 0, 0, 0); let updatesMade = 0;
  for (let i = 0; i < allData.length; i++) {
    const currentRow = i + 2; const currentStatus = allData[i][statusIdx]; const dueDateValue = allData[i][dueDateIdx]; const invNum = allData[i][invNumIdx] || `Row ${currentRow}`;
    if (currentStatus === 'Unpaid' || currentStatus === 'Partially Paid') {
        const dueDate = parseDateString(dueDateValue instanceof Date ? formatDate(dueDateValue) : String(dueDateValue));
        if (dueDate) { dueDate.setHours(0, 0, 0, 0);
            if (dueDate < today) { if (currentStatus !== 'Overdue') { statusesToUpdate[i][0] = 'Overdue'; Logger.log(`Inv# ${invNum} (Row ${currentRow}): Status changing '${currentStatus}' to 'Overdue' (Due: ${formatDate(dueDate, 'yyyy-MM-dd')})`); updatesMade++; }}
        } else if (dueDateValue) { Logger.log(`W: Inv# ${invNum} (Row ${currentRow}): Unparseable Due Date ('${dueDateValue}') for overdue check.`);}
    }
  }
  if (updatesMade > 0) { sheet.getRange(2, statusIdx + 1, statusesToUpdate.length, 1).setValues(statusesToUpdate); Logger.log(`Updated status to 'Overdue' for ${updatesMade} invoice(s).`); }
  else { Logger.log("No status updates needed for overdue invoices."); }
  Logger.log("checkAndUpdateOverdueStatuses finished.");
}


// =============== FUNCTION: SEND OVERDUE REMINDERS ===============
function sendOverdueReminders() {
  Logger.log("Starting sendOverdueReminders run (V.0.7.6h)...");
  if (!SPREADSHEET_ID || SPREADSHEET_ID.startsWith('YOUR_')) { Logger.log("E: SPREADSHEET_ID not configured."); return; }
  if (!MY_PRIMARY_INVOICING_EMAIL || !YOUR_COMPANY_NAME || YOUR_COMPANY_NAME === "Your Company Name Here" || !YOUR_SUPPORT_EMAIL_OR_CONTACT_INFO || YOUR_SUPPORT_EMAIL_OR_CONTACT_INFO === "billing@example.com") {
      Logger.log("E: MY_PRIMARY_INVOICING_EMAIL, YOUR_COMPANY_NAME, or YOUR_SUPPORT_EMAIL_OR_CONTACT_INFO not configured."); return;
  }
  let spreadsheet, sheet, errorLogSheet;
  try { spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); sheet = spreadsheet.getSheetByName(INVOICE_LOG_SHEET_NAME); errorLogSheet = spreadsheet.getSheetByName(ERROR_LOG_SHEET_NAME); if (!sheet || !errorLogSheet) throw new Error("Log sheet(s) missing.");}
  catch (e) { Logger.log(`E opening sheet: ${e.message}. Run initialSetup.`); return; }

  const sortedTiers = [...REMINDER_TIERS].sort((a, b) => a.daysOverdue - b.daysOverdue);
  const tierStagesInOrder = sortedTiers.map(t => t.stage);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = { status: headers.indexOf(COL_STATUS), dueDate: headers.indexOf(COL_DUE_DATE), reminderSent: headers.indexOf(COL_REMINDER_STAGE_SENT), invNum: headers.indexOf(COL_INVOICE_NUM), clientEmail: headers.indexOf(COL_CLIENT_EMAIL), clientName: headers.indexOf(COL_CLIENT_NAME), amountDue: headers.indexOf(COL_AMOUNT_DUE), currency: headers.indexOf(COL_CURRENCY), gmailLink: headers.indexOf(COL_LINK_GMAIL) };
  if (Object.values(idx).includes(-1)) { Logger.log(`E: One or more required columns missing for reminders. Check COL_ constants and sheet headers.`); return; }

  const lastRow = sheet.getLastRow(); if (lastRow < 2) { Logger.log("No data rows for reminders."); return; }
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const allData = dataRange.getValues(); const reminderStagesToUpdate = allData.map(row => [row[idx.reminderSent]]);
  const today = new Date(); today.setHours(0, 0, 0, 0); let emailsSentThisRun = 0; const initialEmailQuota = MailApp.getRemainingDailyQuota();
  Logger.log(`Remaining daily email quota: ${initialEmailQuota}`); if (initialEmailQuota < 1) { Logger.log("W: Email quota reached."); return; }

  for (let i = 0; i < allData.length; i++) {
    const currentRow = i + 2; const currentStatus = allData[i][idx.status]; const lastReminderSent = (allData[i][idx.reminderSent] || "").toString().trim();
    const dueDateValue = allData[i][idx.dueDate]; const invNum = allData[i][idx.invNum] || `Row ${currentRow}`;
    let clientEmailRaw = allData[i][idx.clientEmail]; let clientEmail = null;
    if (clientEmailRaw && typeof clientEmailRaw === 'string') { const firstEmail = clientEmailRaw.split(',')[0].trim(); const emailMatch = firstEmail.match(/<([^>]+)>/); clientEmail = emailMatch ? emailMatch[1].trim() : firstEmail.trim(); if (!clientEmail.includes('@')) clientEmail = null; }
    if (currentStatus !== 'Overdue') continue;
    const finalTierStage = tierStagesInOrder.length > 0 ? tierStagesInOrder[tierStagesInOrder.length - 1] : null;
    if (finalTierStage && lastReminderSent === finalTierStage) continue;
    if (!clientEmail) { Logger.log(`W: Inv# ${invNum} (Row ${currentRow}): Skipping - missing/invalid client email ('${clientEmailRaw}').`); continue; }
    const dueDate = parseDateString(dueDateValue instanceof Date ? formatDate(dueDateValue) : String(dueDateValue));
    if (!dueDate) { Logger.log(`W: Inv# ${invNum} (Row ${currentRow}): Skipping - missing/invalid Due Date ('${dueDateValue}').`); continue; }
    dueDate.setHours(0, 0, 0, 0); const daysOverdue = Math.floor((today.getTime() - dueDate.getTime()) / (1000 * 60 * 60 * 24));
    if (daysOverdue < 0) continue;
    let tierToSend = null; const lastSentIndex = tierStagesInOrder.indexOf(lastReminderSent);
    for (let j = 0; j < sortedTiers.length; j++) { const tier = sortedTiers[j]; if (daysOverdue >= tier.daysOverdue && j > lastSentIndex) { tierToSend = tier; break; }}
    if (tierToSend) {
        if (emailsSentThisRun >= initialEmailQuota) { Logger.log("W: Email quota hit during run. Stopping."); break; }
        const clientName = allData[i][idx.clientName] || "Valued Customer"; const amountDueRaw = allData[i][idx.amountDue];
        const amountDue = (amountDueRaw !== null && amountDueRaw !== "" && !isNaN(parseFloat(amountDueRaw))) ? parseFloat(amountDueRaw).toFixed(2) : "N/A";
        const currency = allData[i][idx.currency] || ""; const dueDateFormatted = formatDate(dueDate, 'yyyy-MM-dd');
        let gmailLinkValue = allData[i][idx.gmailLink]; let gmailLinkForError = '';
        if (gmailLinkValue && typeof gmailLinkValue === 'string' && gmailLinkValue.toUpperCase().startsWith('=HYPERLINK("')) { const formulaMatch = gmailLinkValue.match(/#inbox\/([^"]+)"/); if (formulaMatch && formulaMatch[1]) gmailLinkForError = formulaMatch[1].trim(); else gmailLinkForError = gmailLinkValue; } else { gmailLinkForError = gmailLinkValue; }
        let subject = tierToSend.subject.replace(/\[Client Name\]/g, clientName).replace(/\[Invoice #\]/g, invNum).replace(/\[Currency\]/g, currency).replace(/\[Amount Due\]/g, amountDue).replace(/\[Due Date\]/g, dueDateFormatted).replace(/\[Your Company Name\]/g, YOUR_COMPANY_NAME);
        let body = tierToSend.bodyTemplate.replace(/\[Client Name\]/g, clientName).replace(/\[Invoice #\]/g, invNum).replace(/\[Currency\]/g, currency).replace(/\[Amount Due\]/g, amountDue).replace(/\[Due Date\]/g, dueDateFormatted).replace(/\[Your Company Name\]/g, YOUR_COMPANY_NAME).replace(/\[Your Support Email or Preferred Contact Method\]/g, YOUR_SUPPORT_EMAIL_OR_CONTACT_INFO);
        try {
            Logger.log(`Sending ${tierToSend.stage} for Inv# ${invNum} (${daysOverdue} days overdue) to ${clientEmail}. LastSent: '${lastReminderSent}'`);
            MailApp.sendEmail({ to: clientEmail, replyTo: MY_PRIMARY_INVOICING_EMAIL, subject: subject, body: body, name: YOUR_COMPANY_NAME });
            reminderStagesToUpdate[i][0] = tierToSend.stage; emailsSentThisRun++;
            Logger.log(`--> SUCCESS: Sent ${tierToSend.stage} for Inv# ${invNum}. Total sent: ${emailsSentThisRun}.`);
        } catch (e) {
            Logger.log(`E: Sending ${tierToSend.stage} for Inv# ${invNum} to ${clientEmail}. Error: ${e.message}`);
            logErrorToSheetAdvanced(errorLogSheet, { subject: `Reminder Fail: ${tierToSend.stage} for Inv# ${invNum}`, date: new Date(), recipients: clientEmail, functionName: 'sendOverdueReminders', errorMessage: `MailApp Error: ${e.message}. Quota Left: ${MailApp.getRemainingDailyQuota()}`, threadLink: gmailLinkForError, rawBodySnippet: `Days Overdue: ${daysOverdue}, Last Sent: ${lastReminderSent}` });
        }
    }
  }
  if (emailsSentThisRun > 0) { sheet.getRange(2, idx.reminderSent + 1, reminderStagesToUpdate.length, 1).setValues(reminderStagesToUpdate); Logger.log(`Updated reminder stages. Total emails sent: ${emailsSentThisRun}.`); }
  else { Logger.log("No reminder emails sent/stages updated."); }
  Logger.log(`sendOverdueReminders finished. Email quota left: ${MailApp.getRemainingDailyQuota()}`);
}


// =============== ALL OTHER HELPER FUNCTIONS ===============
// =============== FUNCTION: SETUP TRIGGERS =============== // Corrected placement from before
function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const functionsToTrigger = {
    'processSentInvoices': { type: 'hours', value: 4 },
    'processPaymentReplies': { type: 'hours', value: 1 },
    'checkAndUpdateOverdueStatuses': { type: 'days', value: 1, hour: 2 },
    'sendOverdueReminders': { type: 'days', value: 1, hour: 9 }
  };

  let triggersConfiguredCount = 0;
  Logger.log("Checking and configuring project triggers...");

  for (const functionName in functionsToTrigger) {
    let triggerExists = false;
    // Check if a trigger for this function already exists
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === functionName) {
        // Optional: Delete existing to always apply new schedule/settings
        // ScriptApp.deleteTrigger(triggers[i]);
        // Logger.log(`Deleted existing trigger for ${functionName} to ensure latest settings.`);
        // triggerExists = false; // Uncomment if deleting
        triggerExists = true; // Keep if just verifying
        Logger.log(`Trigger for ${functionName} already exists. Schedule may or may not be current. To force update, manually delete trigger and re-run setup.`);
        break; 
      }
    }

    if (!triggerExists) {
      const config = functionsToTrigger[functionName];
      let newTriggerBuilder = ScriptApp.newTrigger(functionName).timeBased();

      if (config.type === 'hours') {
        newTriggerBuilder.everyHours(config.value);
        Logger.log(`Creating trigger for ${functionName} to run every ${config.value} hour(s).`);
      } else if (config.type === 'days') {
        newTriggerBuilder.everyDays(config.value).atHour(config.hour);
        Logger.log(`Creating trigger for ${functionName} to run daily around ${config.hour}:00-${config.hour+1}:00.`);
      }
      
      // Add randomization to minute to avoid all scripts firing at exact top of hour
      if (config.type === 'days' && config.hour != undefined) {
          newTriggerBuilder.nearMinute(Math.floor(Math.random() * 30) + 15); // e.g., for atHour(2), 2:15-2:45
      } else if (config.type === 'hours') {
          newTriggerBuilder.nearMinute(Math.floor(Math.random() * 50) + 5); // Randomize minute (5-55)
      }

      try {
          newTriggerBuilder.create();
          Logger.log(`Successfully created trigger for ${functionName}.`);
          triggersConfiguredCount++;
      } catch (e) {
          Logger.log(`E: Could not create trigger for ${functionName}. Error: ${e.message}. You may need to manually authorize or create it via Edit > Current project's triggers.`);
      }
    } else {
        triggersConfiguredCount++; // Still counts as "configured" if it exists
    }
  } // End for...in loop

  if(triggersConfiguredCount === Object.keys(functionsToTrigger).length && triggers.length >= Object.keys(functionsToTrigger).length){
    Logger.log("All expected triggers appear to be configured.");
  } else {
    Logger.log(`W: Triggers check complete. Expected: ${Object.keys(functionsToTrigger).length}, Found/Verified: ${triggers.length} existing, ${triggersConfiguredCount} processed. Some triggers might need manual review via 'Edit > Current project's triggers'.`);
  }
  Logger.log("Trigger setup process completed.");
}

function getOrCreateLabel(labelName) {
 if(!labelName||typeof labelName!=='string'){Logger.log(`E: Invalid label name: ${labelName}.`);return null}
 let label = GmailApp.getUserLabelByName(labelName);
 if(!label){ try{ label = GmailApp.createLabel(labelName); Logger.log(`Label "${labelName}" created.`); }
 catch(e){ Logger.log(`W: Create label "${labelName}" fail: ${e.message}. Retrying fetch.`); Utilities.sleep(500); try{ label = GmailApp.getUserLabelByName(labelName); } catch(e2){ Logger.log(`E: Get/Create label "${labelName}" retry fail: ${e2.message}`); label=null; }}}
 return label;
}

function parseInvoiceDetailsFromEmail(emailBody, emailSubject, attachments, recipientsHeader, isDebugTargetThread) {
 if (isDebugTargetThread) Logger.log(`DEBUG (parseInvoiceDetailsFromEmail): Body(500):\n${emailBody ? emailBody.substring(0,500) : "NULL"}`);
 const clientName = findClientName(emailBody, emailSubject, recipientsHeader);
 const invoiceNumber = findInvoiceNumber(emailBody, emailSubject);
 const currencyResult = findCurrency(emailBody);
 const grossTotal = findGrossTotalAmount(emailBody, isDebugTargetThread);
 const discount = findDiscountAmount(emailBody, isDebugTargetThread);
 const invoiceDate = findInvoiceDate(emailBody);
 const dueDate = findDueDate(emailBody);
 const streetAddress = findStreetAddress(emailBody);
 const finalCurrency = currencyResult || (grossTotal !== null || discount !== 0 ? DEFAULT_CURRENCY : null);
 let parsingNotes = [];
 if (clientName === "Client Not Parsed" && (invoiceNumber || grossTotal !== null )) parsingNotes.push("Client Name NP.");
 if (!invoiceNumber && (emailSubject.toLowerCase().includes(INVOICE_SUBJECT_KEYWORD) || (emailBody && emailBody.toLowerCase().includes(INVOICE_SUBJECT_KEYWORD)))) parsingNotes.push("Invoice # NP.");
 if (grossTotal === null && invoiceNumber) parsingNotes.push("Gross Total NP.");
 if (discount === 0 && emailBody && emailBody.match(/discount|credit|less|saving/i) && grossTotal !== null && !emailBody.match(/(?:discount|credit|less|saving)[\s:]*0(?:\.00)?\b/i) ) parsingNotes.push("Discount/Credit keyword present but 0 parsed.");
 if (!invoiceDate && invoiceNumber) parsingNotes.push("Inv Date NP."); if (!dueDate && invoiceNumber) parsingNotes.push("Due Date NP.");
 if (!currencyResult && (grossTotal !== null || discount !== 0 )) parsingNotes.push(`Currency NP, default: ${DEFAULT_CURRENCY}.`);
 return { clientName, invoiceNumber, grossTotal, discount, invoiceDate, dueDate, streetAddress, currency: finalCurrency, parsingNotes: parsingNotes.join('; ') };
}

function findClientName(body, subject, recipientsHeader) {
  if (!body) { Logger.log("findClientName: No body."); return "Client Not Parsed"; }
  const sameLinePatterns = [ /(?:Client Name|CUSTOMER NAME|Bill To|ACCOUNT|Client|To|Recipient)\s*[:]\s*(.+)/i ];
  for (const pattern of sameLinePatterns) { const match = body.match(pattern); if (match && match[1]) { let clientName = cleanAndFilterClientName(match[1].trim()); if (clientName) { Logger.log(`findClientName (Same Line): Pat "${pattern.source.substring(0,30)}..." -> "${clientName}"`); return clientName; }}}
  const lines = body.split(/\r?\n/); const labels = ["Client Name", "CUSTOMER NAME", "Bill To", "ACCOUNT", "Client", "To", "Recipient"];
  for (let i = 0; i < lines.length - 1; i++) { const currentLine = lines[i].trim(); for (const label of labels) { const labelPattern = new RegExp(`^${label}\\s*[:]?\\s*$`, "i"); if (labelPattern.test(currentLine)) { let nextLineIndex = i + 1; while(nextLineIndex < lines.length && lines[nextLineIndex].trim() === '') nextLineIndex++; if (nextLineIndex < lines.length) { let clientName = cleanAndFilterClientName(lines[nextLineIndex].trim()); if (clientName) { Logger.log(`findClientName (Next Line): Label "${label}" -> "${clientName}"`); return clientName; } break; }}}}
  if (recipientsHeader) { try { const recipientsArray = String(recipientsHeader).split(','); for (let recipient of recipientsArray) { const recipientName = parseRecipientName(recipient.trim()); if (recipientName && !/^(support|billing|info|admin|sales|accounts|no.?reply|undisclosed|group|team)/i.test(recipientName)){ Logger.log(`findClientName (Fallback To:): "${recipientName}"`); return recipientName.replace(/^['"]+|['"]+$/g, ''); }}} catch(e) { Logger.log(`findClientName Fallback Error: ${e.message}`); }}
  return "Client Not Parsed";
}

function cleanAndFilterClientName(clientNameRaw) {
  if (!clientNameRaw) return null; let clientName = clientNameRaw.trim();
  clientName = clientName.split(/\n|\r|\s{3,}| Attn:| Attention:| Customer ID:| Tel:| Fax:| Email:| Website:/i)[0].trim();
  clientName = clientName.replace(/(?:\s+(?:Street|St|Ave|Rd|Dr|Ln|Ct|Plz|Sq|Ter|Pl|Blvd|Pkwy|Suite|Ste|Apt|Unit|Floor|Fl|,?\s*(?:PO|P\.O\.) Box|\d{5,}|Total|Amount|Invoice|Due|Date|Account No|Customer ID|Tax ID|VAT ID|Ref|Payment|Services))+.*/i, '').trim();
  clientName = clientName.replace(/,\s*(?:Esq\.?|Inc\.?|Ltd\.?|LLC|Corp\.?|Co\.?)$/i, '').trim();
  clientName = clientName.replace(/^[^a-zA-Z0-9\(\)&\-'.]+|[^a-zA-Z0-9\(\)&\-'.\s]+$/g, "").trim();
  if (clientName && clientName.length >= 3 && clientName.length < 70) { const firstWord = clientName.toLowerCase().split(/[\s-(]/)[0]; const genericTerms = /^(support|billing|info|admin|sales|team|dept|group|office|customer|attn|total|amount|invoice|date|due|subject|ref|no|id|balance|payable|subtotal|dear|hi|hello|greetings|thank|best|regards|francis|mycompany|the|company|services|project|for|from)$/i; const nonNamePatterns = /\d{4,}|^-?INV-\d+|^\w{3,}-\w+-|Account Summary|Invoice Details|Date Generated|Services Rendered|http:|https:|\.com|\.org|\.net|@|<|>|was issued on|schedule a call|click here|please find/i;
      if (!genericTerms.test(firstWord) && !nonNamePatterns.test(clientName) && clientName.replace(/[^a-zA-Z]/g, "").length > 1) return clientName; }
  return null;
}

function findInvoiceNumber(body, subject) {
  const bodyPatterns = [ /Invoice\s*#\s*:?\s*((?:INV-)?[\p{L}\p{N}.-]+)/iu,/REFERENCE\s+NUMBER\s*[:#]?\s*([\p{L}\p{N}.-]{3,})/iu,/Invoice\s+No\.?[:\s#]*((?:INV-)?[\p{L}\p{N}.-]+)/iu,/Invoice\s+Number[:\s#]*((?:INV-)?[\p{L}\p{N}.-]+)/iu,/Invoice\s+ID[:\s#]*((?:INV-)?[\p{L}\p{N}.-]+)/iu,/Invoice\s*:\s*((?:INV-)?[\p{L}\p{N}.-]+)/iu,/(?:Invoice\s)?Statement\s+No\.?\s*([\p{L}\p{N}.-]{3,})/iu,/(?:The\s+project\s+reference\s+is|Project\s+Ref(?:erence|\.)?|Our\s+Ref(?:erence|\.)?|Invoice\s+Ref(?:erence|\.)?)\s+([\p{L}\p{N}][\p{L}\p{N}.-]*[\p{L}\p{N}]{1,})/iu,/\bRef(?:erence|\.)?(?:\s*No\.?)?\s*[:#]?\s*([\p{L}\p{N}][\p{L}\p{N}.-]*[\p{L}\p{N}]{1,})/iu,/\bRef #\s*([\p{L}\p{N}.-]{3,})/iu,/(?<!Order\s|Quote\s|Account\s|Phone\s|Fax\s|Date\s|Due\s)No\.?[:\s#]*((?=.*[\p{N}-])[\p{L}\p{N}.-]{4,})/iu ];
  for (const pattern of bodyPatterns) { try { const match = body.match(pattern); if (match && match[1]) { const invNum = match[1].trim(); if (invNum && !/^\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}$/.test(invNum) && !/^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/.test(invNum) && invNum.length >=3 && !/^(Date|Due|Total|Amount)$/i.test(invNum)) { Logger.log(`findInvoiceNumber (Body): Pat "${pattern.source.substring(0,30)}..." -> "${invNum}"`); return invNum; }}} catch (e) {}}
  const subjectPatterns = [ /(?:Invoice|Facture|INV)\s*#?\s*:?\s*((?:INV-?)?[\p{L}\p{N}.-]{3,})/iu, /\b(INV-[\p{L}\p{N}.-]{3,})\b/iu ];
  for (const pattern of subjectPatterns) { try { const match = subject.match(pattern); if (match && match[1]) { const potentialNum = match[1].trim(); if (potentialNum.toLowerCase()!=='invoice'&&potentialNum.toLowerCase()!=='inv'&&potentialNum.length>=3&&/\d/.test(potentialNum)&&!/^(RE|FW|FWD)$/i.test(potentialNum)) { Logger.log(`findInvoiceNumber (Subject): Pat "${pattern.source.substring(0,30)}..." -> "${potentialNum}"`); return potentialNum; }}} catch (e) {}}
  Logger.log(`findInvoiceNumber V7: No invoice number found.`); return null;
}

// V.0.7.6h - Refining general pattern and Sac for negative amounts
function findDiscountAmount(body, isDebugTargetThread) {
  if(!body)return 0;
  const A_PATTERN = "((?:\\d{1,3}(?:[,.]\\d{3})+[,.]\\d{1,2})|(?:\\d+[,.]\\d{1,2})|(?:\\d{1,3}(?:[,.]\\d{3})+)|(?:\\d+))";
  const fs = "[\\s.:*|_\\-]*"; 
  const Sac_Pattern_Detailed =  "(?:([\\$€£¥])?(\\s*)(\\-)?(\\s*)" + A_PATTERN + ")";
  const dls=[ {label: "Discount Amount", type: "general"}, {label: "Discount Applied", type: "general"}, {label: "Credits Applied", type: "special_credit"}, {label: "Credit", type: "general"}, {label: "Discount", type: "general"}, {label: "New Client Discount", type: "general"}, {label: "Savings", type: "general"}, {label: "Promotion", type: "general"}, {label: "Credit Note Value", type: "general"}, {label: "Less", type: "general"} ];
  let fds=[];
  if (isDebugTargetThread) Logger.log(`findDiscountAmount: Searching body (first 500 chars):\n${body.substring(0,500)}`);
  for(const item of dls){
    let pt = ''; let amountCaptureGroupIndex; let labelForRegex = item.label.replace(/[().]/g, '\\$&'); 
    if (item.type === 'special_credit') { pt = `(?:${labelForRegex})${fs}(?:\\$\\s*)?-{1}\\s*${A_PATTERN}`; amountCaptureGroupIndex = 1; Logger.log(`findDiscountAmount: Using SPECIAL pattern for "${item.label}"`);}
    else { pt = `(?:${labelForRegex})(?:\\s*\\([^)]*\\))?[:]?${fs}${Sac_Pattern_Detailed}`; amountCaptureGroupIndex = 5; } // Added optional colon [:]? here as well
    Logger.log(`findDiscountAmount: Testing Label: "${item.label}", Pattern: ${pt}`);
    try { const p = new RegExp(pt,"gi"); p.lastIndex = 0; let m;
        while((m = p.exec(body)) !== null) {
            if (item.type === 'general' && m.length > 1) { let groupsLog = ""; for (let gi = 1; gi < m.length; gi++) { groupsLog += ` G${gi}:"${m[gi]}"`; } Logger.log(`findDiscountAmount: ---> General Match for "${item.label}": Full: "${m[0]}" ${groupsLog}`);}
            if(m[amountCaptureGroupIndex]){  if (item.type !== 'general') Logger.log(`findDiscountAmount: ---> Match Found for "${item.label}": Full Match: "${m[0]}", Amount Group (${amountCaptureGroupIndex}): "${m[amountCaptureGroupIndex]}"`);
                let cv = cleanCurrencyValue(m[amountCaptureGroupIndex]);
                if(cv !== null){ cv = Math.abs(cv); if(item.label.toLowerCase()==="less" && cv>1900 && cv<2100 && !m[amountCaptureGroupIndex].includes('.') && !m[amountCaptureGroupIndex].includes(',')) { if (isDebugTargetThread) Logger.log(`findDiscountAmount: Ignoring year ${cv} for label "${item.label}"`); continue; } fds.push(cv); Logger.log(`findDiscountAmount: Lbl "${item.label}" -> Parsed Disc: ${cv}`);
                } else { Logger.log(`findDiscountAmount: Lbl "${item.label}" -> cleanCurrencyValue null for "${m[amountCaptureGroupIndex]}"`);}}
            else if (item.type !== 'general') { Logger.log(`findDiscountAmount: ---> Match for "${item.label}", but Amount Group (${amountCaptureGroupIndex}) empty. Full: "${m[0]}"`);}
        }} catch (regexError) { Logger.log(`E: Regex error for label "${item.label}", Pat: ${pt} - Err: ${regexError.message}`);}
  }
  if(fds.length > 0){ fds.sort((a,b)=>b-a); Logger.log(`findDiscountAmount: Discounts found: [${fds.join(', ')}]. Returning: ${fds[0]}`); return fds[0];}
  Logger.log("findDiscountAmount: No discount values parsed."); return 0;
}

function findGrossTotalAmount(body, isDebugTargetThread) {
 if (isDebugTargetThread && body) Logger.log(`DEBUG (findGrossTotalAmount): Body(500):\n${body ? body.substring(0,500) : "NULL"}`); if (!body) return null;
 const A_PATTERN="((?:\\d{1,3}(?:[,.]\\d{3})+[,.]\\d{1,2})|(?:\\d+[,.]\\d{1,2})|(?:\\d{1,3}(?:[,.]\\d{3})+)|(?:\\d+))"; const sep="[\\s.:*_-]*?(?:[A-Z]{3}[\\s.:*_-]*)?"; const S_PATTERN="([\\$€£¥]?)\\s*"+A_PATTERN;
 const pLabels=[ {l:"Grand Total",p:0}, {l:"Total Amount",p:0}, {l:"Amount Payable",p:0}, {l:"Total Due",p:0}, {l:"Balance Due",p:1}, {l:"Subtotal",p:2}, {l:"Sub Total",p:2}, {l:"Sub-Total",p:2}, {l:"TOTAL",p:3}, {l:"AMOUNT",p:3} ]; let pMatches=[];
 for(const c of pLabels){ const pt=`(?:^|\\n)\\s*(${c.l})${sep}${S_PATTERN}`; try{ const p=new RegExp(pt,"gim"); p.lastIndex=0; let m; while((m=p.exec(body))!==null){ if(m[3]){ const cv=cleanCurrencyValue(m[3]); if(cv!==null) pMatches.push({v:cv, t:m[0], pr:c.p, lu:m[1]});}}}catch(e){if(isDebugTargetThread)Logger.log(`DEBUG (findGross): Regex E for ${c.l}: ${e.message}`);}}
 if(isDebugTargetThread && pMatches.length > 0) Logger.log(`DEBUG (findGross): Primary label matches: ${JSON.stringify(pMatches.map(m=>({lbl:m.lu, val:m.v, pri:m.pr, txt:m.t.substring(0,30).replace(/\n/g," ")})))}`);
 else if (isDebugTargetThread) Logger.log(`DEBUG (findGross): No primary label matches.`);
 let bestPA=null; if(pMatches.length>0){ pMatches.sort((a,b)=>{ if(a.pr!==b.pr) return a.pr-b.pr; return b.v-a.v; }); const validMatches = pMatches.filter(m => m.v < 10000000 && !(m.v > 1990 && m.v < 2100 && !m.t.includes('.') && !m.t.includes(','))); if (validMatches.length > 0) bestPA = validMatches[0].v; }
 if(bestPA!==null){ if(isDebugTargetThread)Logger.log(`DEBUG (findGross): Primary Best: ${bestPA}`); return bestPA;}
 if(isDebugTargetThread) Logger.log("DEBUG (findGross): No suitable primary. Trying fallback.");
 const Gf_PATTERN="("+"(?:\\d{1,3}(?:[,.]\\d{3})*[,.]\\d{1,2})|(?:\\d+[,.]\\d{1,2})|(?:\\d+)"+")";
 const Rf=new RegExp("(?:" + "(?:([\\$€£¥])\\s*("+Gf_PATTERN.slice(1,-1)+"))" + "|" + "(?:("+Gf_PATTERN.slice(1,-1)+")\\s*(USD|EUR|GBP|CAD|AUD|JPY|CHF|CNY|INR|NZD|ZAR))" + "|" + "(?:(USD|EUR|GBP|CAD|AUD|JPY|CHF|CNY|INR|NZD|ZAR)\\s*("+Gf_PATTERN.slice(1,-1)+"))" + ")", "gi");
 let fAmts=[]; Rf.lastIndex=0; let fM; while((fM=Rf.exec(body))!==null){ let amountStr=fM[2]||fM[3]||fM[6]; if(amountStr){ const cv=cleanCurrencyValue(amountStr); if(cv!==null) fAmts.push({val:cv, text:fM[0]});}}
 if(isDebugTargetThread&&fAmts.length>0) Logger.log(`DEBUG (findGross): Fallback amounts: ${JSON.stringify(fAmts.map(f=>({val:f.val,txt:f.text.substring(0,30).replace(/\n/g," ")})))}`);
 else if(isDebugTargetThread) Logger.log(`DEBUG (findGross): No fallback amounts.`);
 if(fAmts.length>0){ const sfA=fAmts.map(f=>f.val).filter(x=>Math.abs(x)>=0.01 && Math.abs(x)<10000000 && !(x>1990&&x<2100&&!String(x).includes('.'))); if(sfA.length>0){ sfA.sort((a,b)=>Math.abs(b)-Math.abs(a)); if(isDebugTargetThread)Logger.log(`DEBUG (findGross): Fallback Best: ${sfA[0]}`); return sfA[0];}}
 if(isDebugTargetThread)Logger.log("DEBUG (findGross): No gross total amount found."); return null;
}

function cleanCurrencyValue(valueStr){ if(valueStr==null||typeof valueStr=='undefined')return null; let s=String(valueStr).trim(); if(s==="")return null; s=s.replace(/[\$€£¥\s]/g,''); s=s.replace(/(.*?)(USD|EUR|GBP|CAD|AUD|JPY|CHF|CNY|INR|NZD|ZAR)$/i,'$1'); s=s.trim(); const h=s.includes(','),d=s.includes('.'); let p='.'; if(h&&d){if(s.lastIndexOf(',')>s.lastIndexOf('.'))p=','} else if(h&&!d){if(s.substring(s.lastIndexOf(',')+1).match(/^\d{1,2}$/) && s.indexOf(',')===s.lastIndexOf(',')) p=','; else {s=s.replace(/,/g,'');p='none'}} else if(!h&&!d)p='none'; let n; if(p===',')n=s.replace(/\./g,'').replace(',','.'); else if(p==='.')n=s.replace(/,/g,''); else n=s; const u=parseFloat(n); return isNaN(u)?null:u;}
function findInvoiceDate(body) { if(!body) return null; const dc="(\\d{4}[-/]\\d{1,2}[-/]\\d{1,2}|\\d{1,2}[-/]\\d{1,2}[-/]\\d{4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\s+\\d{1,2},?\\s+\\d{4}|\\d{1,2}\\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\s+\\d{4}|\\d{4}\\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\s+\\d{1,2})"; const ps=[ new RegExp("(?:Invoice Date|Date Issued|Date Generated|Issued|Date|DATE OF INVOICE)[:\\s]*"+dc,"i"), new RegExp("\\b(?:generated on|issued on)\\s+"+dc,"i"), new RegExp("^\\s*"+dc+"\\s*$","im") ]; for(const r of ps){ try{ const m=body.match(r); let ds=m?m[1]:null; if(ds){ ds=ds.trim(); const Dp=parseDateString(ds); if(Dp){ Logger.log(`findInvoiceDate: Pat "${r.source.substring(0,20)}..." -> "${ds}" -> ${formatDate(Dp)}`); return formatDate(Dp);}}}catch(e){}} Logger.log("findInvoiceDate: No date.");return null;}
function findDueDate(body) { if(!body) return null; const dc="(\\d{4}[-/]\\d{1,2}[-/]\\d{1,2}|\\d{1,2}[-/]\\d{1,2}[-/]\\d{4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\s+\\d{1,2},?\\s+\\d{4}|\\d{1,2}\\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\s+\\d{4}|\\d{4}\\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\\s+\\d{1,2})"; const ps=[ new RegExp("(?:Due Date|Payment Due|Due By|PAY BY|Due)[:\\s]*"+dc,"i"), new RegExp("(?:Payment Terms|Terms)[:\\s]*(?!Due upon receipt|Net\\s*\\d+|immediate)(?:Due\\s+)?(?:"+dc+")","i") ]; for(const r of ps){ try{ const m=r.exec(body); let ds=m?m[1]:null; if(ds){ ds=ds.trim(); if(/upon receipt|immediate|net\s*\d+/i.test(ds)) continue; const D=parseDateString(ds); if(D){ Logger.log(`findDueDate: Pat "${r.source.substring(0,20)}..." -> "${ds}" -> ${formatDate(D)}`); return formatDate(D);}}}catch(e){}} Logger.log("findDueDate: No due date.");return null;}
function findCurrency(body) { if(!body)return null; const cm={'$':'USD','USD':'USD', '€':'EUR','EUR':'EUR', '£':'GBP','GBP':'GBP', 'CAD':'CAD', 'JPY':'JPY','¥':'JPY', 'AUD':'AUD','CHF':'CHF','CNY':'CNY','INR':'INR','NZD':'NZD','ZAR':'ZAR'}; for(const sc in cm){ try{ let p; if(sc.length > 1) p = new RegExp(`\\b${sc}\\b`,'i'); else { const es = sc.replace(/[-\/\\^$*+?.()|[\]{}]/g,'\\$&'); p = new RegExp(`${es}(?:\\s*(?=\\d))`,'i');} if(body.match(p)){ Logger.log(`findCurrency: Found "${cm[sc]}" from "${sc}"`); return cm[sc];}} catch(e){}} Logger.log("findCurrency: No specific currency found."); return null;}
function findStreetAddress(body) { return "Address Not Parsed"; }
function parseRecipientName(recipientString) { if (!recipientString || typeof recipientString !== 'string') return null; const match = recipientString.match(/^\s*"?(.+?)"?\s*(?:<.*>)?$/); if (match && match[1]) { let name = match[1].trim().replace(/^"/, '').replace(/"$/, '').trim(); if (name.length > 1 && name.toLowerCase() !== 'null' && !name.includes('@') && !/^\d+$/.test(name)) return name; } return null; }
function parseDateString(dateString) { if(!dateString || typeof dateString!=='string') return null; let normalizedDateString = dateString.trim().replace(/(\d)(st|nd|rd|th)\b/gi,'$1').replace(/\s+/g, ' '); let parsedDate = new Date(normalizedDateString); if (isNaN(parsedDate.getTime())) { const europeanDateMatch = normalizedDateString.match(/^(\d{1,2})[-/. ](\d{1,2}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*)[-/. ](\d{4})$/i); if (europeanDateMatch) { let day = parseInt(europeanDateMatch[1], 10); let monthStr = europeanDateMatch[2]; let year = parseInt(europeanDateMatch[3], 10); let month = -1; if (!isNaN(parseInt(monthStr, 10))) { month = parseInt(monthStr, 10) - 1; } else { const monthParseAttempt = Date.parse(monthStr + " 1, 2000"); if (!isNaN(monthParseAttempt)) month = new Date(monthParseAttempt).getMonth(); } if (day > 0 && day <= 31 && month >= 0 && month <= 11 && year > 1990 && year < 2100) { parsedDate = new Date(Date.UTC(year, month, day)); if (!isNaN(parsedDate.getTime())) { Logger.log(`parseDateString: DD/MM/YYYY "${dateString}" -> ${formatDate(parsedDate)}`); return parsedDate; }}}} if (isNaN(parsedDate.getTime())) { const yyyymmddMatch = normalizedDateString.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/); if (yyyymmddMatch) { let year = parseInt(yyyymmddMatch[1], 10); let month = parseInt(yyyymmddMatch[2], 10) - 1; let day = parseInt(yyyymmddMatch[3], 10); if (day > 0 && day <= 31 && month >= 0 && month <= 11 && year > 1990 && year < 2100) { parsedDate = new Date(Date.UTC(year, month, day)); if (!isNaN(parsedDate.getTime())) { Logger.log(`parseDateString: YYYY/MM/DD "${dateString}" -> ${formatDate(parsedDate)}`); return parsedDate; }}}} if (!isNaN(parsedDate.getTime())) { Logger.log(`parseDateString: Directly/MM/DD/YYYY "${dateString}" -> ${formatDate(parsedDate)}`); return parsedDate; } else { Logger.log(`W: parseDateString FAIL: "${dateString}" (Norm: "${normalizedDateString}")`); return null; }}
// Function: setupSheet (Version V.0.7.6m - Handles newSheetCreatedJustNow and corrects potential brace mismatch)
// To be placed in your Code.gs file, replacing any older version of setupSheet.
function setupSheet(spreadsheet, sheetName, headersArray, isInvoiceLogSheet = false, newSheetCreatedJustNowFlag = false) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  let actualNewSheetCreatedThisCall = false;

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    Logger.log(`Sheet "${sheetName}" created by setupSheet call.`);
    actualNewSheetCreatedThisCall = true;
  }

  let currentHeadersOnSheet = [];
  if (!actualNewSheetCreatedThisCall && sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
    try {
      currentHeadersOnSheet = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    } catch (e) { Logger.log(`W: Could not get headers for existing sheet "${sheetName}": ${e.message}`); }
  }

  let headersNeedUpdate = actualNewSheetCreatedThisCall || (sheet.getLastRow() === 0) ||
                          (currentHeadersOnSheet.length !== headersArray.length) ||
                          headersArray.some((h, i) => h !== currentHeadersOnSheet[i]);

  let applyFullFormatting = actualNewSheetCreatedThisCall || newSheetCreatedJustNowFlag || headersNeedUpdate;

  if (sheetName === DASHBOARD_SHEET_NAME) {
      if (headersNeedUpdate && actualNewSheetCreatedThisCall) {
           Logger.log(`Sheet "${sheetName}" is Dashboard (newly created here). Headers handled by initialSetup's layout.`);
           // Basic header style can be set if desired, but initialSetup handles layout.
           try {
             sheet.getRange(1, 1, 1, headersArray[0].length > 0 ? headersArray[0].length : 1).setFontWeight("bold"); // Example basic style
             sheet.setFrozenRows(1);
           } catch(e) { Logger.log(`W: Min dashboard header style for new sheet: ${e.message}`);}
      }
      applyFullFormatting = false; 
  } else if (headersNeedUpdate) {
      Logger.log(`Sheet "${sheetName}": Headers require update or sheet is new. Applying headers.`);
      const colsToClear = Math.max(currentHeadersOnSheet.length, headersArray.length, 1);
      try { sheet.getRange(1, 1, 1, colsToClear).clearContent().clearFormat().clearNote().clearDataValidations().clearConditionalFormatRules(); } catch (e) {}
      try {
          sheet.getRange(1, 1, 1, headersArray.length).setValues([headersArray]).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
          sheet.setFrozenRows(1);
          Logger.log(`Header row set for "${sheetName}".`);
      } catch (e) { Logger.log(`E setting header for "${sheetName}": ${e.message}.`); }
  }

  if (applyFullFormatting && (sheetName === INVOICE_LOG_SHEET_NAME || sheetName === ERROR_LOG_SHEET_NAME)) {
    Logger.log(`Sheet "${sheetName}": Applying detailed column formatting, rules, and banding.`);
    try {
        if (isInvoiceLogSheet) {
            let gtIdx = headersArray.indexOf(COL_PARSED_GROSS);
            let dIdx = headersArray.indexOf(COL_PARSED_DISCOUNT);
            if (gtIdx > -1 && sheet.getMaxColumns() >= (gtIdx +1) ) { try {sheet.hideColumns(gtIdx + 1);} catch(e){Logger.log(`W: Hiding col ${gtIdx+1} failed: ${e.toString()}`)} }
            if (dIdx > -1 && sheet.getMaxColumns() >= (dIdx+1)) { try {sheet.hideColumns(dIdx + 1);} catch(e){Logger.log(`W: Hiding col ${dIdx+1} failed: ${e.toString()}`)} }
        }

        headersArray.forEach((header, index) => {
            const column = index + 1;
            // Get range for data rows. Ensure numDataRows is at least 1 for getRange if sheet is totally empty.
            const numDataRows = Math.max(1, sheet.getMaxRows() - 1); 
            const dataRowsRange = sheet.getRange(2, column, numDataRows, 1);
            
            // Clear existing formats from data rows for this column to prevent overlaps
            dataRowsRange.clearFormat().clearDataValidations().clearNote(); 

            switch (header) {
                case COL_PROCESSED_TS: case COL_SENT_DATE: case COL_INVOICE_DATE: case COL_DUE_DATE: case COL_REPLY_PAID_DATE: case 'Error Timestamp': case 'Email Date':
                    dataRowsRange.setNumberFormat("yyyy-MM-dd HH:mm:ss"); sheet.setColumnWidth(column, 170); break;
                case COL_EMAIL_SUBJECT: case COL_CLIENT_EMAIL: case COL_ATTACH_NAMES: case COL_INVOICE_SENDER_EMAIL:
                    sheet.setColumnWidth(column, 250); dataRowsRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("top"); break;
                case COL_PARSED_GROSS: case COL_PARSED_DISCOUNT: case COL_AMOUNT_DUE: case COL_REPLY_MENTIONED_AMOUNT:
                    dataRowsRange.setNumberFormat("0.00##").setHorizontalAlignment("right"); sheet.setColumnWidth(column, 120); break;
                case COL_INVOICE_NUM: case COL_CURRENCY: case COL_REPLY_MENTIONED_CURRENCY:
                    sheet.setColumnWidth(column, 100); dataRowsRange.setHorizontalAlignment("left"); break;
                case COL_STATUS:
                    sheet.setColumnWidth(column, 120);
                    dataRowsRange.setHorizontalAlignment("left");
                    const statusValues = ["Unpaid", "Paid", "Partially Paid", "Overdue", "Pending Confirmation", "Cancelled", "NEEDS REVIEW"];
                    const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statusValues, true).setAllowInvalid(false).setHelpText("Select status.").build();
                    dataRowsRange.setDataValidation(statusRule);
                    break;
                case COL_CLIENT_NAME: case COL_STREET_ADDRESS:
                    sheet.setColumnWidth(column, 200); dataRowsRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("top"); break;
                case COL_PAYMENT_METHOD: sheet.setColumnWidth(column, 130); dataRowsRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("top").setHorizontalAlignment("left"); break;
                case COL_PAYMENT_NOTES: sheet.setColumnWidth(column, 350); dataRowsRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("top").setHorizontalAlignment("left"); break;
                case COL_LINK_GMAIL: sheet.setColumnWidth(column, 200); dataRowsRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setHorizontalAlignment("left"); break;
                case COL_PARSING_NOTES: case 'Error Message': case 'Raw Email Body (Snippet)': sheet.setColumnWidth(column, 350); dataRowsRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("top"); break;
                case COL_REMINDER_STAGE_SENT: sheet.setColumnWidth(column, 150); dataRowsRange.setHorizontalAlignment("left").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); break;
                default: sheet.setColumnWidth(column, 150); break;
            }
        }); // End forEach header

        if (isInvoiceLogSheet && sheet.getMaxRows() > 1) {
            try { // Try block for conditional formatting
                sheet.setConditionalFormatRules([]); // Clear all existing rules first
                const fullDataRange = sheet.getRange(2, 1, sheet.getMaxRows() -1, headersArray.length);
                let newRules = [];
                const greenerGreen = "#6aa84f"; const pastelGreen = "#b6d7a8"; const orangeWarning = "#f6b26b"; const redAlert = "#e06666"; const darkRedAlert = "#cc4125"; const pendingConfirmationYellow = "#fff2cc"; const cancelledGrey = "#efefef"; const cancelledFontColor = "#757575";
                const statusColIdx = headersArray.indexOf(COL_STATUS); const dueDateColIdx = headersArray.indexOf(COL_DUE_DATE); const parsingNotesColIdx = headersArray.indexOf(COL_PARSING_NOTES);
                if (statusColIdx > -1) { 
                    const statusColLetter = columnToLetter(statusColIdx + 1); 
                    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=INDIRECT("${statusColLetter}"&ROW())="Paid"`).setBackground(greenerGreen).setRanges([fullDataRange]).build()); 
                    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=INDIRECT("${statusColLetter}"&ROW())="Partially Paid"`).setBackground(pastelGreen).setRanges([fullDataRange]).build()); 
                    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=INDIRECT("${statusColLetter}"&ROW())="Cancelled"`).setBackground(cancelledGrey).setFontColor(cancelledFontColor).setRanges([fullDataRange]).build()); 
                    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=INDIRECT("${statusColLetter}"&ROW())="Pending Confirmation"`).setBackground(pendingConfirmationYellow).setRanges([fullDataRange]).build()); 
                    if (dueDateColIdx > -1) { 
                        const dueDateColLetter = columnToLetter(dueDateColIdx + 1); 
                        const dueDateValidCondition = `AND(ISBLANK(INDIRECT("${dueDateColLetter}"&ROW()))=FALSE, ISDATE(INDIRECT("${dueDateColLetter}"&ROW())))`; 
                        const statusOverdueOrUnpaid = `OR(INDIRECT("${statusColLetter}"&ROW())="Overdue", INDIRECT("${statusColLetter}"&ROW())="Unpaid")`; 
                        const commonOverdueCondition = `AND(${dueDateValidCondition}, ${statusOverdueOrUnpaid}, INDIRECT("${dueDateColLetter}"&ROW())<TODAY())`; 
                        newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${commonOverdueCondition}, INDIRECT("${dueDateColLetter}"&ROW())<TODAY()-30)`).setBackground(darkRedAlert).setRanges([fullDataRange]).build()); 
                        newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${commonOverdueCondition}, INDIRECT("${dueDateColLetter}"&ROW())<TODAY()-15, INDIRECT("${dueDateColLetter}"&ROW())>=TODAY()-30)`).setBackground(redAlert).setRanges([fullDataRange]).build()); 
                        newRules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND(${commonOverdueCondition}, INDIRECT("${dueDateColLetter}"&ROW())>=TODAY()-15)`).setBackground(orangeWarning).setRanges([fullDataRange]).build()); 
                    } // end if dueDateColIdx
                } // end if statusColIdx
                if (parsingNotesColIdx > -1) { 
                    const parsingNotesRange = sheet.getRange(2, parsingNotesColIdx + 1, sheet.getMaxRows() - 1, 1); 
                    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("NEEDS KEYWORD REVIEW").setBackground("#f4cccc").setBold(true).setRanges([parsingNotesRange]).build()); 
                    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("SCRIPT ERROR").setBackground("#ffd966").setBold(true).setRanges([parsingNotesRange]).build()); 
                    newRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("Potential Overpayment").setBackground("#cfe2f3").setRanges([parsingNotesRange]).build()); 
                } // end if parsingNotesColIdx
                if(newRules.length > 0) sheet.setConditionalFormatRules(newRules); 
                Logger.log(`Conditional formats applied to "${sheetName}".`);
            } catch (cfError) {Logger.log(`E: Conditional formatting for "${sheetName}": ${cfError.message}\n${cfError.stack}`);}
        } // end if isInvoiceLogSheet

        try { // Row Banding
            const existingBands = sheet.getBandings(); 
            for (let i = 0; i < existingBands.length; i++) existingBands[i].remove();
            const bandingDataRange = sheet.getRange(2, 1, Math.max(1, sheet.getMaxRows() -1 ), headersArray.length);
            if (bandingDataRange.getWidth() > 0 && bandingDataRange.getHeight() > 0) bandingDataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
            Logger.log(`Row banding applied to "${sheetName}".`);
        } catch (bandErr) { Logger.log(`W: Banding error for "${sheetName}": ${bandErr.message}.`); }
    } catch(formatError) { // Catch errors from the broader detailed formatting block
        Logger.log(`E: Applying detailed formatting for sheet "${sheetName}": ${formatError.message}\n${formatError.stack}`);
    }
  } else { // End if applyFullFormatting AND (isInvoiceLogSheet OR isErrorLogSheet)
    Logger.log(`Sheet "${sheetName}" headers are current or is Dashboard. Detailed column formatting/rules pass skipped this time.`);
  }
  Logger.log(`setupSheet completed for "${sheetName}".`);
} // <<<< THIS IS THE CORRECT CLOSING BRACE FOR setupSheet FUNCTION
function columnToLetter(column) { let temp, letter = ''; while (column > 0) { temp = (column - 1) % 26; letter = String.fromCharCode(temp + 65) + letter; column = (column - temp - 1) / 26; } return letter; }
function formatDate(dateObject, formatString = 'yyyy-MM-dd HH:mm:ss') { if(!dateObject || !(dateObject instanceof Date) || isNaN(dateObject.getTime())) return ''; try { return Utilities.formatDate(dateObject, SCRIPT_TIMEZONE, formatString); } catch (e) { Logger.log(`W: formatDate error: ${e.message}`); return ''; }}
function logErrorToSheetAdvanced(errorLogSheet, errorDetails) { if(!errorLogSheet){ Logger.log("E: Error log sheet null. Cannot log."); return; } try{ const timestamp = formatDate(new Date()); const emailDate = errorDetails.date instanceof Date ? formatDate(errorDetails.date) : (errorDetails.date||''); const rowData=[ timestamp, errorDetails.subject||'', emailDate, String(errorDetails.recipients||''), errorDetails.functionName||'Unknown', errorDetails.errorMessage||'No message', errorDetails.threadLink||'', errorDetails.rawBodySnippet||'' ]; errorLogSheet.appendRow(rowData); } catch(e){ Logger.log(`CRIT FAIL: Write to error log: ${e.message}. Original Error: ${errorDetails.errorMessage || 'N/A'}.`);}}
// =============== SCRIPT END ===============
