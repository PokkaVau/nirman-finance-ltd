# Nirman Finance LTD — Professional Financial Management Platform

A modern, corporate, high-end, and fully responsive financial management system built for **Nirman Finance LTD**.

Built with a modern corporate aesthetic adhering to agency-grade visual design standards: custom **Plus Jakarta Sans** typography, double-bezel (Doppelrand) card architectures, button-in-button nested CTAs, ৳ Bangladeshi Taka formatting, and an interest-free (0% interest) financial management engine.

---

## 🎨 Brand Identity & Design System

The system is styled strictly to the official **Nirman Finance** brand palette:

| Color Role | Hex Code | Purpose / Usage |
| :--- | :--- | :--- |
| **Primary Navy Blue** | `#174879` | Primary navigation, headers, brand typography, primary buttons |
| **Secondary Blue** | `#245B8E` | Active navigation states, subheadings, accent borders |
| **Primary Green** | `#2BA653` | Positive cash flows, active status badges, growth indicators |
| **Light Green** | `#EAF7EE` | Positive highlights, success pill backgrounds |
| **Light Blue** | `#EDF4FA` | Input focus backgrounds, subtle pill tags |
| **Background** | `#F5F8FC` | Ultra-clean ambient workspace background |
| **Card White** | `#FFFFFF` | Nested inner core cards |
| **Dark Text** | `#172B3A` | High-contrast readable typography |
| **Secondary Text** | `#687789` | Secondary captions, timestamps, table meta info |
| **Borders & Dividers** | `#E3EAF1` | Hairline borders, structural separation lines |

---

## 📂 Project Architecture

```
Nirman_Finance_LTD/
│
├── Code.gs             # Google Apps Script backend engine (Google Sheets database operations)
├── Index.html          # Semantic HTML5 corporate dashboard structure (GAS template)
├── Styles.html         # High-End Design System (Double-bezel, Plus Jakarta Sans, responsiveness)
├── Scripts.html        # Reactive JavaScript client engine (Chart.js, CRUD, validation, GAS bridge)
├── appsscript.json     # Apps Script manifest configuration (Asia/Dhaka timezone, V8 runtime)
├── index_local.html    # Standalone bundled version for immediate local preview/testing
├── build_local.js      # Utility script to re-compile index_local.html from source templates
└── Logo.jpg            # Original Nirman Finance corporate identity logo
```

---

## 💼 Core Financial Modules & Business Rules

### 1. Financial Dashboard
- **8 KPI Summary Cards (Double-Bezel Architecture):**
  1. Total Shareholders
  2. Total Shares (`1 Share = ৳ 1,000`)
  3. Total Deposits (Gross member deposits)
  4. Total Withdrawals (Processed capital payouts)
  5. Total Available Balance (`Deposits - Withdrawals`)
  6. Total Active Loans
  7. Total Loan Amount Disbursed
  8. Total Outstanding Loan Amount
- **Monthly Cash Flow Chart:** Monthly deposits vs withdrawals over time.
- **Loan Portfolio Chart:** Capital recovered vs outstanding exposure.
- **Recent Transaction Audit Stream:** Live view of latest deposits, withdrawals, disbursements, and repayments.
- **Upcoming Installments Stream:** Highlighting payments due and overdue warnings.

### 2. Shareholder Management
- Registry of shareholders with ID (`SH-1001`), Full Name, Phone, Email, Address, NID Number, Joining Date, Total Deposited, Total Withdrawn, Available Balance, Number of Shares, Total Share Value (`Shares × ৳1,000`), Status, and Notes.
- View individual shareholder financial statements detailing all their deposits and withdrawals.
- Search by name, phone, NID, or ID; filter by status; sort by name, shares, or balance.

### 3. Deposit Management
- Record member deposits with automated shareholder balance update and shares calculation (`Shares = Available Balance ÷ ৳1,000`).
- Generates transaction ledger entry (`TXN-XXXX`) and audit log.

### 4. Withdrawal Operations
- Strict balance check: **Withdrawals exceeding the shareholder's available balance are automatically blocked**.
- Automatically updates remaining shares and logs transaction.

### 5. Borrower Profiles
- Supports three borrower categories:
  1. `Shareholder`
  2. `Outside Individual`
  3. `Company / Business` (with dedicated Contact Person and Registration/Trade License fields).

### 6. Loan Management (0% Interest)
- **100% Interest-Free:** No interest or hidden fees are ever added.
- **Repayment Duration:** Configurable in months (e.g., 6, 12, 24, 36).
- **Exact Monthly Installment Formula:**
  $$\text{Monthly Installment} = \lfloor \text{Total Loan Amount} \div \text{Duration} \rfloor$$
- **Penny-Adjustment Rule:** The final installment automatically absorbs any rounding differences so the sum of installments equals the original loan amount.
- **Dedicated Loan Schedule View:** Interactive visual schedule with progress bar, payment receipts, and installment statuses (`Pending`, `Paid`, `Partially Paid`, `Overdue`).

### 7. Installment Management & Collections
- Track pending, upcoming, overdue, and paid installments.
- Record full or partial payments.
- Automatic updates to loan remaining balance and status (`Completed` when remaining = 0).

### 8. Central Transactions Ledger
- Single source of truth recording `DEPOSIT`, `WITHDRAWAL`, `LOAN_DISBURSEMENT`, and `LOAN_REPAYMENT`.
- Search by reference or keyword; filter by transaction type, flow direction (`IN` / `OUT`), and date ranges.

### 9. Reports Engine
- Generates 8 distinct financial statements:
  - Monthly Financial Summary
  - Shareholder Equity & Shares Report
  - Deposit Transactions Report
  - Withdrawal Operations Report
  - Loan Portfolio & Repayment Status
  - Borrower Exposure Report
  - Installment Collection Performance
  - Outstanding Loans Aging
- Date filtering, CSV export with automatic file download, and clean printable layout.

### 10. Database Settings & Administration
- Configure company name, official phone, email, and address.
- One-click Google Sheets database table initialization (`initializeDatabase()`).
- One-click demo data seeder (`seedDemoData()`) with realistic Bangladeshi corporate profiles.

---

## 🗄️ Google Sheets Database Structure

The system uses Google Sheets as a relational database across 10 tables:

1. **`Shareholders`** — `[Shareholder_ID, Full_Name, Phone, Email, Address, NID_Number, Joining_Date, Total_Deposited, Total_Withdrawn, Available_Balance, Number_Of_Shares, Total_Share_Value, Status, Notes, Created_At]`
2. **`Deposits`** — `[Deposit_ID, Shareholder_ID, Shareholder_Name, Amount, Deposit_Date, Reference_No, Notes, Created_At]`
3. **`Withdrawals`** — `[Withdrawal_ID, Shareholder_ID, Shareholder_Name, Amount, Withdrawal_Date, Reference_No, Notes, Created_At]`
4. **`Borrowers`** — `[Borrower_ID, Borrower_Type, Full_Name, Phone, Email, Address, NID_Or_Reg_No, Contact_Person, Registration_Date, Status, Notes, Created_At]`
5. **`Loans`** — `[Loan_ID, Borrower_ID, Borrower_Name, Borrower_Type, Loan_Amount, Loan_Date, Duration_Months, Monthly_Installment, First_Installment_Date, Total_Paid, Remaining_Amount, Loan_Status, Notes, Created_At]`
6. **`Loan_Installments`** — `[Installment_ID, Loan_ID, Borrower_ID, Borrower_Name, Installment_Number, Due_Date, Due_Amount, Paid_Amount, Balance_Due, Status, Paid_Date, Payment_Reference]`
7. **`Loan_Payments`** — `[Payment_ID, Loan_ID, Borrower_ID, Borrower_Name, Installment_ID, Amount_Paid, Payment_Date, Reference_No, Notes, Created_At]`
8. **`Transactions`** — `[Transaction_ID, Date, Type, Category, Entity_Type, Entity_ID, Entity_Name, Amount, Direction, Reference_No, Notes, Created_At]`
9. **`Audit_Logs`** — `[Log_ID, Timestamp, Action, Entity, Entity_ID, User, Details]`
10. **`Settings`** — `[Setting_Key, Setting_Value, Description]`

---

## 🚀 Deployment Guide (Google Apps Script & Google Sheets)

### Step 1: Create or Open a Google Sheet
1. Go to [Google Sheets](https://sheets.new) and create a new blank spreadsheet.
2. Name the spreadsheet: **`Nirman Finance LTD — Database`**.

### Step 2: Open the Apps Script Editor
1. In your Google Sheet, click **Extensions** &rarr; **Apps Script**.
2. Rename the project to **`Nirman Finance Web App`**.

### Step 3: Add the Files to Apps Script
In the Apps Script editor file explorer (left pane):

1. **`Code.gs`**:
   - Replace the default code in `Code.gs` with the complete contents of `Code.gs` from this project.
2. **`Index.html`**:
   - Click the **+** (Add a file) icon &rarr; choose **HTML**.
   - Name it `Index` (Apps Script will append `.html`).
   - Paste the contents of `Index.html`.
3. **`Styles.html`**:
   - Click the **+** icon &rarr; choose **HTML**.
   - Name it `Styles`.
   - Paste the contents of `Styles.html`.
4. **`Scripts.html`**:
   - Click the **+** icon &rarr; choose **HTML**.
   - Name it `Scripts`.
   - Paste the contents of `Scripts.html`.
5. **`appsscript.json`**:
   - In Project Settings (gear icon on the left), check **"Show 'appsscript.json' manifest file in editor"**.
   - Return to the Editor tab, click `appsscript.json`, and paste the contents from `appsscript.json`.

### Step 4: Initialize the Database Tables
1. At the top of the Apps Script editor, in the function dropdown, select **`initializeDatabase`**.
2. Click **Run**.
3. Google will ask for permission on first run &rarr; click **Review Permissions** &rarr; choose your Google account &rarr; click **Advanced** &rarr; click **Go to Nirman Finance Web App (unsafe)** &rarr; click **Allow**.
4. Check your Google Sheet: All 10 sheets are automatically created, formatted with corporate Navy Blue headers, frozen title rows, and default settings!

*(Optional: To populate realistic sample data, select `seedDemoData` and click **Run**).*

### Step 5: Deploy as a Web App
1. At the top right of the Apps Script editor, click **Deploy** &rarr; **New deployment**.
2. Click the gear icon next to "Select type" and select **Web app**.
3. Fill in the deployment settings:
   - **Description:** `Production v2.0 - Nirman Finance Portal`
   - **Execute as:** `Me (your email)`
   - **Who has access:** `Anyone` (or `Only myself` / domain users depending on security needs)
4. Click **Deploy**.
5. Copy the **Web App URL**. This is your live production link!

---

## ⚡ Deployment on Vercel

You can deploy the frontend on **Vercel** with global CDN speed and custom domains, connected seamlessly to your Google Sheets database:

### Step 1: Push Repository to GitHub
```bash
git remote add origin https://github.com/YOUR_USERNAME/nirman-finance-ltd.git
git branch -M main
git push -u origin main
```

### Step 2: Import into Vercel
1. Go to [vercel.com](https://vercel.com) and log in.
2. Click **Add New...** &rarr; **Project**.
3. Import your **`nirman-finance-ltd`** repository.
4. Framework Preset: **Other** (Root directory: `./`).
5. Build command: `npm run build` (or leave default: will automatically execute `node build_local.js`).
6. Output directory: `./`.
7. Click **Deploy**.

### Step 3: Connect Vercel Frontend to Google Sheets Backend
1. Deploy your Google Apps Script as a Web App (as described in the GAS guide above) and copy your **Web App URL** (`https://script.google.com/macros/s/.../exec`).
2. Open your live Vercel app in the browser.
3. In the sidebar, navigate to **Settings**.
4. Scroll to **Cloud Hosting & Live Google Sheets Integration**.
5. Paste your Web App URL into the field and click **Save API URL**.
6. Click **Test Ping** to verify the connection.
7. Your Vercel frontend is now connected live to Google Sheets!

---

## 🖥️ Local Preview (Zero Setup)

You can also test the full application immediately in your browser:
- Double-click `index_local.html` or `index.html` to open it in Chrome, Edge, or Firefox.
- All 10 views, 8 summary metrics, interactive charts, modals, and filters are fully functional with the built-in offline test adapter.
