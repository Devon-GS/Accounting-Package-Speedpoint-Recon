<div align="center">
  
# 💳 SpeedpointRecon

**Automated Financial Reconciliation & Auditing Tool**

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?style=for-the-badge&logo=python)
![CustomTkinter](https://img.shields.io/badge/GUI-CustomTkinter-blueviolet?style=for-the-badge)
![PostgreSQL](https://img.shields.io/badge/PostgreSQL-Ready-blue?style=for-the-badge&logo=postgresql)
![SQLite](https://img.shields.io/badge/SQLite-State_Tracking-lightblue?style=for-the-badge&logo=sqlite)
![Excel](https://img.shields.io/badge/Excel-openpyxl-green?style=for-the-badge&logo=microsoftexcel)

</div>

---

## ✨ Overview

**SpeedpointRecon** is an automated desktop reconciliation utility designed to match bank statement line items against Speedpoint terminal batch settlements and OFX accounting payments. 

Featuring a modern **CustomTkinter GUI**, multi-tier heuristic matching, and direct PostgreSQL integration, it eliminates manual cross-referencing and highlights matched and unresolved transactions directly inside your `.xlsx` workbooks.

---

## 🚀 Key Features

* **Modern Desktop GUI:** Built with CustomTkinter, featuring real-time database connection health monitoring, multi-file queues, background thread execution, and clean dark/light system adaptation.
* **Multi-Workbook Processing:** Queue and batch process multiple Speedpoint terminal sheets and Bank Statements simultaneously with in-place saving.
* **Multi-Tier Terminal Matching:**
  * **Speedpoint Batch Matching:** Resolves batch numbers and amounts against dedicated terminal columns (`946389` & `946390`).
  * **Amount Fallback Matching:** Scans Speedpoint summary columns (Column 19 / `S`) for direct amount hits while preventing duplicate row reuse.
  * **NEDLNK DP Settlements:** Dynamically calculates row totals across Columns `G` through `O` (columns 7–15) to match consolidated batch deposits.
* **OFX Database Integration:** Connects to PostgreSQL to cross-reference transactions against cleared OFX supplier payments by matching normalized dates, supplier descriptors, and payment amounts.
* **Discrepancy Resolution & State Tracking:** Unresolved issues are listed in an interactive tree table. Flagged items can be selected and cleared using **Fix Issues**, which marks the bank statement rows and persists resolutions into a local SQLite database (`fixed_issues.sqlite3`) so they are ignored on subsequent runs.
* **Visual Auditing (Auto-Highlighting):** Modifies your `.xlsx` files directly, color-coding cells so reviewers can immediately see which records cleared and which did not.
* **Robust Diagnostic Logging:** Built-in self-truncating rotating logger (`speedpoint.log`, capped at 10 MB) captures near-misses, date/amount mismatches, and execution stats.

---

## 🎨 Color Coding Legend

When workbooks are processed, matching rows and cells receive distinct background highlights:

| Status | Target Workbook | Target Column / Area | Color | Hex Code |
| :--- | :--- | :--- | :--- | :--- |
| **Speedpoint Match** | Speedpoint Report | Batch / Terminal Amount Cell | 🟪 Pink | `#FF86B9` |
| **NEDLNK DP Match** | Speedpoint Report | Columns `G:O` (Columns 7–15) | 🟪 Pink | `#FF86B9` |
| **Bank Match** | Bank Statement | Column `D` (Credit) | 🟨 Yellow | `#FFFF00` |
| **OFX Payment Match** | Bank Statement | Column `C` (Debit) | 🟦 Blue | `#00B0F0` |

---

## ⚙️ Prerequisites & Installation

Ensure you have Python 3.9 or newer installed. Install the dependencies via pip:

```bash
pip install openpyxl python-dotenv keyring customtkinter psycopg
```

*(Note: `psycopg` is required for connecting to PostgreSQL and performing OFX verification).*

---

## 🔐 Database & Directory Configuration

SpeedpointRecon supports secure credentials via the **Windows Credential Manager** (using the `keyring` library) with fallback to a `.env` file.

### Option 1: CLI Configuration (Recommended)
You can set credentials and default file picker directories directly through the application's built-in CLI flags:

```bash
# Save PostgreSQL connection string
python app.py --set DATABASE_URL "postgresql://user:password@localhost:5432/accounting"

# Set default directory shortcuts for the GUI file pickers
python app.py --set SPEEDPOINT_PATH "C:\path\to\speedpoint\reports"
python app.py --set BANK_STATEMENT_PATH "C:\path\to\bank\statements"
```

or

```cmd
app.exe --set DATABASE_URL "postgresql://user:password@localhost:5432/accounting"
app.exe --set SPEEDPOINT_PATH "C:\path\to\speedpoint\reports"
app.exe --set BANK_STATEMENT_PATH "C:\path\to\bank\statements"
```

### Option 2: Environment File (`.env`)
Alternatively, create a `.env` file in the project root:

```env
# Full connection string
DATABASE_URL=postgresql://accounting:password@localhost:5432/accounting

# Or individual connection parameters
PGHOST=localhost
PGPORT=5432
PGDATABASE=accounting
PGUSER=accounting
PGPASSWORD=your_password

# Default picker directories
SPEEDPOINT_PATH="C:/Accounting/Speedpoints"
BANK_STATEMENT_PATH="C:/Accounting/BankStatements"
```

---

## 💡 How It Works

1. **Load Workbooks:** The application takes one or more Speedpoint reports (`.xlsx`) and Bank Statement workbooks.
2. **Fetch OFX Signatures:** It queries the PostgreSQL database for recent `OFX` payments and builds normalized signatures `(date, supplier_name, amount)`.
3. **Parse Transactions:** The bank statement parser extracts transaction lines, identifying Speedpoint terminals, batch numbers, and consolidated NEDLNK deposits.
4. **Match & Highlight:**
   * OFX payments match directly on Date, Supplier, and Amount.
   * Speedpoint transactions match against batch and amount combinations in the terminal report.
   * Both files are highlighted and saved in place.
5. **Review Discrepancies:** Any bank row without a corresponding terminal record or OFX entry is logged in the UI table showing the file, row number, transaction description, credit amount, and failure reason.
6. **Fix Issues:** Selected items can be reconciled manually through the GUI, updating the files and saving the resolved signatures to `fixed_issues.sqlite3`.

---

## 🖥️ Running the Application

Launch the desktop GUI:

```bash
python app.py
```

### GUI Usage Steps:
1. Verify the **Database** pill indicates **Connected** (green).
2. Click **Add Files** under **Speedpoints** to select your terminal `.xlsx` workbooks.
3. Click **Add Files** under **Bank Statements** to select your bank statement `.xlsx` workbooks.
4. Click **Process Files** and confirm the dialog to proceed with in-place reconciliation.
5. If unresolved rows appear in the results list:
   * Review the identified transaction and discrepancy reason.
   * Click the checkbox next to any issue you wish to manually override/clear.
   * Click **Fix Issues** to mark them in the bank statement and prevent them from showing up again in subsequent runs.

---

## 📁 File Structure

```text
├── app.py                  # CustomTkinter GUI & CLI configuration entrypoint
├── speedpoint_core.py     # Reconciliation engine, database connectors & Excel logic
├── fixed_issues.sqlite3   # Local SQLite database tracking manually resolved issues
├── speedpoint.log         # Diagnostic log file (automatically capped at 10 MB)
└── README.md              # Project documentation
```

---

<div align="center">
	<b>Coded by Devon — Redfire IT</b>
</div>