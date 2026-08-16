<div align="center">
  
# 💳 SpeedpointRecon

**Automated Financial Reconciliation & Auditing Tool**

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?style=for-the-badge&logo=python)
![PostgreSQL](https://img.shields.io/badge/PostgreSQL-Ready-blue?style=for-the-badge&logo=postgresql)
![Excel](https://img.shields.io/badge/Excel-openpyxl-green?style=for-the-badge&logo=microsoftexcel)

</div>

---

## ✨ Overview

**SpeedpointRecon** is a specialized, automated Python utility designed to reconcile bank statements with Speedpoint terminal batches and OFX payments. By leveraging intelligent text parsing and direct database integration, it eliminates tedious manual cross-referencing and visually highlights matched transactions directly within your Excel workbooks.

## 🚀 Key Features

*   **Intelligent Reconciliation:** Automatically identifies and matches `SPEEDPOINT` and `NEDLNK DP` transactions using Terminal IDs, Batch Numbers, and amounts.
*   **Database Integration:** Connects seamlessly to your PostgreSQL accounting database to fetch and cross-reference OFX payments.
*   **Visual Auditing (Auto-Highlighting):** Modifies your `.xlsx` files directly, color-coding cells so you know exactly what cleared and what didn't.
*   **Smart Parsing:** Handles multiple date formats, safely sanitizes filenames, and intelligently parses text-based currency strings.
*   **Detailed Reporting:** Generates a comprehensive reconciliation manifest containing matched counts, ignored rows (e.g., SASOL OIL), and unresolved discrepancies.

---

## 🎨 Color Coding Legend

When the script processes your Excel workbooks, it uses the following highlight colors to indicate matches:

| Status | Target Workbook | Color | Hex Code |
| :--- | :--- | :--- | :--- |
| **Speedpoint Match** | Terminal Report | 🟪 Pink | `#FF86B9` |
| **Bank Match** | Bank Statement | 🟨 Yellow | `#FFFF00` |
| **OFX Payment Match**| Bank Statement | 🟦 Blue | `#00B0F0` |

---

## ⚙️ Prerequisites

Before running the application, ensure you have the required Python dependencies installed. You can install them via pip:

```bash
pip install openpyxl python-dotenv keyring psycopg
```
*(Note: `psycopg` is strictly required for the database OFX matching functionality).*

---

## 🔐 Database Configuration & Setup

SpeedpointRecon requires secure access to your PostgreSQL database to fetch OFX payments. To avoid hardcoding passwords, the application securely interacts with the **Windows Credential Manager** via the `keyring` library, or falls back to `.env` variables.

👉 **[Click here to view the Setup Guide (SETUP.md)](SETUP.md)** for detailed instructions on how to configure your Windows Generic Credentials securely.

---

## 💡 How it Works

1. **Load Workbooks:** The core engine takes a path to your Speedpoint report (`.xlsx`) and a list of Bank Statement workbooks.
2. **Fetch OFX Signatures:** It queries the connected database for recent OFX payments and builds matchable signatures.
3. **Cross-Reference:** It loops through bank statement rows, extracting transaction types and expected amounts.
4. **Match & Highlight:** It searches the Speedpoint file for exact Batch/Amount matches. If an exact match is found, both workbooks are highlighted and saved.
5. **Report Generation:** A JSON-like manifest is returned containing the UUID of the run, processed files, and any unresolved rows for manual review.

---

<div align="center">
	<b><a href="https://devon-gs.github.io/" target="_blank">Coded by Devon - Redfire IT</a></b>
</div>

---