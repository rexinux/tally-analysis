# 📊 Rexinux's Tally Analyser

An intelligent, 100% offline, privacy-first financial viewer, voucher editor, and 12-month GSTR-2B reconciliation studio for Tally XML exports and GSTR-2B Excel/JSON files. Developed by **Rexinux**.

---

## ✨ Features at a Glance

- 🚀 **High-Performance XML Engine**: Loads ~150MB Tally XML exports into an in-memory SQLite database in ~3 seconds.
- 📑 **Comprehensive Financial Statements**:
  - **Daybook**: Searchable, paginated transaction grid with voucher type and date filters.
  - **Trial Balance**: Expandable hierarchical tree (Groups $\rightarrow$ Subgroups $\rightarrow$ Ledgers).
  - **Profit & Loss**: Gross Profit, Operating Expenses, Direct/Indirect Expenses, and Net Income.
  - **Ledger Account Statement**: Detailed transaction history with running balance calculations.
- 🔍 **5-Pass Intelligent GSTR-2B Reconciliation Engine**:
  - Automatically parses 12 months of GSTR-2B Excel (`.xlsx`) and JSON files (`B2B`, `B2BA`, `CDNR`, `CDNRA` sheets).
  - **Pass 1**: Exact Invoice Number & Amount Match ($\pm\text{₹}2$).
  - **Pass 2 (Shifted Bill Detection)**: Flags invoices where the accountant typed the wrong bill number in Tally (`SHIFTED BILL`).
  - **Pass 3 (Split-Voucher Aggregation)**: Combines multi-rate vouchers entered under the same bill number in Tally (e.g. 12% + 18% lines).
  - **Pass 4 (Tax Difference Analysis)**: Calculates exact `Tax Diff (2B - Tally)`.
  - **Pass 5 (Missing Analysis)**: Identifies **Unclaimed ITC** (`MISSING_IN_TALLY`) vs **Unfiled Supplier Invoices** (`MISSING_IN_2B`).
  - **Late Filing Tracker**: Tracks **2B Return Period** and **GSTR-1 Filing Date** to spot cross-period filings.
- ✏️ **Voucher Editor & XML Exporter**:
  - Edit voucher dates, numbers, party names, amounts, and narrations directly in your browser.
  - One-click **Export Tally XML** to generate clean XML ready for re-import into Tally.
- 🔒 **100% Offline & Private**: Zero external cloud or API dependencies. Runs completely on `localhost:8000`.

---

## 📁 Directory Structure

```
Tally analysis/
├── AQUAXMLFY26.xml         # Primary Tally XML Data Export (~150MB)
├── R2B/                    # Folder containing 12 Months of GSTR-2B Excel files (.xlsx)
│   ├── 042025_...xlsx
│   ├── ...
│   └── 032026_...xlsx
├── server.py               # Standalone Python HTTP Server & Reconciliation Backend
├── index.html              # Responsive Web Application UI (Tailwind CSS + Chart.js)
├── run_viewer.sh           # Executable Launcher Script
└── README.md               # Documentation
```

---

## 🚀 Getting Started

### Prerequisites
- **Python 3.8+** (Uses standard Python libraries: `sqlite3`, `zipfile`, `xml.etree`, `urllib`, `http.server`). **No external `pip install` required!**

### Starting the Server

#### On Linux / Mac:
Open terminal in the project directory and run:
```bash
./run_viewer.sh
```
*(Or run: `python3 server.py 8000`)*

#### On Windows:
Double-click **`run_viewer.bat`** (or run `python server.py 8000` in Command Prompt / PowerShell).

Then open your browser and navigate to:
👉 **[http://localhost:8000](http://localhost:8000)**

---

## 💻 How to Use on Another Computer

### Option A: Transfer the Folder (USB / Email / Drive)
Since the program has **zero third-party dependencies**, you can simply copy the entire `Tally analysis` folder to any Linux, Mac, or Windows computer:
1. Copy the `Tally analysis` folder to the target computer.
2. Make sure Python 3 is installed on that computer.
3. Run `./run_viewer.sh` (Linux/Mac) or double-click `run_viewer.bat` (Windows).
4. Open **`http://localhost:8000`** in any web browser!

### Option B: Share Across Local Network (Wi-Fi / Office LAN)
If both computers are on the same Wi-Fi or office network, you don't even need to move the files!
1. Find your IP address on the host computer (e.g. `192.168.1.50`).
2. Start the server on the host computer (`python3 server.py 8000`).
3. On any other computer/phone/tablet on the same Wi-Fi, open:
   👉 **`http://192.168.1.50:8000`**

---

## 🛑 Stopping the Server & Cooling System

To free system resources and cool down your computer when finished:

1. **In Terminal**: Press `Ctrl + C` in the window running the server.
2. **Or via Command**: Run `fuser -k 8000/tcp` from any terminal.

---

## 💡 How GSTR-2B Matching Works

| Status | Badge | Description |
| :--- | :---: | :--- |
| **MATCHED** | `MATCHED` | Invoice Number, Supplier, and Tax/Total Amount match within $\pm\text{₹}2$. |
| **SHIFTED BILL** | `SHIFTED BILL` | Supplier GSTIN and Amount match, but Tally Voucher Number differs from GSTR-2B. |
| **TAX MISMATCH** | `TAX MISMATCH` | Invoice Number matches, but Tax/Value differs beyond $\pm\text{₹}2$. |
| **MISSING IN TALLY** | `MISSING TALLY` | Invoice reported in GSTR-2B, but not booked in Tally (Unclaimed ITC). |
| **MISSING IN 2B** | `MISSING 2B` | Voucher booked in Tally, but missing in GSTR-2B (Risk of ITC reversal). |

---

## 📊 CSV Export

Click **Export Reco CSV** in the GSTR-2B tab to download a complete CSV report containing:
- Status & Audit Remarks
- Supplier GSTIN & Legal Name
- Invoice Number (2B & Tally)
- Invoice Date & 2B Return Period
- GSTR-1 Filing Date
- Tax Amount (2B & Tally)
- Tax Difference (`2B - Tally`)
- Total Invoice Value (2B & Tally)

---

## 📜 License & Privacy

This software runs entirely locally on your machine. No financial data, vouchers, or GST records are ever transmitted over the network.
