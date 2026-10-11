# 💳 Buku

**Buku** is a privacy-first web application that converts encrypted Safaricom M-Pesa PDF statements into clean, structured, and audit-ready Excel spreadsheets (`.xlsx`) in seconds.

Built with Python, Streamlit, `pdfplumber`, and `openpyxl`, Buku runs 100% locally on your machine to ensure sensitive financial data never leaves your device.

---

## ✨ Key Features

* 🔒 **100% Local Data Privacy:** Your PDF statement and password are decrypted and parsed entirely on your device.
* ⚡ **Multi-Line Parsing Engine:** Flawlessly maps M-Pesa transaction patterns, handling wrapped descriptions, status flags, and numeric casting.
* 📊 **Instant Financial Analytics:** Displays high-level statement metrics including Total Inflows (`Cr`), Total Outflows (`Dr`), and Net Cash Flow.
* 🔍 **Smart Search & Filtering:** Filter your ledger instantly by keyword (names, Paybills, Tills) or exact transaction amounts.
* 📥 **Professional Excel Export:** Generates auto-fitted, styled `.xlsx` workbooks with proper number formatting (`#,##0.00`) and dark header themes.

---

## 🚀 Quick Start

### Prerequisites

Ensure you have Python 3.9+ installed on your system.

### Installation

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/Edmound776/buku.git](https://github.com/Edmound776/buku.git)
   cd buku
