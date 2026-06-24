# invoice-extractor
Batch extract key invoice data from PDFs and generate an Excel ledger with automatic color‑coded warnings – no more manual typing.
# 📄 Invoice Extractor · Bulk PDF Invoice Parser

> Stop copying invoice numbers and amounts one by one. This tool turns a folder full of PDF invoices into a clean, structured Excel ledger – with built‑in date alerts.

---

## ✨ What It Does

Drop a bunch of invoice PDFs into a folder, run the script, and get a ready‑to‑use Excel file containing:

- **Invoice Number**
- **Invoice Date**
- **Total Amount (VAT‑inclusive)**
- **Supplier Name**

Plus: **any invoice dated before a threshold you set (e.g., 2026‑06‑16) gets its amount cell highlighted in yellow** – so you can instantly spot older invoices that need priority handling.

---

## 🎯 Why You’ll Love It

| The Pain | How This Tool Helps |
|----------|----------------------|
| Manually opening each PDF and copy‑pasting data is tedious and error‑prone | Batch process all PDFs in one go – one command, all done |
| Invoices come in different layouts and formats | Multi‑level regex + smart fallbacks handle various styles |
| Month‑end reconciliation requires filtering by date – a huge time sink | Automatic colour‑coding highlights critical entries at a glance |
| Your finance team keeps asking for structured data | Generates a standard Excel file that opens in Excel or WPS |

---

## 🚀 Perfect For

- Finance teams consolidating input invoices at month‑end
- Admin staff archiving expense reports
- Procurement departments verifying supplier invoice ledgers
- Anyone who regularly converts PDF invoices into spreadsheet form

---

## 💡 Key Features

- **Zero configuration** – place the script in your PDF folder and run it; no need to hardcode paths
- **Robust extraction** – if a PDF is just a scanned image, the script falls back to the filename, so you never lose supplier info
- **Smart date highlighting** – automatically marks old invoices so you can prioritise
- **Open source & customizable** – tweak the regex patterns to match your specific invoice layouts

---

## 📊 Sample Output

| Invoice Number | Invoice Date     | Total Amount | Supplier Name                     |
|----------------|------------------|--------------|-----------------------------------|
| 2632...4886    | 2026‑05‑22       | 175,868.36   | **Apex Hydraulics Co., Ltd.**     |
| 2632...8941    | 2026‑05‑12       | 🟡 11,390.40 | **NexGen Fluid Power Inc.**       |
| 2643...2844    | 2026‑05‑30       | 7,322.40     | **Summit Machinery & Hydraulics** |

*(🟡 Yellow highlight indicates the invoice date is before the warning threshold.)*

---

> If you’re tired of manual data entry for invoices, give this repository a ⭐ – and help others escape the same drudgery.
