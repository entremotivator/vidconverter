# 🏡 Videmi – Booking to CSV

Convert client reservation `.xlsx` spreadsheets into website-ready import CSVs.

---

## Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the app
```bash
streamlit run app.py
```

The app opens at **http://localhost:8501**

---

## What it does

Upload one or more client `.xlsx` booking files. The app auto-detects and displays:

| Panel | Info shown |
|---|---|
| **Properties & Settings** | Property names, addresses, clean hours, stay-over hours, Keys / Codes / Amenities / Laundry toggles, Check-in & Check-out times |
| **Clean Types** | Full legend table — code, type name, and description (CI, SO, CO/CI, FU, DC, COC) |
| **Amenities & Linens** | Amenity item list with quantities, linen & towel details per guest |
| **Preview & Export** | Live month preview with status badges, toggle empty rows, download options |

---

## Export options

Per client:
- **📄 Single month CSV** — one sheet at a time
- **📦 All months combined** — single flat CSV
- **🗜️ All months ZIP** — one CSV per month, zipped

Multi-client (when 2+ files uploaded):
- **🗂️ Master ZIP** — all clients, all months, organised in subfolders

---

## Output CSV format

Exact website import format:

```
Client Name: | Property list | Reservation ID | DATE: | VILLA: | TYPE CLEAN: |
PAX: | START TIME: | END TIME: | STATUS: | LAUNDRY : | Key: | Code: |
Ameneties: | COMMENTS: | QB SHIFT ID | LAST SYNC
```

---

## Expected spreadsheet structure

Each `.xlsx` file should contain:

- **A client info/profile sheet** — named anything containing *client*, *profile*, *info*, or *general*
  - Row with `Check-out time:` → value
  - Row with `Check-in time:` → value
  - Row with `Keys Yes/No:` → value
  - Row with `Codes Yes/No:` → value
  - Row with `Amenities Yes/No:` → value
  - Row with `Laundry Services Yes/No:` → value
  - `Villas/Appertments Name:` section with property rows
  - `Type of cleans:` legend in columns E–G
  - `List of Amenities:` section
  - `Linnens:` section

- **Monthly booking sheets** — named with a 3-letter month abbreviation + 2-digit year:
  `Jan26`, `Feb26`, `Mrt26`, `Apl26`, `Mei26`, `Jun26`, `Jul26`, `Aug26`, `Sep26`, `Oct26`, `Nov26`, `Dec26`
  (Dutch abbreviations supported: Mrt, Apl, Mei, Okt)

  Each sheet must have a header row containing `DATE:`, `VILLA:`, `TYPE CLEAN:`, `PAX:`, `START TIME:`, `RESERVATION STATUS:`, `LAUNDRY:`, `COMMENTS:`

---

## Notes

- Villa names are auto-normalised: `App A` → `Apartment A`
- Date format in output: `D/M/YYYY`
- Works with any number of clients simultaneously
- Empty (NONE) rows are preserved in exported CSVs to match the import format exactly
