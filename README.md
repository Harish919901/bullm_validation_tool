# Excel Validation Tool

A comprehensive validation tool for **Quote Win** and **BOM Matrix** Excel files with a modern React frontend and FastAPI backend.

## Features

- **Two Validator Types**: Quote Win (4 Rules) and BOM Matrix (18 Rules)
- **Modern Web UI**: React + Vite + Tailwind CSS
- **REST API Backend**: FastAPI with automatic OpenAPI documentation
- **Visual Dashboard**: Charts, statistics, and detailed results
- **Export Options**: CSV and Excel report generation
- **File Output**: Save validated files with highlighted issues and comments

---

## Project Structure

```
QW-Validation/
├── backend/                 # FastAPI Backend
│   ├── main.py              # Entry point
│   ├── requirements.txt     # Python dependencies
│   ├── models/              # Pydantic schemas
│   ├── routers/             # API endpoints
│   └── validators/          # Validation logic
│       ├── quote_win.py     # Quote Win validator
│       └── bom_matrix.py    # BOM Matrix validator
│
├── frontend/                # React Vite Frontend
│   ├── src/
│   │   ├── components/      # UI components
│   │   ├── hooks/           # Custom React hooks
│   │   └── api/             # API client
│   ├── package.json
│   └── vite.config.js
│
└── README.md
```

---

## Quick Start

### 1. Start Backend

```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8000
```

Backend runs at: http://localhost:8000
API Docs: http://localhost:8000/docs

### 2. Start Frontend

```bash
cd frontend
npm install
npm run dev
```

Frontend runs at: http://localhost:5173

---

## Quote Win Validation (4 Rules)

Validates Quote Win Excel files for CAM with the following checks:

| Rule | Name | Description |
|------|------|-------------|
| **Rule 1** | Header Validation | Validates all required headers in summary row (Row 12) and main header row (Row 16) |
| **Rule 2** | Project Name Validation | Ensures project name matches between header and data section |
| **Rule 3** | Filter Validation | Checks that no filters are applied to the spreadsheet |
| **Rule 4** | Award Validation | Validates each unique part number has Award = 100 for each Award column |

### Quote Win Template Configuration

- Summary header row: 12
- Main header row: 17
- Data start row: 18
- Project row: 3

### Required Headers

**Static Headers:**
- Project, Part Number, Part Description, Commodity, Mfg Name, Mfg Part Number
- Currency (Original), Supp Name, Pkg Qty, MOQ, Lead Time, Part Qty
- Corrected MPN, Long Comment, Price Type, No Bid Reason, Short Comment
- NCNR, RFQ Number, Eff Date, Exp Date, Quote Validity, Part Status

**Dynamic Headers (Pattern-based):**
- `Cost #X (Conv.)`, `Price (Original) #X`, `Awarded Volume #X`, `Award #X`, `Source #X`

---

## BOM Matrix Validation (18 Rules)

Validates BOM Matrix Excel files with comprehensive checks across multiple sheets:

| Rule | Name | Description |
|------|------|-------------|
| **Rule 1** | Header Validation | Validates `Ext Part Vol Price (Splits) #{X} (Conv.)` headers in CBOM VL sheets |
| **Rule 2** | Quoted MPN Validation | Validates `Quoted MPN` presence under `Corrected MPN Mentioned` section |
| **Rule 3** | Currency (CBOM) | Validates currency formatting in Ext Price and Ext Part Vol Price columns |
| **Rule 4** | MOQ Cost % | Checks percentage values appear above `MOQ Cost` headers |
| **Rule 5** | Currency (Ex Inv) | Validates currency symbols in Ex Inv VL sheets price columns |
| **Rule 6** | Net Ordering qty | Ensures `Net Ordering qty` header exists in Ex Inv sheets |
| **Rule 7** | Currency (A CLASS) | Validates currency in Cost, Ext Price, and Ext Vol Cost columns |
| **Rule 8** | Quoted MFR | Validates `Quoted MFR` presence in Manufacturer Name Mismatch section |
| **Rule 9** | NRFND | Checks for `#. NRFND` sections with values |
| **Rule 10** | Currency (BOM MATRIX) | Validates currency in Unit Price, Grand Total, Net Excess Cost, VL-{X} columns |
| **Rule 11** | CBOM Proto Sheet | Validates Proto sheet existence based on Proto headers |
| **Rule 12** | Proto Volume Summary | Ensures Proto Volume header matches Proto specification |
| **Rule 13** | Proto Pricing Notes | Validates `#. Proto Pricing No Cost` sections |
| **Rule 14** | Ex Inv Proto Sheet | Validates Ex Inv Proto sheet existence |
| **Rule 15** | Proto Columns BOM | Validates Proto Qty and Proto Price column presence |
| **Rule 16** | Serial Number Standardization | Validates serial numbers in Missing Notes are sequential (1, 2, 3...) |
| **Rule 17** | Price Validity Date Format | Validates Effective Date column uses date format (not weeks) |
| **Rule 18** | Uncosted Parts Count | Validates part count matches between Missing Notes and CBOM |

### BOM Matrix Sheets Validated

- `7.0 CBOM VL-{X}` sheets
- `Ex Inv VL-{X}` sheets
- `A CLASS PARTS` sheet
- `BOM MATRIX` sheet
- `Missing Notes` sheet

---

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/api/validate` | Upload and validate Excel file |
| `GET` | `/api/rules/{validator_type}` | Get rules information |
| `POST` | `/api/export/csv` | Export results to CSV |
| `POST` | `/api/export/excel` | Export results to Excel |
| `POST` | `/api/save` | Save validated file (Quote Win only) |

---

## Output

The tool generates:
- **Dashboard**: Visual statistics with pass/fail charts
- **Detailed Results**: Grouped by rule with status badges
- **Export Reports**: CSV and Excel formats
- **Validated File**: Original file with yellow highlights and comments on issues

---

## Tech Stack

**Frontend:**
- React 18
- Vite
- Tailwind CSS
- Recharts (charts)
- Axios (HTTP client)

**Backend:**
- Python 3.10+
- FastAPI
- Pydantic
- openpyxl

---

## License

MIT License
