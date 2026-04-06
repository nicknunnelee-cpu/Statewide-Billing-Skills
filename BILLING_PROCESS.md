# Statewide Interpreters — Weekly Billing Automation

## Overview
Weekly invoice billing pipeline for State Wide Interpreters Corp. Takes a batch PDF export from QuickBooks Desktop (QBD), splits it into individual invoices, classifies each by market rate type, and merges classified invoices with their corresponding market rate PDF attachment.

## File Locations
- **Scripts**: `Statewide Interpreters Automation/` (the project folder)
- **Working data**: `Nico Data/` (OneDrive-synced)
- **MR type PDFs**: `Nico Data/MR type 1.pdf` through `MR type 15.pdf`
- **Invoice workbook**: User prepares this weekly in the INV PDF folder
  - Must contain sheets: "Invoice key" and "Market Rate Key"
  - Column layout (Invoice key sheet): A=Type, B=Name, C=Name Address, D=Memo, E=Date, F=Num (Invoice Number), G=Open Balance, H=MR Type (populated by match script)
  - Header in row 1, data starts row 2

## Pipeline Steps (execute in order)

### Step 0: Preparation (User does manually)
1. Export invoices from QuickBooks Desktop as a single batch PDF.
2. Prepare the invoice workbook (.xlsx) with "Invoice key" and "Market Rate Key" sheets.
3. Create dated folder: `INV PDF {date}` inside `Nico Data/`.
4. Place the batch PDF in that folder (named like `{date} Inv to send whole.pdf`).
5. Place the workbook in the same folder.

### Step 1: Split Batch PDF into Individual Invoice PDFs
The batch PDF from QBD contains all invoices concatenated (one page per invoice).

**Tool**: Python (pypdf) — Claude runs this interactively each week.

**Method**:
- Extract text from each page, find invoice number via regex: `Invoice No.\s*\n?\s*(\d{5,6})`
- Save each page as `{InvoiceNum}.pdf` in the same INV PDF folder.
- Verify: page count should equal number of unique invoice numbers.

**Input**: `Nico Data/INV PDF {date}/{date} Inv to send whole.pdf`
**Output**: Individual PDFs in same folder, e.g. `224231.pdf`, `224232.pdf`, etc.

### Step 2: Match & Classify Rate Types
Classifies each invoice by market rate type and writes the result to Column H (MR Type) in the workbook.

**Script**: `match_invoices.py`
**Usage**: `python match_invoices.py "<path_to_workbook.xlsx>"`

**Current column mapping** (as of 4.1.26):
- MEMO_COL = Column D (index 3)
- ADDR_COL = Column C (index 2)
- NUM_COL = Column F (index 5)
- RATE_COL = Column H (writes here)
- Sheet: "Invoice key", data starts row 2

**What it does**:
- Reads "Market Rate Key" sheet for type definitions (Type 1–15), memo descriptions, and address filters.
- Matches each invoice's memo against the key: exact match → 60-char prefix match → regex fallbacks.
- Writes result to Column H: `Type N`, `?` (unresolved), `excluded`, or blank (supplementary).
- Also outputs a CSV alongside the workbook: `{workbook_name}_matched.csv`

**Market Rate Key layout** (unchanged): Col A=Type label, B=Description, C=Rate, D=MR PDF, E=Address Filter, F+=Memo descriptions

**Known type mappings**:
- Type 1: Spanish medical interpreting (hourly)
- Type 2: Panel QME / AME (Spanish), SIBTF
- Type 3: C&R and Stipulation readings
- Type 4: Deposition transcript readings
- Type 5: Half-day Spanish legal
- Type 6: Half-day Cantonese/Mandarin legal
- Type 7: Cantonese/Mandarin medical (standard)
- Type 8: Cantonese/Mandarin medical (Allied Managed Care — address filter)
- Type 9: Vietnamese medical
- Type 10: Other non-Spanish languages (Farsi, Laotian, Punjabi, Armenian, etc.)
- Type 11: Tagalog medical
- Type 12: Korean medical
- Type 13: Arabic medical
- Type 14: Tagalog legal (half/full day)
- Type 15: Arabic legal (half/full day)

### Step 3: Merge Invoices with Market Rate PDFs
Takes classified workbook + individual invoice PDFs + MR type PDFs → produces final output folder.

**Script**: `merge_invoices.py`
**Usage**:
```
python merge_invoices.py \
    "<classified_workbook.xlsx>" \
    "<INV PDF folder>" \
    "<Nico Data folder (contains MR type PDFs)>" \
    "<output Merged Inv folder>"
```

**Current column mapping** (as of 4.1.26):
- NUM_COL = Column F (index 5)
- ADDR_COL = Column C (index 2)
- RATE_COL = Column H (index 7)
- Sheet: "Invoice key", data starts row 2

**Example (4.1.26 run)**:
```
python merge_invoices.py \
    "Nico Data/INV PDF 4.1.26/4.1.26 Recovered Inv to send inv key.xlsx" \
    "Nico Data/INV PDF 4.1.26" \
    "Nico Data" \
    "Nico Data/INV PDF 4.1.26/Merged Inv"
```

**Critical invariant**: X invoices in = X invoices out (for invoices present in the workbook). Every invoice PDF gets a renamed copy in the output folder regardless of whether it has a rate type.

**What it does for each invoice PDF**:
1. Extracts invoice number from filename (first token of stem).
2. Looks up Name Address (Col C) and Rate Type (Col H) from the workbook.
3. Copies the invoice PDF to the output folder, renamed to `{NameAddress} {InvoiceNum}.pdf`.
4. If the invoice has a Rate Type (Type 1–15): merges the renamed copy IN-PLACE with the corresponding `MR type N.pdf`.
5. If no rate type (excluded, blank): renamed copy stays as-is.

**Note**: The batch PDF (`{date} Inv to send whole.pdf`) will be skipped since it has no workbook entry. This is expected.

### Step 4: Print (optional)
**Script**: `Print_PDFs.ps1` (PowerShell, runs on Windows only)
- Hardcoded to Brother MFC-L5850DW printer.
- Uses Adobe Acrobat or SumatraPDF.
- File list must be manually updated each week.

## Edge Cases and Troubleshooting

### OneDrive/Excel Corrupt xlsx
Both `match_invoices.py` and `merge_invoices.py` have built-in xlsx repair logic:
- Strategy A: Rebuild EOCD from existing central directory entries
- Strategy B: Reconstruct from local file headers (no CD at all)
- **Fix**: Close Excel before running scripts. Wait for OneDrive sync to complete.

### "?" (Unresolved) Rate Types
Service memos that don't match any type in the Market Rate Key and don't match regex fallbacks.
- **Fix**: Add new memo patterns to the Market Rate Key sheet, then rerun match.

### Filename Sanitization
Windows-illegal characters (`\ / * ? : " < > |`) are stripped from Name Address.

### Exclusions
Clients listed in Market Rate Key exclusion row ("List of Name Address values...") get `excluded` in Column H. Currently includes Farahi Law, Setareh Firm, and others.

### Invoices Not in Workbook
If a PDF exists in the INV PDF folder but has no matching invoice number in the workbook, it is skipped with a warning. This commonly happens with the unsplit batch PDF or invoices from prior billing cycles.

## Dependencies
```
pip install openpyxl pypdf
```

## Quick Reference — Typical Weekly Run
```bash
# 1. Split the batch PDF (Claude runs interactively with pypdf)
#    Extract invoice numbers from each page, save as {InvoiceNum}.pdf

# 2. Match & classify (writes Column H in workbook + CSV)
python match_invoices.py "Nico Data/INV PDF {date}/{workbook}.xlsx"

# 3. Merge
python merge_invoices.py \
    "Nico Data/INV PDF {date}/{workbook}.xlsx" \
    "Nico Data/INV PDF {date}" \
    "Nico Data" \
    "Nico Data/INV PDF {date}/Merged Inv"
```

## History
- 3.26.26: Initial pipeline built. Old layout: "Invoices report" sheet, Col H=Num, L=Address, N=Memo, T=Rate Type.
- 4.1.26: Migrated to new layout: "Invoice key" sheet, Col F=Num, C=Address, D=Memo, H=MR Type. match_invoices.py now writes directly to Column H. Output folder moved inside INV PDF folder as "Merged Inv".
