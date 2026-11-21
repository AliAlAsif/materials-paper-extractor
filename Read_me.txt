# Instructions for the Paper Data Extractor v2

## What this script does

It scans every PDF in a folder and automatically pulls out:

| Output | What's in it |
|---|---|
| `extraction_results.xlsx` | **Open this first.** 5 sheets: Summary, Values, Property tables, Table captions, Figures |
| `summary.csv` | One row per PDF: status, DOI, page count, table/figure/value counts |
| `values_found.csv` | Sentences that mention a property (yield strength, UTS, hardness…) **and** contain a number+unit |
| `property_table_rows.csv` | Rows of tables that look like property tables (with `table_no`, full `table_title`, header, engine used) |
| `table_captions.csv` | Every "Table N" with its full title and page |
| `figures_index.csv` | Every figure, its caption, page and the PNG filename (feed to WebPlotDigitizer) |
| `tables/` | Every detected table as its own CSV file; the first lines give the table number, full title, page and source |
| `table_images/` | A PNG of each table with the detected caption and paper reference |
| `figures/` | Cropped figure PNGs with the detected caption and paper reference, named `file_p3_Fig_2.png` |

PNG headings show the source PDF filename, PDF page, and the paper title and DOI when available. A table without a matched caption is marked as such. Headings sit above the cropped content. Set `ADD_CROP_HEADINGS = False` to save plain crops.

Tables and figures are kept **separate** — figures are never cut from table areas.

---

## Setup (one time)

1. Install Python 3.10+ if you don't have it.
2. Open a terminal / PyCharm terminal in the folder containing this script and run:

```bash
python -m pip install -r requirements.txt
```

## Folder layout

```
your_project/
├── extract_pdfs.py           ← this script
├── pdfs/                     ← put your PDFs here (sub-folders are OK)
│   ├── paper1.pdf
│   └── subfolder/paper2.pdf
└── extracted/                ← created automatically on first run
```

## Run it

**In PyCharm:** right-click the file → **Run 'extract_pdfs'**

**In a terminal:**
```bash
python extract_pdfs.py
```

You'll see a progress line per PDF:
```
[3/25] subfolder/paper3.pdf    ok                  tables=4 (property tables=2) figures=6 values=18
```

---

## How to use the results (important!)

⚠️ **This is a helper, not an oracle.** Always verify against the PDF.

1. **Open `extraction_results.xlsx` → Summary sheet.**
   - `status = ok` → good.
   - `status = NO TEXT (scanned PDF…)` → the PDF is an image; the script found nothing. You must read it by hand (or OCR it first).
   - `status = ERROR …` → the file is corrupted or password-protected.
   - `ok, but N scanned page(s)` → some pages are images; results for those pages are missing.

2. **Values sheet** — a *sentence* that mentions a property *and* has a number+unit. Check:
   - `alloy_in_sentence` — which material the value belongs to (may be empty if the alloy is named in the previous sentence).
   - The `sentence` column — read the full context. The regex can pick up wrong numbers.
   - `values_found` may contain several numbers separated by `;` — pick the right one.

3. **Property tables sheet** — rows of tables judged to contain property data.
   - `table_no` is the real table number from the paper (e.g. `Table 3`), not the detection order.
   - `table_title` is the full caption title. `caption` repeats the number + title together.
   - `engine`:
     - `pymupdf-lines` → table had ruling lines. **Most reliable.**
     - `pymupdf-textmode` / `pdfplumber-textmode` → table had no lines, reconstructed from text position. **Check carefully** — columns may be misaligned.
   - Columns `c1…c14` hold the cells. The `header` field shows the top row.

4. **Table images** — open `table_images/` and compare the PNG with the CSV of the same name. This is the fastest way to spot a bad extraction. The matching `tables/*.csv` starts with its table number and title.

   If a table's title is empty in the sheet, the PDF text did not contain a readable "Table N" caption near it; open the table image and read the title from the PDF.

5. **Figures** — open `figures/` to see the crops. Check they aren't cut off; then drag the PNG into **WebPlotDigitizer** to digitize curves. `figures_index.csv` gives you the caption and page for each.

6. **Table captions sheet** — `table_no`, `table_title`, and the combined `caption` for every "Table N" the script found. If a caption is missing here, the table wasn't detected at all (go look at the PDF page).

---

## Tuning the settings

Edit the block near the top of the file:

```python
PDF_DIR = "pdfs"            # input folder
OUT_DIR = "extracted"       # output folder
MAX_COLS_SAVED = 14         # max columns kept in Property tables sheet
SAVE_FIGURES = True         # turn off to skip figure cropping
SAVE_TABLE_IMAGES = True    # turn off to skip table PNGs
ADD_CROP_HEADINGS = True    # put the detected caption and paper reference above each crop
CSV_COMMENT_HEADER = True   # put table number + title at the top of each tables/*.csv
FIG_DPI = 200               # figure PNG resolution (200 = good for WebPlotDigitizer)
TABLE_DPI = 150             # table PNG resolution
```

To recognise **new properties** (e.g. "creep rate", "corrosion rate"), add a regex to the `PROPS` dict:

```python
"creep rate": r"creep\s+(rate|strain)",
"corrosion rate": r"corrosion\s+rate",
```

To recognise **new units** (e.g. `μm`, `nm`), extend the `UNIT` regex — it's one long string near the top.

---

## Tips & known limitations

- **Scanned PDFs:** the script cannot read them. OCR them first (e.g. with Adobe, `ocrmypdf`, or ABBYY) and re-run.
- **Two-column layout:** figure cropping works best on single-column and clean two-column journals; occasional crops may include neighbouring text — re-run on that page or trim the PNG manually.
- **Regex false positives:** a sentence like *"the yield strength increased by 20 %"* will appear in Values with `20 %`. Cross-check with the table sheets and the PDF.
- **Numbers with commas** (European decimal: `8,2 MPa`) are captured, but check the decimal separator in the source.
- **Sub-folders** inside `pdfs/` are searched recursively; the relative path is kept in the `file` column so you know where each came from.
- **Re-runs** replace the workbook and index CSVs. Use a fresh `--out-dir` for a new batch to keep it free of earlier crop files.
- If `pandas`/`openpyxl` are missing the script still writes all the CSVs — just no Excel file. It prints a warning.

---

## Quick workflow per paper

1. Run the script over the whole `pdfs/` folder.
2. In the **Summary** sheet, sort by status — fix/handle the non-`ok` files manually.
3. For each `ok` paper: read **Values** sheet, verify each number against the PDF, copy the good ones into your own data sheet.
4. For property **tables**: open the matching `table_images/*.png` next to the `property_table_rows.csv` row, copy values across, recording the caption + page for citation.
5. For **figures**: open the PNG, load into WebPlotDigitizer, digitize, and cite `figures_index.csv` (page + caption).

That's it — the script does the boring 80 %, and you do the verifying 20 % that actually matters.
