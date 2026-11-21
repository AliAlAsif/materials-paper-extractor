# Materials Paper Extractor

Extract tables and figures from research PDFs, with specialized screening for thermal and mechanical properties in materials science. The extractor runs locally and records the source paper and page alongside each result.

## Install and run

1. Clone or download this repository and open a terminal in the folder containing `extract_pdfs.py`.
2. Create a folder named `pdfs` next to `extract_pdfs.py` and put your research PDFs inside it. Subfolders are supported.
3. With Python 3.10 or newer, run these commands from the project folder.

```sh
python -m pip install -r requirements.txt
python extract_pdfs.py --pdf-dir pdfs --out-dir extracted
```

4. Open `extracted/extraction_results.xlsx` and start with the **Summary** sheet. Use **Property tables** and **Figures** to open the linked crops.

If your PDFs are already in another folder, pass its path to `--pdf-dir`. Put paths containing spaces in quotes. For a new batch, choose a fresh `--out-dir` to keep its files together.

## Verify extracted results

The Excel workbook contains five sheets.

| Sheet | What it contains |
| --- | --- |
| Summary | One row per PDF with its extraction status and counts. |
| Values | Sentences containing a recognized materials property and candidate values with units. |
| Property tables | Rows from tables flagged by the property rules, with links to the corresponding PNG and CSV. |
| Table captions | Detected table numbers, captions, and PDF page numbers. |
| Figures | Detected figure captions and links to the saved PNGs. Numeric values plotted in figures require manual digitization. |

Every detected table is saved as a separate CSV in `tables/`, which you can open in Excel. This includes tables that do not appear in the **Property tables** sheet. Table and figure detection can also be used on papers from other fields. The **Values** and **Property tables** sheets use materials-specific screening rules.

- Open `extraction_results.xlsx` and check the **Summary** sheet for papers that need manual review.
- Use **Property tables** to open the matching PNG and CSV. Review notes identify reconstructed columns and missing captions.
- Use **Figures** to open figure PNGs. Keep the workbook and its sibling folders together so the relative links can resolve.
- Each image heading shows the detected caption and source PDF page. It includes the original filename and the paper title and DOI when found. Set `ADD_CROP_HEADINGS = False` in the script for plain crops.
- Confirm the specimen, test conditions, units, and values against the original paper before using them for analysis.

The extractor uses patterns and page layout. Scanned pages need OCR or manual review. Complex layouts can produce incomplete crops, mismatched captions, or misaligned cells. Excel output requires pandas and openpyxl. The script also writes CSV files.

## Try the synthetic example

```sh
python make_demo_pdf.py
python extract_pdfs.py --pdf-dir demo-input --out-dir demo-output
```

The demo contains a ruled tensile-property table, a yield-strength sentence, and an illustrative figure. All values are fictional. Journal PDFs and generated output stay local.

See [Read_me.txt](Read_me.txt) for more extraction settings.

## Ownership and license

Copyright © 2026 [ASk (@AliAlAsif)](https://github.com/AliAlAsif), project owner and maintainer.

The project source code is licensed under the [GNU Affero General Public License v3.0](LICENSE) (`AGPL-3.0-only`). Third-party dependencies and source papers retain their own terms.
