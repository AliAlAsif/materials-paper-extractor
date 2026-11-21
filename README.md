# Materials Paper Extractor

Screen materials research PDFs for candidate thermal and mechanical property data. The extractor runs locally and records the source paper and page alongside each result.

## Install and run

Use Python 3.10 or newer.

```sh
python -m pip install -r requirements.txt
python extract_pdfs.py --pdf-dir pdfs --out-dir extracted
```

Input folders may contain subfolders. For a new batch, choose a fresh output directory to keep its files together.

## Verify extracted results

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
