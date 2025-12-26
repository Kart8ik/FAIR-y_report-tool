## FAIR-y

Desktop Tkinter tool for placing numbered “bubbles” on a PDF part-design drawing, capturing characteristic/requirement/tolerance data, exporting a bubbled PDF and an Excel FAIR report based on `FORMAT.xlsx`.

### Features
- Load any PDF, pan/zoom, right‑click to place bubbles with immediate feedback.
- Prompt collects characteristic, requirement, -Tol, +Tol per bubble.
- Undo last bubble on the current page; navigate pages.
- Export bubbled drawing to PDF (user picks path).
- Export FAIR report: copies `FORMAT.xlsx`, inserts extra rows before the footer, preserves styles/merges/heights, fills columns (Pg, Balloon, Char, Req, -Tol, +Tol).
- PyInstaller-friendly resource lookup via `resource_path` for `FORMAT.xlsx`.

### Controls
- Buttons: Open PDF, Prev/Next page, Undo, Save PDF, Save Report.
- Mouse: right‑click adds bubble, left‑click drag pans, mouse wheel zooms.
- Bubble size slider affects drawn radius (and preview).

### Data flow
1. Open PDF → resets state.
2. Right‑click → bubble drawn immediately → popup for data → keeps or discards bubble.
3. Save PDF → redraws bubbles onto a copy of the PDF.
4. Save Report → copies template, inserts rows if more than 15 data rows, re-merges footer, restores row heights, writes data rows top-down.

### Requirements
- Python 3 with: `fitz` (PyMuPDF), `tkinter`, `Pillow`, `openpyxl`.
- Template file `FORMAT.xlsx` in the working directory (or bundled with PyInstaller; `resource_path` resolves it).

### Running
```
python test.py
```
If packaged with PyInstaller (`--onefile`), ensure `FORMAT.xlsx` is included (`--add-data FORMAT.xlsx;.`) so `resource_path` can find it.

### Notes
- Initial globals set `PDF_IN = None`; opening a PDF via the UI is required before rendering/exporting.
- The export logic preserves footer merges/heights after row insertion; the first data row is used to restyle inserted rows to keep borders/alignment consistent.

