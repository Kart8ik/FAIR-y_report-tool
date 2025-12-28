## FAIR-y

Desktop Tkinter tool for placing numbered “bubbles” on a PDF part-design drawing, capturing characteristic/requirement/tolerance/equipment data, exporting a bubbled PDF and an Excel FAIR report based on `FORMAT.xlsx`.

### Features
- Load any PDF; start maximized; pan/zoom; right‑click to place bubbles with immediate feedback.
- Prompt collects characteristic designator, requirement, -Tol, +Tol, and equipment; list view shows aligned columns.
- Undo last bubble on the current page; navigate pages with smooth cached rendering for multi-page PDFs.
- Export bubbled drawing to PDF (user picks path).
- Export FAIR report: copies `FORMAT.xlsx`, inserts extra rows before the footer, preserves styles/merges/heights, fills columns (Pg, Balloon, Char, Req, -Tol, +Tol, Equip).
- PyInstaller-friendly resource lookup via `resource_path` for `FORMAT.xlsx` and icons.

### Controls / Shortcuts
- Buttons: Open PDF, Prev/Next page, Undo, Save PDF, Save Report, Shortcuts.
- Mouse: right‑click adds bubble, left‑click drag pans, mouse wheel zooms.
- Keyboard:
  - Ctrl+O open, Ctrl+S save PDF, Ctrl+Shift+S save report
  - Ctrl+Z undo
  - Arrow Left/Right page nav; Arrow Up/Down move selection in list (auto-highlights bubble)
  - Ctrl+= / Ctrl++ / Ctrl+- (incl. keypad) zoom in/out; Ctrl+/ show shortcuts
  - Shift+Up/Down change bubble size
- Bubble size slider affects drawn radius (and preview).

### Data flow
1. Open PDF → resets state and cache.
2. Right‑click → bubble drawn immediately → popup for data → keeps or discards bubble.
3. Save PDF → redraws bubbles onto a copy of the PDF.
4. Save Report → copies template, inserts rows if more than 15 data rows, re-merges footer, restores row heights, writes data rows top-down.

### Requirements
- Python 3 with: `fitz` (PyMuPDF), `tkinter`, `Pillow`, `openpyxl`.
- Template file `FORMAT.xlsx` in the working directory (or bundled with PyInstaller; `resource_path` resolves it). Optional `app-icon.ico`.

### Running
```
python test.py
```
If packaged with PyInstaller (`--onefile`), ensure `FORMAT.xlsx` is included (`--add-data "FORMAT.xlsx;." ^ --add-data "app-icon.ico;." `) so `resource_path` can find it.

### Notes
- Opens fullscreen on launch.
- Initial globals set `PDF_IN = None`; opening a PDF via the UI is required before rendering/exporting.
- The export logic preserves footer merges/heights after row insertion; the first data row is used to restyle inserted rows to keep borders/alignment consistent.
- List focus and selection are initialized so arrow keys work immediately and highlight the chosen bubble.

