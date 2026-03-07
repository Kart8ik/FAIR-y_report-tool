## FAIR-y

Professional desktop tool for placing numbered balloons on PDF engineering drawings, capturing dimensional inspection data, and generating FAIR documentation.

This repo has two company-specific variants of the same app:
- `test.py`: Company A workflow using `FORMAT.xlsx`
- `test2.py`: Company B (WORBYN) workflow using `FORMAT_WORBYN_2.xlsx`

Both variants share the same UI patterns, project file format (`.fairy`), and annotation workflow.

### Core Features

#### PDF Annotation and Review
- Open and render multi-page PDFs with page caching for smooth navigation
- Add numbered balloons with right-click
- Two-point mode: place a start point and balloon end point with connector line
- Pan with left-click drag
- Zoom with mouse wheel, keyboard, or zoom slider
- Rotate page view left/right from toolbar or keyboard
- Undo last balloon on current page (`Ctrl+Z`)

#### Balloon Metadata Entry
- Per-balloon fields: Zone, Characteristic, Requirement, `-Tol`, `+Tol`, Equipment
- Characteristic and Equipment dropdown catalogs
- Placeholder memory for repeated entries (fast data entry)
- Inline edit/delete from list view (double click, Enter, Delete)
- Temporary highlight when editing selected balloon

#### Color Features (New)
- `Pick Balloon Color` toolbar action
- Color swatch preview in toolbar
- Balloon preview updates with current zoom and selected color
- Balloon color stored per balloon
- Selected default balloon color restored when loading a project

#### Header Management
- Dedicated Headers popup for report metadata
- Header values stored inside `.fairy` project files
- Missing header confirmation before report export
- `Ctrl+H` shortcut to open headers quickly

#### Project Management
- Save and load projects as `.fairy` JSON
- Auto-restore last project on startup (AppData state file)
- Dirty state tracking with unsaved-change close prompt
- Save current project silently when path already exists
- Relink PDF if original path is missing during project load

#### Export Options
- Ballooned PDF export (includes circles, numbers, connectors, and colors)
- Excel FAIR report export with dynamic row insertion and style preservation
- Automatic tolerance-derived lower/upper values in list/report

### Controls and Shortcuts

#### Mouse
- Right-click: add balloon (or start/end in two-point mode)
- Left-click drag: pan canvas
- Mouse wheel: zoom centered on cursor
- Double-click list row: edit balloon

#### Keyboard
- `Ctrl+O`: Open PDF
- `Ctrl+P`: Open Project
- `Ctrl+Shift+P`: Save Project
- `Ctrl+S`: Save PDF
- `Ctrl+Shift+S`: Save Report
- `Ctrl+H`: Open Headers
- `Shift+C`: Pick Balloon Color
- `Ctrl+T`: Toggle two-point mode
- `Ctrl+Z`: Undo last balloon on current page
- `Shift+Left` / `Shift+Right`: Rotate page view
- `Ctrl+=` / `Ctrl++`: Zoom in
- `Ctrl+-`: Zoom out
- `Shift+Up` / `Shift+Down`: Increase/decrease balloon size
- `Enter`: Edit selected balloon in list
- `Delete`: Delete selected balloon in list
- `Ctrl+/`: Show shortcuts window
- `Ctrl+Q`: Close app (with unsaved-changes flow)
- `Esc`: Close active popup

### `.fairy` Project Format

Projects are JSON with versioned structure. Key sections include:
- `pdf`: source path and page count
- `view`: rotation and selected default balloon color
- `headers`: report metadata for active variant
- `balloons`: geometry, annotation values, optional connector points, color

Example:

```json
{
  "version": 1,
  "pdf": {
    "path": "C:/path/sample.pdf",
    "page_count": 2
  },
  "view": {
    "rotation": 0,
    "selected_balloon_color": "#0095ff"
  },
  "headers": {
    "part_number": "PN-001"
  },
  "balloons": [
    {
      "page": 0,
      "no": 1,
      "x": 305.33,
      "y": 156.66,
      "r": 6,
      "zone": "A1",
      "char": "Length",
      "req": 120,
      "neg": 0.1,
      "pos": 0.2,
      "equip": "Digital vernier caliper",
      "color": "#ff0000",
      "start_x": 280.1,
      "start_y": 190.4
    }
  ]
}
```

### Requirements

Install dependencies:

```bash
pip install pymupdf pillow openpyxl
```

Notes:
- `tkinter` is required and usually bundled with standard Python distributions.
- Keep the correct template file for the variant you run.

### Run

Company A:

```bash
python test.py
```

Company B (WORBYN):

```bash
python test2.py
```

### Build (PyInstaller)

Example build for `test.py`:

```bash
pyinstaller --onefile --windowed ^
  --add-data "FORMAT.xlsx;." ^
  --add-data "FORMAT_WORBYN_2.xlsx;." ^
  --add-data "app-icon.ico;." ^
  --icon=app-icon.ico ^
  --name FAIR-y ^
  test.py
```

For `test2.py`, keep the same options and replace the entry script with `test2.py`.

### Session Persistence

App state path:
- `%APPDATA%\FAIR-y\state.json`

Behavior:
- Last saved project path is tracked
- App tries to restore that project at startup
- If the project/PDF is missing, app starts fresh safely

### Notes

- Rotation and selected balloon color are now part of saved/restored project view state.
- Guard checks prevent rotate/two-point actions before any PDF is opened.
- Excel export logic differs between `test.py` and `test2.py` only in template layout and header mapping.
