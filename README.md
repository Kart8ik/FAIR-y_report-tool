## FAIR-y

Professional desktop tool for placing numbered balloons on PDF engineering drawings, capturing dimensional inspection data (characteristic/requirement/tolerance/equipment), and generating FAIR (First Article Inspection Report) documentation with full project management capabilities.

### Core Features

#### PDF Annotation
- Load multi-page engineering drawings with smooth cached rendering
- Place numbered balloons via right-click with immediate visual feedback
- Two-point mode: draw connector lines from detail to balloon
- Pan (left-click drag) and zoom (mouse wheel) with real-time preview
- Undo last balloon on current page (Ctrl+Z)
- Navigate pages with arrow keys or toolbar buttons

#### Data Entry
- Rich dropdown lists for characteristic designators and equipment types
- Capture: Characteristic, Requirement, -Tolerance, +Tolerance, Equipment
- Smart placeholder system: remembers your last entries across balloons
- Gray placeholders auto-fill, press Tab to accept or type to replace
- Speeds up repetitive data entry significantly

#### Project Management (NEW)
- **Save/Load Projects**: Work saved as `.fairy` files (JSON format)
- **Session Persistence**: Automatically reopens your last saved project on startup
- **Smart Close Dialog**: Warns about unsaved changes with 3 options:
  - Save (silent if path exists, Save As if new)
  - Don't Save
  - Cancel
- **Dirty State Tracking**: Only warns when actual project data changes (not PDF/Report exports)
- **Missing PDF Recovery**: If project's PDF is moved, prompts to locate it

#### Export Options
- **Bubbled PDF**: Original drawing with overlaid balloons and connectors
- **FAIR Report**: Auto-generated Excel using `FORMAT.xlsx` template
  - Dynamically inserts rows for any number of balloons
  - Preserves template styling, merges, and row heights
  - Fills columns: Page, Balloon #, Characteristic, Req, -Tol, +Tol, Equipment
- **Project Files**: Save work-in-progress as `.fairy` for later

### Controls & Shortcuts

#### Toolbar Buttons
- **Open PDF** - Load a new drawing (clears project)
- **Open Project** - Resume saved work from `.fairy` file
- **Prev/Next Page** - Navigate multi-page drawings
- **Undo** - Remove last balloon on current page
- **Save PDF** - Export bubbled drawing
- **Save Report** - Generate FAIR Excel report
- **Save Project** - Save current work as `.fairy` file
- **Help** - Show keyboard shortcuts

#### Mouse Controls
- **Right-click**: Place balloon (or start/end connector in two-point mode)
- **Left-click + drag**: Pan the canvas
- **Mouse wheel**: Zoom in/out centered on cursor
- **Double-click** (in list): Edit balloon data
- **Ctrl+T**: Toggle two-point mode (balloon with/without line)

#### Keyboard Shortcuts
**File Operations:**
- `Ctrl+O` - Open PDF
- `Ctrl+P` - Open Project (.fairy)
- `Ctrl+Shift+P` - Save Project (.fairy)
- `Ctrl+S` - Save PDF
- `Ctrl+Shift+S` - Save Report
- `Ctrl+Q` - Exit (with unsaved changes warning)

**Editing:**
- `Ctrl+Z` - Undo last balloon
- `Enter` - Edit selected balloon
- `Delete` - Delete selected balloon

**Navigation:**
- `←/→` - Previous/Next page
- `↑/↓` - Move selection in balloon list (auto-highlights)

**View:**
- `Ctrl+=/+` - Zoom in
- `Ctrl+-` - Zoom out
- `Shift+↑/↓` - Increase/Decrease balloon size

**Other:**
- `Ctrl+/` - Show keyboard shortcuts
- `Escape` - Close any popup

### Data Flow

1. **Open PDF** → Starts fresh session, clears all state
2. **Open Project** → Loads PDF + balloon data, restores session
3. **Right-click** → Places balloon → Opens data entry popup → Saves or discards
4. **Save Project** → Stores balloon data + PDF reference as `.fairy` JSON
5. **Save PDF** → Renders balloons onto PDF copy
6. **Save Report** → Generates Excel from template with all balloon data
7. **Close App** → Saves session state, warns if unsaved changes

### Project File Format

Projects are saved as human-readable JSON (`.fairy` extension):

```json
{
  "version": 1,
  "pdf": {
    "path": "C:/Engineering/Bracket_Rev3.pdf",
    "page_count": 3
  },
  "balloons": [
    {
      "page": 0,
      "no": 1,
      "x": 152.34,
      "y": 420.91,
      "r": 6,
      "char": "Length",
      "req": 120.0,
      "neg": 0.1,
      "pos": 0.2,
      "equip": "Digital vernier caliper",
      "start_x": 100.0,
      "start_y": 200.0
    }
  ]
}
```

**Benefits:**
- Readable and debuggable
- Version control friendly (Git diff works)
- No vendor lock-in
- Won't corrupt your PDF files
- Easy to migrate/script if needed

### Session Persistence

The app automatically remembers your last saved project:
- State stored in: `%APPDATA%\FAIR-y\state.json`
- On startup: attempts to reopen last project
- If PDF is missing: shows notification and starts fresh
- If project is deleted: silently starts fresh
- Crash-resistant: state saved immediately after every save

### Requirements

**Python Dependencies:**
```bash
pip install pymupdf pillow openpyxl
```
- `pymupdf` (fitz) - PDF rendering and manipulation
- `tkinter` - GUI framework (usually included with Python)
- `Pillow` - Image processing
- `openpyxl` - Excel file generation

**Template Files:**
- `FORMAT.xlsx` - FAIR report template (required for Excel export)
- `app-icon.ico` - Application icon (optional)

### Running

**Development:**
```bash
python test.py
```

**PyInstaller Build:**
```bash
pyinstaller --onefile --windowed ^
  --add-data "FORMAT.xlsx;." ^
  --add-data "app-icon.ico;." ^
  --icon=app-icon.ico ^
  --name FAIR-y ^
  test.py
```

The `resource_path()` function handles both development and bundled execution.

### Technical Notes

#### State Management
- **Project state**: Balloon positions, data, PDF reference
- **Session state**: Last opened project path (auto-restore)
- **UI state**: Zoom, pan, selection (not saved - resets on load)
- **Dirty flag**: Tracks unsaved changes to prevent data loss

#### Smart Placeholders
- Remembers last 5 values: char, req, neg, pos, equip
- Shows as gray text in new balloon popups
- Tab accepts placeholder, any key replaces it
- Cache persists for entire session
- Empty fields don't overwrite cache

#### PDF Coordinate System
- Balloons stored in PDF coordinate space (not screen pixels)
- Survives zoom/pan operations
- Connector lines computed relative to balloon radius
- Proper rendering at any zoom level

#### Export Logic
- **PDF**: PyMuPDF draws circles + text + lines onto page copy
- **Excel**: Dynamic row insertion before footer block, style preservation
- **Project**: JSON serialization with version field for future migration

#### Error Handling
- Corrupt `.fairy` files → error dialog, safe abort
- Missing PDF on load → locate dialog or cancel
- Invalid balloon data → skipped with warning
- Page count mismatch → warning, continues anyway
- Silent failures for non-critical state saves

### Architecture

- **Single-file application**: `test.py` (~1600 lines)
- **No database**: All data in memory + `.fairy` files
- **Cross-platform**: Works on Windows, macOS, Linux
- **Stateless exports**: PDF and Excel generation are pure functions
- **Clean separation**: UI ↔ Business Logic ↔ File I/O

### Use Cases

1. **Manufacturing QA**: First article inspection documentation
2. **Engineering**: Dimensional verification of prototypes
3. **Supplier Audit**: Part qualification and certification
4. **Continuous Work**: Save mid-inspection, resume later
5. **Collaboration**: Share `.fairy` files with team (just ensure PDF access)

### Future Enhancements

The `.fairy` JSON format enables:
- Autosave (periodic background saves)
- Version diffing (Git integration)
- Cloud sync (Dropbox, OneDrive)
- Multi-user workflows (merge conflicts)
- Audit trails (change history)
- Batch processing (scripted reports)

---

**Made with Python + Tkinter**  
No databases, no servers, no subscriptions. Just a solid desktop tool that does one thing well.
