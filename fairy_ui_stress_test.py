"""
FAIR-y GUI Stress Test Script

Black-box GUI automation test for the FAIR-y application using pyautogui and pygetwindow.
Simulates a real user: launches exe, opens PDF, fills headers, adds bubbles, saves outputs,
closes app, and verifies files on disk. Fails loudly on crash, blocking dialogs, or missing outputs.

Dependencies: pyautogui, pygetwindow
Install: pip install pyautogui pygetwindow

Usage: python fairy_ui_stress_test.py

Assume Windows OS. Do not run while using the computer - automation will control keyboard/mouse.
"""

from __future__ import annotations

import glob
import os
import random
import subprocess
import sys
import time
from datetime import datetime

try:
    import pyautogui
    import pygetwindow as gw
except ImportError as e:
    print("ERROR: Missing dependency. Install with: pip install pyautogui pygetwindow")
    sys.exit(1)

# Disable pyautogui fail-safe (moving mouse to corner) during automated run
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05

# =============================================================================
# CONFIGURABLE CONSTANTS - adjust paths and parameters as needed
# =============================================================================

# Path to FAIR-y.exe (PyInstaller typically outputs to dist/ or build/)
FAIR_EXE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "dist", "FAIR-y.exe"
)

# Full path to a sample PDF (multi-page recommended for stress; must exist)
SAMPLE_PDF_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "sample.pdf"
)

# Directory for test outputs (PDF, xlsx, .fairy); created if missing
OUTPUT_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "stress_test_output"
)

# Number of bubbles to add
BUBBLE_COUNT = 37

# Loop entire test this many times (1 = single run)
ITERATIONS = 1

# Base delay in seconds between major actions; increase if flaky
PAUSE_SEC = 0.3

# Clear app state before launch for reproducible clean start (avoids session restore)
CLEAR_STATE_BEFORE_LAUNCH = True

# Window title to search for
APP_WINDOW_TITLE = "FAIR-y"

# Timeouts (seconds)
LAUNCH_TIMEOUT = 15
CLOSE_TIMEOUT = 5
# ADD_BUBBLES_TIMEOUT = 120  # ~80 bubbles * ~1.5s each
ADD_BUBBLES_TIMEOUT = 500  # ~80 bubbles * ~1.5s each

# Field options for bubble popup (inspired by test.py); 5 options each for variety
ZONE_OPTIONS = ["A1", "A2", "B1", "B2", "C1"]
CHAR_OPTIONS = ["Length", "Width", "Diameter", "Slot length", "Flatness"]
REQ_OPTIONS = ["10", "25.5", "0.5", "100", "1.25"]
NEG_OPTIONS = ["0.1", "0.05", "0.2", "0.01", "0.5"]
POS_OPTIONS = ["0.1", "0.05", "0.2", "0.01", "0.5"]
EQUIP_OPTIONS = ["Vernier caliper", "Digital vernier caliper", "Digital micrometer", "Pin gauge", "CMM"]

# =============================================================================
# HELPERS
# =============================================================================


def log(msg: str) -> None:
    """Log progress to console."""
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")


def fail_loud(msg: str, screenshot: bool = True) -> None:
    """Exit with error message; optionally save screenshot for debugging."""
    log(f"FAIL: {msg}")
    if screenshot:
        try:
            path = os.path.join(OUTPUT_DIR, "failure_screenshot.png")
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            pyautogui.screenshot(path)
            log(f"Screenshot saved to {path}")
        except Exception as e:
            log(f"Could not save screenshot: {e}")
    sys.exit(1)


def wait_for_window(title: str, timeout: float) -> "gw.Window":
    """Poll for window with given title. Returns window or fails."""
    start = time.time()
    while time.time() - start < timeout:
        wins = gw.getWindowsWithTitle(title)
        if wins:
            return wins[0]
        time.sleep(0.3)
    fail_loud(f"Window '{title}' not found within {timeout}s", screenshot=True)


def ensure_output_dir() -> None:
    """Create output directory if it does not exist."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)


def type_in_file_dialog(path: str) -> None:
    """
    Type a path into a Windows file dialog (Open or Save).
    1. Address bar (top): parent dir path, Enter to navigate
    2. File name field (bottom): filename only, Enter to confirm
    """
    time.sleep(PAUSE_SEC)
    parent_dir = os.path.dirname(path)
    filename = os.path.basename(path)

    # Step 1: Focus address bar (top), type parent dir path, Enter to navigate
    pyautogui.hotkey("alt", "d")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    parent_norm = parent_dir.replace("\\", "/")
    pyautogui.write(parent_norm, interval=0.02)
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.5)  # Wait for folder to load

    # Step 2: Focus File name field (bottom), type filename only, Enter
    pyautogui.hotkey("alt", "n")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.write(filename, interval=0.02)
    time.sleep(0.2)
    pyautogui.press("enter")

def navigate_and_save_with_default(output_dir: str, turn: int = 1) -> None:
    """
    Navigate to output_dir, accept pre-filled default filename, save.
    num_enters: 3 when in different dir (first save), 2 when already in same dir.
    """
    time.sleep(PAUSE_SEC)
    if turn == 1:
        pyautogui.hotkey("alt", "d")
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        path_norm = output_dir.replace("\\", "/")
        pyautogui.write(path_norm, interval=0.02)
        time.sleep(0.2)
        for _ in range(4):
            pyautogui.press("enter")
            time.sleep(1)  # Longer wait after first (folder load)
    else:
        pyautogui.press("enter")
        time.sleep(1)


def discover_output_files(output_dir: str) -> tuple[str | None, str | None, str | None]:
    """
    Scan output_dir for most recent files matching app's default patterns.
    Returns (pdf_path, report_path, project_path); None for any missing.
    """
    patterns = [
        "FAIR_QA_drawing_*.pdf",
        "FAIR_Report_*.xlsx",
        "FAIR_Project_*.fairy",
    ]
    result = [None, None, None]
    for i, pattern in enumerate(patterns):
        matches = glob.glob(os.path.join(output_dir, pattern))
        if matches:
            result[i] = max(matches, key=os.path.getmtime)
    return result[0], result[1], result[2]


def get_canvas_region() -> tuple[int, int, int, int]:
    """
    Get approximate canvas click region from FAIR-y window.
    Canvas is below toolbar (~50px) and listbox (~120px) in vertical PanedWindow.
    Returns (x_min, y_min, x_max, y_max) in screen coordinates.
    """
    wins = gw.getWindowsWithTitle(APP_WINDOW_TITLE)
    if not wins:
        fail_loud("FAIR-y window not found when computing canvas region")
    win = wins[0]
    left, top, width, height = win.left, win.top, win.width, win.height
    # Canvas starts below toolbar + listbox; add margin for window chrome
    canvas_top = top + 170
    canvas_left = left + 20
    canvas_right = left + width - 20
    canvas_bottom = top + height - 20
    # Ensure valid region
    if canvas_bottom <= canvas_top or canvas_right <= canvas_left:
        fail_loud("Invalid canvas region computed")
    return canvas_left, canvas_top, canvas_right, canvas_bottom


# =============================================================================
# PHASE 1: LAUNCH APP
# =============================================================================


def launch_app() -> subprocess.Popen:
    """Launch FAIR-y exe, wait for window, bring to front. Returns process handle."""
    log("Phase 1: Launching app...")

    if not os.path.exists(FAIR_EXE_PATH):
        fail_loud(f"FAIR-y exe not found: {FAIR_EXE_PATH}")

    if CLEAR_STATE_BEFORE_LAUNCH:
        state_path = os.path.join(
            os.environ.get("APPDATA", ""), "FAIR-y", "state.json"
        )
        try:
            if os.path.exists(state_path):
                os.remove(state_path)
                log("Cleared state.json for clean start")
        except Exception as e:
            log(f"Warning: Could not clear state: {e}")

    proc = subprocess.Popen(
        [FAIR_EXE_PATH],
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if sys.platform == "win32" else 0,
    )
    log("Process started, waiting for window...")

    win = wait_for_window(APP_WINDOW_TITLE, LAUNCH_TIMEOUT)
    win.activate()
    time.sleep(2)  # Full load of toolbar, paned layout, etc.

    log("Phase 1 complete: App launched and ready")
    return proc


# =============================================================================
# PHASE 2: OPEN PDF
# =============================================================================


def open_pdf() -> None:
    """Open sample PDF via Ctrl+O and file dialog."""
    log("Phase 2: Opening PDF...")

    if not os.path.exists(SAMPLE_PDF_PATH):
        fail_loud(f"Sample PDF not found: {SAMPLE_PDF_PATH}")

    # Ensure app has focus
    wins = gw.getWindowsWithTitle(APP_WINDOW_TITLE)
    if wins:
        wins[0].activate()
    time.sleep(0.5)

    pyautogui.hotkey("ctrl", "o")
    time.sleep(1.5)

    type_in_file_dialog(SAMPLE_PDF_PATH)
    time.sleep(2)  # PDF load and canvas render

    log("Phase 2 complete: PDF opened")
    time.sleep(PAUSE_SEC)


# =============================================================================
# PHASE 3: FILL HEADERS
# =============================================================================


def fill_headers() -> None:
    """Open Headers popup via Ctrl+H and fill all fields with keyboard."""
    log("Phase 3: Filling headers...")

    wins = gw.getWindowsWithTitle(APP_WINDOW_TITLE)
    if wins:
        wins[0].activate()
    time.sleep(0.3)

    pyautogui.hotkey("ctrl", "h")
    time.sleep(0.5)

    # Field order: part_number, drawing_rev, rm_used, date_inspected,
    # part_description, customer, project_name, accepted_qty
    header_values = [
        "STRESS-TEST-001",
        "A",
        "Test RM",
        datetime.now().strftime("%Y-%m-%d"),
        "Stress test part",
        "Test Customer",
        "Stress Test",
        "80",
    ]

    for val in header_values:
        pyautogui.hotkey("ctrl", "a")  # Overwrite any placeholder/default
        time.sleep(0.05)
        pyautogui.write(val, interval=0.02)
        time.sleep(0.05)
        pyautogui.press("enter")
        time.sleep(0.1)

    time.sleep(0.5)
    log("Phase 3 complete: Headers filled")
    time.sleep(PAUSE_SEC)


# =============================================================================
# PHASE 4: ADD BUBBLES
# =============================================================================


def add_bubbles() -> None:
    """Add bubbles via right-click on canvas, fill requirement popup with keyboard."""
    log(f"Phase 4: Adding {BUBBLE_COUNT} bubbles...")

    x_min, y_min, x_max, y_max = get_canvas_region()
    # Use center-ish region with padding to avoid edges
    pad_x = (x_max - x_min) // 6
    pad_y = (y_max - y_min) // 6
    cx_min, cx_max = x_min + pad_x, x_max - pad_x
    cy_min, cy_max = y_min + pad_y, y_max - pad_y
    rev = 0
    start = time.time()
    for i in range(BUBBLE_COUNT):
        if time.time() - start > ADD_BUBBLES_TIMEOUT:
            fail_loud(f"Add bubbles phase exceeded timeout ({ADD_BUBBLES_TIMEOUT}s)")

        # Random but valid canvas coordinates
        x = random.randint(cx_min, cx_max) if cx_max > cx_min else (cx_min + cx_max) // 2
        y = random.randint(cy_min, cy_max) if cy_max > cy_min else (cy_min + cy_max) // 2

        # Right-click to add balloon
        pyautogui.rightClick(x, y)
        time.sleep(0.3)

        # Fill requirement popup: Zone, Char, Req, neg, pos, equip
        # Enter advances; last Enter saves; pick random option per field to hit many cases
        fields = {
            "zone": random.choice(ZONE_OPTIONS),
            "char": random.choice(CHAR_OPTIONS),
            "req": random.choice(REQ_OPTIONS),
            "neg": random.choice(NEG_OPTIONS),
            "pos": random.choice(POS_OPTIONS),
            "equip": random.choice(EQUIP_OPTIONS),
        }
        for j, (_, val) in enumerate(fields.items()):
            pyautogui.write(str(val), interval=0.02)
            pyautogui.press("enter")
            time.sleep(0.3 if j == len(fields) - 1 else 0.05)

        if (i + 1) % 5 == 0:
            if rev == 0 :
                pyautogui.press("left")
                rev = 1
                time.sleep(0.3)
            elif rev == 1 :
                pyautogui.press("right")
                rev = 0
                time.sleep(0.3)
            log(f"  Added {i + 1}/{BUBBLE_COUNT} bubbles")

    log(f"Phase 4 complete: {BUBBLE_COUNT} bubbles added")
    time.sleep(PAUSE_SEC)


# =============================================================================
# PHASE 5: SAVE OUTPUTS
# =============================================================================


def _dismiss_saved_dialog() -> None:
    """Click OK on 'Saved' info dialog."""
    time.sleep(0.5)
    pyautogui.press("enter")


def _dismiss_open_file_prompt() -> None:
    """Click No on 'Open the saved file now?' prompt."""
    time.sleep(0.5)
    pyautogui.press("N")


def save_outputs(iteration: int) -> tuple[str, str, str]:
    """Save PDF, Report, Project using app defaults. Returns discovered paths."""
    log("Phase 5: Saving outputs...")
    ensure_output_dir()

    wins = gw.getWindowsWithTitle(APP_WINDOW_TITLE)
    if wins:
        wins[0].activate()
    time.sleep(0.3)

    # 5a. Save PDF (uses default FAIR_QA_drawing_*.pdf) - 3 enters (different dir)
    pyautogui.hotkey("ctrl", "s")
    time.sleep(1.5)
    navigate_and_save_with_default(OUTPUT_DIR, turn=1)
    time.sleep(1)
    _dismiss_saved_dialog()
    _dismiss_open_file_prompt()
    time.sleep(0.5)

    # 5b. Save Report (uses default FAIR_Report_*.xlsx) - 2 enters (same dir)
    pyautogui.hotkey("ctrl", "shift", "s")
    time.sleep(1.5)
    navigate_and_save_with_default(OUTPUT_DIR, turn=2)
    time.sleep(1)
    _dismiss_saved_dialog()
    _dismiss_open_file_prompt()
    time.sleep(0.5)

    # 5c. Save Project (uses default FAIR_Project_*.fairy) - 2 enters (same dir)
    pyautogui.hotkey("ctrl", "shift", "p")
    time.sleep(1.5)
    navigate_and_save_with_default(OUTPUT_DIR, turn=3)
    time.sleep(1)
    _dismiss_saved_dialog()
    time.sleep(0.5)

    # Discover actual paths (app uses timestamped default names)
    saved_pdf, saved_report, saved_project = discover_output_files(OUTPUT_DIR)
    log("Phase 5 complete: All outputs saved")
    return saved_pdf, saved_report, saved_project


# =============================================================================
# PHASE 6: CLOSE APP
# =============================================================================


def close_app(proc: subprocess.Popen) -> None:
    """Close app via Alt+F4; handle 'Save before closing?' if it appears."""
    log("Phase 6: Closing app...")

    wins = gw.getWindowsWithTitle(APP_WINDOW_TITLE)
    if wins:
        wins[0].activate()
    time.sleep(0.3)

    pyautogui.hotkey("alt", "f4")
    time.sleep(0.8)

    # If "Save before closing?" dialog (title "Unsaved Changes") appeared, press No
    dialogs = [w for w in gw.getAllWindows() if "Unsaved" in (w.title or "")]
    if dialogs:
        pyautogui.press("n")
        time.sleep(0.5)

    # Wait for process to exit
    start = time.time()
    while proc.poll() is None and time.time() - start < CLOSE_TIMEOUT:
        time.sleep(0.3)
        # Retry No if dialog appeared late
        dialogs = [w for w in gw.getAllWindows() if "Unsaved" in (w.title or "")]
        if dialogs:
            pyautogui.press("n")
            time.sleep(0.3)

    if proc.poll() is None:
        proc.terminate()
        time.sleep(1)
        if proc.poll() is None:
            proc.kill()

    log("Phase 6 complete: App closed")


# =============================================================================
# PHASE 7: VERIFY OUTPUTS
# =============================================================================


def verify_outputs(
    saved_pdf: str | None, saved_report: str | None, saved_project: str | None
) -> None:
    """Verify output files exist; fail loudly if any missing."""
    log("Phase 7: Verifying outputs...")

    missing = []
    if not saved_pdf or not os.path.exists(saved_pdf):
        missing.append(saved_pdf or "(PDF not found)")
    if not saved_report or not os.path.exists(saved_report):
        missing.append(saved_report or "(Report not found)")
    if not saved_project or not os.path.exists(saved_project):
        missing.append(saved_project or "(Project not found)")

    if missing:
        fail_loud(f"Expected output files missing: {missing}")

    for path in [saved_pdf, saved_report, saved_project]:
        if path and os.path.getsize(path) == 0:
            fail_loud(f"Output file is empty: {path}")

    log("Phase 7 complete: All outputs verified")
    log(f"  PDF:    {saved_pdf}")
    log(f"  Report: {saved_report}")
    log(f"  Project: {saved_project}")


# =============================================================================
# MAIN
# =============================================================================


def run_single_iteration(iteration: int) -> None:
    """Run one full stress test iteration."""
    log(f"=== Iteration {iteration + 1}/{ITERATIONS} ===")
    proc = launch_app()
    try:
        open_pdf()
        fill_headers()
        add_bubbles()
        saved_pdf, saved_report, saved_project = save_outputs(iteration)
        close_app(proc)
        verify_outputs(saved_pdf, saved_report, saved_project)
    except Exception as e:
        close_app(proc)
        fail_loud(f"Exception during test: {e}", screenshot=True)
        raise

    ret = proc.poll()
    if ret is not None and ret != 0:
        fail_loud(f"App exited with code {ret}")


def main() -> None:
    """Main entry point."""
    log("FAIR-y GUI Stress Test starting...")
    log(f"  EXE:    {FAIR_EXE_PATH}")
    log(f"  PDF:    {SAMPLE_PDF_PATH}")
    log(f"  Output: {OUTPUT_DIR}")
    log(f"  Bubbles: {BUBBLE_COUNT}, Iterations: {ITERATIONS}")
    log("")

    ensure_output_dir()

    for i in range(ITERATIONS):
        run_single_iteration(i)
        if i < ITERATIONS - 1:
            log("")
            time.sleep(2)

    log("")
    log("SUCCESS: All stress test iterations passed.")


if __name__ == "__main__":
    main()
