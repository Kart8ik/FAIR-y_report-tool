import fitz
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from PIL import Image, ImageTk
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
from datetime import datetime
import shutil
from copy import copy
import sys, os

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS   # PyInstaller temp dir
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# ================ to apply icon ===================
def apply_icon(window):
    try:
        window.iconbitmap(resource_path("app-icon.ico"))
    except:
        pass

# ================ handle data types ===================
def to_number(value):
    value = value.strip()
    if value == "":
        return None
    try:
        return int(value) if value.isdigit() else float(value)
    except ValueError:
        return None



# ================= FILES =================
PDF_IN = None
PDF_OUT = None
doc = None
num_pages = 0
current_page_index = 0
current_page = None  # set after a PDF is opened via open_pdf()
TEMPLATE_XLSX = resource_path("FORMAT.xlsx")

# ================= STATE =================
zoom = 1.5
bubble_no = 1
bubbles = []
rendered_img = None
rendered_zoom = None
rendered_page_index = None
page_cache = {}  # cache rendered pages per (page_index, zoom)

pan_start = None
offset_x, offset_y = 0, 0

# =====================================================
# RENDER PDF
# =====================================================
def render_pdf():
    global rendered_img, rendered_zoom, rendered_page_index

    cache_key = (current_page_index, zoom)
    cached = page_cache.get(cache_key)
    if cached:
        rendered_img = cached
    else:
        page = doc[current_page_index]
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        rendered_img = ImageTk.PhotoImage(img)
        page_cache[cache_key] = rendered_img
    rendered_zoom = zoom
    rendered_page_index = current_page_index

    canvas.delete("pdf")
    canvas.create_image(offset_x, offset_y, anchor="nw",
                        image=rendered_img, tags="pdf")

def render_overlays():
    canvas.delete("overlay")

    for b in bubbles:
        if b["page"] != current_page_index:
            continue

        x = b["x"] * zoom + offset_x
        y = b["y"] * zoom + offset_y
        r = b["r"] * zoom

        if b.get("highlight"):
            outline = "red"
            fill = "yellow"
            width = max(3, r / 6)
        else:
            outline = "red"
            fill = ""
            width = max(2, r / 10)

        canvas.create_oval(
            x - r, y - r, x + r, y + r,
            outline=outline,
            fill=fill,
            width=width,
            tags="overlay"
        )

        canvas.create_text(
            x, y,
            text=str(b["no"]),
            font=("Arial", int(r)),
            fill=outline,
            tags="overlay"
        )


def render(force=False):
    if not doc:
        return

    if force or rendered_zoom != zoom or rendered_page_index != current_page_index:
        render_pdf()

    render_overlays()
    update_bubble_list()




def open_pdf():
    global PDF_IN, doc, num_pages, current_page_index, bubbles, bubble_no
    global offset_x, offset_y, page_cache

    path = filedialog.askopenfilename(
        title="Select PDF",
        filetypes=[("PDF files", "*.pdf")]
    )

    if not path:
        return

    if doc:
        doc.close()

    PDF_IN = path
    doc = fitz.open(PDF_IN)
    num_pages = len(doc)
    current_page_index = 0
    bubbles.clear()
    bubble_no = 1
    offset_x = offset_y = 0
    page_cache.clear()

    render()

# =====================================================
# POPUP (Requirement / Tolerance)
# =====================================================
def requirement_popup(existing=None):
    if not doc:
        messagebox.showwarning("No File", "Open file to work on")
        return

    popup = tk.Toplevel(root)
    popup.title("Edit FAIR Dimension" if existing else "FAIR Dimension Entry")
    apply_icon(popup)
    popup.transient(root)
    popup.grab_set()

    # Let the grid stretch so buttons can anchor to opposite sides
    popup.columnconfigure(0, weight=1)
    popup.columnconfigure(1, weight=1)

    result = {"action": None}

    tk.Label(popup, text="Characteristic Designator").grid(row=0, column=0, padx=8, pady=5, sticky="e")
    tk.Label(popup, text="Requirement").grid(row=1, column=0, padx=8, pady=5, sticky="e")
    tk.Label(popup, text="- Tol").grid(row=2, column=0, padx=8, pady=5, sticky="e")
    tk.Label(popup, text="+ Tol").grid(row=3, column=0, padx=8, pady=5, sticky="e")
    tk.Label(popup, text="Equipment Used").grid(row=4, column=0, padx=8, pady=5, sticky="e")

    char = ttk.Combobox(
        popup,
        values = [
            "--- BASIC DIMENSIONS ---",
            "Length",
            "Width",
            "Thickness",
            "Diameter",
            "Radius",
            "Chamfer",
            "Angle",

            "--- SLOT / CUTOUT ---",
            "Slot length",
            "Slot width",

            "--- HOLE / LOCATION ---",
            "Edge to edge",
            "Edge to hole center",
            "Hole center to edge",
            "Hole center to hole center",
            "Edge to slot",
            "Slot to edge",

            "--- BEND ---",
            "Edge to bend",
            "Bend to edge",
            "Bend to bend",
            "Bending radius",

            "--- EXTRUSION ---",
            "Extrusion height",
            "Extrusion diameter",

            "--- FASTENING ---",
            "Tapping",
            "Riveting",
            "PEM",
            "Torque test",

            "--- FINISH ---",
            "Coating thickness",
            "Aesthetic parameters",

            "--- GD&T ---",
            "Flatness",
            "Parallelism",
            "Perpendicularity",
            "Concentricity",

            "--- MISC ---",
            "Dimension",
            "Ref dimension",
            "Note",
        ],
        width=28
    )
    vcmd = popup.register(lambda P: P == "" or P.replace(".", "", 1).isdigit())

    req = tk.Entry(popup, validate="key", validatecommand=(vcmd, "%P"))
    neg = tk.Entry(popup, validate="key", validatecommand=(vcmd, "%P"))
    pos = tk.Entry(popup, validate="key", validatecommand=(vcmd, "%P"))

    equip = ttk.Combobox(
        popup,
        values = [
            "--- BASIC TOOLS ---",
            "Visual",
            "Vernier caliper",
            "Digital vernier caliper",
            "Digital micrometer",

            "--- HEIGHT MEASUREMENT ---",
            "Digital height gauge",
            "Micro height gauge",

            "--- REFERENCE ---",
            "Slip gauge",

            "--- FORM & ANGLE ---",
            "Radius gauge",
            "Feeler gauge",
            "Bevel protractor",

            "--- GO / NO-GO ---",
            "Pin gauge",
            "Thread plug gauge",

            "--- SURFACE ---",
            "DFT meter",

            "--- ASSEMBLY ---",
            "Torque wrench",

            "--- ADVANCED ---",
            "Profile projector",
            "CMM",
        ],
        width=28
    )

    if existing:
        char.insert(0, existing["char"])
        req.insert(0, existing["req"])
        neg.insert(0, existing["neg"])
        pos.insert(0, existing["pos"])
        equip.insert(0, existing["equip"])

    char.grid(row=0, column=1, padx=(0, 16))
    req.grid(row=1, column=1, sticky="w")
    neg.grid(row=2, column=1, sticky="w")
    pos.grid(row=3, column=1, sticky="w")
    equip.grid(row=4, column=1, sticky="w")

    def save():
        req_val = to_number(req.get())
        neg_val = to_number(neg.get())
        pos_val = to_number(pos.get())

        if not req.get().strip():
            messagebox.showwarning("Missing", "Requirement is mandatory")
            return
        result.update({
            "action": "save",
            "char": char.get().strip(),
            "req": req_val,
            "neg": neg_val,
            "pos": pos_val,
            "equip": equip.get().strip()
        })
        popup.destroy()

    def delete():
        if messagebox.askyesno("Delete", "Delete this bubble?"):
            result["action"] = "delete"
            popup.destroy()

    if not existing:
        tk.Button(popup, text="Save", width=10, command=save).grid(
            row=5, column=1, padx=(0, 20), pady=8, sticky="e"
        )

    if existing:
        tk.Button(popup, text="Save", width=10, command=save).grid(
            row=5, column=0, padx=(20, 0), pady=8, sticky="w"
        )
        tk.Button(popup, text="Delete", width=10, fg="red", command=delete).grid(
            row=5, column=1, padx=(0, 20), pady=8, sticky="e"
        )

    popup.wait_window()
    return result



# =====================================================
# ADD BUBBLE (IMMEDIATE)
# =====================================================
def add_bubble(event):
    if not doc:
        messagebox.showwarning("No File", "Open file to work on")
        return

    global bubble_no

    pdf_x = (event.x - offset_x) / zoom
    pdf_y = (event.y - offset_y) / zoom

    bubbles.append({
        "page": current_page_index,
        "no": bubble_no,
        "x": pdf_x,
        "y": pdf_y,
        "r": bubble_radius_slider.get(),
        "char": "",
        "req": "",
        "neg": "",
        "pos": "",
        "equip": "",
        "highlight": False
    })


    render()

    data = requirement_popup()

    # requirement_popup returns {"action": "save"|"delete"|None, ...}
    if data.get("action") == "save":
        bubbles[-1]["char"] = data["char"]
        bubbles[-1]["req"]  = data["req"]
        bubbles[-1]["neg"]  = data["neg"]
        bubbles[-1]["pos"]  = data["pos"]
        bubbles[-1]["equip"]  = data["equip"]
        bubble_no += 1
    else:
        # If the dialog was closed or cancelled, discard the pending bubble
        bubbles.pop()

    render()



def highlight_bubble(bubble, duration=5000):
    bubble["highlight"] = True
    render_overlays()

    def clear():
        bubble["highlight"] = False
        render_overlays()

    root.after(duration, clear)


# =====================================================
# EDIT BUBBLE (from list)
# =====================================================

def on_bubble_edit(event):
    lb = event.widget
    idx = lb.nearest(event.y)

    if idx >= 0:
        lb.selection_clear(0, tk.END)
        lb.selection_set(idx)
        lb.activate(idx)

    if idx < 2:
        return  # header rows

    page_bubbles = [b for b in bubbles if b["page"] == current_page_index]
    bubble_idx = idx - 2

    if bubble_idx >= len(page_bubbles):
        return

    bubble = page_bubbles[bubble_idx]

    highlight_bubble(bubble)

    result = requirement_popup(existing=bubble)

    if result["action"] == "save":
        bubble["char"] = result["char"]
        bubble["req"]  = result["req"]
        bubble["neg"]  = result["neg"]
        bubble["pos"]  = result["pos"]
        bubble["equip"]  = result["equip"]

    elif result["action"] == "delete":
        bubbles.remove(bubble)
        for i, b in enumerate(bubbles, start=1):
            b["no"] = i
        global bubble_no
        bubble_no = len(bubbles) + 1

    update_bubble_list()
    render(force=True)



# =====================================================
# ZOOM / PAN
# =====================================================
zoom_job = None

def zoom_canvas(event):
    global zoom, offset_x, offset_y, zoom_job, page_cache

    factor = 1.25 if event.delta > 0 else 1 / 1.25
    new_zoom = zoom * factor

    if new_zoom > 10.0 or new_zoom < 0.5 :
        return

    mx, my = event.x, event.y
    offset_x = mx - factor * (mx - offset_x)
    offset_y = my - factor * (my - offset_y)
    if new_zoom != zoom:
        page_cache.clear()
    zoom = new_zoom

    update_preview(bubble_radius_slider.get())

    if zoom_job:
        root.after_cancel(zoom_job)

    zoom_job = root.after(120, lambda: render(force=True))


def start_pan(event):
    global pan_start
    pan_start = (event.x, event.y)

def do_pan(event):
    global offset_x, offset_y, pan_start
    if pan_start:
        dx = event.x - pan_start[0]
        dy = event.y - pan_start[1]
        offset_x += dx
        offset_y += dy
        pan_start = (event.x, event.y)

        canvas.move("pdf", dx, dy)
        canvas.move("overlay", dx, dy)


def end_pan(event):
    global pan_start
    pan_start = None

# =====================================================
# PAGE NAV / UNDO
# =====================================================
def next_page():
    global current_page_index, offset_x, offset_y
    if current_page_index < num_pages - 1:
        current_page_index += 1
        offset_x = offset_y = 0
        render()

def prev_page():
    global current_page_index, offset_x, offset_y
    if current_page_index > 0:
        current_page_index -= 1
        offset_x = offset_y = 0
        render()

def undo():
    global bubble_no
    for i in reversed(range(len(bubbles))):
        if bubbles[i]["page"] == current_page_index:
            bubbles.pop(i)
            bubble_no -= 1
            break
    render()

# =====================================================
# LIST VIEW
# =====================================================
def update_bubble_list():
    # Fixed column widths for neat alignment in the list view
    w_no, w_char, w_req, w_tol, w_equip = 3, 28, 8, 6, 22
    header = (
        f"{'No':<{w_no}} | "
        f"{'Char':<{w_char}} | "
        f"{'Req':<{w_req}} | "
        f"{'-Tol':<{w_tol}} | "
        f"{'+Tol':<{w_tol}} | "
        f"{'Equip':<{w_equip}}"
    )
    sep = "-" * len(header)

    bubble_listbox.delete(0, tk.END)
    bubble_listbox.insert(tk.END, header)
    bubble_listbox.insert(tk.END, sep)

    for b in bubbles:
        if b["page"] == current_page_index:
            bubble_listbox.insert(
                tk.END,
                f"{str(b['no']):<{w_no}} | "
                f"{str(b['char']):<{w_char}} | "
                f"{str(b['req']):<{w_req}} | "
                f"{str(b['neg']):<{w_tol}} | "
                f"{str(b['pos']):<{w_tol}} | "
                f"{str(b['equip']):<{w_equip}}"
            )

# =====================================================
# SAVE BUBBLED PDF
# =====================================================
def save_pdf():
    if not doc:
        messagebox.showwarning("No File", "No file to save PDF")
        return

    if not bubbles:
        messagebox.showwarning("No data", "No bubbles to export")
        return

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    default_name = f"FAIR_QA_drawing_{timestamp}.pdf"

    PDF_OUT = filedialog.asksaveasfilename(
        initialfile=default_name,
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Save QA Report As"
    )

    if not PDF_OUT:
        return

    out = fitz.open()
    for i in range(num_pages):
        p = out.new_page(width=doc[i].rect.width, height=doc[i].rect.height)
        p.show_pdf_page(p.rect, doc, i)
        for b in bubbles:
            if b["page"] == i:
                x, y, r = b["x"], b["y"], b["r"]
                p.draw_oval(fitz.Rect(x-r, y-r, x+r, y+r), color=(1,0,0), width=r/10)
                text = str(b["no"])
                font_size = r
                try:
                    text_width = fitz.get_text_length(text, fontsize=font_size)
                except Exception:
                    text_width = font_size * 0.6 * len(text)  # fallback estimate
                # center horizontally; adjust baseline to visually center vertically
                tx = x - text_width / 2
                ty = y + font_size * 0.35
                p.insert_text(fitz.Point(tx, ty), text, fontsize=font_size, color=(1,0,0))
    out.save(PDF_OUT)
    out.close()
    messagebox.showinfo("Saved", f"Bubbled drawing saved as {PDF_OUT}")


# =====================================================
# SAVE REPORT (UNCHANGED)
# =====================================================
def save_report():
    if not doc:
        messagebox.showwarning("No File", "No file to save Report")
        return

    if not bubbles:
        messagebox.showwarning("No data", "No bubbles to export")
        return

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    default_name = f"FAIR_Report_{timestamp}.xlsx"

    report_file = filedialog.asksaveasfilename(
        title="Save FAIR Report",
        initialfile=default_name,
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not report_file:
        return

    shutil.copy(TEMPLATE_XLSX, report_file)
    wb = load_workbook(report_file)
    ws = wb.active

    # The template has 15 data rows (8-22). Row 23 participates in the merged
    # Legends block (A23:B24, C23:M24), so we must insert before that at row 23.
    START_ROW = 8
    TEMPLATE_ROWS = 15
    LAST_TEMPLATE_ROW = START_ROW + TEMPLATE_ROWS - 1  # 22
    INSERT_AT = LAST_TEMPLATE_ROW + 1                  # 23 (just before merged legends)
    template_row_idx = LAST_TEMPLATE_ROW
    footer_rows = {24: ws.row_dimensions[24].height,
                   25: ws.row_dimensions[25].height,
                   26: ws.row_dimensions[26].height}

    def copy_row_style(src_row_idx, dest_row_idx, clear_values=True):
        # Copy styles across all columns to preserve borders/alignment
        for col_idx in range(1, ws.max_column + 1):
            src_cell = ws.cell(row=src_row_idx, column=col_idx)
            dest_cell = ws.cell(row=dest_row_idx, column=col_idx)
            dest_cell._style = copy(src_cell._style)
            if clear_values:
                dest_cell.value = None
        # Copy row height to keep grid consistent
        ws.row_dimensions[dest_row_idx].height = ws.row_dimensions[src_row_idx].height

    bubble_count = len(bubbles)
    extra_rows = max(0, bubble_count - TEMPLATE_ROWS)

    # Insert new rows just above the merged footer block (row 23+)
    if extra_rows > 0:
        ws.insert_rows(INSERT_AT, amount=extra_rows)

        for i in range(extra_rows):
            target_row = INSERT_AT + i
            copy_row_style(template_row_idx, target_row)

        # Re-merge the footer block (Legends/Observations/Sign-off) so it
        # stays intact after the shift, and restore footer row heights.
        footer_merges = [
            "A24:B24", "C24:M24",
            "D25:M25",
            "A26:B26", "C26:D26", "E26:F26", "G26:H26", "I26:J26", "K26:M26",
        ]
        for rng in footer_merges:
            if rng in ws.merged_cells:
                ws.unmerge_cells(rng)
        for rng in footer_merges:
            min_col, min_row, max_col, max_row = range_boundaries(rng)
            ws.merge_cells(
                start_row=min_row + extra_rows,
                start_column=min_col,
                end_row=max_row + extra_rows,
                end_column=max_col,
            )
        for src_row, height in footer_rows.items():
            ws.row_dimensions[src_row + extra_rows].height = height

    # Refresh styles on all data rows (including originals) to keep borders/alignments
    style_src_row = START_ROW  # first data row has canonical formatting
    total_rows = bubble_count + 1
    for r in range(START_ROW, START_ROW + total_rows):
        copy_row_style(style_src_row, r, clear_values=False)

    # WRITE DATA (flows naturally)
    row = START_ROW
    for b in bubbles:
        ws.cell(row=row, column=1).value = b["page"] + 1
        ws.cell(row=row, column=2).value = b["no"]
        ws.cell(row=row, column=3).value = b['char']
        ws.cell(row=row, column=4).value = b["req"]
        ws.cell(row=row, column=5).value = b["neg"]
        ws.cell(row=row, column=6).value = b["pos"]
        ws.cell(row=row, column=7).value = b["equip"]
        row += 1

    wb.save(report_file)
    messagebox.showinfo("Saved", f"FAIR report created:\n{report_file}")


# =====================================================
# UI
# =====================================================
root = tk.Tk()
root.title("FAIR-y")
root.iconbitmap(resource_path("app-icon.ico"))


toolbar = tk.Frame(root)
toolbar.pack(fill="x", padx=5)

tk.Button(toolbar, text="Open PDF", command=open_pdf).pack(side="left")
tk.Button(toolbar, text="Prev Page", command=prev_page).pack(side="left")
tk.Button(toolbar, text="Next Page", command=next_page).pack(side="left")
tk.Button(toolbar, text="Undo Bubble", command=undo).pack(side="left")
tk.Button(toolbar, text="Save PDF", command=save_pdf).pack(side="left")
tk.Button(toolbar, text="Save Report", command=save_report).pack(side="left")

bubble_radius_slider = tk.Scale(toolbar, from_=3, to=25, orient="horizontal", label="Bubble Size")
bubble_radius_slider.set(6)
bubble_radius_slider.pack(side="right")

preview_canvas = tk.Canvas(toolbar, width=50, height=50, bg="gray")
preview_canvas.pack(side="right", padx=5)

def update_preview(val):
    preview_canvas.delete("all")
    r = int(val) * zoom
    preview_canvas.create_oval(25-r, 25-r, 25+r, 25+r, outline="red", width=2)

bubble_radius_slider.config(command=update_preview)
update_preview(bubble_radius_slider.get())

bubble_listbox = tk.Listbox(
    root,
    height=7,
    font=("Consolas", 10),
    selectmode="browse",
    exportselection=False
)

bubble_listbox.pack(fill="x", padx=5, pady=3)
bubble_listbox.bind("<Double-Button-1>", on_bubble_edit)


canvas = tk.Canvas(root, bg="gray")
canvas.pack(fill="both", expand=True)

canvas.bind("<Button-3>", add_bubble)
canvas.bind("<Button-1>", start_pan)
canvas.bind("<B1-Motion>", do_pan)
canvas.bind("<ButtonRelease-1>", end_pan)
canvas.bind("<MouseWheel>", zoom_canvas)

render()
root.mainloop()
doc.close()
