import fitz
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from openpyxl import Workbook
from datetime import datetime

# ================= FILES =================
PDF_IN = "demo.pdf"
PDF_OUT = "output_bubbled.pdf"

# ================= PDF =================
doc = fitz.open(PDF_IN)
num_pages = len(doc)
current_page_index = 0

# ================= STATE =================
zoom = 1.5
bubble_no = 1
bubbles = []

pan_start = None
offset_x, offset_y = 0, 0

# =====================================================
# RENDER PDF
# =====================================================
def render():
    page = doc[current_page_index]
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    tk_img = ImageTk.PhotoImage(img)
    canvas.image = tk_img
    canvas.delete("all")
    canvas.create_image(offset_x, offset_y, anchor="nw", image=tk_img)

    for b in bubbles:
        if b["page"] == current_page_index:
            x = b["x"] * zoom + offset_x
            y = b["y"] * zoom + offset_y
            r = b["r"] * zoom
            canvas.create_oval(x-r, y-r, x+r, y+r, outline="red", width=2)
            canvas.create_text(x, y, text=str(b["no"]), fill="red")

    update_bubble_list()

# =====================================================
# POPUP
# =====================================================
def requirement_popup():
    popup = tk.Toplevel(root)
    popup.title("FAIR Dimension Entry")
    popup.grab_set()

    result = {"ok": False}

    tk.Label(popup, text="Characteristic").grid(row=0, column=0, padx=8, pady=5, sticky="e")
    tk.Label(popup, text="Requirement").grid(row=1, column=0, padx=8, pady=5, sticky="e")
    tk.Label(popup, text="- Tol").grid(row=2, column=0, padx=8, pady=5, sticky="e")
    tk.Label(popup, text="+ Tol").grid(row=3, column=0, padx=8, pady=5, sticky="e")

    char = tk.Entry(popup, width=25)
    req  = tk.Entry(popup, width=20)
    neg  = tk.Entry(popup, width=10)
    pos  = tk.Entry(popup, width=10)

    char.grid(row=0, column=1)
    req.grid(row=1, column=1)
    neg.grid(row=2, column=1, sticky="w")
    pos.grid(row=3, column=1, sticky="w")

    def submit():
        if not req.get().strip():
            messagebox.showwarning("Missing", "Requirement is mandatory")
            return
        result["ok"] = True
        result["char"] = char.get().strip()
        result["req"]  = req.get().strip()
        result["neg"]  = neg.get().strip()
        result["pos"]  = pos.get().strip()
        popup.destroy()

    tk.Button(popup, text="OK", width=10, command=submit).grid(
        row=4, column=0, columnspan=2, pady=10
    )

    popup.wait_window()
    return result

# =====================================================
# ADD BUBBLE (IMMEDIATE)
# =====================================================
def add_bubble(event):
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
        "pos": ""
    })

    render()

    data = requirement_popup()

    if data.get("ok"):
        bubbles[-1]["char"] = data["char"]
        bubbles[-1]["req"]  = data["req"]
        bubbles[-1]["neg"]  = data["neg"]
        bubbles[-1]["pos"]  = data["pos"]
        bubble_no += 1
    else:
        bubbles.pop()

    render()

# =====================================================
# ZOOM / PAN
# =====================================================
def zoom_canvas(event):
    global zoom, offset_x, offset_y
    factor = 1.25 if event.delta > 0 else 1 / 1.25
    mx, my = event.x, event.y
    offset_x = mx - factor * (mx - offset_x)
    offset_y = my - factor * (my - offset_y)
    zoom *= factor
    update_preview(bubble_radius_slider.get())
    render()

def start_pan(event):
    global pan_start
    pan_start = (event.x, event.y)

def do_pan(event):
    global offset_x, offset_y, pan_start
    if pan_start:
        offset_x += event.x - pan_start[0]
        offset_y += event.y - pan_start[1]
        pan_start = (event.x, event.y)
        render()

def end_pan(event):
    global pan_start
    pan_start = None

# =====================================================
# NAV / UNDO
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
    bubble_listbox.delete(0, tk.END)
    bubble_listbox.insert(tk.END, "No | Char | Req | -Tol | +Tol")
    bubble_listbox.insert(tk.END, "-" * 40)

    for b in bubbles:
        if b["page"] == current_page_index:
            bubble_listbox.insert(
                tk.END,
                f"{str(b['no']).ljust(2)} | "
                f"{str(b['char']).ljust(5)} | "
                f"{str(b['req']).ljust(5)} | "
                f"{str(b['neg']).ljust(4)} | "
                f"{str(b['pos']).ljust(4)}"
            )

# =====================================================
# SAVE DRAWING (PDF)
# =====================================================
def save_drawing():
    out = fitz.open()
    for i in range(num_pages):
        p = out.new_page(width=doc[i].rect.width, height=doc[i].rect.height)
        p.show_pdf_page(p.rect, doc, i)
        for b in bubbles:
            if b["page"] == i:
                x, y, r = b["x"], b["y"], b["r"]
                p.draw_oval(fitz.Rect(x-r, y-r, x+r, y+r), color=(1,0,0), width=1.2)
                p.insert_text(fitz.Point(x-4, y+4), str(b["no"]), fontsize=8, color=(1,0,0))
    out.save(PDF_OUT)
    out.close()
    messagebox.showinfo("Saved", f"Drawing saved as {PDF_OUT}")

# =====================================================
# SAVE REPORT (EXCEL)
# =====================================================
def save_report():
    wb = Workbook()
    ws = wb.active
    ws.title = "FAIR"

    ws.append([
        "Pg",
        "Ballon No",
        "Characteristic Designator",
        "Requirement",
        "- Tol",
        "+ Tol"
    ])

    for b in bubbles:
        ws.append([
            b["page"] + 1,
            b["no"],
            b["char"],
            b["req"],
            b["neg"],
            b["pos"]
        ])

    filename = f"FAIR_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx"
    wb.save(filename)
    messagebox.showinfo("Saved", f"Report saved as:\n{filename}")

# =====================================================
# UI
# =====================================================
root = tk.Tk()
root.title("FAIR GENERATOR")

toolbar = tk.Frame(root)
toolbar.pack(fill="x")

tk.Button(toolbar, text="Prev", command=prev_page).pack(side="left")
tk.Button(toolbar, text="Next", command=next_page).pack(side="left")
tk.Button(toolbar, text="Undo", command=undo).pack(side="left")
tk.Button(toolbar, text="Save Drawing", command=save_drawing).pack(side="left")
tk.Button(toolbar, text="Save Report", command=save_report).pack(side="left")

bubble_radius_slider = tk.Scale(toolbar, from_=6, to=25, orient="horizontal", label="Bubble Size")
bubble_radius_slider.set(12)
bubble_radius_slider.pack(side="right")

preview_canvas = tk.Canvas(toolbar, width=50, height=50, bg="gray")
preview_canvas.pack(side="right", padx=5)

def update_preview(val):
    preview_canvas.delete("all")
    r = int(val) * zoom
    preview_canvas.create_oval(25-r, 25-r, 25+r, 25+r, outline="red", width=2)

bubble_radius_slider.config(command=update_preview)
update_preview(bubble_radius_slider.get())

bubble_listbox = tk.Listbox(root, height=7, font=("Consolas", 10))
bubble_listbox.pack(fill="x", padx=5, pady=3)

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
