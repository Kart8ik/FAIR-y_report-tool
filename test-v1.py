import fitz
import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog

from PIL import Image, ImageTk

PDF_IN = None
PDF_OUT = None
doc = None
num_pages = 0
current_page_index = 0
current_page = None  # set after a PDF is opened via open_pdf()

zoom = 1.5
bubble_no = 1
bubbles = []

pan_start = None
offset_x, offset_y = 0, 0

# ---------- RENDER ----------
def render():
    if not doc:
        return
    global current_page
    current_page = doc[current_page_index]

    mat = fitz.Matrix(zoom, zoom)
    pix = current_page.get_pixmap(matrix=mat)
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
            canvas.create_text(x, y, text=str(b["no"]), fill="red",
                               font=("Arial", max(8, int(8 * zoom))))

    update_bubble_list()

def open_pdf():
    global PDF_IN, doc, num_pages, current_page_index, bubbles, bubble_no
    global offset_x, offset_y

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

    render()

# ---------- ADD BUBBLE ----------
def add_bubble(event):
    global bubble_no
    pdf_x = (event.x - offset_x) / zoom
    pdf_y = (event.y - offset_y) / zoom
    r = bubble_radius_slider.get()

    bubbles.append({
        "page": current_page_index,
        "no": bubble_no,
        "x": pdf_x,
        "y": pdf_y,
        "r": r,
        "desc": "..."
    })
    render()

    desc = simpledialog.askstring(
        "QA Description",
        f"Enter description for Bubble {bubble_no}:"
    )

    if desc:
        bubbles[-1]["desc"] = desc
        bubble_no += 1
    else:
        bubbles.pop()

    render()

# ---------- ZOOM ----------
def zoom_canvas(event):
    global zoom, offset_x, offset_y
    factor = 1.25 if event.delta > 0 else 1 / 1.25

    mx, my = event.x, event.y
    offset_x = mx - factor * (mx - offset_x)
    offset_y = my - factor * (my - offset_y)

    zoom *= factor
    update_preview(bubble_radius_slider.get())
    render()

# ---------- PAN ----------
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
        render()

def end_pan(event):
    global pan_start
    pan_start = None

# ---------- PAGE NAV ----------
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

# ---------- UNDO ----------
def undo():
    global bubble_no
    for i in reversed(range(len(bubbles))):
        if bubbles[i]["page"] == current_page_index:
            bubbles.pop(i)
            bubble_no -= 1
            break
    render()

# ---------- DELETE ----------
def delete_bubble():
    sel = bubble_listbox.curselection()
    if not sel:
        return
    b = bubbles_on_current_page()[sel[0]]
    bubbles.remove(b)
    render()

def bubbles_on_current_page():
    return [b for b in bubbles if b["page"] == current_page_index]

def update_bubble_list():
    bubble_listbox.delete(0, tk.END)
    for b in bubbles_on_current_page():
        bubble_listbox.insert(tk.END, f"{b['no']}  -  {b['desc']}")

# ---------- SAVE ----------
def save_pdf():
    PDF_OUT = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Save QA Report As"
    )

    if not PDF_OUT:
        return

    out = fitz.open()

    qa_page = out.new_page(width=595, height=842)
    y = 50
    qa_page.insert_text((50, y), "QA DIMENSION LIST", fontsize=14)
    y += 40

    qa_page.insert_text((50, y), "S.No", fontsize=10)
    qa_page.insert_text((100, y), "Description", fontsize=10)
    qa_page.insert_text((450, y), "Remarks", fontsize=10)
    y += 20

    for b in bubbles:
        qa_page.insert_text((50, y), str(b["no"]), fontsize=9)
        qa_page.insert_text((100, y), b["desc"], fontsize=9)
        qa_page.draw_rect(fitz.Rect(430, y-5, 560, y+10))
        y += 18
        if y > 800:
            qa_page = out.new_page(width=595, height=842)
            y = 50

    for i in range(num_pages):
        p = out.new_page(width=doc[i].rect.width, height=doc[i].rect.height)
        p.show_pdf_page(p.rect, doc, i)
        for b in [b for b in bubbles if b["page"] == i]:
            x, y, r = b["x"], b["y"], b["r"]
            p.draw_oval(fitz.Rect(x-r, y-r, x+r, y+r), color=(1,0,0), width=1.2)
            p.insert_text(fitz.Point(x-4, y+4), str(b["no"]), fontsize=8, color=(1,0,0))

    out.save(PDF_OUT)
    out.close()
    messagebox.showinfo("Saved", f"PDF saved as {PDF_OUT}")

# ---------- UI ----------
root = tk.Tk()
root.title("QA REPORT TOOL")

# Header
tk.Label(root, text="QA REPORT TOOL",
         font=("Arial", 14, "bold")).pack(pady=3)

# Instructions bar
instructions = (
    "Right-Click : Add Bubble    |    "
    "Left-Click + Drag : Pan Drawing    |    "
    "Mouse Scroll : Zoom In / Out"
)

tk.Label(
    root,
    text=instructions,
    font=("Arial", 9),
    fg="#333",
    bg="#E6E6E6",
    padx=10,
    pady=4
).pack(fill="x", padx=5, pady=(0, 4))

toolbar = tk.Frame(root)
toolbar.pack(fill="x")

tk.Button(toolbar, text="Open PDF", command=open_pdf).pack(side="left")
tk.Button(toolbar, text="Prev Page", command=prev_page).pack(side="left")
tk.Button(toolbar, text="Next Page", command=next_page).pack(side="left")
tk.Button(toolbar, text="Undo", command=undo).pack(side="left")
tk.Button(toolbar, text="Delete Bubble", command=delete_bubble).pack(side="left")
tk.Button(toolbar, text="Save PDF", command=save_pdf).pack(side="left")

bubble_radius_slider = tk.Scale(
    toolbar, from_=6, to=25, orient="horizontal", label="Bubble Size"
)
bubble_radius_slider.set(12)
bubble_radius_slider.pack(side="right")

preview_canvas = tk.Canvas(
    toolbar, width=50, height=50, bg="gray",
    highlightthickness=1, highlightbackground="black"
)
preview_canvas.pack(side="right", padx=5)

def update_preview(val):
    preview_canvas.delete("all")
    r = int(val) * zoom
    preview_canvas.create_oval(25-r, 25-r, 25+r, 25+r, outline="red", width=2)

bubble_radius_slider.config(command=update_preview)
update_preview(bubble_radius_slider.get())

bubble_listbox = tk.Listbox(root, height=5)
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
