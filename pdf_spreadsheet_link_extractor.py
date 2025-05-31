import os
import sys
import fitz
import csv
import re
import io
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from PIL import Image, ImageTk

# Tray support
from pystray import Icon, Menu, MenuItem
from PIL import Image as PILImage, ImageDraw

# Drag-and-drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_ENABLED = True
except ImportError:
    DND_ENABLED = False

try:
    import openpyxl
    EXCEL_ENABLED = True
except ImportError:
    EXCEL_ENABLED = False

# Globals
default_save_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
save_location = None
link_regions = []
tray_icon = None
root = None
output_area = None

# Create tray icon image
def create_tray_icon():
    icon = PILImage.new('RGB', (64, 64), (70, 130, 180))
    d = ImageDraw.Draw(icon)
    d.rectangle((8, 24, 56, 40), fill=(255, 255, 255))
    d.rectangle((12, 28, 52, 36), fill=(70, 130, 180))
    return icon

# === Extractors ===
def extract_links_from_pdf_manual(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]
    links = page.get_links()
    table_links = []
    for region in link_regions:
        for link in links:
            if 'uri' in link and region.intersects(link['from']):
                table_links.append(link['uri'])
    return table_links

def extract_links_from_csv(csv_path):
    links = []
    pattern = re.compile(r'https?://[^\s,"]+')
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            for cell in row:
                links.extend(pattern.findall(cell))
    return links

def extract_links_from_excel(xlsx_path):
    if not EXCEL_ENABLED:
        return []
    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    links = []
    pattern = re.compile(r'https?://[^\s,"]+')
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                urls = []
                if cell.value and isinstance(cell.value, str):
                    urls += pattern.findall(cell.value)
                if cell.hyperlink and cell.hyperlink.target:
                    urls.append(cell.hyperlink.target)
                links.extend(urls)
    return links

def save_links_to_file(base_name, urls):
    global save_location
    if not save_location:
        base_path = os.path.join(default_save_dir, f"{base_name}_extracted_links.txt")
    else:
        base_path = save_location
    path = base_path
    i = 1
    while os.path.exists(path):
        root_path, ext = os.path.splitext(base_path)
        path = f"{root_path} ({i}){ext}"
        i += 1
    with open(path, "w", encoding="utf-8") as f:
        for url in urls:
            f.write(f"{url}\n")
    messagebox.showinfo("Saved", f"✅ Links saved to:\n{path}")

# === Manual region selection ===
def show_pdf_and_select_tables(pdf_path):
    global link_regions
    link_regions.clear()
    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(dpi=150)
    img_data = Image.open(io.BytesIO(pix.tobytes()))

    selector = tk.Toplevel(root)
    selector.title("Select Table Regions")

    canvas_frame = tk.Frame(selector)
    canvas_frame.pack(fill='both', expand=True)
    v_scroll = tk.Scrollbar(canvas_frame, orient='vertical')
    h_scroll = tk.Scrollbar(canvas_frame, orient='horizontal')
    canvas = tk.Canvas(canvas_frame, yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set,
                       width=700, height=500, scrollregion=(0, 0, img_data.width, img_data.height))
    v_scroll.config(command=canvas.yview)
    h_scroll.config(command=canvas.xview)
    v_scroll.pack(side='right', fill='y')
    h_scroll.pack(side='bottom', fill='x')
    canvas.pack(side='left', fill='both', expand=True)

    photo = ImageTk.PhotoImage(img_data)
    canvas.create_image(0, 0, anchor="nw", image=photo)
    canvas.image = photo

    selections = []
    current_rect = [0, 0, 0, 0]
    temp_id = None

    def start_rect(event):
        current_rect[0], current_rect[1] = canvas.canvasx(event.x), canvas.canvasy(event.y)

    def draw_rect(event):
        nonlocal temp_id
        current_rect[2], current_rect[3] = canvas.canvasx(event.x), canvas.canvasy(event.y)
        if temp_id:
            canvas.delete(temp_id)
        temp_id = canvas.create_rectangle(*current_rect, outline='red')

    def finish_rect(event):
        x0, y0, x1, y1 = current_rect
        rect = fitz.Rect(
            x0 * 72 / 150,
            y0 * 72 / 150,
            x1 * 72 / 150,
            y1 * 72 / 150
        )
        link_regions.append(rect)
        selections.append(canvas.create_rectangle(*current_rect, outline='green'))

    def extract_and_close():
        selector.destroy()
        result = extract_links_from_pdf_manual(pdf_path)
        output_area.config(state='normal')
        output_area.delete("1.0", tk.END)
        if result:
            output_area.insert(tk.END, "\n".join(result))
            filename = os.path.splitext(os.path.basename(pdf_path))[0]
            save = messagebox.askyesno("Save?", f"Found {len(result)} link(s). Save to file?")
            if save:
                save_links_to_file(filename, result)
        else:
            output_area.insert(tk.END, "No links found in selected regions.")
            messagebox.showinfo("No Links", "No hyperlinks found.")
        output_area.config(state='disabled')

    canvas.bind("<Button-1>", start_rect)
    canvas.bind("<B1-Motion>", draw_rect)
    canvas.bind("<ButtonRelease-1>", finish_rect)
    tk.Button(selector, text="✅ Done - Extract Links", command=extract_and_close).pack(pady=10)

# === Main logic ===
def process_file(file_path):
    ext = file_path.lower()
    if ext.endswith(".pdf"):
        show_pdf_and_select_tables(file_path)
    elif ext.endswith(".csv"):
        result = extract_links_from_csv(file_path)
        display_and_save(file_path, result)
    elif ext.endswith(".xlsx"):
        result = extract_links_from_excel(file_path)
        display_and_save(file_path, result)
    else:
        messagebox.showerror("Unsupported", "Only PDF, CSV, and XLSX files are supported.")

def display_and_save(file_path, result):
    output_area.config(state='normal')
    output_area.delete("1.0", tk.END)
    if result:
        output_area.insert(tk.END, "\n".join(result))
        filename = os.path.splitext(os.path.basename(file_path))[0]
        save = messagebox.askyesno("Save?", f"Found {len(result)} link(s). Save to file?")
        if save:
            save_links_to_file(filename, result)
    else:
        output_area.insert(tk.END, "No links found.")
        messagebox.showinfo("No Links", "No hyperlinks found.")
    output_area.config(state='disabled')

def browse_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Supported Files", "*.pdf *.csv *.xlsx")],
        title="Select PDF/CSV/XLSX File"
    )
    if file_path:
        process_file(file_path)

def choose_save_location():
    global save_location
    path = filedialog.asksaveasfilename(
        title="Choose where to save extracted links",
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")]
    )
    if path:
        save_location = path
        messagebox.showinfo("Save Location Set", f"📁 Save location set to:\n{save_location}")

def on_drop(event):
    file_path = event.data.strip("{ }")
    process_file(file_path)

# === GUI ===
def show_gui():
    global root, output_area
    if root is None or not root.winfo_exists():
        root = TkinterDnD.Tk() if DND_ENABLED else tk.Tk()
        root.title("📄 Universal Link Extractor")
        root.geometry("740x520")
        root.resizable(False, False)

        tk.Label(root, text="🔗 Extract Links from PDF, Excel or CSV", font=("Helvetica", 16)).pack(pady=10)
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="📁 Browse File", command=browse_file, font=("Helvetica", 11)).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="📂 Choose Save Location", command=choose_save_location, font=("Helvetica", 11)).grid(row=0, column=1, padx=10)

        output_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=90, height=20, font=("Courier", 10))
        output_area.pack(padx=10, pady=10)
        output_area.insert(tk.END, "📎 Drag a file here or click 'Browse' to begin.\n")
        output_area.config(state='disabled')

        if DND_ENABLED:
            root.drop_target_register(DND_FILES)
            root.dnd_bind('<<Drop>>', on_drop)
        else:
            output_area.config(state='normal')
            output_area.insert(tk.END, "\n[!] Drag-and-drop not available.\n")
            output_area.config(state='disabled')

        root.protocol("WM_DELETE_WINDOW", hide_gui)
    else:
        root.deiconify()
        root.lift()

# === Tray Functions ===
def hide_gui():
    if root:
        root.withdraw()

def quit_app(icon, item):
    icon.stop()
    if root:
        root.destroy()

def run_tray():
    menu = Menu(
        MenuItem("Open Link Extractor", lambda icon, item: show_gui()),
        MenuItem("Quit", quit_app)
    )
    icon = Icon("LinkExtractor", create_tray_icon(), "Link Extractor", menu)
    threading.Thread(target=icon.run, daemon=True).start()
    return icon

# === Entry ===
if __name__ == '__main__':
    silent = '--silent' in sys.argv
    tray_icon = run_tray()
    if not silent:
        show_gui()
    tk.mainloop()
