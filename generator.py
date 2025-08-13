import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import string
import qrcode
from reportlab.lib.pagesizes import mm, A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
from PIL import Image, ImageTk, ImageDraw, ImageFont

# -------------------------------------------------------
# Hilfsfunktionen für Daten
# -------------------------------------------------------

def generate_lagerliste(lagerort, regal_typ, regale, faecher, ebenen, besondere_orte):
    """Erzeugt Lagerliste mit Regal/Fach/Ebene in absteigender Reihenfolge"""
    data = []
    if regale > 0 and faecher > 0 and ebenen > 0:
        if regal_typ == "Buchstaben":
            regal_labels = list(string.ascii_uppercase)[:regale]
        else:
            regal_labels = [str(i+1) for i in range(regale)]
        regal_labels = regal_labels[::-1]  # absteigend
        for regal in regal_labels:
            for fach in range(faecher, 0, -1):
                for ebene in range(ebenen, 0, -1):
                    qr_data = f"{lagerort};{regal}-{fach}-{ebene}"
                    data.append([lagerort, regal, fach, ebene, "", qr_data])
    for ort in besondere_orte:
        ort = ort.strip()
        if ort:
            qr_data = f"{ort};"
            data.append([ort, "", "", "", ort, qr_data])
    df = pd.DataFrame(data, columns=[
        "Lagerort", "Regal", "Fach", "Ebene", "besondere Lagerorte", "Daten für QR-Code"
    ])
    return df

def save_excel(df):
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel-Dateien", "*.xlsx")]
    )
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Erfolg", f"Excel-Datei gespeichert:\n{file_path}")

def format_val(val):
    if pd.notna(val):
        try:
            if float(val).is_integer():
                return str(int(val))
        except:
            pass
        return str(val)
    return ""

# -------------------------------------------------------
# Hilfsfunktionen für PDF
# -------------------------------------------------------

def fit_text_to_width(c, text, max_width, max_font_size, min_font_size=6, font_name="Helvetica-Bold"):
    if not text:
        return min_font_size
    font_size = max_font_size
    while font_size >= min_font_size:
        c.setFont(font_name, font_size)
        if c.stringWidth(text, font_name, font_size) <= max_width:
            return font_size
        font_size -= 0.5
    return min_font_size

# -------------------------------------------------------
# Punkt 1: Komplette PDF (Einzel-Etiketten 70x32mm)
# -------------------------------------------------------

def create_qr_labels_from_excel(excel_path, output_pdf):
    df = pd.read_excel(excel_path)
    df = df[df["besondere Lagerorte"].isna() | (df["besondere Lagerorte"].astype(str).str.strip() == "")]
    df = df.iloc[::-1]  # umgekehrte Reihenfolge

    page_w = 70 * mm
    page_h = 32 * mm
    qr_size = 22 * mm
    text_x = 26 * mm
    text_w = page_w - text_x - 2*mm

    c = canvas.Canvas(output_pdf, pagesize=(page_w, page_h))

    for _, row in df.iterrows():
        qr_value = str(row["Daten für QR-Code"])
        lagerort = str(row["Lagerort"])
        regal = format_val(row["Regal"])
        fach = format_val(row["Fach"])
        ebene = format_val(row["Ebene"])
        lagerplatz = f"{regal}-{fach}-{ebene}"

        qr_img = qrcode.make(qr_value).convert("RGB")
        buf = BytesIO()
        qr_img.save(buf, format="PNG")
        buf.seek(0)
        c.drawImage(ImageReader(buf), 2*mm, 5*mm, qr_size, qr_size)

        fs_lo = fit_text_to_width(c, lagerort, text_w, 10, font_name="Helvetica")
        fs_lp = fit_text_to_width(c, lagerplatz, text_w, 22, font_name="Helvetica-Bold")
        total_h = fs_lo + fs_lp + fs_lo
        start_y = (page_h - total_h)/2

        c.setFont("Helvetica", fs_lo)
        tw = c.stringWidth(lagerort, "Helvetica", fs_lo)
        c.drawString(text_x + (text_w - tw)/2, start_y + fs_lp + fs_lo, lagerort)

        c.setFont("Helvetica-Bold", fs_lp)
        tw = c.stringWidth(lagerplatz, "Helvetica-Bold", fs_lp)
        c.drawString(text_x + (text_w - tw)/2, start_y, lagerplatz)

        c.showPage()

    c.save()
    messagebox.showinfo("Erfolg", f"PDF erstellt: {output_pdf}")

# -------------------------------------------------------
# Punkt 2: 2x8 Raster mit festen 70x32mm Feldern
# -------------------------------------------------------

def create_qr_labels_a4(excel_path, output_pdf):
    df = pd.read_excel(excel_path)
    df = df[df["besondere Lagerorte"].isna() | (df["besondere Lagerorte"].astype(str).str.strip() == "")]
    df = df.iloc[::-1]

    label_w = 70 * mm
    label_h = 32 * mm
    cols = 2
    rows = 8
    page_w, page_h = A4
    x_margin = (page_w - cols*label_w) / 2
    y_margin = (page_h - rows*label_h) / 2
    qr_size = 22 * mm
    text_x_offset = 26 * mm
    text_w = label_w - text_x_offset - 2*mm

    c = canvas.Canvas(output_pdf, pagesize=A4)

    col = 0
    row_i = 0
    x = x_margin
    y = page_h - y_margin - label_h

    for _, r in df.iterrows():
        qr_value = str(r["Daten für QR-Code"])
        lagerort = str(r["Lagerort"])
        regal = format_val(r["Regal"])
        fach = format_val(r["Fach"])
        ebene = format_val(r["Ebene"])
        lagerplatz = f"{regal}-{fach}-{ebene}"

        qr_img = qrcode.make(qr_value).convert("RGB")
        buf = BytesIO()
        qr_img.save(buf, format="PNG")
        buf.seek(0)
        c.drawImage(ImageReader(buf), x + 2*mm, y + 5*mm, qr_size, qr_size)

        fs_lo = fit_text_to_width(c, lagerort, text_w, 10, font_name="Helvetica")
        fs_lp = fit_text_to_width(c, lagerplatz, text_w, 22, font_name="Helvetica-Bold")
        total_h = fs_lo + fs_lp + fs_lo
        start_y = y + (label_h - total_h)/2

        c.setFont("Helvetica", fs_lo)
        tw = c.stringWidth(lagerort, "Helvetica", fs_lo)
        c.drawString(x + text_x_offset + (text_w - tw)/2, start_y + fs_lp + fs_lo, lagerort)

        c.setFont("Helvetica-Bold", fs_lp)
        tw = c.stringWidth(lagerplatz, "Helvetica-Bold", fs_lp)
        c.drawString(x + text_x_offset + (text_w - tw)/2, start_y, lagerplatz)

        col += 1
        x += label_w
        if col >= cols:
            col = 0
            x = x_margin
            row_i += 1
            y -= label_h
            if row_i >= rows:
                row_i = 0
                y = page_h - y_margin - label_h
                c.showPage()

    c.save()
    messagebox.showinfo("Erfolg", f"A4 PDF erstellt: {output_pdf}")

# -------------------------------------------------------
# Punkt 3: Einzelnes Etikett
# -------------------------------------------------------

def create_single_qr(input_text, output_pdf):
    page_w = 70 * mm
    page_h = 32 * mm
    qr_size = 22 * mm
    text_x = 26 * mm
    text_w = page_w - text_x - 2*mm

    c = canvas.Canvas(output_pdf, pagesize=(page_w, page_h))

    qr_img = qrcode.make(input_text).convert("RGB")
    buf = BytesIO()
    qr_img.save(buf, format="PNG")
    buf.seek(0)
    c.drawImage(ImageReader(buf), 2*mm, 5*mm, qr_size, qr_size)

    try:
        lagerort, lagerplatz = input_text.split(";", 1)
    except:
        lagerort = input_text
        lagerplatz = ""

    fs_lo = fit_text_to_width(c, lagerort, text_w, 10, font_name="Helvetica")
    fs_lp = fit_text_to_width(c, lagerplatz, text_w, 22, font_name="Helvetica-Bold")
    total_h = fs_lo + fs_lp + fs_lo
    start_y = (page_h - total_h)/2

    c.setFont("Helvetica", fs_lo)
    tw = c.stringWidth(lagerort, "Helvetica", fs_lo)
    c.drawString(text_x + (text_w - tw)/2, start_y + fs_lp + fs_lo, lagerort)

    c.setFont("Helvetica-Bold", fs_lp)
    tw = c.stringWidth(lagerplatz, "Helvetica-Bold", fs_lp)
    c.drawString(text_x + (text_w - tw)/2, start_y, lagerplatz)

    c.showPage()
    c.save()
    messagebox.showinfo("Erfolg", f"Einzelnes QR-Etikett erstellt: {output_pdf}")

# -------------------------------------------------------
# Punkt 4: PDF für besondere Lagerorte
# -------------------------------------------------------

def create_special_locations_pdf(excel_path, output_pdf):
    df = pd.read_excel(excel_path)
    df = df[df["besondere Lagerorte"].notna() & (df["besondere Lagerorte"].astype(str).str.strip() != "")]
    c = canvas.Canvas(output_pdf, pagesize=A4)
    page_w, page_h = A4

    for _, row in df.iterrows():
        qr_value = str(row["Daten für QR-Code"])
        lagerort = str(row["Lagerort"])
        regal = format_val(row["Regal"])
        fach = format_val(row["Fach"])
        ebene = format_val(row["Ebene"])
        lagerplatz = f"{regal}-{fach}-{ebene}" if regal else ""

        # 1/3 Seite für Text
        text_height = page_h / 3
        fs_lo = fit_text_to_width(c, lagerort, page_w - 40, 48, font_name="Helvetica")
        fs_lp = fit_text_to_width(c, lagerplatz, page_w - 40, 60, font_name="Helvetica-Bold")

        # Lagerort oben
        c.setFont("Helvetica", fs_lo)
        tw = c.stringWidth(lagerort, "Helvetica", fs_lo)
        y_cursor = page_h - fs_lo*2
        c.drawString((page_w - tw)/2, y_cursor, lagerort)

        # Lagerplatz darunter
        c.setFont("Helvetica-Bold", fs_lp)
        tw = c.stringWidth(lagerplatz, "Helvetica-Bold", fs_lp)
        y_cursor -= (fs_lp + 10)
        c.drawString((page_w - tw)/2, y_cursor, lagerplatz)

        # QR-Code in den restlichen 2/3 der Seite
        available_h = y_cursor - 50
        qr_size = min(available_h, page_w - 100)
        qr_img = qrcode.make(qr_value).convert("RGB")
        buf = BytesIO()
        qr_img.save(buf, format="PNG")
        buf.seek(0)
        c.drawImage(ImageReader(buf), (page_w - qr_size)/2, 50, qr_size, qr_size)

        c.showPage()

    c.save()
    messagebox.showinfo("Erfolg", f"Besondere Lagerorte PDF erstellt: {output_pdf}")

# -------------------------------------------------------
# Vorschau (korrekte Reihenfolge & Größen)
# -------------------------------------------------------

def get_ttf():
    for name in ("arial.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(name, 20)
        except:
            continue
    return ImageFont.load_default()

def pil_fit_text(draw, text, font_base, max_width_px, max_pt):
    pt = max_pt
    while pt >= 6:
        try:
            font = ImageFont.truetype(font_base.path, int(pt))
        except:
            font = font_base
        w, _ = draw.textsize(text, font=font)
        if w <= max_width_px:
            return font
        pt -= 1
    return font_base

def render_preview(input_text):
    w, h = 700, 320
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    mm_to_px = lambda mm: int(mm * (w / 70))
    qr_img = qrcode.make(input_text).convert("RGB")
    qr_h = mm_to_px(22)
    qr_img = qr_img.resize((qr_h, qr_h))
    img.paste(qr_img, (mm_to_px(2), mm_to_px(5)))

    try:
        lagerort, lagerplatz = input_text.split(";", 1)
    except:
        lagerort = input_text
        lagerplatz = ""
    text_x = mm_to_px(26)
    text_w = w - text_x - mm_to_px(2)
    base_font = get_ttf()
    font_lo = pil_fit_text(draw, lagerort, base_font, text_w, 36)
    font_lp = pil_fit_text(draw, lagerplatz, base_font, text_w, 50)
    h_lo = draw.textsize(lagerort, font=font_lo)[1]
    h_lp = draw.textsize(lagerplatz, font=font_lp)[1]
    total_h = h_lo + h_lp + h_lo
    start_y = (h - total_h)//2
    w_lo = draw.textsize(lagerort, font=font_lo)[0]
    draw.text((text_x + (text_w - w_lo)//2, start_y + h_lp + h_lo), lagerort, font=font_lo, fill="black")
    w_lp = draw.textsize(lagerplatz, font=font_lp)[0]
    draw.text((text_x + (text_w - w_lp)//2, start_y), lagerplatz, font=font_lp, fill="black")
    return img

def update_preview(*args):
    global preview_photo
    text = entry_single_qr.get()
    img = render_preview(text)
    preview_photo = ImageTk.PhotoImage(img)
    preview_label.configure(image=preview_photo)

# -------------------------------------------------------
# GUI
# -------------------------------------------------------

root = tk.Tk()
root.title("Lagerkoordinaten & QR-Code Generator")
root.geometry("850x800")
tabs = ttk.Notebook(root)

# Tab 1
tab1 = ttk.Frame(tabs)
tabs.add(tab1, text="Lagerliste erstellen")
ttk.Label(tab1, text="Lagerort:").pack()
entry_lagerort = ttk.Entry(tab1); entry_lagerort.pack()
regal_typ_var = tk.StringVar(value="Buchstaben")
ttk.Radiobutton(tab1, text="Buchstaben", variable=regal_typ_var, value="Buchstaben").pack()
ttk.Radiobutton(tab1, text="Zahlen", variable=regal_typ_var, value="Zahlen").pack()
ttk.Label(tab1, text="Anzahl Regale:").pack()
entry_regale = ttk.Entry(tab1); entry_regale.pack()
ttk.Label(tab1, text="Anzahl Fächer:").pack()
entry_faecher = ttk.Entry(tab1); entry_faecher.pack()
ttk.Label(tab1, text="Anzahl Ebenen:").pack()
entry_ebenen = ttk.Entry(tab1); entry_ebenen.pack()
ttk.Label(tab1, text="Besondere Lagerorte:").pack()
text_besondere = tk.Text(tab1, height=5); text_besondere.pack()
ttk.Button(tab1, text="Excel erstellen", command=lambda: save_excel(
    generate_lagerliste(entry_lagerort.get(), regal_typ_var.get(),
                        int(entry_regale.get() or 0),
                        int(entry_faecher.get() or 0),
                        int(entry_ebenen.get() or 0),
                        text_besondere.get("1.0", tk.END).splitlines())
)).pack()

# Tab 2
tab2 = ttk.Frame(tabs)
tabs.add(tab2, text="QR-Codes erzeugen")
ttk.Button(tab2, text="1. Komplette PDF (ohne besondere Lagerorte)", command=lambda: create_qr_labels_from_excel(
    filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")]),
    filedialog.asksaveasfilename(defaultextension=".pdf")
)).pack(pady=5)
ttk.Label(tab2, text="(Benötigt vorher erstellte Excel-Liste)", foreground="grey").pack()

ttk.Button(tab2, text="2. A4 PDF (2×8 Felder à 70×32mm)", command=lambda: create_qr_labels_a4(
    filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")]),
    filedialog.asksaveasfilename(defaultextension=".pdf")
)).pack(pady=5)

ttk.Label(tab2, text="3. Einzelnes Lagerplatzetikett:").pack()
entry_single_qr = ttk.Entry(tab2, width=40); entry_single_qr.insert(0, "Lagerort;Lagerplatz"); entry_single_qr.pack()
entry_single_qr.bind("<KeyRelease>", update_preview)
preview_label = ttk.Label(tab2); preview_label.pack()
ttk.Button(tab2, text="PDF erzeugen", command=lambda: create_single_qr(entry_single_qr.get(),
    filedialog.asksaveasfilename(defaultextension=".pdf"))).pack()

ttk.Button(tab2, text="4. PDF für besondere Lagerorte", command=lambda: create_special_locations_pdf(
    filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")]),
    filedialog.asksaveasfilename(defaultextension=".pdf")
)).pack(pady=5)
ttk.Label(tab2, text="(Benötigt vorher erstellte Excel-Liste)", foreground="grey").pack()

tabs.pack(expand=1, fill="both")
update_preview()
root.mainloop()
