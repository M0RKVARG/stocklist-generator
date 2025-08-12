import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import string
import qrcode
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
from PIL import Image

# ------------------------
# Hilfsfunktionen
# ------------------------

def generate_lagerliste(lagerort, regal_typ, regale, faecher, ebenen, besondere_orte):
    data = []
    if regale > 0 and faecher > 0 and ebenen > 0:
        if regal_typ == "Buchstaben":
            regal_labels = list(string.ascii_uppercase)[:regale]
        else:
            regal_labels = [str(i+1) for i in range(regale)]

        # Absteigend sortieren
        regal_labels = regal_labels[::-1]

        for regal in regal_labels:
            for fach in range(faecher, 0, -1):  # absteigend
                for ebene in range(ebenen, 0, -1):  # absteigend
                    qr_data = f"{lagerort};{regal}-{fach}-{ebene}"
                    data.append([lagerort, regal, fach, ebene, "", qr_data])

    # Besondere Lagerorte hinzufügen
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

def fit_text_to_width(canvas_obj, text, max_width, max_font_size, min_font_size=6, font_name="Helvetica-Bold"):
    """Findet die größte Schriftgröße, bei der der Text in max_width passt."""
    font_size = max_font_size
    while font_size >= min_font_size:
        canvas_obj.setFont(font_name, font_size)
        if canvas_obj.stringWidth(text, font_name, font_size) <= max_width:
            return font_size
        font_size -= 0.5
    return min_font_size

def create_qr_labels_from_excel(excel_path, output_pdf):
    df = pd.read_excel(excel_path)

    page_width = 70 * mm
    page_height = 32 * mm
    qr_size = 22 * mm
    text_area_x = 26 * mm
    text_area_width = page_width - text_area_x - 2 * mm

    c = canvas.Canvas(output_pdf, pagesize=(page_width, page_height))

    for _, row in df.iterrows():
        qr_value = str(row["Daten für QR-Code"])
        lagerort = str(row["Lagerort"])

        # Zahlen korrekt formatieren (kein .0 mehr)
        def format_val(val):
            if pd.notna(val):
                try:
                    if float(val).is_integer():
                        return str(int(val))
                except:
                    pass
                return str(val)
            return ""

        regal = format_val(row["Regal"])
        fach = format_val(row["Fach"])
        ebene = format_val(row["Ebene"])
        besondere = str(row["besondere Lagerorte"]) if pd.notna(row["besondere Lagerorte"]) else ""

        # QR-Code generieren
        qr_img = qrcode.make(qr_value).convert("RGB")
        qr_buffer = BytesIO()
        qr_img.save(qr_buffer, format="PNG")
        qr_buffer.seek(0)
        qr_reader = ImageReader(qr_buffer)

        # QR-Code links platzieren
        c.drawImage(qr_reader, 2 * mm, 5 * mm, qr_size, qr_size)

        # Lagerplatz-Text
        if besondere.strip():
            lagerplatz_text = ""
        else:
            lagerplatz_text = f"{regal}-{fach}-{ebene}"

        # Schriftgrößen anpassen
        lagerort_font_size = fit_text_to_width(c, lagerort, text_area_width, max_font_size=10, font_name="Helvetica")
        lagerplatz_font_size = fit_text_to_width(c, lagerplatz_text, text_area_width, max_font_size=22, font_name="Helvetica-Bold")

        # Gesamthöhe inkl. Leerzeile
        total_text_height = lagerort_font_size + lagerplatz_font_size + lagerort_font_size  # zweite Zeile = Leerzeilenhöhe
        start_y = (page_height - total_text_height) / 2

        # Lagerort zeichnen
        c.setFont("Helvetica", lagerort_font_size)
        text_width = c.stringWidth(lagerort, "Helvetica", lagerort_font_size)
        c.drawString(text_area_x + (text_area_width - text_width) / 2, start_y + lagerplatz_font_size + lagerort_font_size, lagerort)

        # Leerzeile (Platzhalterhöhe = lagerort_font_size)
        # Lagerplatz zeichnen
        c.setFont("Helvetica-Bold", lagerplatz_font_size)
        text_width = c.stringWidth(lagerplatz_text, "Helvetica-Bold", lagerplatz_font_size)
        c.drawString(text_area_x + (text_area_width - text_width) / 2, start_y, lagerplatz_text)

        c.showPage()

    c.save()
    messagebox.showinfo("Erfolg", f"PDF mit QR-Codes gespeichert:\n{output_pdf}")

# ------------------------
# GUI-Funktionen
# ------------------------

def create_lagerliste_gui():
    lagerort = entry_lagerort.get()
    regal_typ = regal_typ_var.get()
    try:
        regale = int(entry_regale.get())
    except:
        regale = 0
    try:
        faecher = int(entry_faecher.get())
    except:
        faecher = 0
    try:
        ebenen = int(entry_ebenen.get())
    except:
        ebenen = 0

    besondere_orte = text_besondere.get("1.0", tk.END).splitlines()

    df = generate_lagerliste(lagerort, regal_typ, regale, faecher, ebenen, besondere_orte)
    save_excel(df)

def generate_qr_gui():
    excel_path = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xlsx")])
    if not excel_path:
        return
    output_pdf = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF-Dateien", "*.pdf")]
    )
    if not output_pdf:
        return
    create_qr_labels_from_excel(excel_path, output_pdf)

# ------------------------
# Haupt-GUI
# ------------------------

root = tk.Tk()
root.title("Lagerkoordinaten & QR-Code Generator")
root.geometry("500x500")

tabControl = ttk.Notebook(root)

# Tab 1: Lagerliste erstellen
tab1 = ttk.Frame(tabControl)
tabControl.add(tab1, text="Lagerliste erstellen")

ttk.Label(tab1, text="Lagerort:").pack(pady=5)
entry_lagerort = ttk.Entry(tab1)
entry_lagerort.pack(pady=5)

ttk.Label(tab1, text="Regaltyp:").pack(pady=5)
regal_typ_var = tk.StringVar(value="Buchstaben")
ttk.Radiobutton(tab1, text="Buchstaben", variable=regal_typ_var, value="Buchstaben").pack()
ttk.Radiobutton(tab1, text="Zahlen", variable=regal_typ_var, value="Zahlen").pack()

ttk.Label(tab1, text="Anzahl Regale:").pack(pady=5)
entry_regale = ttk.Entry(tab1)
entry_regale.pack(pady=5)

ttk.Label(tab1, text="Anzahl Fächer:").pack(pady=5)
entry_faecher = ttk.Entry(tab1)
entry_faecher.pack(pady=5)

ttk.Label(tab1, text="Anzahl Ebenen:").pack(pady=5)
entry_ebenen = ttk.Entry(tab1)
entry_ebenen.pack(pady=5)

ttk.Label(tab1, text="Besondere Lagerorte (jeweils in neuer Zeile):").pack(pady=5)
text_besondere = tk.Text(tab1, height=5)
text_besondere.pack(pady=5)

ttk.Button(tab1, text="Excel erstellen", command=create_lagerliste_gui).pack(pady=10)

# Tab 2: QR-Codes aus Excel generieren
tab2 = ttk.Frame(tabControl)
tabControl.add(tab2, text="QR-Codes erzeugen")

ttk.Button(tab2, text="Excel-Datei wählen und PDF erzeugen", command=generate_qr_gui).pack(pady=20)

tabControl.pack(expand=1, fill="both")

root.mainloop()