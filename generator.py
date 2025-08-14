import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont
import pandas as pd
import string
import qrcode
from reportlab.lib.pagesizes import mm, A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
from PIL import Image, ImageTk, ImageDraw, ImageFont


# --------------------------------
# Konfiguration: Hintergrundgrafik für Sonderlagerorte A4-PDF
# --------------------------------
# Lege die Datei unter diesem Pfad ab (Standard: im gleichen Ordner wie das Skript).
BACKGROUND_IMAGE_PATH = "QR_A4_Hintergrund_hoch.png"  # z.B. "sonderlagerorte_bg.jpg" oder "assets/bg.png"
BACKGROUND_FIT_MODE = "cover"  # "cover" (seitenfüllend, Beschnitt möglich) oder "contain" (komplett sichtbar, evtl. Ränder)


# ----------------------------
# Hilfsfunktionen für Daten
# ----------------------------

def generate_lagerliste(lagerort, regal_typ, regale, faecher, ebenen, sonderorte):
    """
    Erzeugt die Excel-Liste wie gewünscht:
    - Zuerst normale Lagerplätze (Spalten A-G gefüllt, inkl. QR-Daten in E sowie LO/LP).
    - Sonderlagerorte werden von ganz oben beginnend in Spalte H eingetragen (= über vorhandene Zeilen),
      ohne neue Zeilen zu erzwingen. Falls mehr Sonderlagerorte existieren als normale Zeilen, werden
      zusätzliche Zeilen am Ende angefügt, in denen ausschließlich Spalte H gefüllt ist.
    Spalten: A Lagerort, B Regal, C Fach, D Ebene, E Daten für QR-Code, F LO, G LP, H Sonderlagerorte
    """
    normal_rows = []

    # 1) Normale Lagerplätze (wenn definiert)
    if regale > 0 and faecher > 0 and ebenen > 0 and lagerort:
        if regal_typ == "Buchstaben":
            regal_labels = list(string.ascii_uppercase)[:regale]
        else:
            regal_labels = [str(i + 1) for i in range(regale)]
        regal_labels = regal_labels[::-1]  # absteigend

        for regal in regal_labels:
            for fach in range(faecher, 0, -1):
                for ebene in range(ebenen, 0, -1):
                    qr_data = f"{lagerort};{regal}-{fach}-{ebene}"
                    # ["Lagerort","Regal","Fach","Ebene","Sonderlagerorte","Daten für QR-Code"]
                    normal_rows.append([lagerort, regal, fach, ebene, "", qr_data])

    # DataFrame zuerst nur mit normalen Zeilen
    df = pd.DataFrame(
        normal_rows,
        columns=["Lagerort", "Regal", "Fach", "Ebene", "Sonderlagerorte", "Daten für QR-Code"],
    )

    # LO/LP-Spalten anhand "Daten für QR-Code" erzeugen
    if "Daten für QR-Code" in df.columns:
        if df.empty:
            df["LO"] = []
            df["LP"] = []
        else:
            split_vals = df["Daten für QR-Code"].astype(str).str.split(";", n=1, expand=True)
            df["LO"] = split_vals[0].where(split_vals[0].notna(), "")
            df["LP"] = split_vals[1].fillna("")

    # Sonderorte bereinigen (leere Zeilen raus)
    sonder_clean = [o.strip() for o in sonderorte if o.strip()]

    # 2) Sonderlagerorte in Spalte H von oben eintragen
    for i, ort in enumerate(sonder_clean):
        if i < len(df):
            df.at[i, "Sonderlagerorte"] = ort
        else:
            # Zusätzliche Zeile anhängen, nur Spalte H befüllt
            df = pd.concat(
                [
                    df,
                    pd.DataFrame(
                        [["", "", "", "", ort, ""]],  # A-D leer, H=ort, E=QR-Daten leer
                        columns=["Lagerort", "Regal", "Fach", "Ebene", "Sonderlagerorte", "Daten für QR-Code"],
                    ),
                ],
                ignore_index=True,
            )

    # LO/LP für etwaige neu angehängte Zeilen ohne QR-Daten sicher leer halten
    if "LO" not in df.columns:
        df["LO"] = ""
    if "LP" not in df.columns:
        df["LP"] = ""
    if not df.empty:
        mask_no_qr = df["Daten für QR-Code"].astype(str).str.strip() == ""
        df.loc[mask_no_qr, ["LO", "LP"]] = ""

    # "Sonderlagerorte" ganz nach rechts neben LP verschieben
    col_order = ["Lagerort", "Regal", "Fach", "Ebene", "Daten für QR-Code", "LO", "LP", "Sonderlagerorte"]
    df = df[col_order]

    return df


def save_excel(df):
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel-Dateien", "*.xlsx")],
    )
    if not file_path:
        return

    # Versuche Pixelbreiten mit XlsxWriter zu setzen; Fallback auf OpenPyXL mit Zeichenbreiten
    try:
        with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
            sheet_name = "Tabelle1"
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]
            # Spaltenbreiten exakt in Pixeln setzen
            worksheet.set_column_pixels('A:A', 150)  # Lagerort
            worksheet.set_column_pixels('B:D', 80)   # Regal, Fach, Ebene
            worksheet.set_column_pixels('E:E', 250)  # Daten für QR-Code
            worksheet.set_column_pixels('F:G', 150)  # LO, LP
            worksheet.set_column_pixels('H:H', 250)  # Sonderlagerorte
        messagebox.showinfo("Erfolg", f"Excel-Datei gespeichert:\n{file_path}")
        return
    except Exception:
        pass

    # Fallback (OpenPyXL): Breite in "Zeichen" annähern (~ 1 Zeichen ≈ 7 Pixel)
    def px_to_chars(px):
        return round(px / 7.0, 1)

    try:
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            sheet_name = "Tabelle1"
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            ws.column_dimensions['A'].width = px_to_chars(150)
            ws.column_dimensions['B'].width = px_to_chars(80)
            ws.column_dimensions['C'].width = px_to_chars(80)
            ws.column_dimensions['D'].width = px_to_chars(80)
            ws.column_dimensions['E'].width = px_to_chars(250)
            ws.column_dimensions['F'].width = px_to_chars(150)
            ws.column_dimensions['G'].width = px_to_chars(150)
            ws.column_dimensions['H'].width = px_to_chars(250)
        messagebox.showinfo("Erfolg", f"Excel-Datei gespeichert:\n{file_path}")
    except Exception:
        # Letzter Fallback: ohne Formatierung speichern
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Erfolg", f"Excel-Datei gespeichert (ohne Spaltenbreiten-Formatierung):\n{file_path}")


def format_val(val):
    if pd.notna(val):
        try:
            if float(val).is_integer():
                return str(int(val))
        except Exception:
            pass
        return str(val)
    return ""


# ----------------------------
# Textmessung (Pillow-kompatibel)
# ----------------------------

def pil_measure_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont):
    """
    Misst Textbreite und -höhe robust, kompatibel zu neuen/alten Pillow-Versionen.
    """
    if hasattr(draw, "textbbox"):
        bbox = draw.textbbox((0, 0), text, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        return w, h
    # Fallback
    return draw.textsize(text, font=font)


# --------------------------------
# PDF-Helfer (ReportLab)
# --------------------------------

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


def _draw_background_fullpage(c, page_w, page_h, bg_path, mode="cover"):
    """
    Zeichnet eine Hintergrundgrafik seitenfüllend (cover) oder komplett sichtbar (contain) auf die Seite.
    """
    if not bg_path:
        return
    try:
        img = Image.open(bg_path)
        iw, ih = img.size
        ir = ImageReader(img)

        page_ratio = page_w / page_h
        img_ratio = iw / ih

        if mode == "contain":
            # Bild vollständig sichtbar, passt in die Seite (Ränder möglich)
            if img_ratio >= page_ratio:
                draw_w = page_w
                draw_h = page_w / img_ratio
            else:
                draw_h = page_h
                draw_w = page_h * img_ratio
            x = (page_w - draw_w) / 2
            y = (page_h - draw_h) / 2
        else:
            # cover: seitenfüllend, ggf. Beschnitt
            if img_ratio >= page_ratio:
                # durch Höhe begrenzt (Bild breiter als Seite)
                draw_h = page_h
                draw_w = page_h * img_ratio
            else:
                # durch Breite begrenzt (Bild schmaler als Seite)
                draw_w = page_w
                draw_h = page_w / img_ratio
            x = (page_w - draw_w) / 2
            y = (page_h - draw_h) / 2

        # Hintergrund zeichnen (unter allen späteren Inhalten)
        c.drawImage(ir, x, y, width=draw_w, height=draw_h, mask='auto')
    except Exception:
        # Optional: logging/Messagebox
        pass


# --------------------------------
# Punkt 1: Komplette PDF (Einzel-Etiketten dynamisch nach Format)
# --------------------------------

def create_qr_labels_from_excel(excel_path, output_pdf, fmt_value):
    if not excel_path or not output_pdf:
        return

    df = pd.read_excel(excel_path)

    # Wenn alte Excel mit "besondere Lagerorte" geöffnet wird: in "Sonderlagerorte" umbenennen
    if "besondere Lagerorte" in df.columns and "Sonderlagerorte" not in df.columns:
        df.rename(columns={"besondere Lagerorte": "Sonderlagerorte"}, inplace=True)

    # Alle Zeilen berücksichtigen; Spalte H ignorieren. Nur Zeilen mit QR-Daten verwenden.
    df = df[df["Daten für QR-Code"].astype(str).str.strip() != ""]
    df = df.iloc[::-1]  # umgekehrte Reihenfolge

    label_w_mm, label_h_mm, _, _ = get_label_specs(fmt_value)
    page_w = label_w_mm * mm
    page_h = label_h_mm * mm
    qr_size = 22 * mm
    text_x = 26 * mm
    text_w = page_w - text_x - 2 * mm

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
        y_qr = (page_h - qr_size) / 2
        c.drawImage(ImageReader(buf), 2 * mm, y_qr, qr_size, qr_size)

        fs_lo = fit_text_to_width(c, lagerort, text_w, 10, font_name="Helvetica")
        fs_lp = fit_text_to_width(c, lagerplatz, text_w, 22, font_name="Helvetica-Bold")
        total_h = fs_lo + fs_lp + fs_lo
        start_y = (page_h - total_h) / 2

        c.setFont("Helvetica", fs_lo)
        tw = c.stringWidth(lagerort, "Helvetica", fs_lo)
        c.drawString(text_x + (text_w - tw) / 2, start_y + fs_lp + fs_lo, lagerort)

        c.setFont("Helvetica-Bold", fs_lp)
        tw = c.stringWidth(lagerplatz, "Helvetica-Bold", fs_lp)
        c.drawString(text_x + (text_w - tw) / 2, start_y, lagerplatz)

        c.showPage()

    c.save()
    messagebox.showinfo("Erfolg", f"PDF erstellt: {output_pdf}")


# --------------------------------
# A4-PDF dynamisch nach Format
# --------------------------------

def get_label_specs(fmt_value: str):
    """
    Liefert (label_w_mm, label_h_mm, cols, rows) je nach gewähltem Etikettenformat.
    """
    if fmt_value == "75x25 mm":
        return 75, 25, 2, 10  # 2 Spalten x 10 Reihen
    # Default: 70x32 mm
    return 70, 32, 2, 8      # 2 Spalten x 8 Reihen


def create_qr_labels_a4(excel_path, output_pdf, fmt_value):
    if not excel_path or not output_pdf:
        return

    df = pd.read_excel(excel_path)
    # Abwärtskompatibel: Spalte umbenennen, falls nötig
    if "besondere Lagerorte" in df.columns and "Sonderlagerorte" not in df.columns:
        df.rename(columns={"besondere Lagerorte": "Sonderlagerorte"}, inplace=True)

    # Alle Zeilen berücksichtigen; Spalte H ignorieren. Nur Zeilen mit QR-Daten verwenden.
    df = df[df["Daten für QR-Code"].astype(str).str.strip() != ""]
    df = df.iloc[::-1]

    label_w_mm, label_h_mm, cols, rows = get_label_specs(fmt_value)
    label_w = label_w_mm * mm
    label_h = label_h_mm * mm

    page_w, page_h = A4
    x_margin = (page_w - cols * label_w) / 2
    y_margin = (page_h - rows * label_h) / 2
    qr_size = 22 * mm
    text_x_offset = 26 * mm
    text_w = label_w - text_x_offset - 2 * mm

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

        # QR links, vertikal zentriert
        qr_img = qrcode.make(qr_value).convert("RGB")
        buf = BytesIO()
        qr_img.save(buf, format="PNG")
        buf.seek(0)
        c.drawImage(ImageReader(buf), x + 2 * mm, y + (label_h - qr_size) / 2, qr_size, qr_size)

        # Texte (vertikal zentriert)
        fs_lo = fit_text_to_width(c, lagerort, text_w, 10, font_name="Helvetica")
        fs_lp = fit_text_to_width(c, lagerplatz, text_w, 22, font_name="Helvetica-Bold")
        total_h = fs_lo + fs_lp + fs_lo
        start_y = y + (label_h - total_h) / 2

        c.setFont("Helvetica", fs_lo)
        tw = c.stringWidth(lagerort, "Helvetica", fs_lo)
        c.drawString(x + text_x_offset + (text_w - tw) / 2, start_y + fs_lp + fs_lo, lagerort)

        c.setFont("Helvetica-Bold", fs_lp)
        tw = c.stringWidth(lagerplatz, "Helvetica-Bold", fs_lp)
        c.drawString(x + text_x_offset + (text_w - tw) / 2, start_y, lagerplatz)

        # Rahmen um das Label (1px-ähnlich, dünn, für Ausschneiden)
        c.setLineWidth(0.25)
        c.rect(x, y, label_w, label_h, stroke=1, fill=0)

        # nächste Zelle
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


# --------------------------------
# Punkt 3: Einzelnes Lagerplatzetikett
# --------------------------------

def create_single_qr(input_text, output_pdf, fmt_value):
    """
    Erzeugt ein einzelnes Etikett (je nach Auswahl 70x32mm oder 75x25mm) mit QR links und Texten rechts.
    input_text erwartet Format 'Lagerort;Lagerplatz'. Der komplette input_text wird als QR-Inhalt genutzt.
    """
    if not input_text or not output_pdf:
        return

    label_w_mm, label_h_mm, _, _ = get_label_specs(fmt_value)
    page_w = label_w_mm * mm
    page_h = label_h_mm * mm
    qr_size = 22 * mm
    text_x = 26 * mm
    text_w = page_w - text_x - 2 * mm

    # Zerlege Anzeige-Texte
    if ";" in input_text:
        lagerort, lagerplatz = input_text.split(";", 1)
    else:
        lagerort, lagerplatz = input_text, ""

    c = canvas.Canvas(output_pdf, pagesize=(page_w, page_h))

    # QR-Code
    qr_img = qrcode.make(input_text).convert("RGB")
    buf = BytesIO()
    qr_img.save(buf, format="PNG")
    buf.seek(0)
    y_qr = (page_h - qr_size) / 2
    c.drawImage(ImageReader(buf), 2 * mm, y_qr, qr_size, qr_size)

    # Texte vertikal zentriert
    fs_lo = fit_text_to_width(c, lagerort, text_w, 10, font_name="Helvetica")
    fs_lp = fit_text_to_width(c, lagerplatz, text_w, 22, font_name="Helvetica-Bold")
    total_h = fs_lo + fs_lp + fs_lo
    start_y = (page_h - total_h) / 2

    c.setFont("Helvetica", fs_lo)
    tw = c.stringWidth(lagerort, "Helvetica", fs_lo)
    c.drawString(text_x + (text_w - tw) / 2, start_y + fs_lp + fs_lo, lagerort)

    c.setFont("Helvetica-Bold", fs_lp)
    tw = c.stringWidth(lagerplatz, "Helvetica-Bold", fs_lp)
    c.drawString(text_x + (text_w - tw) / 2, start_y, lagerplatz)

    c.showPage()
    c.save()
    messagebox.showinfo("Erfolg", f"PDF erstellt: {output_pdf}")


# --------------------------------
# Punkt 4: PDF für Sonderlagerorte (mit festem Hintergrund)
# --------------------------------

def create_special_locations_pdf(excel_path, output_pdf):
    """
    Nutzt ausschließlich Spalte H ("Sonderlagerorte") aus der Excel-Datei.
    Auf jeder Seite wird eine feste Hintergrundgrafik gezeichnet (siehe BACKGROUND_IMAGE_PATH / BACKGROUND_FIT_MODE).
    QR-Code-Daten werden als "<Text>;" aufgebaut, alle anderen Felder sind irrelevant/leergelassen.
    """
    if not excel_path or not output_pdf:
        return

    df = pd.read_excel(excel_path)
    # Abwärtskompatibel: Spalte umbenennen, falls nötig
    if "besondere Lagerorte" in df.columns and "Sonderlagerorte" not in df.columns:
        df.rename(columns={"besondere Lagerorte": "Sonderlagerorte"}, inplace=True)

    # Nur Zeilen, die einen Eintrag in "Sonderlagerorte" haben
    df = df[df["Sonderlagerorte"].notna() & (df["Sonderlagerorte"].astype(str).str.strip() != "")]
    c = canvas.Canvas(output_pdf, pagesize=A4)
    page_w, page_h = A4

    for _, row in df.iterrows():
        # 1) Hintergrund pro Seite (hart eingebunden)
        _draw_background_fullpage(c, page_w, page_h, BACKGROUND_IMAGE_PATH, mode=BACKGROUND_FIT_MODE)

        special_text = str(row["Sonderlagerorte"]).strip()
        qr_value = f"{special_text};"  # QR aus Spalte H
        lagerort = special_text        # Obere Textzeile = spezieller Ort
        lagerplatz = ""                # Kein Lagerplatz für Sonderlagerorte

        # 1/3 Seite für Text
        fs_lo = fit_text_to_width(c, lagerort, page_w - 40, 48, font_name="Helvetica")
        fs_lp = fit_text_to_width(c, lagerplatz, page_w - 40, 60, font_name="Helvetica-Bold")

        # Lagerort oben
        c.setFont("Helvetica", fs_lo)
        tw = c.stringWidth(lagerort, "Helvetica", fs_lo)
        y_cursor = page_h - fs_lo * 4
        c.drawString((page_w - tw) / 2, y_cursor, lagerort)

        # Lagerplatz darunter (leer)
        c.setFont("Helvetica-Bold", fs_lp)
        tw = c.stringWidth(lagerplatz, "Helvetica-Bold", fs_lp)
        y_cursor -= (fs_lp + 10)
        c.drawString((page_w - tw) / 2, y_cursor, lagerplatz)

        # QR-Code in den restlichen 2/3 der Seite
        available_h = y_cursor - 50
        qr_size = min(available_h, page_w - 100)
        qr_img = qrcode.make(qr_value).convert("RGB")
        buf = BytesIO()
        qr_img.save(buf, format="PNG")
        buf.seek(0)
        c.drawImage(ImageReader(buf), (page_w - qr_size) / 2, 50, qr_size, qr_size)

        c.showPage()

    c.save()
    messagebox.showinfo("Erfolg", f"Sonderlagerorte PDF erstellt: {output_pdf}")


# --------------------------------
# Vorschau (korrekte Reihenfolge & Größen) + Rahmen
# --------------------------------

def get_ttf():
    for name in ("arial.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(name, 20)
        except Exception:
            continue
    return ImageFont.load_default()


def pil_fit_text(draw, text, font_base, max_width_px, max_pt):
    pt = max_pt
    while pt >= 6:
        try:
            font = ImageFont.truetype(font_base.path, int(pt))
        except Exception:
            font = font_base
        w, _ = pil_measure_text(draw, text, font)
        if w <= max_width_px:
            return font
        pt -= 1
    return font_base


def render_preview(input_text, fmt_value):
    # Maße abhängig vom Etikettenformat; 10 px pro mm
    if fmt_value == "75x25 mm":
        w_mm, h_mm = 75, 25
    else:
        w_mm, h_mm = 70, 32

    w, h = w_mm * 10, h_mm * 10
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)

    # 1px Rahmen um das gesamte Etikett
    draw.rectangle([(0, 0), (w - 1, h - 1)], outline="black", width=1)

    mm_to_px = lambda mm_v: int(mm_v * 10)

    # QR links, vertikal zentriert
    qr_h = mm_to_px(22)
    qr_img = qrcode.make(input_text).convert("RGB")
    qr_img = qr_img.resize((qr_h, qr_h))
    qr_x = mm_to_px(2)
    qr_y = (h - qr_h) // 2
    img.paste(qr_img, (qr_x, qr_y))

    # Texte
    try:
        lagerort, lagerplatz = input_text.split(";", 1)
    except Exception:
        lagerort = input_text
        lagerplatz = ""

    text_x = mm_to_px(26)
    text_w = w - text_x - mm_to_px(2)
    base_font = get_ttf()
    font_lo = pil_fit_text(draw, lagerort, base_font, text_w, 36)
    font_lp = pil_fit_text(draw, lagerplatz, base_font, text_w, 50)

    w_lo, h_lo = pil_measure_text(draw, lagerort, font_lo)
    w_lp, h_lp = pil_measure_text(draw, lagerplatz, font_lp)

    # Leerzeile/Abstand zwischen Lagerort und Lagerplatz
    gap_px = mm_to_px(4)

    # Gesamthöhe der beiden Textzeilen inkl. Abstand und vertikal zentrieren
    total_h = h_lo + gap_px + h_lp
    start_y = (h - total_h) // 2
    y_lo = start_y
    y_lp = y_lo + h_lo + gap_px

    # Lagerort oben
    draw.text((text_x + (text_w - w_lo) // 2, y_lo), lagerort, font=font_lo, fill="black")
    # Lagerplatz darunter
    draw.text((text_x + (text_w - w_lp) // 2, y_lp), lagerplatz, font=font_lp, fill="black")

    return img


def update_preview(*args):
    global preview_photo, a4_btn_text
    text = entry_single_qr.get()
    fmt_value = format_var.get()
    img = render_preview(text, fmt_value)

    # Vorschau kleiner skalieren (60%)
    scale = 0.6
    new_size = (max(1, int(img.width * scale)), max(1, int(img.height * scale)))
    img_small = img.resize(new_size, Image.LANCZOS)

    preview_photo = ImageTk.PhotoImage(img_small)
    preview_label.configure(image=preview_photo)

    # Button-Label für A4 dynamisch setzen
    if fmt_value == "75x25 mm":
        a4_btn_text.set("A4 PDF erzeugen (2 Spalten à 10 Reihen mit 75x25mm)")
    else:
        a4_btn_text.set("A4 PDF erzeugen (2 Spalten à 8 Reihen mit 70x32mm)")


# --------------------------------
# GUI
# --------------------------------

root = tk.Tk()
root.title("Lagerlisten & QR-Code Generator")
root.geometry("850x735")
root.resizable(False, False)

tabs = ttk.Notebook(root)

# Tab 1 (zentriert + Leerzeilen nach jedem Eingabefeld)
tab1 = ttk.Frame(tabs)
tabs.add(tab1, text="Lagerliste erstellen")

# Innerer zentrierter Container
tab1_center = ttk.Frame(tab1)
tab1_center.pack(anchor="center", pady=(10, 10))

# Grid im zentrierten Container
tab1_center.columnconfigure(0, weight=0)
tab1_center.columnconfigure(1, weight=0)
tab1_center.columnconfigure(2, weight=0)

# Lagerort
ttk.Label(tab1_center, text="Lagerort:").grid(row=0, column=0, columnspan=3, pady=(0, 4))
entry_lagerort = ttk.Entry(tab1_center, width=54)
entry_lagerort.grid(row=1, column=0, columnspan=3)
ttk.Label(tab1_center, text="").grid(row=2, column=0, columnspan=3)  # Leerzeile

# Anzahl Regale: Label über dem Feld, Eingabe + Radio-Buttons zentriert in einer Reihe
ttk.Label(tab1_center, text="Anzahl Regale:").grid(row=3, column=0, columnspan=3, pady=(0, 4))

regal_row = ttk.Frame(tab1_center)
regal_row.grid(row=4, column=0, columnspan=3)

entry_regale = ttk.Entry(regal_row, width=30)  # etwas breiter, Ausrichtung mit "Zahlen"
entry_regale.pack(side="left")

regaltyp_frame = ttk.Frame(regal_row)
regaltyp_frame.pack(side="left", padx=(10, 0))

regal_typ_var = tk.StringVar(value="Buchstaben")
ttk.Radiobutton(regaltyp_frame, text="Buchstaben", variable=regal_typ_var, value="Buchstaben").pack(side="left", padx=(0, 8))
ttk.Radiobutton(regaltyp_frame, text="Zahlen", variable=regal_typ_var, value="Zahlen").pack(side="left")

ttk.Label(tab1_center, text="").grid(row=5, column=0, columnspan=3)  # Leerzeile

# Anzahl Fächer
ttk.Label(tab1_center, text="Anzahl Fächer:").grid(row=6, column=0, columnspan=3, pady=(0, 4))
entry_faecher = ttk.Entry(tab1_center, width=54)
entry_faecher.grid(row=7, column=0, columnspan=3)
ttk.Label(tab1_center, text="").grid(row=8, column=0, columnspan=3)  # Leerzeile

# Anzahl Ebenen
ttk.Label(tab1_center, text="Anzahl Ebenen:").grid(row=9, column=0, columnspan=3, pady=(0, 4))
entry_ebenen = ttk.Entry(tab1_center, width=54)
entry_ebenen.grid(row=10, column=0, columnspan=3)
ttk.Label(tab1_center, text="").grid(row=11, column=0, columnspan=3)  # Leerzeile

# Sonderlagerorte (doppelte Höhe)
ttk.Label(tab1_center, text="Sonderlagerorte:").grid(row=12, column=0, columnspan=3, pady=(0, 4))
text_sonder = tk.Text(tab1_center, height=10, width=54)
text_sonder.grid(row=13, column=0, columnspan=3)

# Hinweis unter dem Eingabefeld
ttk.Label(tab1_center, text="Hinweis: Jeder Lagerort in eine separate Zeile.", foreground="grey").grid(
    row=14, column=0, columnspan=3, pady=(4, 0)
)

# Leerzeile
ttk.Label(tab1_center, text="").grid(row=15, column=0, columnspan=3)  # Leerzeile

# Excel erstellen Button
ttk.Button(
    tab1_center,
    text="Excel erstellen",
    command=lambda: save_excel(
        generate_lagerliste(
            entry_lagerort.get(),
            regal_typ_var.get(),
            int(entry_regale.get() or 0),
            int(entry_faecher.get() or 0),
            int(entry_ebenen.get() or 0),
            text_sonder.get("1.0", tk.END).splitlines(),
        )
    ),
).grid(row=16, column=0, columnspan=3, pady=(4, 12))

# Tab 2
tab2 = ttk.Frame(tabs)
tabs.add(tab2, text="QR-Codes erzeugen")

# Format-Auswahl ganz oben (zentriert & auffälliger)
format_section = ttk.Frame(tab2)
format_section.pack(fill="x", pady=(14, 10))

title_font = tkfont.Font(size=12, weight="bold")
ttk.Label(format_section, text="Etikettenformat wählen", font=title_font).pack(anchor="center")

format_var = tk.StringVar(value="70x32 mm")
fmt_opts = ttk.Frame(format_section)
fmt_opts.pack(anchor="center", pady=6)

rb1 = ttk.Radiobutton(
    fmt_opts,
    text="70×32 mm",
    variable=format_var,
    value="70x32 mm",
    command=update_preview,
)
rb1.pack(side="left", padx=12)

rb2 = ttk.Radiobutton(
    fmt_opts,
    text="75×25 mm",
    variable=format_var,
    value="75x25 mm",
    command=update_preview,
)
rb2.pack(side="left", padx=12)

# Leerzeile + Button: Komplette PDF (ohne Sonderlagerorte)
ttk.Label(tab2, text="").pack()
ttk.Button(
    tab2,
    text="Komplette PDF (ohne Sonderlagerorte)",
    command=lambda: create_qr_labels_from_excel(
        filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")]),
        filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF-Datei", "*.pdf")]),
        format_var.get(),
    ),
).pack(pady=8)
ttk.Label(tab2, text="(Benötigt vorher erstellte Excel-Liste)", foreground="grey").pack()

# Dynamisches Label für A4-Button
a4_btn_text = tk.StringVar(value="A4 PDF erzeugen (2 Spalten à 8 Reihen mit 70x32mm)")

# Leerzeile + Button: A4 PDF (dynamisch je nach Auswahl)
ttk.Label(tab2, text="").pack()
ttk.Button(
    tab2,
    textvariable=a4_btn_text,
    command=lambda: create_qr_labels_a4(
        filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")]),
        filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF-Datei", "*.pdf")]),
        format_var.get(),
    ),
).pack(pady=8)
ttk.Label(tab2, text="(Benötigt vorher erstellte Excel-Liste)", foreground="grey").pack()

# Leerzeile + Einzelnes Lagerplatzetikett
ttk.Label(tab2, text="").pack()
ttk.Label(tab2, text="Einzelnes Lagerplatzetikett:").pack(pady=(12, 4))
entry_single_qr = ttk.Entry(tab2, width=40)
entry_single_qr.insert(0, "Lagerort;Lagerplatz")
entry_single_qr.pack()
entry_single_qr.bind("<KeyRelease>", update_preview)

# Vorschau
preview_label = ttk.Label(tab2)
preview_label.pack(pady=(10, 8))

ttk.Button(
    tab2,
    text="PDF für einzelnes Lagerplatzetikett erzeugen (nutzt gewähltes Format)",
    command=lambda: create_single_qr(
        entry_single_qr.get(),
        filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF-Datei", "*.pdf")]),
        format_var.get(),
    ),
).pack(pady=(4, 12))

# Leerzeile + PDF für Sonderlagerorte
ttk.Label(tab2, text="").pack()
ttk.Button(
    tab2,
    text="A4 PDF für Sonderlagerorte",
    command=lambda: create_special_locations_pdf(
        filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")]),
        filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF-Datei", "*.pdf")]),
    ),
).pack(pady=8)
ttk.Label(tab2, text="(Benötigt vorher erstellte Excel-Liste)", foreground="grey").pack()

tabs.pack(expand=1, fill="both")

# Footer
footer = ttk.Label(root, text="© copyright 2025 - Jonas Müller - efleetcon®", foreground="grey")
footer.pack(side="bottom", pady=(6, 8))

# Initiale Vorschau (setzt auch initial den Button-Text korrekt)
def _init_preview():
    try:
        update_preview()
    except Exception:
        pass

_init_preview()

root.mainloop()