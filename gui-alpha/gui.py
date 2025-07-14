import dearpygui.dearpygui as dpg
import xlsxwriter
import time
import sys
import os

path = os.path.dirname(os.path.abspath(__file__))
timestr = time.strftime("%Y-%m-%d_%H-%M")

# Variablen initial festlegen
regale = []
faecher = []
ebenen = []
sonder = []
anzahl_rs = []

dpg.create_context()

def print_value(sender, data): # Ausgabe des aktuellen Wertes eines Eingabefelds im Terminal
    print(sender, " returned: ", dpg.get_value(sender))

def list_generate(sender, data):
    input_lagerort = dpg.get_value("lagerort")
    input_regale = dpg.get_value("regal_bst")
    input_faecher = dpg.get_value("faecher")
    input_ebenen = dpg.get_value("ebenen")
    input_rs = dpg.get_value("anz_rs")
    input_sonder = dpg.get_value("sonder")
    print(input_lagerort)
    print(input_faecher)
    print(input_ebenen)
    print(input_rs)
    print(input_sonder)

    # function für Regalbeschriftung mit Buchstaben
    # def regale_bst(input_regale):
    regale = []
    regale = input_regale.split(", ")
    regale.reverse()
    #   return regale

    # function für Regalbeschriftung mit Zahlen
    # def regale_zahlen():
    #     regale = []
    #     anzahl_regale = int(input("Wie viele Regale sind zu beschriften?\nEingabe: "))
    #     print("")
    #     for r in range(1,anzahl_regale+1):
    #         regale.append(r)
    #     regale.reverse()
    #     return regale

    for f in range(1, input_faecher+1):
        faecher.append(f)

    for e in range(1, input_ebenen+1):
        ebenen.append(e)

    for rs in range(1, input_rs+1):
        anzahl_rs.append(rs)

    sonder = input_sonder.split(", ")

    faecher.reverse()
    ebenen.reverse()
    anzahl_rs.reverse()

    workbook = xlsxwriter.Workbook(f"{path}/{timestr}_Lagerliste_{input_lagerort}.xlsx")
    worksheet = workbook.add_worksheet("Lager")

    worksheet.write("A1", "Lagerort")
    worksheet.write("B1", "Regal")
    worksheet.write("C1", "Fach")
    worksheet.write("D1", "Ebene")
    worksheet.write("E1", "Radsatz")
    worksheet.write("G1", "Sonderlagerplätze:")

    zeile = 2

    for r in regale:
        for f in faecher:
            for e in ebenen:
                for rs in anzahl_rs:
                    worksheet.write("B"+str(zeile), r)
                    worksheet.write("C"+str(zeile), f)
                    worksheet.write("D"+str(zeile), e)
                    if input_rs > 1:
                        worksheet.write("E"+str(zeile), rs)
                    if input_lagerort != "":
                        worksheet.write("A"+str(zeile), input_lagerort) 
                    zeile += 1

    if sonder != []:
        zeile = 2
        for s in sonder:
            worksheet.write("G"+str(zeile), s)
            zeile +=1

    workbook.close()



dpg.create_viewport(title='etirehotel Lagerlisten-Generator', width=600, height=800) #Hauptfenster festlegen

width, height, channels, data = dpg.load_image("c:/Users/jonas.mueller/Downloads/Python/Lagerlisten Generator GUI/Logo-etirehotel_klein.png")

with dpg.texture_registry(show=False):
    eth_logo = dpg.add_static_texture(width, height, data) #eth Logo als Grafik einlesen

with dpg.window(label="Generator", width=585, height=770, tag="Hauptfenster"):
    dpg.draw_image(eth_logo, (100, 0), (470, 200))
    dpg.add_spacer(height = 200) #Abstand oberer Bildschirmrand bis zum ersten Eingabefeld

    dpg.add_combo(
        label = "Beschriftung der Regale",
        items=("Buchstaben", "Zahlen"),
        default_value="Buchstaben",
        tag = "regaltyp",
        callback = print_value)
    typ_regale = dpg.get_value("Beschriftung der Regale")
    dpg.add_spacer(height = 20) #Abstand zwischen den Eingabefeldern

    dpg.add_input_text(
        label="Lagerort", 
        default_value="Hier bitte den Lagerort eintragen",
        tag = "lagerort",
        callback = print_value)
    input_lagerort = dpg.get_value("Lagerort")
    dpg.add_spacer(height = 20)

    dpg.add_input_text(
        label="Buchstaben Regale", 
        default_value="Bsp.: A, B, C, D,...",
        tag = "regal_bst",
        callback = print_value)
    input_regale = dpg.get_value("Buchstaben Regale")
    dpg.add_spacer(height = 20)

    dpg.add_slider_int(
        label="Anzahl Fächer (Breite)", 
        default_value = 1, 
        max_value = 100, 
        min_value = 1,
        tag = "faecher",
        callback = print_value)
    input_faecher = dpg.get_value("Anzahl Fächer (Breite")
    dpg.add_spacer(height = 20)

    dpg.add_slider_int(
        label="Anzahl Ebenen (Höhe)", 
        default_value = 1, 
        max_value = 20, 
        min_value = 1,
        tag = "ebenen",
        callback = print_value)
    input_ebenen = dpg.get_value("Anzahl Ebenen (Höhe")
    dpg.add_spacer(height = 20)

    dpg.add_slider_int(
        label="Anzahl Radsätze je Fach", 
        default_value = 1, 
        max_value = 5, 
        min_value = 1,
        tag = "anz_rs",
        callback=print_value)
    input_rs = dpg.get_value("Anzahl Radsätze je Fach")
    dpg.add_spacer(height = 20)

    dpg.add_input_text(
        label="Sonderlagerplätze", 
        default_value="Bsp. Waschplatz, Vorkommissionierung, Kunde,...",
        tag = "sonder",
        callback = print_value)
    input_sonder = dpg.get_value("Sonderlagerplätze")
    dpg.add_spacer(height = 50)

    dpg.add_button(tag = "Btn xlsx", label="Liste erstellen",  callback=list_generate)


dpg.setup_dearpygui()
dpg.show_viewport()
dpg.set_primary_window("Hauptfenster", True)
dpg.start_dearpygui()
dpg.destroy_context()