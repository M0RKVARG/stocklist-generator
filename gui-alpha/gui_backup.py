import dearpygui.dearpygui as dpg
import xlsxwriter
import time
import sys
import os

path = os.path.dirname(os.path.abspath(__file__))
timestr = time.strftime("%Y-%m-%d_%H-%M")

#Variablen initial festlegen
regale = []
faecher = []
ebenen = []
sonder = []
anzahl_rs = []

dpg.create_context()

def print_value(sender): #Rückgabe der Eingabefelder
    print(dpg.get_value(sender)) 

def button_callback(sender, app_data):
    print(f"sender is: {sender}")
    print(f"app_data is: {app_data}")

dpg.create_viewport(title='etirehotel Lagerlisten-Generator', width=600, height=800) #Hauptfenster festlegen

width, height, channels, data = dpg.load_image("./Logo-etirehotel_klein.png")

with dpg.texture_registry(show=False):
    eth_logo = dpg.add_static_texture(width, height, data) #eth Logo als Grafik einlesen

with dpg.window(label="Generator", width=585, height=770, tag="Hauptfenster"):
    dpg.draw_image(eth_logo, (100, 0), (470, 200))
    dpg.add_spacer(height = 200) #Abstand oberer Bildschirmrand bis zum ersten Eingabefeld

    typ_regale = dpg.add_combo(
        label = "Beschriftung der Regale",
        items=("Buchstaben", "Zahlen"),
        default_value="Buchstaben",
        callback = print_value)
    dpg.add_spacer(height = 20) #Abstand zwischen den Eingabefeldern

    input_lagerort = dpg.add_input_text(
        label="Lagerort", 
        default_value="Hier bitte den Lagerort eintragen",
        callback = print_value)
    dpg.add_spacer(height = 20)

    input_regale = dpg.add_input_text(
        label="Buchstaben Regale", 
        default_value="Bsp.: A, B, C, D,...",
        callback = print_value)
    dpg.add_spacer(height = 20)

    input_faecher = dpg.add_slider_int(
        label="Anzahl Fächer (Breite)", 
        default_value = 1, 
        max_value = 100, 
        min_value = 1,
        callback = print_value)
    dpg.add_spacer(height = 20)

    input_ebenen = dpg.add_slider_int(
        label="Anzahl Ebenen (Höhe)", 
        default_value = 1, 
        max_value = 20, 
        min_value = 1,
        callback = print_value)
    dpg.add_spacer(height = 20)

    input_rs = dpg.add_slider_int(
        label="Anzahl Radsätze je Fach", 
        default_value = 1, 
        max_value = 5, 
        min_value = 1,
        callback=print_value)
    dpg.add_spacer(height = 20)

    input_sonder = dpg.add_input_text(
        label="Sonderlagerplätze", 
        default_value="Bsp. Waschplatz, Vorkommissionierung, Kunde,...",
        callback = print_value)
    dpg.add_spacer(height = 50)

    dpg.add_button(label="Liste erstellen", callback=button_callback)


dpg.setup_dearpygui()
dpg.show_viewport()
dpg.set_primary_window("Hauptfenster", True)
dpg.start_dearpygui()
dpg.destroy_context()

