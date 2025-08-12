import xlsxwriter
import time
import sys
import os

print("""
 _        _    ____ _____ ____  _____ _____ ___ _  _______ _____ _____ _____ _   _ 
| |      / \  / ___| ____|  _ \| ____|_   _|_ _| |/ / ____|_   _|_   _| ____| \ | |
| |     / _ \| |  _|  _| | |_) |  _|   | |  | || ' /|  _|   | |   | | |  _| |  \| |
| |___ / ___ \ |_| | |___|  _ <| |___  | |  | || . \| |___  | |   | | | |___| |\  |
|_____/_/   \_\____|_____|_| \_\_____| |_| |___|_|\_\_____| |_|   |_| |_____|_| \_|
                                                                                   
 _     ___ ____ _____            ____ _____ _   _ _____ ____      _  _____ ___  ____  
| |   |_ _/ ___|_   _|          / ___| ____| \ | | ____|  _ \    / \|_   _/ _ \|  _ \ 
| |    | |\___ \ | |    _____  | |  _|  _| |  \| |  _| | |_) |  / _ \ | || | | | |_) |
| |___ | | ___) || |   |_____| | |_| | |___| |\  | |___|  _ <  / ___ \| || |_| |  _ < 
|_____|___|____/ |_|            \____|_____|_| \_|_____|_| \_\/_/   \_\_| \___/|_| \_|

""")

path = os.path.dirname(os.path.abspath(__file__))
timestr = time.strftime("%Y-%m-%d_%H-%M")

#Variablen initial festlegen
regale = []
faecher = []
ebenen = []
sonder = []
anzahl_rs = []

#function für Regalbeschriftung mit Buchstaben
def regale_bst():
    regale = []
    input_regale = input("Wie sind die Regale bezeichnet? Bitte mit Komma getrennt eingeben.\nBsp.: A, B, C, D, E\nEingabe: ")
    print("")
    regale = input_regale.split(", ")
    regale.reverse()
    return regale

#function für Regalbeschriftung mit Zahlen
def regale_zahlen():
    regale = []
    anzahl_regale = int(input("Wie viele Regale sind zu beschriften?\nEingabe: "))
    print("")
    for r in range(1,anzahl_regale+1):
        regale.append(r)
    regale.reverse()
    return regale

typ_regale = input("Bitte angeben, wie die Regale nummieriert werden sollen.\n1 für Buchstaben oder 2 für Zahlen eingeben.\nEingabe: ")
print("")

if typ_regale == "1":
    regale = regale_bst()
elif typ_regale == "2":
    regale = regale_zahlen()
else:
#sys.exit bricht Programm mit entsprechender Fehlermeldung ab
    sys.exit("Falsche Eingabe!!! Programm wird beendet.")

input_lagerort = input("Zu welchem Lagerort gehören diese Regale?\nEingabe: ")
print("")
input_faecher = int(input("Wie viele Fächer haben die Regale jeweils in der Breite?\nEingabe: "))
print("")
input_rs = int(input("Wie viele Radsätze werden pro Fach gelagert?\nEingabe: "))
print("")
input_ebenen = int(input("Wie viele Ebenen haben die Regale in der Höhe?\nEingabe: "))
print("")
input_sonder = input("Welche Sonderlagerplätze werden gewünscht? Bitte mit Komma getrennt eingeben.\nBsp: Waschplatz, Vorkommissionierung, Kunde\nEingabe: ")
print("")



for f in range(1, input_faecher+1):
    faecher.append(f)

for e in range(1, input_ebenen+1):
    ebenen.append(e)

for rs in range(1, input_rs+1):
    anzahl_rs.append(rs)

#regale = input_regale.split(", ")
sonder = input_sonder.split(", ")

#regale.reverse()
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

print(f"Die Lagerliste wurde in {timestr}_Lagerliste_{input_lagerort}.xlsx abgelegt")