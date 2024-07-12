import csv
from datetime import datetime
import pyfiglet
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.chart import LineChart, Reference
import re

def stampa_titolo():
    titolo = pyfiglet.figlet_format("GlycemyStats")
    print(titolo)

def inserisci_dati_glicemia():
    while True:
        data = input("Inserisci la data (YYYY-MM-DD o DD/MM/YYYY o DD\\MM\\YYYY): ")
        ora = input("Inserisci l'ora (HH:MM o HH.MM): ")
        livello = input("Inserisci il livello di glicemia (2 o 3 cifre): ")

        try:
            # Prova a riconoscere diversi formati di data e ora
            if "/" in data:
                data_ora = datetime.strptime(f"{data} {ora.replace('.', ':')}", "%d/%m/%Y %H:%M")
            elif "\\" in data:
                data_ora = datetime.strptime(f"{data} {ora.replace('.', ':')}", "%d\\%m\\%Y %H:%M")
            else:
                data_ora = datetime.strptime(f"{data} {ora.replace('.', ':')}", "%Y-%m-%d %H:%M")
            
            # Verifica che il livello di glicemia sia un numero di 2 o 3 cifre
            if re.match(r"^\d{2,3}$", livello):
                livello = float(livello)
            else:
                raise ValueError("Livello di glicemia non valido")
            
            break
        except ValueError as ve:
            print(f"Errore: {ve}. Riprova.")
        except Exception as e:
            print(f"Errore di formato: {e}. Riprova.")

    with open("livelli_glicemia.csv", mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([data_ora.strftime("%Y-%m-%d"), data_ora.strftime("%H:%M"), livello])
        print("Dati inseriti correttamente.")

def crea_file_excel():
    # Crea un nuovo file Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Livelli di Glicemia"
    
    # Scrivi l'intestazione
    ws.append(["Data", "Ora", "Livello di Glicemia"])
    
    # Leggi i dati dal file CSV e scrivili nel file Excel
    with open("livelli_glicemia.csv", mode='r') as file:
        reader = csv.reader(file)
        for row in reader:
            ws.append(row)
    
    # Applica la formattazione condizionale
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
        for cell in row:
            if 101 <= cell.value <= 129:
                cell.fill = red_fill
            elif 70 <= cell.value <= 99:
                cell.fill = green_fill

    # Crea il grafico
    chart = LineChart()
    chart.title = "Livelli di Glicemia"
    chart.style = 10
    chart.y_axis.title = 'Livello di Glicemia'
    chart.x_axis.title = 'Data e Ora'

    data = Reference(ws, min_col=3, min_row=1, max_row=ws.max_row)
    categories = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    ws.add_chart(chart, "E5")
    
    # Salva il file Excel
    wb.save("livelli_glicemia.xlsx")

def main():
    stampa_titolo()
    while True:
        inserisci_dati_glicemia()
        continua = input("Vuoi inserire un altro dato? (s/n): ").lower()
        if continua != 's':
            break
    crea_file_excel()
    print("File Excel creato correttamente: livelli_glicemia.xlsx")

if __name__ == "__main__":
    main()
