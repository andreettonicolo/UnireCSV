import os
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Border, Side
from merge_csv import merge_csv_files
from convert_to_xlsx import convert_to_xlsx

# Directory contenente i file CSV da unire
directory = '../csv_da_unire'

# Nome del file CSV risultante
output_csv_file = 'merged.csv'
# Nome del file XLSX risultante
output_xlsx_file = 'merged.xlsx'

# Lista per tenere traccia dei nomi dei file CSV da unire
csv_files = []

# Leggi i nomi dei file CSV dalla directory
for filename in os.listdir(directory):
    if filename.endswith('.csv'):
        csv_files.append(os.path.join(directory, filename))

# Unisci i file CSV in un unico DataFrame
merged_data = merge_csv_files(csv_files)

# Scrivi il risultato in un nuovo file CSV
merged_data.to_csv(output_csv_file, index=False)

# Converti il CSV in XLSX
convert_to_xlsx(merged_data, output_xlsx_file)

# Esegui la pulizia dei file
os.system('python cleanup_files.py')

# Esegui l'elaborazione del file XLSX
os.system('python process_xlsx.py')

print('Pulizia dei file e elaborazione del file XLSX completate.')

print('Conversione in XLSX completata. Il file XLSX risultante è:', output_xlsx_file)







