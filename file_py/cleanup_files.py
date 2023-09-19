import os
import shutil

# Percorso del file merged.csv
merged_csv_file = 'merged.csv'

# Percorso del file merged.xlsx
merged_xlsx_file = 'merged.xlsx'

# Directory per l'archiviazione
archive_directory = '../xlsx_finale'

# Crea la directory per l'archiviazione se non esiste
if not os.path.exists(archive_directory):
    os.makedirs(archive_directory)

# Elimina il file merged.csv se esiste
if os.path.exists(merged_csv_file):
    os.remove(merged_csv_file)
    print(f'File {merged_csv_file} eliminato.')

# Sposta il file merged.xlsx nella directory di archiviazione
if os.path.exists(merged_xlsx_file):
    destination_path = os.path.join(archive_directory, 'merged.xlsx')
    shutil.move(merged_xlsx_file, destination_path)
    print(f'File {merged_xlsx_file} spostato in {destination_path}.')
