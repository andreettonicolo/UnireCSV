import pandas as pd

def merge_csv_files(csv_files):
    merged_data = pd.DataFrame()
    for csv_file in csv_files:
        df = pd.read_csv(csv_file, sep='\t')
        merged_data = pd.concat([merged_data, df], ignore_index=True)

    # Rimuovi colonne vuote dal DataFrame
    merged_data = merged_data.dropna(axis=1, how='all')

    return merged_data
