import pandas as pd
import random

def kmzCheck():

    gilang_col = [
    'Result', 'Pole to FAT', 'Pole to FDT',
    'HP to pole 35m', 'Coordinate HP to pole 35m', 'HP to FAT 150m', 'Coordinate HP to FAT 150m',
    'OPM result Check', 'OPM 1310 nm', 'FAT Coordinate', 'FAT Naming', 'FAT Count', 'Format'
]
    
    # Create a DataFrame with column names from column_names
    gilang_df = pd.DataFrame(columns=gilang_col)
    random_number = random.randint(000, 999)

    new_row = {}
    for col_name in gilang_col:
        new_row[col_name] = f"Gilang {random_number}"

    # Append the new row to the DataFrame
    gilang_df = gilang_df._append(new_row, ignore_index=True)

    return pd.DataFrame(gilang_df)