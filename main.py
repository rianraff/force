# Define the directory paths
from datetime import datetime
import os
import random
import pandas as pd
from validator_module import hpdbCheck
from kmz_module import kmzCheck
import threading
import time
import concurrent.futures

def main():
    log_columns = ['Cluster ID', 'Checking date', 'Checking Time']

    print("FORCE is Running...")

    clusters = os.listdir('Input')

    for cluster in clusters:
        input_dir = f"Input\{cluster}"
        summary_dir = f"Summary\{cluster}"
        kmz_file_path = os.path.join(input_dir, f"ABD - {cluster}.kmz")
        hpdb_file_path = os.path.join(input_dir, f"HPDB - {cluster}.xlsx")
        summary_file_path = os.path.join(summary_dir, f"Summary_{cluster}.xlsx")

        # Membuat direktori jika belum ada
        os.makedirs(summary_dir, exist_ok=True)

        print(hpdb_file_path)

        start_time = time.time()

        # hdpb_df = hpdbCheck(hpdb_file_path)
        # kmz_df = kmzCheck(kmz_file_path)

        # Jalankan kedua fungsi secara paralel
        with concurrent.futures.ThreadPoolExecutor() as executor:
            future1 = executor.submit(hpdbCheck, hpdb_file_path)
            future2 = executor.submit(kmzCheck, kmz_file_path)

            # Ambil hasil kembali dari kedua fungsi
            hdpb_df = future1.result()
            kmz_df = future2.result()

        end_time = time.time()

        execution_time = end_time - start_time
        print("Execution time:", execution_time, "seconds")

        # Get current date
        checking_date = datetime.today().strftime('%Y-%m-%d')

        # Get current time
        checking_time = datetime.now().strftime('%H:%M:%S')

        # Create a DataFrame with the log_columns and the new row
        log_columns_df = pd.DataFrame(columns=log_columns)
        log_columns_df.loc[0] = [cluster, checking_date, checking_time]

        summary_df = pd.concat([log_columns_df, kmz_df, hdpb_df], axis=1, ignore_index=False)

        # Create temp_master.xlsx with headers from master_temp_df if it doesn't exist
        if not os.path.exists(summary_file_path):
            summary_df.to_excel(summary_file_path, index=False, engine='openpyxl')
        
        else:
            with pd.ExcelWriter(summary_file_path, 'openpyxl', mode='a',  if_sheet_exists="overlay") as writer:
                # fix line
                reader = pd.read_excel(summary_file_path, sheet_name="Sheet1")
                summary_df.to_excel(writer, sheet_name="Sheet1", index=False, engine='openpyxl', header=False, startrow=len(reader)+1)

if __name__ == "__main__":
    main()