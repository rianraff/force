# Define the directory paths
from datetime import datetime
import os
import random
import pandas as pd
from validator_module import hpdbCheck
from kmz_module import kmzCheck

def main():
    log_columns = ['Cluster ID', 'Checking date', 'Checking Time']

    print("FORCE is Running...")

    clusters = ['SIT-00016', 'SIT-00017', 'SIT-00018']

    for cluster in clusters:
        input_dir = f"Input\{cluster}"
        summary_dir = f"Summary\{cluster}"
        kmz_file_path = os.path.join(input_dir, f"ABD - {cluster}.kmz")
        hpdb_file_path = os.path.join(input_dir, f"HPDB - {cluster}.xlsx")
        summary_file_path = os.path.join(summary_dir, f"Summary_{cluster}.xlsx")

        # Membuat direktori jika belum ada
        os.makedirs(summary_dir, exist_ok=True)

        print(hpdb_file_path)

        hdpb_df = hpdbCheck(hpdb_file_path)
        kmz_df = kmzCheck(kmz_file_path)

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