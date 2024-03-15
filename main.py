# Define the directory paths
from datetime import datetime
import os
import random
import pandas as pd
import validator_module
import kmz_module

def main():
    input_dir = 'Input'
    output_dir = 'Output'
    isStandby = False
    log_columns = ['Cluster ID', 'Checking date', 'Checking Time']

    print("FORCE is Running...")

    while True:
        # List all files in the raw directory
        files = os.listdir(input_dir)
        # Filter files that start with "HPDB" and end with ".xlsx"
        hpdb_files = [file for file in files if file.startswith("HPDB") and file.endswith(".xlsx")]
        

        if len(hpdb_files) == 0 and isStandby == False:
            isStandby = True
            continue

        else:
            for file in hpdb_files:
                isStandby = False
                hdpb_df = validator_module.hpdbCheck(file)
                kmz_df = kmz_module.kmzCheck()

                # Generate random Cluster ID
                cluster_id = 'XL-' + ''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=5))

                # Get current date
                checking_date = datetime.today().strftime('%Y-%m-%d')

                # Get current time
                checking_time = datetime.now().strftime('%H:%M:%S')

                # Create a DataFrame with the log_columns and the new row
                log_columns_df = pd.DataFrame(columns=log_columns)
                log_columns_df.loc[0] = [cluster_id, checking_date, checking_time]

                master_temp_df = pd.concat([log_columns_df, kmz_df, hdpb_df], axis=1, ignore_index=False)

                # Create temp_master.xlsx with headers from master_temp_df if it doesn't exist
                if not os.path.exists('temp_master.xlsx'):
                    master_temp_df = pd.concat([kmz_df, hdpb_df], axis=1, ignore_index=False)
                    master_temp_df.to_excel('temp_master.xlsx', index=False)

                with pd.ExcelWriter("temp_master.xlsx", 'openpyxl', mode='a',  if_sheet_exists="overlay") as writer:
                    # fix line
                    reader = pd.read_excel(r'temp_master.xlsx', sheet_name="Sheet1")
                    master_temp_df.to_excel(writer, "Sheet1", index=False, header=False, startrow=len(reader)+1)

if __name__ == "__main__":
    main()