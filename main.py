# Define the directory paths
from datetime import datetime
import json
import os
import shutil
import pandas as pd
from hpdb_module import hpdbCheck
from kmz_module import kmzCheck
import time
import requests
import fnmatch
from colorama import Fore, Style
text = '''
 __________  ____  ____________   
/ ____/ __ \/ __ \/ ____/ ____/   
/ /_  / / / / /_/ / /   / __/      
/ __/ / /_/ / _, _/ /___/ /___      
/_/    \____/_/ |_|\____/_____/      
'''

def main():
    print(Fore.BLUE + text + Style.RESET_ALL)
    print("FORCE is Running...\n")
    standby = False
    force_base_url = 'http://localhost:5000'

    # Pola pencocokan nama file
    hpdb_pattern_rpa = '*RPA*.xlsx'
    hpdb_pattern_hpdb = '*HPDB*.xlsx'
    kmz_pattern = '*ABD*.kmz'

    while True:
        url = f'{force_base_url}/get_path_cluster_id'
        response = requests.get(url)

        if response.status_code == 200:
            data = response.json()
            if data.get('file_path') == None:  # Cek jika data kosong
                if standby:
                    time.sleep(60)  # Jika data kosong, tunggu 60 detik sebelum cek lagi
                    continue
                print('No Documents to Check. Stand by')
                standby = True
            else:
                standby = False
                file_path = data.get('file_path')
                cluster_id = data.get('cluster_id')
                print(f'File Path: {file_path}, Cluster ID: {cluster_id}')
        else:
            print('Failed to get data. Status code:', response.status_code)
        
        if not standby:

            input_dir = file_path
            files = os.listdir(input_dir)

            # Iterasi semua file dalam direktori
            for file in files:
                if fnmatch.fnmatch(file, hpdb_pattern_hpdb) or fnmatch.fnmatch(file, hpdb_pattern_rpa):
                    hpdb_file_path = os.path.join(input_dir, file)
                if fnmatch.fnmatch(file, kmz_pattern):
                    kmz_file_path = os.path.join(input_dir, file)

            summary_dir = f"Summary/{cluster_id}"

            # Membuat direktori jika belum ada
            os.makedirs(summary_dir, exist_ok=True)

            print(hpdb_file_path)

            start_time = time.time()

            # Get current date
            checking_date = datetime.today().strftime('%Y-%m-%d')

            # Get current time
            checking_time = datetime.now().strftime('%H:%M:%S')

            summary_file_path = os.path.join("Summary", f"Checking Summary.xlsx")
            hpdb_check = None
            kmz_check = None
            
            try:
                hpdb_check = hpdbCheck(hpdb_file_path, kmz_file_path, cluster_id, checking_date, checking_time)
            except:
                df_error = pd.DataFrame({"Cluster ID": [cluster_id], "Error":"HPDB file error"})
                with pd.ExcelWriter(summary_file_path, 'openpyxl', mode='a',  if_sheet_exists="overlay") as writer:
            # fix line
                    reader = pd.read_excel(summary_file_path, sheet_name="Invalid Files")
                    df_error.to_excel(writer, sheet_name="Invalid Files", index=False, engine='openpyxl', header=False, startrow=len(reader)+1)

            try:        
                kmz_check = kmzCheck(kmz_file_path, cluster_id, checking_date, checking_time)
            except:
                df_error = pd.DataFrame({"Cluster ID": [cluster_id], "Error":"KMZ file error"})
                with pd.ExcelWriter(summary_file_path, 'openpyxl', mode='a',  if_sheet_exists="overlay") as writer:
            # fix line
                    reader = pd.read_excel(summary_file_path, sheet_name="Invalid Files")
                    df_error.to_excel(writer, sheet_name="Invalid Files", index=False, engine='openpyxl', header=False, startrow=len(reader)+1)

            url = f'{force_base_url}/update_processed'

            data = {
                'cluster_id': cluster_id
            }
            headers = {'Content-Type': 'application/json'}
            response = requests.post(url, headers=headers, data=json.dumps(data))

            if response.status_code == 200:
                print('Data inserted successfully.')
            else:
                print('Failed to insert data. Status code:', response.status_code)

            # Check if HPDB or KMZ file need to be revised
            if hpdb_check or kmz_check:
                revise = True
            else:
                revise = False

            # force update revise
            url = f'{force_base_url}/update_revise'

            data = {
                'cluster_id': cluster_id,
                'revise': revise
            }
            headers = {'Content-Type': 'application/json'}
            response = requests.post(url, headers=headers, data=json.dumps(data))

            if response.status_code == 200:
                print('Data inserted successfully.')
            else:
                print('Failed to insert data. Status code:', response.status_code)

            #shutil.rmtree(file_path)
            
            end_time = time.time()

            execution_time = end_time - start_time
            print("Execution time:", execution_time, "seconds")
            time.sleep(5)

if __name__ == "__main__":
    main()
