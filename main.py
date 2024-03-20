# Define the directory paths
from datetime import datetime
import os
import random
import pandas as pd
from hpdb_module import hpdbCheck
from kmz_module import kmzCheck
import threading
import time
import concurrent.futures

def main():
    log_columns = ['Cluster ID', 'Checking date', 'Checking Time']

    print("FORCE is Running...")

    clusters = [cluster for cluster in os.listdir('Input') if os.path.isdir(os.path.join('Input', cluster))]

    for cluster in clusters:
        input_dir = f"Input\{cluster}"
        summary_dir = f"Summary\{cluster}"
        kmz_file_path = os.path.join(input_dir, f"ABD - {cluster}.kmz")
        hpdb_file_path = os.path.join(input_dir, f"HPDB - {cluster}.xlsx")

        # Membuat direktori jika belum ada
        os.makedirs(summary_dir, exist_ok=True)

        print(hpdb_file_path)

        start_time = time.time()

        # Get current date
        checking_date = datetime.today().strftime('%Y-%m-%d')

        # Get current time
        checking_time = datetime.now().strftime('%H:%M:%S')

        hpdbCheck(hpdb_file_path, kmz_file_path, cluster, checking_date, checking_time)
        # kmz_df = kmzCheck(kmz_file_path)

        # Jalankan kedua fungsi secara paralel
        # with concurrent.futures.ThreadPoolExecutor() as executor:
        #     future1 = executor.submit(hpdbCheck, hpdb_file_path, kmz_file_path)
        #     future2 = executor.submit(kmzCheck, kmz_file_path)

        #     # Ambil hasil kembali dari kedua fungsi
        #     hdpb_df = future1.result()
        #     kmz_df = future2.result()

        end_time = time.time()

        execution_time = end_time - start_time
        print("Execution time:", execution_time, "seconds")

if __name__ == "__main__":
    main()