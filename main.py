# Define the directory paths
import os
import validator_module

def main():
    input_dir = 'Input'
    output_dir = 'Output'
    isStandby = False

    print("System is Running...")

    while True:
        # List all files in the raw directory
        files = os.listdir(input_dir)
        # Filter files that start with "HPDB" and end with ".xlsx"
        hpdb_files = [file for file in files if file.startswith("HPDB") and file.endswith(".xlsx")]

        if len(hpdb_files) == 0 and isStandby == False:
            print("No files found. Standby...")
            isStandby = True
            continue

        for file in hpdb_files:
            isStandby = False
            hdpb_df = validator_module.hpdbCheck(file)

if __name__ == "__main__":
    main()