# FORCE (Fiber Optic Readiness Check Engine)

## Overview
FORCE (Fiber Optic Readiness Check Engine) is a Python-based tool designed to automate the verification and validation of data related to fiber optic network readiness. The tool processes input data files, performs various checks and validations, and generates output files with summary reports highlighting any issues or discrepancies found in the input data.

## Features
- **Automated Data Processing**: FORCE automates the processing of input data files, reducing manual effort and ensuring consistency in data processing.
- **Data Validation**: The tool performs various validation checks on the input data to ensure its accuracy and completeness.
- **Error Reporting**: FORCE generates detailed summary reports highlighting any errors or issues found in the input data, making it easier for users to identify and rectify them.
- **Scalability**: The tool is designed to handle large volumes of data efficiently, making it suitable for use in large-scale fiber optic network deployments.

## Installation
FORCE can be installed using pip:

```
pip install force
```

## Usage
To use FORCE, simply run the main script `force.py` and provide the path to the input data files as arguments. For example:

```
python force.py --input_dir=Input --output_dir=Output
```

## Dependencies
FORCE relies on the following Python libraries:
- pandas
- openpyxl
- re
- datetime
- os
- shutil

These dependencies can be installed using pip:

```
pip install pandas openpyxl
```

## Collaborator
- Aldrian Raffi (GitHub: rianraff)
- Gilang Banyu (GitHub: gilangbbe)

## License
This project is under supervision of XL Axiata, Core Project Division

---