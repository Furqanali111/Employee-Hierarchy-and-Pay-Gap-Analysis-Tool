This repository contains a Python script that processes employee data to generate hierarchical reports and analyze pay gaps using Excel files. The script utilizes libraries such as pandas, xlwings, and os to manipulate and organize data from multiple Excel sheets. It creates hierarchical relationships between employees and supervisors, checks for specific conditions, and generates detailed Excel reports.

# Features
Data Extraction and Processing: Reads data from multiple Excel files, including employee details, supervisor details, and template files.
Hierarchical Relationship Mapping: Constructs hierarchical relationships between supervisors and their subordinates.
Data Filtering and Analysis: Filters employees based on specific conditions and analyzes data to identify key metrics.
Excel Report Generation: Uses xlwings to generate and manipulate Excel reports, creating folders and files based on supervisor IDs and names.
Flexible and Scalable: Handles large datasets efficiently and can be customized for various data processing and reporting needs.

# Installation
Install Dependencies: Ensure you have Python installed, and install the required libraries using pip:
pip install pandas xlwings openpyxl

# Download and Setup:

Clone this repository to your local machine.
Place your input Excel files (Calibration1.xlsx, Directors.xlsx, VP.xlsx) in the specified paths or update the script with the correct paths.

# Prepare Input Files:

Calibration1.xlsx: Main data file containing employee details.
Directors.xlsx: File containing employee IDs.
VP.xlsx: File containing supervisor IDs.

# Output:

The script generates hierarchical reports in the form of Excel files, organized in folders based on VP IDs and names.
