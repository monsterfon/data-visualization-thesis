# Data Visualization Thesis

This repository contains the code for my **Data Visualization Thesis**, which focuses on automating the creation of calibration graphs for **electromagnetic compatibility** testing. The project was designed for the company **Malle** to simplify the tedious and error-prone process of generating graphs manually during calibration. The repository includes three files, one of which is the main project for the thesis, while the others are small utility scripts for office automation.

## Projects Overview

### 1. **Comparison of Reference Calibrations (Primerjave referenčnih kalibracij)**

The main part of the thesis involves automating the creation of calibration graphs used in **electromagnetic compatibility (EMC)** testing. The company Malle needed these graphs for calibration purposes, but generating them manually was time-consuming and inefficient. 

The solution was a Python program that:
- **Automatically generates calibration graphs** from provided data, removing the need for manual graph creation.
- **Runs indefinitely** with the ability to generate new graphs with each set of calibration data received multiple times a year.
- Includes a **user interface** for tweaking graph parameters in the rare case of problematic graphs, though the program ensures high-quality results most of the time.

This project demonstrates knowledge in **data visualization**, **data analysis**, and **electromagnetic compatibility**—the key area of interest. The automation process streamlined the workflow, saving time and reducing errors during the calibration process.

### 2. **Office Automation Utilities**

The repository also includes two smaller projects:
- **xlm2xls**: A utility that converts **XML** files to **XLS** (Excel) files.
- **xlm2xls.toFile**: A slight variation of the previous utility, designed to directly export the converted files.

These tools were created for office automation purposes, assisting the same company (Malle) with their file format conversions but are not the main focus of the thesis.

## Technologies Used:
- **Python** for the automation and data visualization program.
- **Matplotlib** and **Seaborn** for generating the calibration graphs.
- **Tkinter** for the user interface (UI) to tweak parameters.
- **pandas** for handling the data and preparing it for visualization.

## How to Run

1. Clone the repository.
2. Install the required libraries (use `requirements.txt` to install dependencies).
3. Run the main Python script for the calibration graph generation. Provide the necessary data in the specified format.
4. Use the UI for any adjustments to graph parameters if needed.

---

This project is a comprehensive demonstration of **data visualization**, **automation**, and an in-depth understanding of **electromagnetic compatibility**. The main objective was to significantly reduce the time and effort required to generate calibration graphs, while ensuring accuracy and quality. The utility scripts for office automation were an additional enhancement for the company.
