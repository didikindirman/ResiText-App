# ResiText: Logistics Waybill Automation Tool

## Overview

**ResiText** is a powerful, localized Python application designed to eliminate manual data entry in high-volume e-commerce logistics and warehouse operations. Developed to solve the common bottleneck of transcribing order data onto printed waybill receipts, this tool automates the process of reading product descriptions and special notes from an Excel file and accurately inserting them onto existing PDF waybills.

The application includes robust pre-processing checks to ensure data integrity, preventing mis-shipments by validating file counts and waybill IDs before processing.

## Key Features

* **Custom Text Insertion:** Automatically reads data from specific Excel columns (Item Description and Special Notes) and inserts the text onto individual pages of multiple PDF waybills.
* **Automatic Data Validation:** Performs critical **three-point checks** to ensure accuracy:
    1.  **Count Match:** Verifies that the total number of data rows in Excel perfectly matches the total number of pages across all selected PDF files.
    2.  **File Name Check:** Validates that selected PDF files adhere to the required sorting standard (`ELEVEN` for Ascending or `FAMI` for Descending mode).
    3.  **Waybill ID Check (Ascending Only):** Automatically validates the last 5 digits of the Waybill ID extracted from the PDF against the corresponding data in Excel (Column G/7).
* **Dual Sorting Modes:** Supports two primary logistics sorting requirements: **Ascending (Top-to-Bottom)** and **Descending (Bottom-to-Top)** order for data processing.
* **Special Instruction Handling:** Detects specific keywords (e.g., '60') in the special notes column (Column C) and inserts a highly visible warning text (`(INSERT 60 NT !!)`) before the product description.
* **PDF Manipulation:** Uses `reportlab` to precisely insert, rotate, and center text at a specific coordinate on the waybill documents.
* **User-Friendly GUI:** Features a clean graphical user interface built with `tkinter` for easy file selection, status monitoring, and drag-and-drop file reordering.

## Prerequisites

To run **ResiText**, you must have Python installed, along with the required libraries.

* Python (3.x recommended)

## Installation

1.  **Clone the Repository (or download the script):**
    ```bash
    git clone [Your Repository URL Here]
    cd ResiText
    ```

2.  **Install Dependencies:**
    The application relies on several essential Python libraries for data processing and PDF manipulation.
    ```bash
    pip install pandas PyPDF2 reportlab pdfplumber openpyxl
    ```
    *(Note: `openpyxl` is required by `pandas` to read `.xlsx` files.)*

## Usage Guide

1.  **Place Files:** Ensure your primary Excel file (`.xlsx` or `.xls`) is placed in the same directory as the `ResiText.py` script.
2.  **Run the Application:**
    ```bash
    python ResiText.py
    ```

3.  **GUI Steps:**
    * **Step 1: Sort Order Type:** Select the required sorting logic (`Ascending` for 7-Eleven or `Descending` for Family-Mart). This selection triggers the file name validation check.
    * **Step 2: Excel Data:** The tool will automatically detect your Excel file and display the total number of data rows (Column A). You can click "Edit Excel file" to open it.
    * **Step 3: Select Waybill PDF(s):** Click **"Add PDF(s)"** and select all waybill files you wish to process.
        * The file names will be validated against your selected sort order keyword (`ELEVEN` or `FAMI`).
        * Use the **▲ Up / ▼ Down** buttons to correct the processing order if necessary.
        * The **"Status Check"** box will confirm if the Excel data rows, PDF pages, and Waybill IDs match. **The process will not start unless all checks pass (green).**
    * **Start:** Click **"Start"**, choose the save location for your processed PDF, and the automation will begin.

## Required Excel Structure

The automation script relies on specific columns in your Excel file (assuming 0-indexed columns, or A, B, C...):

| Column Index | Column Name | Purpose | Required? |
| :--- | :--- | :--- | :--- |
| **0 (A)** | Item Description | The main text to be inserted onto the PDF waybill. | YES |
| **2 (C)** | Special Notes / Slip Data | Used for checking specific instructions (e.g., if it contains '60', special text is added). | YES (Used for logic) |
| **6 (G)** | Waybill Number | Used for the automatic validation check (last 5 digits) when **Ascending** mode is selected. | YES (For Ascending check) |