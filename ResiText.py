import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import io
import glob
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import sys
import subprocess
import platform
import re
import pdfplumber 
import time 

# --- GLOBAL GUI & STATUS VARIABLES ---
pdf_file_path_list = [] 
is_name_match = False # NEW: File Name Check Status
is_count_match = False
is_resi_match = False  
# Updated: global list to store Column 1 AND Column 3 data
description_data_global = None 
pdf_path_label = None 
output_text = None    
excel_count_label = None
pdf_count_label = None
last_excel_modified_time = 0 
status_label_name = None # NEW: File Name Status Label
status_label_count = None 
status_label_resi = None  
sort_var = None 
pdf_list_display = None 
current_selected_pdf_index = -1 
open_file_var = None


# --- REDIRECT OUTPUT FUNCTION ---
class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str_to_write):
        self.widget.configure(state="normal")
        self.widget.insert(tk.END, str_to_write, self.tag)
        self.widget.see(tk.END)
        self.widget.configure(state="disabled")

    def flush(self):
        pass

# --- UTILITY FUNCTION (USED FOR AUTO-OPEN) ---
# Enhanced: Added force_open argument to ignore the open_file_var status
def open_file_in_os(path, force_open=False): 
    """Opens the file using the OS default program."""
    global open_file_var 
    
    # KEY FIX: Always open if force_open=True (for Excel),
    # OR if open_file_var is set and its value is 1 (for Output PDF)
    should_open = force_open or (open_file_var and open_file_var.get() == 1)
    
    if should_open:
        if os.name == 'nt': # Windows
            try:
                os.startfile(path)
            except Exception as e:
                print(f"Warning: Failed to open file automatically: {e}")

        elif platform.system() == 'Darwin': # macOS (Using platform.system() is more robust)
            try:
                os.system(f"open {path}")
            except Exception as e:
                print(f"Warning: Failed to open file automatically: {e}")

        else: # Linux/Unix
            try:
                os.system(f"xdg-open {path}")
            except Exception as e:
                print(f"Warning: Failed to open file automatically: {e}")


def get_excel_filename():
    excel_files = glob.glob('*.xlsx') + glob.glob('*.xls')
    return excel_files[0] if excel_files else None

def edit_excel_file():
    excel_path = get_excel_filename()
    if excel_path:
        print(f"Opening Excel file: {os.path.basename(excel_path)}")
        # FIX: Calling open_file_in_os with force_open=True
        open_file_in_os(excel_path, force_open=True) 
    else:
        messagebox.showerror("Error", "No Excel file found in the same folder.")

def update_excel_count_label(total_rows):
    if excel_count_label:
        excel_count_label.config(text=f"{total_rows}")

def update_pdf_count_label(total_pages):
    if pdf_count_label:
        pdf_count_label.config(text=f"{total_pages}")

# Function to display the selected PDF list
def update_pdf_list_display(pdf_paths, highlight_index=-1):
    global pdf_list_display
    if pdf_list_display:
        pdf_list_display.configure(state='normal')
        pdf_list_display.delete('1.0', tk.END)
        
        pdf_list_display.tag_delete('highlight')
        pdf_list_display.tag_configure('highlight', background='#cceeff')
        
        if pdf_paths:
            for i, path in enumerate(pdf_paths):
                file_name = os.path.basename(path)
                line = f"{i+1}. {file_name}\n"
                pdf_list_display.insert(tk.END, line)
                
                if i == highlight_index:
                    start_index = f"{i + 1}.0"
                    end_index = f"{i + 1}.{len(line.strip())}"
                    pdf_list_display.tag_add('highlight', start_index, end_index)

        else:
            pdf_list_display.insert(tk.END, "No PDF files selected yet.")
        pdf_list_display.configure(state='disabled')


# --- PDF ORDER MANIPULATION FUNCTIONS ---
def get_selected_pdf_index(event):
    global pdf_list_display, current_selected_pdf_index
    if not pdf_list_display:
        current_selected_pdf_index = -1
        return
        
    try:
        line_number = int(pdf_list_display.index(f"@{event.x},{event.y}").split('.')[0])
        selected_index = line_number - 1
        
        if 0 <= selected_index < len(pdf_file_path_list):
            current_selected_pdf_index = selected_index
            update_pdf_list_display(pdf_file_path_list, current_selected_pdf_index)
            print(f"Selected File: {os.path.basename(pdf_file_path_list[current_selected_pdf_index])} (Index: {current_selected_pdf_index})")
        else:
            current_selected_pdf_index = -1
            update_pdf_list_display(pdf_file_path_list, -1)
            
    except Exception:
        current_selected_pdf_index = -1
        update_pdf_list_display(pdf_file_path_list, -1)


def move_pdf_up():
    global pdf_file_path_list, current_selected_pdf_index
    
    if current_selected_pdf_index > 0 and current_selected_pdf_index != -1:
        i = current_selected_pdf_index
        pdf_file_path_list[i], pdf_file_path_list[i-1] = pdf_file_path_list[i-1], pdf_file_path_list[i]
        
        current_selected_pdf_index = i - 1
        
        update_pdf_list_display(pdf_file_path_list, current_selected_pdf_index)
        check_on_select(pdf_file_path_list, show_print=True)
        print(f"Moving file up. File order changed.")
    elif current_selected_pdf_index == 0:
        print("File is already at the top.")
    else:
        print("Please select a file from the list first.")


def move_pdf_down():
    global pdf_file_path_list, current_selected_pdf_index
    
    if current_selected_pdf_index != -1 and current_selected_pdf_index < len(pdf_file_path_list) - 1:
        i = current_selected_pdf_index
        pdf_file_path_list[i], pdf_file_path_list[i+1] = pdf_file_path_list[i+1], pdf_file_path_list[i]
        
        current_selected_pdf_index = i + 1
        
        update_pdf_list_display(pdf_file_path_list, current_selected_pdf_index)
        check_on_select(pdf_file_path_list, show_print=True)
        print(f"Moving file down. File order changed.")
    elif current_selected_pdf_index == len(pdf_file_path_list) - 1 and current_selected_pdf_index != -1:
        print("File is already at the bottom.")
    else:
        print("Please select a file from the list first.")
# --- END PDF ORDER MANIPULATION FUNCTIONS ---


# --- STATUS DISPLAY UPDATE FUNCTIONS ---
# Updated: Receives 3 statuses (Name, Count, Waybill)
def update_check_status_display(match_name, match_count, match_resi):
    global status_label_name, status_label_count, status_label_resi, pdf_file_path_list, sort_var

    if sort_var is None:
        return

    expected_name = "ELEVEN" if sort_var.get() == "Ascending" else "FAMI"

    if not pdf_file_path_list: 
        status_label_name.config(text=f"1. File Name Check ('{expected_name}') ⚪", foreground="black")
        status_label_count.config(text="2. Data Rows (Excel) and Pages (PDF) ⚪", foreground="black")
        status_label_resi.config(text="3. Waybill Number Check (Column 7) ⚪", foreground="black")
        return

    # 1. File Name Check
    if match_name:
        status_label_name.config(text=f"1. File Name Matches ('{expected_name}') ✅", foreground="green")
    else:
        status_label_name.config(text=f"1. File Name DOES NOT Match ('{expected_name}') 🚨", foreground="red")

    # 2. Row/Page Check (only if file name matches)
    if not match_name:
        status_label_count.config(text="2. (Waiting for File Name Match) 🟡", foreground="orange")
        status_label_resi.config(text="3. (Waiting for File Name Match) 🟡", foreground="orange")
        return

    if match_count:
        status_label_count.config(text="2. Data Rows and Pages match ✅", foreground="green")
    else:
        status_label_count.config(text="2. Data Rows and Pages DO NOT match 🚨", foreground="red")

    # 3. Waybill Check (only if file name and count match)
    if not match_count:
        status_label_resi.config(text="3. (Waiting for Row/Page Match) 🟡", foreground="orange")
    else:
        if sort_var.get() == "Ascending":
            if match_resi:
                status_label_resi.config(text="3. Waybill Number (Last 5 Digits) Match ✅", foreground="green")
            else:
                status_label_resi.config(text="3. Waybill Number DOES NOT Match 🚨", foreground="red")
        else:
            # Descending does not need waybill check
            status_label_resi.config(text="3. Descending Order selected (Waybill OK) ✅", foreground="green")


# --- EXCEL AUTO-REFRESH FUNCTION ---
def check_excel_modified(root):
    global last_excel_modified_time, pdf_file_path_list
    
    excel_path = get_excel_filename()
    
    if excel_path:
        try:
            current_modified_time = os.path.getmtime(excel_path)
            
            if current_modified_time > last_excel_modified_time:
                last_excel_modified_time = current_modified_time
                
                if pdf_file_path_list:
                    check_on_select(pdf_file_path_list, show_print=True)
                else:
                    df = pd.read_excel(excel_path, header=None)
                    valid_rows = df.iloc[:, 0].dropna().shape[0]
                    update_excel_count_label(valid_rows)
                    print(f"✅ Refresh: Total Excel Data Rows updated to {valid_rows}.")
            
        except Exception as e:
            print(f"Auto-Refresh Warning: Failed to read Excel file status: {e}")
            
    root.after(2000, lambda: check_excel_modified(root))


# --- WAYBILL VALIDATION FUNCTION (Ascending Mode) ---
def validate_resi_number(pdf_paths, excel_path):
    print("🔬 Waybill Validation (Ascending Mode)...")
    
    try:
        df = pd.read_excel(excel_path, header=None) 
        # Get data from column 7 (index 6)
        resi_excel_list = [str(item).strip() for item in df.iloc[:, 6].dropna().tolist() if pd.notna(item)]
        
        if not resi_excel_list:
            print("Warning: Column 7 (Waybill) in Excel is empty. Proceeding without waybill validation.")
            return True 

        mismatches = []
        excel_idx = 0

        for pdf_path in pdf_paths:
            try:
                # Using PdfReader as pdfplumber is not needed here
                pdf_reader = PdfReader(pdf_path)
                pdf_page_count = len(pdf_reader.pages)
                
                for i in range(pdf_page_count):
                    # extract_resi_number_from_pdf function uses pdfplumber
                    full_waybill_pdf = extract_resi_number_from_pdf(pdf_path, i)
                    
                    clean_waybill_pdf = re.sub(r'\D', '', full_waybill_pdf)
                    waybill_5_digit_pdf = clean_waybill_pdf[-5:] if len(clean_waybill_pdf) >= 5 else clean_waybill_pdf
                    
                    if excel_idx < len(resi_excel_list):
                        resi_excel_value = resi_excel_list[excel_idx]
                        clean_waybill_excel = re.sub(r'\D', '', resi_excel_value)
                        waybill_5_digit_excel = clean_waybill_excel[-5:] if len(clean_waybill_excel) >= 5 else clean_waybill_excel

                        if waybill_5_digit_pdf != waybill_5_digit_excel:
                            mismatches.append({
                                "File PDF": os.path.basename(pdf_path),
                                "Halaman PDF": i + 1,
                                "Baris Excel": excel_idx + 1,
                                "Resi PDF (5 Digit Terakhir)": waybill_5_digit_pdf,
                                "Resi Excel (Kolom 7)": resi_excel_value 
                            })
                        excel_idx += 1
                    if excel_idx >= len(resi_excel_list):
                        break

            except Exception as e:
                print(f"ERROR: Failed to process file {os.path.basename(pdf_path)} during waybill validation: {e}")
                continue 
        
        if mismatches:
            print("🚨 FAILED: Mismatches found in the last 5 Digits of the waybill number:")
            for m in mismatches:
                print(f"-> File {m['File PDF']}, Page {m['Halaman PDF']} (Excel Row {m['Baris Excel']}): PDF Last 5 Digits '{m['Resi PDF (5 Digit Terakhir)']}' DO NOT MATCH Excel '{m['Resi Excel (Kolom 7)']}'")
            return False

        print("[STATUS] SUCCESS: Last 5 Digits of waybill in PDF match Excel Column 7.")
        return True

    except Exception as e:
        print(f"ERROR: An error occurred during waybill validation: {e}")
        messagebox.showerror("Validation Error", f"An error occurred while comparing waybill data: {e}. Ensure Column 7 contains the correct data.")
        return False


# --- WAYBILL EXTRACTION FUNCTION ---
def extract_resi_number_from_pdf(pdf_path, page_num):
    target_string = "交貨便服務代碼"
    
    # This function requires the pdfplumber library
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_num]
            text = page.extract_text()
            
            if text is None:
                return "ERROR: Text Could Not Be Extracted"
                
            cleaned_text = " ".join(text.split())

            if target_string in cleaned_text:
                specific_pattern = re.compile(
                    r'交貨便服務代碼\s*[:：]\s*([\w\d\-]+)', 
                    re.IGNORECASE
                )
                match = specific_pattern.search(cleaned_text)
                
                if match:
                    return match.group(1).strip()
                else:
                    return f"DEBUG: '{target_string}' Found, BUT REGEX PATTERN FAILED" 
            
            # --- FALLBACK PATTERNS ---
            long_number_pattern = re.compile(r'(\d{4}\s\d{4}\s\d{4}\s\d{4})')
            match = long_number_pattern.search(text)
            if match:
                return match.group(1).replace(' ', '')
                
            short_id_pattern = re.compile(r'(DRE\s*\d{3}\s*\d{4})')
            match = short_id_pattern.search(text)
            if match:
                return match.group(1).replace(' ', '')
                    
            return f"DEBUG: '{target_string}' NOT FOUND"
            
    except Exception as e:
        print(f"ERROR: Failed to Read/Process PDF on page {page_num + 1}: {e}")
        return f"ERROR: Failed to Read PDF"

# --- INITIAL BALANCE CHECK FUNCTION (Loading Column 3 Data) ---
def check_on_select(pdf_paths, show_print=True):
    global is_name_match, is_count_match, is_resi_match, description_data_global, sort_var
    
    # Reset count and waybill status, but is_name_match will be recalculated
    is_count_match = False
    is_resi_match = True 
    description_data_global = None
    excel_path = get_excel_filename()

    # PHASE 0: File Name Check (recalculated here to support arrow buttons & auto-refresh)
    current_sort = sort_var.get()
    expected_name = "ELEVEN" if current_sort == "Ascending" else "FAMI"
    is_name_valid_current = True

    if pdf_paths:
        for path in pdf_paths:
            if expected_name not in os.path.basename(path).upper():
                is_name_valid_current = False
                break
    else:
        is_name_valid_current = False
    
    is_name_match = is_name_valid_current # Update global status
    
    if not is_name_match or not excel_path:
        is_count_match = False
        is_resi_match = False
        update_check_status_display(is_name_match, is_count_match, is_resi_match)
        return False
        
    # If file name matches, proceed to count check
    try:
        # PHASE 1: READ ONLY COLUMN A (Item Description)
        df_description = pd.read_excel(excel_path, header=None, usecols=[0])
        
        # Clean and prepare data (Column 1 - Item Description)
        description_series = df_description.iloc[:, 0].dropna() 
        description_list = [str(item).replace('\n', ' ') for item in description_series.tolist()]
        
        total_excel_descriptions = len(description_list)
        
        # PHASE 2: OPTIONALLY READ COLUMN C (Column 3)
        column3_list = [''] * total_excel_descriptions 
        
        try:
            df_column3 = pd.read_excel(excel_path, header=None, usecols=[2])
            
            column3_series = df_column3.iloc[:total_excel_descriptions, 0].fillna('')
            column3_list = [str(item).replace('\n', ' ') for item in column3_series.tolist()]

            if show_print:
                print("✅ Column 3 (Slip Data) found and loaded.")
                
        except ValueError as ve:
            if "out-of-bounds indices" in str(ve) or "does not exist" in str(ve):
                if show_print:
                    print("⚠️ Warning: Column 3 (Notes) not found in Excel. Program will ignore it.")
            else:
                raise ve
        
        description_data_global = list(zip(description_list, column3_list))

        # 3. Calculate Total PDF Pages from ALL files
        total_pdf_pages = 0
        for pdf_path in pdf_paths:
            try:
                pdf_reader = PdfReader(pdf_path)
                total_pdf_pages += len(pdf_reader.pages)
            except Exception as e:
                print(f"Warning: Failed to read PDF file {os.path.basename(pdf_path)}: {e}")
                
        
        update_excel_count_label(total_excel_descriptions)
        update_pdf_count_label(total_pdf_pages)

        # 4. Balance Check Logic 
        if show_print:
            print("\n" + "=" * 50)
            print(f"✅ Data Balance Check...")
            print(f"-> Total Data Rows (Column 1) in Excel: {total_excel_descriptions}")
            print(f"-> Total Pages in PDF: {total_pdf_pages}")

        if total_excel_descriptions == total_pdf_pages:
            is_count_match = True
            if show_print:
                print(f"[STATUS] SUCCESS: Number of data rows and pages MATCH EXACTLY.")

            # Waybill Validation Logic (Column B/7). Always active if Ascending.
            if sort_var.get() == "Ascending":
                is_resi_match = validate_resi_number(pdf_paths, excel_path)
            else:
                is_resi_match = True # Descending does not need waybill check
                if show_print:
                    print("[STATUS] Descending Order selected. Waybill Number Check IS SKIPPED.")

        else:
            is_count_match = False
            is_resi_match = False
            error_message = (
                f"🚨 WARNING! COUNTS DO NOT MATCH!\n"
                f"Excel Data Rows: {total_excel_descriptions}\n"
                f"PDF Pages: {total_pdf_pages}"
            )
            if show_print:
                print(f"[STATUS] FAILED: {error_message}".replace('\n', ' '))
                
        if show_print:
            print("=" * 50)
            
        # Call update display with 3 statuses
        update_check_status_display(is_name_match, is_count_match, is_resi_match)
        return is_count_match 

    except Exception as e:
        update_excel_count_label(0)
        update_pdf_count_label(0)
        is_count_match = False
        is_resi_match = False
        update_check_status_display(is_name_match, is_count_match, is_resi_match)
        if show_print:
            if "out-of-bounds indices" in str(e):
                 print(f"\n[STATUS] FATAL ERROR: Unexpected Excel index internal error: {e}")
                 messagebox.showerror("Check Error", "Unexpected Excel index internal error.")
            else:
                 print(f"\n[STATUS] ERROR: An error occurred during file check: {e}")
                 messagebox.showerror("Check Error", f"An error occurred while loading data: {e}")
        return False


# Function to select PDF files (REVERSE ORDER)
def choose_pdf_file():
    global pdf_file_path_list, current_selected_pdf_index, sort_var, is_name_match
    
    filepaths = filedialog.askopenfilenames(
        title="Select Waybill File(s) (PDF) - Multiple Selection Allowed",
        filetypes=[("PDF files", "*.pdf")]
    )
    
    if filepaths:
        current_sort = sort_var.get()
        expected_name = ""
        
        # Determine the expected keyword
        if current_sort == "Ascending":
            expected_name = "ELEVEN"
        elif current_sort == "Descending":
            expected_name = "FAMI"
        
        # Perform file name validation
        valid_filepaths = []
        is_valid = True
        for path in filepaths:
            file_name = os.path.basename(path).upper()
            if expected_name in file_name:
                valid_filepaths.append(path)
            else:
                is_valid = False
                messagebox.showerror("Input Error", f"File '{os.path.basename(path)}' is invalid. Because '{current_sort}' mode is selected, the file name must contain the keyword '{expected_name}'.")
                print(f"🚨 REJECTED: File '{os.path.basename(path)}' does not contain the keyword '{expected_name}' for '{current_sort}' mode.")
                break # Cancel all if 1 file fails validation

        if is_valid:
            is_name_match = True
            pdf_file_path_list = list(filepaths)
            
            # REVERSE THE FILE ORDER (Solution for default reversed order)
            pdf_file_path_list.reverse() 
            print(f"Reversing the imported file order. First file to be processed: {os.path.basename(pdf_file_path_list[0])}")
            
            pdf_path_label.config(text=f"{len(pdf_file_path_list)} Files Selected")
            current_selected_pdf_index = 0 
            update_pdf_list_display(pdf_file_path_list, current_selected_pdf_index)
            
            check_on_select(pdf_file_path_list, show_print=True)
        else:
            # If validation failed or break was triggered
            is_name_match = False
            pdf_file_path_list = [] 
            pdf_path_label.config(text="File Validation Failed")
            current_selected_pdf_index = -1
            update_pdf_list_display([])
            update_excel_count_label(0)
            update_pdf_count_label(0)
            # Call update display with 3 statuses (Name, Count, Waybill)
            update_check_status_display(is_name_match, False, False) 

    else:
        # If user cancels the file dialog
        is_name_match = False
        pdf_file_path_list = [] 
        pdf_path_label.config(text="No files selected")
        current_selected_pdf_index = -1
        update_pdf_list_display([])
        update_excel_count_label(0)
        update_pdf_count_label(0)
        # Call update display with 3 statuses
        update_check_status_display(False, False, False) 


def change_sort_order(event=None):
    
    if pdf_file_path_list:
        # Call check_on_select to re-trigger the check (including file name)
        check_on_select(pdf_file_path_list, show_print=True)
    else:
        # Call display update so file name and waybill status refresh
        update_check_status_display(is_name_match, is_count_match, is_resi_match)


# Function to start the main process
def start_process(sort_order):
    global pdf_file_path_list, is_count_match, is_resi_match, is_name_match

    output_text.delete(1.0, tk.END)

    if not pdf_file_path_list:
        messagebox.showerror("Error", "Please select at least one PDF file first.")
        return
    
    # Perform final check before starting the process
    if not check_on_select(pdf_file_path_list, show_print=True):
        messagebox.showerror("Error", "Balance Check FAILED. Check Program Output for details.")
        return

    # Additional Validation 1: Check File Name
    if not is_name_match:
        messagebox.showerror("Error", "File Name Check FAILED. Please ensure all PDF files contain the keyword corresponding to the selected sort order.")
        return

    # Additional Validation 2: Check Waybill (always active if Ascending)
    if sort_order == "Ascending" and not is_resi_match:
        messagebox.showerror("Error", "Waybill Number Check FAILED. Please check Column 7 of your Excel file.")
        return
        
    # --- OUTPUT FILE NAMING LOGIC ---
    first_file_path = pdf_file_path_list[0]
    first_file_name_base = os.path.basename(first_file_path)
    base_name, ext = os.path.splitext(first_file_name_base)
    
    if len(pdf_file_path_list) == 1:
        default_save_name = f"{base_name}{ext}" # Changed for single file consistency
    else:
        default_save_name = f"{base_name}_FULL{ext}" # Changed for multi-file consistency
    
    print(f"Default output name set to: {default_save_name}")

    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        initialdir=os.path.dirname(first_file_path) or os.getcwd(),
        initialfile=default_save_name,
        filetypes=[("PDF files", "*.pdf")]
    )

    if not save_path:
        print("Operation cancelled by user.")
        return

    process_pdf_and_excel(sort_order, pdf_file_path_list, save_path)


# --- MAIN PROCESSING FUNCTION (Adding Custom Text) ---
def process_pdf_and_excel(sort_order, pdf_input_paths, pdf_output_path):
    global description_data_global, is_count_match, is_resi_match, is_name_match

    # Re-check before starting
    if not is_name_match:
        messagebox.showerror("Error", "File Name Check Failed.")
        return
    if not is_count_match:
        messagebox.showerror("Error", "Balance Check Failed.")
        return
    
    # Waybill Validation (Ascending Mode)
    if sort_order == "Ascending" and not is_resi_match:
        messagebox.showerror("Error", "Waybill Check Failed.")
        return


    description_data = description_data_global 

    print("Executing program...")
    print("=" * 50)

    try:
        if sort_order == "Descending":
            description_data.reverse()
            print("Description order is Bottom to Top (Descending).")
        else:
            print("Description order is Top to Bottom (Ascending).")
        
        pdf_writer = PdfWriter()
        description_index = 0

        # --- CUSTOM FONT SETTINGS ---
        FONT_NORMAL = "Helvetica-Bold"
        FONT_SIZE_NORMAL = 9  # Default size 
        FONT_SIZE_NUMBER = 18 # Specific size for pure numbers
        MAX_LINE_WIDTH = 300
        LINE_SPACING = FONT_SIZE_NORMAL + 1 # Spacing between lines
        # -----------------------------

        for pdf_input_path in pdf_input_paths:
            print(f"Processing file: {os.path.basename(pdf_input_path)}...")

            try:
                pdf_reader = PdfReader(pdf_input_path)
            except Exception as e:
                print(f"Warning: Failed to load file {os.path.basename(pdf_input_path)}. Skipping this file. Error: {e}")
                continue

            if 'page_height' not in locals():
                page_height = float(pdf_reader.pages[0].mediabox.height)

            for i, page in enumerate(pdf_reader.pages):
                if description_index >= len(description_data):
                    print(f"Warning: Excel descriptions ran out on Page {i+1} of {os.path.basename(pdf_input_path)}. Stopping PDF page processing.")
                    break

                # Get data from global list: (item_description, column_3_data)
                original_description, original_column3_data = description_data[description_index]
                description_index += 1
                
                # --- SPECIAL TEXT ADDITION LOGIC (INSERT 60 NT) ---
                final_description = original_description
                # Check if the number '60' is in original_column3_data (using string conversion)
                if '60' in str(original_column3_data):
                    special_text = "(INSERT 60 NT !!) "
                    print(f" -> '60' detected in Column 3 (Page {i+1}): Adding '{special_text.strip()}'")
                    final_description = special_text + original_description
                # -----------------------------------------------------------

                
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=A4)

                x_pos_center = 258
                y_pos_from_top = 190
                y_pos_from_bottom = page_height - y_pos_from_top

                # --- 1. WORD WRAPPING LOGIC ---
                # Use FONT_SIZE_NORMAL for line width calculation
                can.setFont(FONT_NORMAL, FONT_SIZE_NORMAL)
                
                # Use final_description
                words = final_description.split(' ')
                lines = []
                current_line = ""
                for word in words:
                    # Use a space for accurate width calculation
                    test_line = current_line + " " + word if current_line else word
                    
                    # Check width using FONT_SIZE_NORMAL
                    if can.stringWidth(test_line, FONT_NORMAL, FONT_SIZE_NORMAL) < MAX_LINE_WIDTH:
                        current_line = test_line
                    else:
                        if current_line:
                            lines.append(current_line.strip())
                        current_line = word
                lines.append(current_line.strip())
                
                # --- 2. INITIAL POSITION DETERMINATION ---
                
                # Estimated block height (using LINE_SPACING)
                text_block_height = len(lines) * LINE_SPACING 
                initial_y_pos = y_pos_from_bottom + (text_block_height / 2)

                can.saveState()
                can.translate(x_pos_center, initial_y_pos)
                can.rotate(90)
                
                # --- 3. PRINTING LOGIC WITH DIFFERENT FONT SIZES & VERTICAL OFFSET ---
                
                # Vertical Offset Calculation to align large numbers in the center
                FONT_SIZE_DIFF = FONT_SIZE_NUMBER - FONT_SIZE_NORMAL
                Y_ADJUSTMENT = FONT_SIZE_DIFF / 2.5 # 2.5 is an empirical adjustment value that works well

                for j, line in enumerate(lines):
                    current_x = 0
                    
                    # Tokenize the line: (text, number, (number))
                    # Regex: (\d+)|(\(.*?\))|([^\s\d()]+) -> Numbers, Parentheses, Text
                    tokens = re.findall(r'(\d+)|(\(.*?\))|([^\s\d()]+)', line)
                    
                    # --- Width Calculation for Centering ---
                    total_width = 0
                    word_segments = []

                    # Phase 1: Tokenization and Width Calculation
                    for token_group in tokens:
                        for segment in token_group:
                            if segment: # Ensure the segment is not empty
                                is_number_only = segment.isdigit()
                                is_in_parentheses = segment.startswith('(') and segment.endswith(')')
                                
                                calc_font_size = FONT_SIZE_NUMBER if is_number_only and not is_in_parentheses else FONT_SIZE_NORMAL
                                
                                word_segments.append({'text': segment, 'size': calc_font_size})
                                
                                total_width += can.stringWidth(segment, FONT_NORMAL, calc_font_size)
                                total_width += can.stringWidth(' ', FONT_NORMAL, FONT_SIZE_NORMAL) 
                    
                    if total_width > 0:
                        # Subtract the width of the last space
                        total_width -= can.stringWidth(' ', FONT_NORMAL, FONT_SIZE_NORMAL) 

                    start_x = -total_width / 2
                    # --- End Width Calculation ---

                    # Phase 2: Printing Based on Segments
                    for segment_data in word_segments:
                        word = segment_data['text']
                        font_size = segment_data['size']
                        
                        can.setFont(FONT_NORMAL, font_size)

                        word_width = can.stringWidth(word, FONT_NORMAL, font_size)
                        
                        # --- KEY: Y Position Adjustment (Vertical Offset) ---
                        # Normal text is printed at the baseline (0). Large text is shifted down.
                        y_pos_in_rotation = j * -LINE_SPACING
                        
                        if font_size == FONT_SIZE_NUMBER:
                            # Apply negative Y offset to shift down
                            final_y_pos = y_pos_in_rotation - Y_ADJUSTMENT
                        else:
                            final_y_pos = y_pos_in_rotation
                        
                        # Print the word/segment with adjusted font size and Y position
                        can.drawString(start_x + current_x, final_y_pos, word)
                        
                        # Move the x position for the next word, add space after the word
                        current_x += word_width + can.stringWidth(' ', FONT_NORMAL, FONT_SIZE_NORMAL)
                
                can.restoreState()
                can.save()
                # --- END REPORTLAB LOGIC ---

                packet.seek(0)
                new_pdf = PdfReader(packet)
                page.merge_page(new_pdf.pages[0])
                pdf_writer.add_page(page)
            
            if description_index >= len(description_data):
                break

        
        with open(pdf_output_path, 'wb') as f:
            pdf_writer.write(f)

        print(f"\nOperation complete. Result file saved as '{os.path.basename(pdf_output_path)}'.")
        
        # --- AUTO-OPEN FILE (ACTIVE) ---
        open_file_in_os(pdf_output_path)
        # ---------------------------

        messagebox.showinfo("Complete", f"Operation complete. Result file saved as '{os.path.basename(pdf_output_path)}' and has been opened.")

    except Exception as e:
        print(f"\nAn error occurred: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")


# --- GUI SETUP ---
def create_ui():
    global pdf_path_label, output_text, excel_count_label, pdf_count_label, last_excel_modified_time, status_label_name, status_label_count, status_label_resi, sort_var, pdf_list_display, open_file_var
    
    root = tk.Tk()
    root.title("ResiText - Waybill Description Input Tool")
    
    APP_BG = '#e6f0ff'
    STATUS_BOX_BG = 'white' 
    STATUS_BOX_RELIEF = 'sunken' 
    STATUS_FRAME_BG = "#d1eef3" 

    root.configure(bg=APP_BG)

    style = ttk.Style()
    style.configure('TFrame', background=APP_BG)
    style.configure('TLabel', background=APP_BG, font=('Helvetica', 10, 'normal'))
    style.configure('TButton', font=('Helvetica', 9, 'normal'), padding=5)
    style.configure('TRadiobutton', background=APP_BG, font=('Helvetica', 10, 'normal'), padding=5)
    style.configure('TCheckbutton', background=APP_BG, font=('Helvetica', 10, 'normal'))
    style.configure('Step.TLabel', background=APP_BG, font=('Helvetica', 12, 'bold'))
    style.configure('Header.TLabel', background=APP_BG, font=('Helvetica', 16, 'bold'))
    style.configure('Start.TButton', font=('Helvetica', 12, 'bold'), background='#00cc66')
    style.map('Start.TButton', background=[('active', '#00b359')])

    main_frame = ttk.Frame(root, padding=20)
    main_frame.pack(fill=tk.BOTH, expand=True)

    header_label = ttk.Label(main_frame, text="Add Text to PDF", style='Header.TLabel')
    header_label.pack(pady=(0, 20))

    top_grid_frame = ttk.Frame(main_frame)
    top_grid_frame.pack(fill=tk.X, pady=(0, 20))
    
    top_grid_frame.columnconfigure(0, weight=1) 
    top_grid_frame.columnconfigure(2, weight=1) 
    top_grid_frame.columnconfigure(4, weight=1) 

    # --- STEP 1: SORT ORDER TYPE (Column 0) ---
    sort_var = tk.StringVar(value="Ascending") 
    
    step1_frame = ttk.Frame(top_grid_frame)
    step1_frame.grid(row=0, column=0, padx=10, sticky='nwes')
    ttk.Label(step1_frame, text="Step 1: Sort Order Type", style='Step.TLabel').pack(pady=(0, 10), anchor='center') 
    
    radio_frame1 = ttk.Frame(step1_frame)
    radio_frame1.pack(anchor='w', pady=(5, 10), padx=20) 

    # Radio Button for Ascending
    asc_radio = ttk.Radiobutton(radio_frame1, text="Ascending (7-Eleven)", variable=sort_var, value="Ascending", command=change_sort_order)
    asc_radio.pack(anchor='w', pady=(0, 5)) 
    
    # Radio Button for Descending
    desc_radio = ttk.Radiobutton(radio_frame1, text="Descending (Family-Mart)", variable=sort_var, value="Descending", command=change_sort_order)
    desc_radio.pack(anchor='w', pady=(0, 10)) 
    
    ttk.Separator(top_grid_frame, orient='vertical').grid(row=0, column=1, sticky='ns', padx=10)


    # --- STEP 2: EXCEL DATA (Column 2) ---
    step2_frame = ttk.Frame(top_grid_frame)
    step2_frame.grid(row=0, column=2, padx=10, sticky='nwes')
    ttk.Label(step2_frame, text="Step 2: Excel Data", style='Step.TLabel').pack(pady=(0, 10), anchor='center')
    
    excel_content_frame = ttk.Frame(step2_frame)
    excel_content_frame.pack(anchor='center')
    
    excel_name = os.path.basename(get_excel_filename()) if get_excel_filename() else "x.xlsx (Not Found)"
    ttk.Label(excel_content_frame, text=f"Excel File Found: {excel_name}").pack(pady=(0, 5), anchor='w')
    
    edit_excel_label = ttk.Label(excel_content_frame, text="Edit Excel file", foreground="#0000ff", cursor="hand2")
    edit_excel_label.pack(pady=(5, 10), anchor='w')
    edit_excel_label.bind("<Button-1>", lambda e: edit_excel_file())

    ttk.Label(excel_content_frame, text="Total Data Rows:").pack(pady=(10, 2), anchor='w')
    excel_count_box = tk.Frame(excel_content_frame, bg=STATUS_BOX_BG, relief=STATUS_BOX_RELIEF, borderwidth=1, width=150, height=30)
    excel_count_box.pack(anchor='w')
    excel_count_box.pack_propagate(False) 
    excel_count_label = tk.Label(excel_count_box, text="0", 
                                 fg='black', bg=STATUS_BOX_BG, font=('Helvetica', 12, 'normal'), 
                                 anchor='center', justify=tk.CENTER)
    excel_count_label.pack(expand=True, fill='both') 

    ttk.Separator(top_grid_frame, orient='vertical').grid(row=0, column=3, sticky='ns', padx=10)


    # --- STEP 3: SELECT WAYBILL PDF (Column 4) ---
    step3_frame = ttk.Frame(top_grid_frame)
    step3_frame.grid(row=0, column=4, padx=10, sticky='nwes')
    
    ttk.Label(step3_frame, text="Step 3: Select Waybill PDF(s) (Multi-file)", style='Step.TLabel').pack(pady=(0, 10), anchor='center')
    
    pdf_controls_frame = ttk.Frame(step3_frame)
    pdf_controls_frame.pack(anchor='center', fill='x', padx=10)

    pdf_path_label = ttk.Label(pdf_controls_frame, text="No files selected")
    pdf_path_label.pack(pady=(5, 5), anchor='center') 
    
    pdf_input_frame = ttk.Frame(pdf_controls_frame)
    pdf_input_frame.pack(pady=5, anchor='center') 
    
    ttk.Button(pdf_input_frame, text="Add PDF(s)", command=choose_pdf_file, style='TButton').pack(side=tk.LEFT)
    
    ttk.Label(pdf_controls_frame, text="Selected File List:").pack(pady=(10, 2), anchor='w')
    
    list_and_control_frame = ttk.Frame(pdf_controls_frame)
    list_and_control_frame.pack(fill=tk.X, expand=True)

    pdf_list_display = ScrolledText(list_and_control_frame, height=5, width=40, state='disabled', wrap=tk.WORD, relief='sunken', borderwidth=1)
    pdf_list_display.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    pdf_list_display.bind('<Button-1>', get_selected_pdf_index) 

    order_buttons_frame = ttk.Frame(list_and_control_frame)
    order_buttons_frame.pack(side=tk.RIGHT, padx=(5, 0))

    ttk.Button(order_buttons_frame, text="▲ Up", command=move_pdf_up).pack(fill=tk.X, pady=2)
    ttk.Button(order_buttons_frame, text="▼ Down", command=move_pdf_down).pack(fill=tk.X, pady=2)
    
    
    # Frame to hold Total Pages label and Open File Checkbox
    bottom_pdf_frame = ttk.Frame(pdf_controls_frame)
    bottom_pdf_frame.pack(fill='x', pady=(10, 5))
    
    
    # Frame for Total PDF Pages (Top Row)
    pdf_count_wrapper_frame = ttk.Frame(bottom_pdf_frame)
    pdf_count_wrapper_frame.pack(anchor='w', pady=(0, 5)) 
    
    # 'Total PDF Pages:' Label (on the left)
    ttk.Label(pdf_count_wrapper_frame, text="Total PDF Pages:").pack(side=tk.LEFT, padx=(0, 5))
    
    # TOTAL PDF PAGES BOX
    pdf_count_box = tk.Frame(pdf_count_wrapper_frame, bg=STATUS_BOX_BG, relief=STATUS_BOX_RELIEF, borderwidth=1, width=150, height=30) 
    
    pdf_count_box.pack(side=tk.LEFT, anchor='w') 
    pdf_count_box.pack_propagate(False) 
    pdf_count_label = tk.Label(pdf_count_box, text="0", 
                                 fg='black', bg=STATUS_BOX_BG, font=('Helvetica', 12, 'normal'), 
                                 anchor='center', justify=tk.CENTER) 
    pdf_count_label.pack(expand=True, fill='both')
    
    
    # Open File Automatically Checkbox (Bottom Row)
    open_file_var = tk.IntVar(value=0) 
    
    open_file_checkbox = ttk.Checkbutton(bottom_pdf_frame, 
                                     text="Open file after creation", 
                                     variable=open_file_var,
                                     style='TCheckbutton')
    open_file_checkbox.pack(anchor='w', padx=5)
    
    
    # --- SMALL STATUS COMMAND ABOVE START (3 Status Check) ---
    step_status_frame = tk.Frame(main_frame, bg=STATUS_FRAME_BG, relief='groove', borderwidth=1, padx=10, pady=5)
    step_status_frame.pack(pady=(10, 15), fill=tk.X)
    
    ttk.Label(step_status_frame, text="Status Check:", font=('Helvetica', 10, 'bold'), background=STATUS_FRAME_BG).pack(anchor='w')
    
    # 1. File Name Check
    status_label_name = ttk.Label(step_status_frame, text="1. File Name Check ⚪", background=STATUS_FRAME_BG)
    status_label_name.pack(anchor='w')
    
    # 2. Count Check
    status_label_count = ttk.Label(step_status_frame, text="2. Data Rows (Excel) and Pages (PDF) ⚪", background=STATUS_FRAME_BG)
    status_label_count.pack(anchor='w')
    
    # 3. Waybill Check
    status_label_resi = ttk.Label(step_status_frame, text="3. Waybill Number Check (Column 7) ⚪", background=STATUS_FRAME_BG)
    status_label_resi.pack(anchor='w')
    # -------------------------------------------------

    start_button = ttk.Button(main_frame, text="Start", command=lambda: start_process(sort_var.get()), style='Start.TButton')
    start_button.pack(pady=(10, 20), ipadx=50)

    ttk.Label(main_frame, text="Program Output :").pack(pady=(0, 5), anchor='w')
    output_text = ScrolledText(main_frame, height=10, width=70, state='disabled', relief='sunken', borderwidth=2)
    output_text.pack(fill=tk.BOTH, expand=True)
    
    sys.stdout = TextRedirector(output_text, "stdout")
    sys.stderr = TextRedirector(output_text, "stderr")

    excel_path_init = get_excel_filename()
    if excel_path_init:
        try:
            df_init = pd.read_excel(excel_path_init, header=None)
            valid_rows = df_init.iloc[:, 0].dropna().shape[0]
            update_excel_count_label(valid_rows)
            last_excel_modified_time = os.path.getmtime(excel_path_init)
        except Exception as e:
            print(f"Warning: Failed to load/count Excel at startup: {e}")
            update_excel_count_label(0)
    
    
    # Initialize status with 3 False values
    update_check_status_display(False, False, False)
    update_pdf_list_display([]) 
    
    check_excel_modified(root) 
    
    credit_label = ttk.Label(root, text="© didk_", font=('Helvetica', 8, 'italic'), background=APP_BG, foreground='#666666')
    credit_label.place(relx=0.0, rely=1.0, anchor='sw', x=10, y=0)
    
    root.mainloop()

# Run UI
if __name__ == '__main__':
    create_ui()