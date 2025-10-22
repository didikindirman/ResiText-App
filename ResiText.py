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

# --- VARIABEL GLOBAL GUI & STATUS ---
pdf_file_path_list = [] 
is_count_match = False
is_resi_match = False  
# Diperbarui: global list untuk menyimpan data Kolom 1 DAN Kolom 3
keterangan_data_global = None 
pdf_path_label = None 
output_text = None    
excel_count_label = None
pdf_count_label = None
last_excel_modified_time = 0 
status_label_count = None 
status_label_resi = None  
sort_var = None 
pdf_list_display = None 
current_selected_pdf_index = -1 
check_resi_var = None 
# BARU: Global referensi untuk widget Checkbutton
resi_check_box_widget = None 


# --- FUNGSI REDIRECT OUTPUT ---
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

# --- FUNGSI UTILITY (DIGUNAKAN UNTUK AUTO-OPEN) ---
def open_file_in_os(file_path):
    """Membuka file di sistem operasi default."""
    try:
        if platform.system() == 'Darwin':       # macOS
            subprocess.call(('open', file_path))
        elif platform.system() == 'Windows':    # Windows
            os.startfile(file_path)
        else:                                   # Linux
            subprocess.call(('xdg-open', file_path))
    except FileNotFoundError:
        messagebox.showerror("Kesalahan", "Aplikasi default untuk membuka file tidak ditemukan.")
    except Exception as e:
        messagebox.showerror("Kesalahan", f"Terjadi kesalahan saat mencoba membuka file: {e}")

def get_excel_filename():
    excel_files = glob.glob('*.xlsx') + glob.glob('*.xls')
    return excel_files[0] if excel_files else None

def edit_excel_file():
    excel_path = get_excel_filename()
    if excel_path:
        print(f"Membuka file Excel: {os.path.basename(excel_path)}")
        open_file_in_os(excel_path)
    else:
        messagebox.showerror("Kesalahan", "Tidak ada file Excel ditemukan di folder yang sama.")

def update_excel_count_label(total_rows):
    if excel_count_label:
        excel_count_label.config(text=f"{total_rows}")

def update_pdf_count_label(total_pages):
    if pdf_count_label:
        pdf_count_label.config(text=f"{total_pages}")

# Fungsi untuk menampilkan daftar PDF yang dipilih
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
            pdf_list_display.insert(tk.END, "Belum ada file PDF yang dipilih.")
        pdf_list_display.configure(state='disabled')


# --- FUNGSI MANIPULASI URUTAN PDF ---
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
            print(f"File terpilih: {os.path.basename(pdf_file_path_list[current_selected_pdf_index])} (Index: {current_selected_pdf_index})")
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
        print(f"Menggeser file ke atas. Urutan file diubah.")
    elif current_selected_pdf_index == 0:
        print("File sudah berada di urutan paling atas.")
    else:
        print("Harap pilih file dari daftar terlebih dahulu.")


def move_pdf_down():
    global pdf_file_path_list, current_selected_pdf_index
    
    if current_selected_pdf_index != -1 and current_selected_pdf_index < len(pdf_file_path_list) - 1:
        i = current_selected_pdf_index
        pdf_file_path_list[i], pdf_file_path_list[i+1] = pdf_file_path_list[i+1], pdf_file_path_list[i]
        
        current_selected_pdf_index = i + 1
        
        update_pdf_list_display(pdf_file_path_list, current_selected_pdf_index)
        check_on_select(pdf_file_path_list, show_print=True)
        print(f"Menggeser file ke bawah. Urutan file diubah.")
    elif current_selected_pdf_index == len(pdf_file_path_list) - 1 and current_selected_pdf_index != -1:
        print("File sudah berada di urutan paling bawah.")
    else:
        print("Harap pilih file dari daftar terlebih dahulu.")
# --- AKHIR FUNGSI MANIPULASI URUTAN PDF ---

# BARU: FUNGSI UNTUK MENGAKTIFKAN/MENONAKTIFKAN CHECKBOX
def toggle_resi_checkbox():
    global sort_var, resi_check_box_widget, check_resi_var
    
    if resi_check_box_widget is None:
        return
        
    current_sort = sort_var.get()
    
    if current_sort == "Descending":
        # Nonaktifkan Checkbox, dan atur nilainya menjadi 0 (tidak dicentang)
        resi_check_box_widget.config(state=tk.DISABLED)
        check_resi_var.set(0) 
        print("Urutan Descending (Family-Mart) dipilih.")
    else:
        # Aktifkan Checkbox untuk Ascending
        resi_check_box_widget.config(state=tk.NORMAL)
        print("Urutan Ascending (7-Eleven) dipilih. Pengecekan Nomor Resi Diaktifkan.")

# --- FUNGSI UPDATE TAMPILAN STATUS ---
def update_check_status_display(match_count, match_resi):
    global status_label_count, status_label_resi, pdf_file_path_list, sort_var, check_resi_var

    if sort_var is None or check_resi_var is None:
        return

    if not pdf_file_path_list: 
        status_label_count.config(text="1. Baris Data (Excel) dan Halaman (PDF) ⚪", foreground="black")
        status_label_resi.config(text="2. Pengecekan Nomor Resi (5 Digit Terakhir) ⚪", foreground="black")
        return

    if match_count:
        status_label_count.config(text="1. Baris Data dan Halaman sama ✅", foreground="green")
    else:
        status_label_count.config(text="1. Baris Data dan Halaman TIDAK sama 🚨", foreground="red")

    if match_count: 
        is_resi_check_enabled = check_resi_var.get() == 1
        
        if sort_var.get() == "Ascending":
            if is_resi_check_enabled:
                if match_resi:
                    status_label_resi.config(text="2. Nomor Resi (5 Digit Terakhir) Sesuai ✅", foreground="green")
                else:
                    status_label_resi.config(text="2. Nomor Resi TIDAK Sesuai 🚨", foreground="red")
            else:
                status_label_resi.config(text="2. Pengecekan Nomor Resi Diabaikan 🟡", foreground="orange")
        else:
            status_label_resi.config(text="2. Urutan Descending dipilih (Resi OK) ✅", foreground="green")
    else:
        status_label_resi.config(text="2. (Menunggu Baris/Halaman Sama) 🟡", foreground="orange")


# --- FUNGSI AUTO-REFRESH EXCEL ---
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
                    print(f"✅ Refresh: Total Baris Data Excel diperbarui menjadi {valid_rows}.")
            
        except Exception as e:
            print(f"Peringatan Auto-Refresh: Gagal membaca status file Excel: {e}")
            
    root.after(2000, lambda: check_excel_modified(root))


# --- FUNGSI VALIDASI RESI (Ascending Mode) ---
def validate_resi_number(pdf_paths, excel_path):
    print("🔬 Validasi Resi (Ascending Mode)...")
    
    try:
        df = pd.read_excel(excel_path, header=None) 
        # Ambil data dari kolom ke-7 (indeks 6)
        resi_excel_list = [str(item).strip() for item in df.iloc[:, 6].dropna().tolist() if pd.notna(item)]
        
        if not resi_excel_list:
            print("Peringatan: Kolom 7 (Resi) di Excel kosong. Lanjut tanpa validasi resi.")
            return True 

        mismatches = []
        excel_idx = 0

        for pdf_path in pdf_paths:
            try:
                # Menggunakan PdfReader karena pdfplumber tidak diperlukan di sini
                pdf_reader = PdfReader(pdf_path)
                pdf_page_count = len(pdf_reader.pages)
                
                for i in range(pdf_page_count):
                    # Fungsi extract_resi_number_from_pdf menggunakan pdfplumber
                    resi_lengkap_pdf = extract_resi_number_from_pdf(pdf_path, i)
                    
                    resi_bersih_pdf = re.sub(r'\D', '', resi_lengkap_pdf)
                    resi_5_digit_pdf = resi_bersih_pdf[-5:] if len(resi_bersih_pdf) >= 5 else resi_bersih_pdf
                    
                    if excel_idx < len(resi_excel_list):
                        resi_excel_value = resi_excel_list[excel_idx]
                        resi_bersih_excel = re.sub(r'\D', '', resi_excel_value)
                        resi_5_digit_excel = resi_bersih_excel[-5:] if len(resi_bersih_excel) >= 5 else resi_bersih_excel

                        if resi_5_digit_pdf != resi_5_digit_excel:
                            mismatches.append({
                                "File PDF": os.path.basename(pdf_path),
                                "Halaman PDF": i + 1,
                                "Baris Excel": excel_idx + 1,
                                "Resi PDF (5 Digit Terakhir)": resi_5_digit_pdf,
                                "Resi Excel (Kolom 7)": resi_excel_value 
                            })
                        excel_idx += 1
                    if excel_idx >= len(resi_excel_list):
                        break

            except Exception as e:
                print(f"ERROR: Gagal memproses file {os.path.basename(pdf_path)} saat validasi resi: {e}")
                continue 
        
        if mismatches:
            print("🚨 GAGAL: Ditemukan ketidakcocokan 5 Digit Terakhir resi:")
            for m in mismatches:
                print(f"-> File {m['File PDF']}, Halaman {m['Halaman PDF']} (Baris Excel {m['Baris Excel']}): PDF 5 Digit Terakhir '{m['Resi PDF (5 Digit Terakhir)']}' TIDAK SAMA dengan Excel '{m['Resi Excel (Kolom 7)']}'")
            return False

        print("[STATUS] SUKSES: 5 Digit Terakhir resi di PDF cocok dengan Kolom 7 Excel.")
        return True

    except Exception as e:
        print(f"ERROR: Terjadi kesalahan saat validasi resi: {e}")
        messagebox.showerror("Kesalahan Validasi", f"Terjadi kesalahan saat membandingkan data resi: {e}. Pastikan Kolom 7 terisi data yang benar.")
        return False


# --- FUNGSI EKSTRAKSI RESI ---
def extract_resi_number_from_pdf(pdf_path, page_num):
    target_string = "交貨便服務代碼"
    
    # Fungsi ini memerlukan library pdfplumber
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[page_num]
            text = page.extract_text()
            
            if text is None:
                return "ERROR: Teks Tidak Dapat Diekstrak"
                
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
                    return f"DEBUG: '{target_string}' Ditemukan, TAPI POLA REGEX GAGAL" 
            
            # --- POLA CADANGAN ---
            long_number_pattern = re.compile(r'(\d{4}\s\d{4}\s\d{4}\s\d{4})')
            match = long_number_pattern.search(text)
            if match:
                return match.group(1).replace(' ', '')
                
            short_id_pattern = re.compile(r'(DRE\s*\d{3}\s*\d{4})')
            match = short_id_pattern.search(text)
            if match:
                return match.group(1).replace(' ', '')
                    
            return f"DEBUG: '{target_string}' TIDAK DITEMUKAN"
            
    except Exception as e:
        print(f"ERROR: Gagal Membaca/Memproses PDF di halaman {page_num + 1}: {e}")
        return f"ERROR: Gagal Membaca PDF"

# --- FUNGSI PENGECEKAN KESEIMBANGAN AWAL (Memuat Data Kolom 3) ---
def check_on_select(pdf_paths, show_print=True):
    global is_count_match, is_resi_match, keterangan_data_global, check_resi_var
    
    is_count_match = False
    is_resi_match = True # Default true, kecuali jika divalidasi dan gagal
    keterangan_data_global = None
    excel_path = get_excel_filename()

    if not excel_path or not pdf_paths:
        is_resi_match = False
        return False


    try:
        # 1. Memuat Data dari Kolom 1 (indeks 0) dan Kolom 3 (indeks 2)
        df = pd.read_excel(excel_path, header=None, usecols=[0, 2]) 
        
        # Bersihkan dan siapkan data (Kolom 1 - Keterangan Barang)
        keterangan_list = [str(item).replace('\n', ' ') for item in df.iloc[:, 0].dropna().tolist()]
        # Bersihkan dan siapkan data (Kolom 3 - Data Tambahan/Selip)
        kolom3_list = [str(item).replace('\n', ' ') for item in df.iloc[:, 1].fillna('').tolist()]
        
        # Pastikan panjang kolom 1 dan 3 sama (hanya untuk data yang sudah di dropna() di kolom 1)
        kolom3_list = kolom3_list[:len(keterangan_list)]
        
        jumlah_keterangan_excel = len(keterangan_list)
        
        # Gabungkan data menjadi list of lists: [(kolom1_data, kolom3_data), ...]
        keterangan_data_global = list(zip(keterangan_list, kolom3_list))

        # 2. Hitung Total Halaman PDF dari SEMUA file
        jumlah_halaman_pdf = 0
        for pdf_path in pdf_paths:
            try:
                pdf_reader = PdfReader(pdf_path)
                jumlah_halaman_pdf += len(pdf_reader.pages)
            except Exception as e:
                print(f"Peringatan: Gagal membaca file PDF {os.path.basename(pdf_path)}: {e}")
                
        
        update_excel_count_label(jumlah_keterangan_excel)
        update_pdf_count_label(jumlah_halaman_pdf)

        if show_print:
            print("\n" + "=" * 50)
            print(f"✅ Pengecekan Keseimbangan Data...")
            print(f"-> Jumlah Baris Data (Kolom 1) di Excel: {jumlah_keterangan_excel}")
            print(f"-> Jumlah Halaman Total dari PDF: {jumlah_halaman_pdf}")

        if jumlah_keterangan_excel == jumlah_halaman_pdf:
            is_count_match = True
            if show_print:
                print(f"[STATUS] SUKSES: Jumlah baris data dan halaman SAMA PERSIS.")

            # Logika Validasi Resi Baru
            is_resi_check_enabled = check_resi_var.get() == 1
            if sort_var.get() == "Ascending" and is_resi_check_enabled:
                is_resi_match = validate_resi_number(pdf_paths, excel_path)
            elif sort_var.get() == "Ascending" and not is_resi_check_enabled:
                 # Diabaikan, anggap true
                 is_resi_match = True
                 if show_print:
                     print("[STATUS] Pengecekan Nomor Resi DIIBAIKAN oleh pengguna.")
            else:
                # Descending/Family Mart selalu OK
                is_resi_match = True 

        else:
            is_count_match = False
            is_resi_match = False
            error_message = (
                f"🚨 PERINGATAN! JUMLAH TIDAK SAMA!\n"
                f"Baris Data Excel: {jumlah_keterangan_excel}\n"
                f"Halaman PDF: {jumlah_halaman_pdf}"
            )
            if show_print:
                print(f"[STATUS] GAGAL: {error_message}".replace('\n', ' '))
                
        if show_print:
            print("=" * 50)
            
        update_check_status_display(is_count_match, is_resi_match)
        return is_count_match and is_resi_match

    except Exception as e:
        update_excel_count_label(0)
        update_pdf_count_label(0)
        is_count_match = False
        is_resi_match = False
        update_check_status_display(False, False)
        if show_print:
            print(f"\n[STATUS] ERROR: Terjadi kesalahan saat pengecekan file: {e}")
            messagebox.showerror("Kesalahan Pengecekan", f"Terjadi kesalahan saat memuat data: {e}")
        return False


# Fungsi untuk memilih file PDF (REVERSE URUTAN)
def choose_pdf_file():
    global pdf_file_path_list, current_selected_pdf_index
    
    filepaths = filedialog.askopenfilenames(
        title="Pilih File Resi (PDF) - Boleh Pilih Lebih Dari Satu",
        filetypes=[("PDF files", "*.pdf")]
    )
    
    if filepaths:
        pdf_file_path_list = list(filepaths)
        
        # MEMBALIKKAN URUTAN FILE (Solusi untuk urutan default yang terbalik)
        pdf_file_path_list.reverse() 
        print(f"Mengubah urutan file yang diimpor (reverse). File pertama yang akan diproses: {os.path.basename(pdf_file_path_list[0])}")
        
        pdf_path_label.config(text=f"{len(pdf_file_path_list)} File Dipilih")
        current_selected_pdf_index = 0 
        update_pdf_list_display(pdf_file_path_list, current_selected_pdf_index)
        
        check_on_select(pdf_file_path_list, show_print=True)
    else:
        pdf_file_path_list = [] 
        pdf_path_label.config(text="Tidak ada file yang dipilih")
        current_selected_pdf_index = -1
        update_pdf_list_display([])
        update_excel_count_label(0)
        update_pdf_count_label(0)
        update_check_status_display(False, False) 


def change_sort_order(event=None):
    # Panggil fungsi toggle untuk mengontrol status checkbox
    toggle_resi_checkbox() 
    
    if pdf_file_path_list:
        check_on_select(pdf_file_path_list, show_print=True)
    else:
        update_check_status_display(is_count_match, is_resi_match)


# Fungsi untuk memulai proses utama
def start_process(sort_order):
    global pdf_file_path_list, is_count_match, is_resi_match, check_resi_var

    output_text.delete(1.0, tk.END)

    if not pdf_file_path_list:
        messagebox.showerror("Kesalahan", "Harap pilih minimal satu file PDF terlebih dahulu.")
        return
    
    # Lakukan pengecekan akhir sebelum memulai proses
    if not check_on_select(pdf_file_path_list, show_print=True):
        messagebox.showerror("Kesalahan", "Pengecekan Keseimbangan GAGAL. Periksa Output Program untuk detail.")
        return

    # Validasi Tambahan (Hanya cek Resi jika Ascending DAN Checkbox dicentang)
    is_resi_check_enabled = check_resi_var.get() == 1
    if sort_order == "Ascending" and is_resi_check_enabled and not is_resi_match:
        messagebox.showerror("Kesalahan", "Pengecekan Nomor Resi GAGAL. Harap periksa Kolom 7 Excel Anda.")
        return
        
    # --- LOGIKA PENAMAAN FILE OUTPUT ---
    first_file_path = pdf_file_path_list[0]
    first_file_name_base = os.path.basename(first_file_path)
    base_name, ext = os.path.splitext(first_file_name_base)
    
    if len(pdf_file_path_list) == 1:
        default_save_name = first_file_name_base
    else:
        default_save_name = f"{base_name}_FULL{ext}"
    
    print(f"Nama default output diatur ke: {default_save_name}")

    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        initialdir=os.path.dirname(first_file_path) or os.getcwd(),
        initialfile=default_save_name,
        filetypes=[("PDF files", "*.pdf")]
    )

    if not save_path:
        print("Operasi dibatalkan oleh pengguna.")
        return

    process_pdf_and_excel(sort_order, pdf_file_path_list, save_path)


# --- FUNGSI PROSES UTAMA (Menambahkan Teks Khusus) ---
def process_pdf_and_excel(sort_order, pdf_input_paths, pdf_output_path):
    # Diperbarui: menggunakan keterangan_data_global
    global keterangan_data_global, is_count_match, is_resi_match, check_resi_var

    # Pengecekan ulang sebelum memulai
    if not is_count_match:
        messagebox.showerror("Kesalahan", "Pengecekan Keseimbangan Gagal.")
        return
    
    is_resi_check_enabled = check_resi_var.get() == 1
    if sort_order == "Ascending" and is_resi_check_enabled and not is_resi_match:
        messagebox.showerror("Kesalahan", "Pengecekan Resi Gagal.")
        return


    keterangan_data = keterangan_data_global 

    print("Mengeksekusi program...")
    print("=" * 50)

    try:
        if sort_order == "Descending":
            keterangan_data.reverse()
            print("Urutan keterangan dari Bawah ke Atas (Descending).")
        else:
            print("Urutan keterangan dari Atas ke Bawah (Ascending).")
        
        pdf_writer = PdfWriter()
        keterangan_index = 0

        # --- PENGATURAN FONT KUSTOM ---
        FONT_NORMAL = "Helvetica-Bold"
        FONT_SIZE_NORMAL = 9  # Ukuran default 
        FONT_SIZE_NUMBER = 12 # Ukuran khusus untuk angka murni
        MAX_LINE_WIDTH = 300
        LINE_SPACING = FONT_SIZE_NORMAL + 1 # Jarak antar baris
        # -----------------------------

        for pdf_input_path in pdf_input_paths:
            print(f"Memproses file: {os.path.basename(pdf_input_path)}...")

            try:
                pdf_reader = PdfReader(pdf_input_path)
            except Exception as e:
                print(f"Peringatan: Gagal memuat file {os.path.basename(pdf_input_path)}. Melewati file ini. Error: {e}")
                continue

            if 'page_height' not in locals():
                page_height = float(pdf_reader.pages[0].mediabox.height)

            for i, page in enumerate(pdf_reader.pages):
                if keterangan_index >= len(keterangan_data):
                    print(f"Peringatan: Keterangan Excel habis di Halaman {i+1} dari {os.path.basename(pdf_input_path)}. Berhenti memproses halaman PDF.")
                    break

                # Ambil data dari list global: (keterangan_barang, data_kolom_3)
                keterangan_barang_asli, data_kolom3_asli = keterangan_data[keterangan_index]
                keterangan_index += 1
                
                # --- LOGIKA PENAMBAHAN TULISAN KHUSUS (SELIPKAN 60 NT) ---
                keterangan_final = keterangan_barang_asli
                # Cek apakah angka '60' ada di data_kolom3_asli (menggunakan string)
                if '60' in str(data_kolom3_asli):
                    teks_khusus = "(SELIPKAN 60 NT) "
                    print(f" -> Deteksi '60' di Kolom 3 (Halaman {i+1}): Menambahkan '{teks_khusus.strip()}'")
                    keterangan_final = teks_khusus + keterangan_barang_asli
                # -----------------------------------------------------------

                
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=A4)

                x_pos_center = 258
                y_pos_from_top = 190
                y_pos_from_bottom = page_height - y_pos_from_top

                # --- 1. LOGIKA WORD WRAPPING ---
                # Menggunakan FONT_SIZE_NORMAL untuk perhitungan lebar baris
                can.setFont(FONT_NORMAL, FONT_SIZE_NORMAL)
                
                # Menggunakan keterangan_final
                words = keterangan_final.split(' ')
                lines = []
                current_line = ""
                for word in words:
                    # Menggunakan spasi agar perhitungan lebar akurat
                    test_line = current_line + " " + word if current_line else word
                    
                    # Cek lebar menggunakan FONT_SIZE_NORMAL
                    if can.stringWidth(test_line, FONT_NORMAL, FONT_SIZE_NORMAL) < MAX_LINE_WIDTH:
                        current_line = test_line
                    else:
                        if current_line:
                            lines.append(current_line.strip())
                        current_line = word
                lines.append(current_line.strip())
                
                # --- 2. PENENTUAN POSISI AWAL ---
                
                # Perkiraan tinggi blok (menggunakan LINE_SPACING)
                text_block_height = len(lines) * LINE_SPACING 
                initial_y_pos = y_pos_from_bottom + (text_block_height / 2)

                can.saveState()
                can.translate(x_pos_center, initial_y_pos)
                can.rotate(90)
                
                # --- 3. LOGIKA CETAK DENGAN FONT SIZE YANG BERBEDA ---
                for j, line in enumerate(lines):
                    current_x = 0
                    line_words = line.split(' ')
                    
                    # --- Kalkulasi Lebar untuk Centering ---
                    total_width = 0
                    for word in line_words:
                        # Pengecekan hanya untuk angka murni
                        is_number = word.isdigit()
                        calc_font_size = FONT_SIZE_NUMBER if is_number else FONT_SIZE_NORMAL
                        
                        total_width += can.stringWidth(word, FONT_NORMAL, calc_font_size)
                        total_width += can.stringWidth(' ', FONT_NORMAL, FONT_SIZE_NORMAL) 
                    
                    if total_width > 0:
                        total_width -= can.stringWidth(' ', FONT_NORMAL, FONT_SIZE_NORMAL) 

                    start_x = -total_width / 2
                    # --- Akhir Kalkulasi Lebar ---

                    for k, word in enumerate(line_words):
                        # Pengecekan hanya untuk angka murni
                        is_number = word.isdigit()
                        
                        # Tentukan font size: 12 jika angka murni, 9 jika bukan
                        font_size = FONT_SIZE_NUMBER if is_number else FONT_SIZE_NORMAL
                        can.setFont(FONT_NORMAL, font_size)

                        word_width = can.stringWidth(word, FONT_NORMAL, font_size)
                        
                        # Menggunakan LINE_SPACING 
                        y_pos_in_rotation = j * -LINE_SPACING
                        
                        # Mencetak kata dengan font size yang telah disesuaikan
                        can.drawString(start_x + current_x, y_pos_in_rotation, word)
                        
                        # Pindahkan posisi x untuk kata berikutnya
                        current_x += word_width + can.stringWidth(' ', FONT_NORMAL, FONT_SIZE_NORMAL)
                
                can.restoreState()
                can.save()
                # --- AKHIR LOGIKA REPORTLAB ---

                packet.seek(0)
                new_pdf = PdfReader(packet)
                page.merge_page(new_pdf.pages[0])
                pdf_writer.add_page(page)
            
            if keterangan_index >= len(keterangan_data):
                break 

        
        with open(pdf_output_path, 'wb') as f:
            pdf_writer.write(f)

        print(f"\nOperasi selesai. File hasil disimpan sebagai '{os.path.basename(pdf_output_path)}'.")
        
        # --- BUKA FILE OTOMATIS (AKTIF) ---
        open_file_in_os(pdf_output_path)
        # ---------------------------

        messagebox.showinfo("Selesai", f"Operasi selesai. File hasil disimpan sebagai '{os.path.basename(pdf_output_path)}' dan telah dibuka.")

    except Exception as e:
        print(f"\nTerjadi kesalahan: {e}")
        messagebox.showerror("Kesalahan", f"Terjadi kesalahan: {e}")


# --- SETUP GUI ---
def create_ui():
    global pdf_path_label, output_text, excel_count_label, pdf_count_label, last_excel_modified_time, status_label_count, status_label_resi, sort_var, pdf_list_display, check_resi_var, resi_check_box_widget
    
    root = tk.Tk()
    root.title("Alat Input Keterangan Resi")
    
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
    style.configure('TCheckbutton', background=APP_BG, font=('Helvetica', 10, 'normal')) # Gaya untuk Checkbox
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

    # --- STEP 1: PILIHAN URUTAN & CHECKBOX (Kolom 0) ---
    sort_var = tk.StringVar(value="Ascending") 
    check_resi_var = tk.IntVar(value=1) # Default dicentang
    
    step1_frame = ttk.Frame(top_grid_frame)
    step1_frame.grid(row=0, column=0, padx=10, sticky='nwes')
    ttk.Label(step1_frame, text="Step 1: Jenis Urutan Data & Opsi", style='Step.TLabel').pack(pady=(0, 10), anchor='center') 
    
    radio_frame1 = ttk.Frame(step1_frame)
    radio_frame1.pack(anchor='w', pady=(5, 10), padx=20) 

    # Radio Button untuk Ascending
    asc_radio = ttk.Radiobutton(radio_frame1, text="Ascending (7-Eleven)", variable=sort_var, value="Ascending", command=change_sort_order)
    asc_radio.pack(anchor='w', pady=(0, 5)) 
    
    # Radio Button untuk Descending
    desc_radio = ttk.Radiobutton(radio_frame1, text="Descending (Family-Mart)", variable=sort_var, value="Descending", command=change_sort_order)
    desc_radio.pack(anchor='w', pady=(0, 10)) 
    
    # Checkbox untuk Validasi Resi
    resi_check_box = ttk.Checkbutton(step1_frame, 
                                     text="Cek No. Resi 7-Eleven", 
                                     variable=check_resi_var,
                                     command=change_sort_order,
                                     style='TCheckbutton')
    resi_check_box.pack(anchor='w', padx=20, pady=(0, 5))
    
    # BARU: Simpan referensi widget untuk dikontrol oleh toggle_resi_checkbox
    resi_check_box_widget = resi_check_box 
    
    ttk.Separator(top_grid_frame, orient='vertical').grid(row=0, column=1, sticky='ns', padx=10)


    # --- STEP 2: EXCEL DATA (Kolom 2) ---
    step2_frame = ttk.Frame(top_grid_frame)
    step2_frame.grid(row=0, column=2, padx=10, sticky='nwes')
    ttk.Label(step2_frame, text="Step 2: Excel Data", style='Step.TLabel').pack(pady=(0, 10), anchor='center')
    
    excel_content_frame = ttk.Frame(step2_frame)
    excel_content_frame.pack(anchor='center')
    
    excel_name = os.path.basename(get_excel_filename()) if get_excel_filename() else "x.xlsx (Tidak Ditemukan)"
    ttk.Label(excel_content_frame, text=f"File Excel Ditemukan: {excel_name}").pack(pady=(0, 5), anchor='w')
    
    edit_excel_label = ttk.Label(excel_content_frame, text="Edit file Excel", foreground="#0000ff", cursor="hand2")
    edit_excel_label.pack(pady=(5, 10), anchor='w')
    edit_excel_label.bind("<Button-1>", lambda e: edit_excel_file())

    ttk.Label(excel_content_frame, text="Total Baris Data:").pack(pady=(10, 2), anchor='w')
    excel_count_box = tk.Frame(excel_content_frame, bg=STATUS_BOX_BG, relief=STATUS_BOX_RELIEF, borderwidth=1, width=150, height=30)
    excel_count_box.pack(anchor='w')
    excel_count_box.pack_propagate(False) 
    excel_count_label = tk.Label(excel_count_box, text="0", 
                                 fg='black', bg=STATUS_BOX_BG, font=('Helvetica', 12, 'normal'), 
                                 anchor='center', justify=tk.CENTER)
    excel_count_label.pack(expand=True, fill='both') 

    ttk.Separator(top_grid_frame, orient='vertical').grid(row=0, column=3, sticky='ns', padx=10)


    # --- STEP 3: PILIH RESI PDF (Kolom 4) ---
    step3_frame = ttk.Frame(top_grid_frame)
    step3_frame.grid(row=0, column=4, padx=10, sticky='nwes')
    
    ttk.Label(step3_frame, text="Step 3: Pilih Resi PDF (Multi-file)", style='Step.TLabel').pack(pady=(0, 10), anchor='center')
    
    pdf_controls_frame = ttk.Frame(step3_frame)
    pdf_controls_frame.pack(anchor='center', fill='x', padx=10)

    pdf_path_label = ttk.Label(pdf_controls_frame, text="Tidak ada file yang dipilih")
    pdf_path_label.pack(pady=(5, 5), anchor='center') 
    
    pdf_input_frame = ttk.Frame(pdf_controls_frame)
    pdf_input_frame.pack(pady=5, anchor='center') 
    
    ttk.Button(pdf_input_frame, text="Add PDF(s)", command=choose_pdf_file, style='TButton').pack(side=tk.LEFT)
    
    ttk.Label(pdf_controls_frame, text="Daftar File Dipilih:").pack(pady=(10, 2), anchor='w')
    
    list_and_control_frame = ttk.Frame(pdf_controls_frame)
    list_and_control_frame.pack(fill=tk.X, expand=True)

    pdf_list_display = ScrolledText(list_and_control_frame, height=5, width=40, state='disabled', wrap=tk.WORD, relief='sunken', borderwidth=1)
    pdf_list_display.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    pdf_list_display.bind('<Button-1>', get_selected_pdf_index) 

    order_buttons_frame = ttk.Frame(list_and_control_frame)
    order_buttons_frame.pack(side=tk.RIGHT, padx=(5, 0))

    ttk.Button(order_buttons_frame, text="▲ Up", command=move_pdf_up).pack(fill=tk.X, pady=2)
    ttk.Button(order_buttons_frame, text="▼ Down", command=move_pdf_down).pack(fill=tk.X, pady=2)
    
    
    ttk.Label(pdf_controls_frame, text="Total Halaman PDF:").pack(pady=(10, 2), anchor='center')
    pdf_count_box = tk.Frame(pdf_controls_frame, bg=STATUS_BOX_BG, relief=STATUS_BOX_RELIEF, borderwidth=1, width=150, height=30)
    pdf_count_box.pack(anchor='center') 
    pdf_count_box.pack_propagate(False) 
    pdf_count_label = tk.Label(pdf_count_box, text="0", 
                                 fg='black', bg=STATUS_BOX_BG, font=('Helvetica', 12, 'normal'), 
                                 anchor='center', justify=tk.CENTER) 
    pdf_count_label.pack(expand=True, fill='both')
    
    
    # --- COMMAND SIMPLE KECIL DI ATAS START ---
    step_status_frame = tk.Frame(main_frame, bg=STATUS_FRAME_BG, relief='groove', borderwidth=1, padx=10, pady=5)
    step_status_frame.pack(pady=(10, 15), fill=tk.X)
    
    ttk.Label(step_status_frame, text="Pengecekan Status:", font=('Helvetica', 10, 'bold'), background=STATUS_FRAME_BG).pack(anchor='w')
    
    status_label_count = ttk.Label(step_status_frame, text="1. Baris Data (Excel) dan Halaman (PDF) ⚪", background=STATUS_FRAME_BG)
    status_label_count.pack(anchor='w')
    
    status_label_resi = ttk.Label(step_status_frame, text="2. Pengecekan Nomor Resi (Kolom 7) ⚪", background=STATUS_FRAME_BG)
    status_label_resi.pack(anchor='w')
    # -------------------------------------------------

    start_button = ttk.Button(main_frame, text="Start", command=lambda: start_process(sort_var.get()), style='Start.TButton')
    start_button.pack(pady=(10, 20), ipadx=50)

    ttk.Label(main_frame, text="Output Program :").pack(pady=(0, 5), anchor='w')
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
            print(f"Peringatan: Gagal memuat/menghitung Excel saat startup: {e}")
            update_excel_count_label(0)
    
    
    # Panggil fungsi toggle untuk inisialisasi status checkbox saat startup
    toggle_resi_checkbox()
    
    update_check_status_display(False, False)
    update_pdf_list_display([]) 
    
    check_excel_modified(root) 
    
    credit_label = ttk.Label(root, text="© didk_", font=('Helvetica', 8, 'italic'), background=APP_BG, foreground='#666666')
    credit_label.place(relx=0.0, rely=1.0, anchor='sw', x=10, y=0)
    
    root.mainloop()

# Jalankan UI
if __name__ == '__main__':
    create_ui()