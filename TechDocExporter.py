import tkinter as tk
from tkinter import scrolledtext
import os
import sys
import io
import shutil
import pythoncom
import win32com.client
from PyPDF2 import PdfReader, PdfWriter

# --- Надёжное определение пути к директории с .py или .exe ---
def get_base_dir():
    if getattr(sys, 'frozen', False):  # Если exe
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# --- Перенаправление stdout в GUI ---
class RedirectText(io.TextIOBase):
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)

    def flush(self):
        pass


# --- PREPRINT ---
def clear_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"Ошибка при удалении '{file_path}': {e}")

def convert_xlsm_to_xlsx(xlsm_file, xlsx_file):
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        workbook = excel.Workbooks.Open(xlsm_file)
        workbook.SaveAs(xlsx_file, FileFormat=51)
        workbook.Close(SaveChanges=False)
        excel.Quit()
        print(f"Файл '{xlsm_file}' преобразован в '{xlsx_file}'")
    except Exception as e:
        print(f"Ошибка Excel: {e}")
    finally:
        pythoncom.CoUninitialize()

def excel_to_pdf(excel_file, pdf_file):
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        workbook = excel.Workbooks.Open(excel_file)
        sheet_count = workbook.Sheets.Count
        filename = os.path.basename(excel_file)

        if filename.startswith("ВСИ"):
            exclude_count = 1
        elif filename.startswith("ВМ"):
            exclude_count = 2
        elif filename[0].isdigit():
            exclude_count = 0
        else:
            exclude_count = 0

        sheets_to_print = list(range(1, sheet_count - exclude_count + 1))
        if sheets_to_print:
            workbook.ExportAsFixedFormat(
                Type=0,
                Filename=pdf_file,
                Quality=1,
                IncludeDocProperties=False,
                IgnorePrintAreas=False,
                From=sheets_to_print[0],
                To=sheets_to_print[-1]
            )
            print(f"Экспортировано в PDF: {pdf_file}")
        else:
            print(f"Пропущен файл (нет листов): {excel_file}")

        workbook.Close(SaveChanges=False)
        excel.Quit()
    except Exception as e:
        print(f"Ошибка конвертации '{excel_file}': {e}")
    finally:
        pythoncom.CoUninitialize()

def merge_pdfs(pdf_files, output_file, mode='full'):
    try:
        writer = PdfWriter()
        for pdf_file in pdf_files:
            reader = PdfReader(pdf_file)
            if mode == 'full':
                pages = range(len(reader.pages))
            elif mode == 'title':
                pages = [0] if reader.pages else []
            elif mode == 'notitle':
                pages = range(1, len(reader.pages))
            else:
                raise ValueError(f"Неизвестный режим: {mode}")
            for i in pages:
                writer.add_page(reader.pages[i])
        with open(output_file, "wb") as f:
            writer.write(f)
        print(f"Собран файл: {output_file}")
    except Exception as e:
        print(f"Ошибка объединения PDF: {e}")

def process_preprint():
    base_dir = get_base_dir()
    excel_input_dir = os.path.join(base_dir, "Excel")
    print_dir = os.path.join(base_dir, "Print")
    export_dir = os.path.join(base_dir, "NotSignedExport")
    service_dir = os.path.join(base_dir, "Service")

    os.makedirs(excel_input_dir, exist_ok=True)
    os.makedirs(print_dir, exist_ok=True)
    os.makedirs(export_dir, exist_ok=True)
    os.makedirs(service_dir, exist_ok=True)

    clear_folder(print_dir)
    clear_folder(export_dir)
    clear_folder(service_dir)

    pdf_files = []

    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    excel.Visible = False

    try:
        for filename in os.listdir(excel_input_dir):
            full_path = os.path.join(excel_input_dir, filename)
            base_filename, ext = os.path.splitext(filename)

            # Обработка .xlsm: конвертируем во временный .xlsx
            if ext.lower() == ".xlsm":
                temp_xlsx_path = os.path.join(excel_input_dir, base_filename + ".xlsx")
                try:
                    wb = excel.Workbooks.Open(full_path)
                    wb.SaveAs(temp_xlsx_path, FileFormat=51)
                    wb.Close(SaveChanges=False)
                    del wb
                    os.remove(full_path)  # удаляем исходный xlsm
                    full_path = temp_xlsx_path
                    ext = ".xlsx"
                    print(f"Преобразован: {filename} → {os.path.basename(temp_xlsx_path)}")
                except Exception as e:
                    print(f"❗ Ошибка при преобразовании {filename}: {e}")
                    continue

            # Обработка .xlsx или .xls
            if ext.lower() in [".xlsx", ".xls"]:
                try:
                    wb = excel.Workbooks.Open(full_path)
                    sheet_count = wb.Sheets.Count
                    name = os.path.basename(full_path)

                    if name.startswith("ВСИ"):
                        exclude_count = 1
                    elif name.startswith("ВМ"):
                        exclude_count = 2
                    elif name[0].isdigit():
                        exclude_count = 0
                    else:
                        exclude_count = 0

                    sheets_to_print = list(range(1, sheet_count - exclude_count + 1))

                    if sheets_to_print:
                        pdf_path = os.path.join(export_dir, base_filename + ".pdf")
                        wb.ExportAsFixedFormat(
                            Type=0,
                            Filename=pdf_path,
                            Quality=1,
                            IncludeDocProperties=False,
                            IgnorePrintAreas=False,
                            From=sheets_to_print[0],
                            To=sheets_to_print[-1]
                        )
                        pdf_files.append(pdf_path)
                        print(f"Экспортировано: {pdf_path}")
                    else:
                        print(f"Пропущен файл без листов: {filename}")

                    wb.Close(SaveChanges=False)
                    del wb

                    # Если это временный .xlsx из .xlsm — удалим
                    if filename.endswith(".xlsm"):
                        os.remove(full_path)

                except Exception as e:
                    print(f"❗ Ошибка при экспорте {filename}: {e}")
                    continue

    finally:
        excel.Quit()
        del excel
        pythoncom.CoUninitialize()

    # Переименование и копирование PDF в Service
    pdf_files.sort()
    renamed_files = []
    for i, pdf_path in enumerate(pdf_files, 1):
        new_path = os.path.join(service_dir, f"{i:03d}_{os.path.basename(pdf_path)}")
        shutil.copy2(pdf_path, new_path)
        renamed_files.append(new_path)
        print(f"Скопирован: {pdf_path} → {new_path}")

    # Объединение PDF
    merge_pdfs(renamed_files, os.path.join(print_dir, "complete_merged.pdf"), mode='full')
    merge_pdfs(renamed_files, os.path.join(print_dir, "title_merged.pdf"), mode='title')
    merge_pdfs(renamed_files, os.path.join(print_dir, "no_title_merged.pdf"), mode='notitle')


# --- POSTPRINT ---
def replace_first_page(source_pdf, new_first_page, output_pdf):
    reader = PdfReader(source_pdf)
    writer = PdfWriter()
    writer.add_page(new_first_page)
    for page in reader.pages[1:]:
        writer.add_page(page)
    with open(output_pdf, 'wb') as f:
        writer.write(f)

def clear_output_directory(output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        return
    for file in os.listdir(output_dir):
        if file.lower().endswith(".pdf"):
            os.remove(os.path.join(output_dir, file))

def process_postprint():
    base_dir = get_base_dir()
    service_dir = os.path.join(base_dir, "Service")
    print_dir = os.path.join(base_dir, "Print")
    output_dir = os.path.join(base_dir, "Final")
    title_scan_pdf = os.path.join(print_dir, "title_scan.pdf")

    if not os.path.exists(title_scan_pdf):
        raise FileNotFoundError("Файл 'title_scan.pdf' не найден в папке Print.")

    title_reader = PdfReader(title_scan_pdf)
    scanned_pages = title_reader.pages

    os.makedirs(output_dir, exist_ok=True)
    clear_output_directory(output_dir)

    numbered_pdfs = sorted([
        f for f in os.listdir(service_dir)
        if f.lower().endswith(".pdf") and f[:3].isdigit() and "_" in f
    ])

    for i, filename in enumerate(numbered_pdfs):
        input_path = os.path.join(service_dir, filename)
        if i >= len(scanned_pages):
            print(f"⚠️ Страница {i + 1} отсутствует в скане.")
            continue
        new_name = filename[4:]
        output_path = os.path.join(output_dir, new_name)
        replace_first_page(input_path, scanned_pages[i], output_path)
        print(f"✅ Обновлён: {filename} → {new_name}")


# --- GUI CALLBACKS ---
def check_files_in_excel():
    log_output.delete(1.0, tk.END)
    base_dir = get_base_dir()
    excel_dir = os.path.join(base_dir, "Excel")
    files = os.listdir(excel_dir) if os.path.exists(excel_dir) else []
    if files:
        log_output.insert(tk.END, "Найдены файлы:\n" + "\n".join(files))
    else:
        log_output.insert(tk.END, "В папке 'Excel' нет файлов.\n")

def run_preprint():
    log_output.delete(1.0, tk.END)
    try:
        log_output.insert(tk.END, "▶ Запуск подготовительного этапа...\n")
        process_preprint()
        log_output.insert(tk.END, "✅ Завершено.\n")
    except Exception as e:
        log_output.insert(tk.END, f"❗ Ошибка: {e}\n")

def run_postprint():
    log_output.delete(1.0, tk.END)
    try:
        log_output.insert(tk.END, "▶ Замена титульников...\n")
        process_postprint()
        log_output.insert(tk.END, "✅ Завершено.\n")
    except FileNotFoundError as e:
        log_output.insert(tk.END, f"❗ Ошибка: {e}\n")
    except Exception as e:
        log_output.insert(tk.END, f"⚠️ Ошибка: {e}\n")


# --- GUI ---
root = tk.Tk()
root.title("Утилита для подготовки документов")
root.geometry("800x600")
root.minsize(700, 500)
root.config(bg="#f0f0f0")
root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)

top_frame = tk.Frame(root, bg="#f0f0f0")
top_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)
top_frame.grid_columnconfigure(0, weight=0)
top_frame.grid_columnconfigure(1, weight=1)
top_frame.grid_rowconfigure(0, weight=1)

button_frame = tk.Frame(top_frame, bg="#f0f0f0")
button_frame.grid(row=0, column=0, sticky="n", pady=(40, 0))

tk.Button(button_frame, text="Проверить файлы Excel", command=check_files_in_excel,
          bg="#2196F3", fg="white", width=30).grid(row=0, column=0, pady=(0, 15), sticky="w")

tk.Button(button_frame, text="Подготовить к печати", command=run_preprint,
          bg="#FF9800", fg="white", width=30).grid(row=1, column=0, pady=15, sticky="w")

tk.Button(button_frame, text="Заменить титульники", command=run_postprint,
          bg="#FF5722", fg="white", width=30).grid(row=2, column=0, pady=15, sticky="w")

instruction_frame = tk.Frame(top_frame, bg="#f0f0f0")
instruction_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 10), pady=(0, 10))
instruction_label = tk.Label(instruction_frame, text="Инструкция:", bg="#f0f0f0", font=("Arial", 11, "bold"))
instruction_label.grid(row=0, column=0, sticky="w")
instruction_text = tk.Text(instruction_frame, height=12, wrap="word", bg="#fff", font=("Arial", 10))
instruction_text.grid(row=1, column=0, sticky="nsew")
instruction_text.insert(tk.END, """
1. Загрузите excel-файлы в папку "Excel".

2. Нажмите "Проверить файлы Excel" и убедитесь что все необходимые файлы найдены.

3. Нажмите "Подготовить к печати".
    - программа создаст папку Print в которой будут лежать PDF
        > все файлы в одном по порядку
        > все титульники в одном
        > все файлы без титульников в одном
    - программа создаст папку NotSignedExport в которой будут все PDF по отдельности
    
4. Распечатайте файл "title_merged" (титульники), отправьте на подпись и отсканируйте.

    ====!ВАЖНО!========!ВАЖНО!========!ВАЖНО!====
    - при сканировании убедитесь что листы лежат в том же порядке что и при печати
    
5. Отсканированный файл назовите "title_scan" и положить в папку Print.

    ====!ВАЖНО!========!ВАЖНО!========!ВАЖНО!====
    - Название файла со сканами СТРОГО "title_scan"
    
6. Нажмите кнопку "Заменить титульники". Программа поместит результат в папку "Final".
""")
instruction_text.config(state='disabled')
instruction_frame.grid_rowconfigure(1, weight=1)
instruction_frame.grid_columnconfigure(0, weight=1)

log_output = scrolledtext.ScrolledText(root, wrap=tk.WORD, bg="#ffffff", fg="#000000", font=("Consolas", 10))
log_output.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=(0, 10))

sys.stdout = RedirectText(log_output)

# Гарантируем наличие папки Excel
excel_dir = os.path.join(get_base_dir(), "Excel")
if not os.path.exists(excel_dir):
    os.mkdir(excel_dir)

root.mainloop()
