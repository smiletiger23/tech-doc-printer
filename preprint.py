import os
import shutil
import pythoncom
import win32com.client
from PyPDF2 import PdfReader, PdfWriter

def clear_folder(folder_path):
    """
    Удаляет все файлы в указанной папке, не удаляя саму папку.
    """
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

        try:
            workbook = excel.Workbooks.Open(xlsm_file)
        except Exception as e:
            print(f"Ошибка при открытии файла '{xlsm_file}': {e}")
            return

        try:
            workbook.SaveAs(xlsx_file, FileFormat=51)
            print(f"Файл '{xlsm_file}' успешно преобразован в '{xlsx_file}'")
        except Exception as e:
            print(f"Ошибка при сохранении файла '{xlsm_file}' как '{xlsx_file}': {e}")
            return
        finally:
            workbook.Close(SaveChanges=False)
            excel.Quit()

    except Exception as e:
        print(f"Ошибка при обработке файла '{xlsm_file}': {e}")
    finally:
        pythoncom.CoUninitialize()

def excel_to_pdf(excel_file, pdf_file):
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        try:
            workbook = excel.Workbooks.Open(excel_file)
        except Exception as e:
            print(f"Ошибка при открытии файла '{excel_file}': {e}")
            return

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
        if not sheets_to_print:
            print(f"Нет листов для печати в файле '{excel_file}'")
            workbook.Close(SaveChanges=False)
            excel.Quit()
            return

        try:
            workbook.ExportAsFixedFormat(
                Type=0,
                Filename=pdf_file,
                Quality=1,
                IncludeDocProperties=False,
                IgnorePrintAreas=False,
                From=sheets_to_print[0],
                To=sheets_to_print[-1]
            )
        except Exception as e:
            print(f"Ошибка при экспорте в PDF файла '{excel_file}': {e}")
            workbook.Close(SaveChanges=False)
            excel.Quit()
            return

        workbook.Close(SaveChanges=False)
        excel.Quit()
        print(f"Файл '{excel_file}' успешно преобразован в PDF '{pdf_file}'")

    except Exception as e:
        print(f"Ошибка при обработке файла '{excel_file}': {e}")
    finally:
        pythoncom.CoUninitialize()

def merge_pdfs(pdf_files, output_file, mode='full'):
    try:
        writer = PdfWriter()

        for pdf_file in pdf_files:
            reader = PdfReader(pdf_file)
            num_pages = len(reader.pages)

            if mode == 'full':
                pages = range(num_pages)
            elif mode == 'title':
                pages = [0] if num_pages > 0 else []
            elif mode == 'notitle':
                pages = range(1, num_pages)
            else:
                raise ValueError(f"Неизвестный режим объединения: {mode}")

            for i in pages:
                writer.add_page(reader.pages[i])

        with open(output_file, "wb") as output_pdf:
            writer.write(output_pdf)

        print(f"PDF-файлы объединены в '{output_file}' (режим: {mode})")

    except Exception as e:
        print(f"Ошибка при объединении PDF-файлов ({mode}): {e}")

def process_files(directory):
    excel_input_dir = os.path.join(directory, "Excel")
    print_dir = os.path.join(directory, "Print")
    export_dir = os.path.join(directory, "NotSignedExport")
    service_dir = os.path.join(directory, "Service")

    os.makedirs(excel_input_dir, exist_ok=True)
    os.makedirs(print_dir, exist_ok=True)
    os.makedirs(export_dir, exist_ok=True)
    os.makedirs(service_dir, exist_ok=True)

    clear_folder(print_dir)
    clear_folder(export_dir)
    clear_folder(service_dir)

    pdf_files = []

    for filename in excel_input_dir:
        full_path = os.path.join(excel_input_dir, filename)

        if filename.endswith(".xlsm"):
            xlsx_file = os.path.join(directory, filename.replace(".xlsm", ".xlsx"))
            convert_xlsm_to_xlsx(full_path, xlsx_file)
            os.remove(full_path)
            print(f"Исходный файл '{full_path}' удален.")

            pdf_path = os.path.join(export_dir, filename.replace(".xlsm", ".pdf"))
            excel_to_pdf(xlsx_file, pdf_path)
            pdf_files.append(pdf_path)

        elif filename.endswith((".xlsx", ".xls")):
            pdf_path = os.path.join(export_dir,
                                    filename.replace(".xlsx", ".pdf").replace(".xls", ".pdf"))
            excel_to_pdf(full_path, pdf_path)
            pdf_files.append(pdf_path)

    pdf_files.sort()

    # Переименовываем копии
    renamed_files = []
    for i, pdf_path in enumerate(pdf_files, start=1):
        base_name = os.path.basename(pdf_path)
        new_path = os.path.join(service_dir, f"{i:03d}_{base_name}")
        shutil.copy2(pdf_path, new_path)
        renamed_files.append(new_path)
        print(f"Файл переименован и перемещён: {pdf_path} -> {new_path}")

    # Итоговые объединения
    merge_pdfs(renamed_files, os.path.join(print_dir, "complete_merged.pdf"), mode='full')
    merge_pdfs(renamed_files, os.path.join(print_dir, "title_merged.pdf"), mode='title')
    merge_pdfs(renamed_files, os.path.join(print_dir, "no_title_merged.pdf"), mode='notitle')

def run():
    script_directory = os.path.dirname(os.path.abspath(__file__))
    process_files(script_directory)
