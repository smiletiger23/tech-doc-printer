import os
import pythoncom
import win32com.client
from PyPDF2 import PdfReader, PdfWriter
from typing import List

constants = win32com.client.constants

def remove_download_block(file_path: str):
    try:
        os.remove(file_path + ":Zone.Identifier")
        print(f"Блокировка удалена: {file_path}")
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"Ошибка снятия блокировки: {e}")

def init_excel_app():
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel  # automation security удалена как причина ошибок

def convert_xlsm_to_xlsx(xlsm_path: str, xlsx_path: str):
    excel = None
    workbook = None
    try:
        remove_download_block(xlsm_path)
        excel = init_excel_app()
        workbook = excel.Workbooks.Open(xlsm_path)
        workbook.SaveAs(xlsx_path, FileFormat=51)
        print(f"Преобразован: {xlsm_path} -> {xlsx_path}")
    except Exception as e:
        print(f"Ошибка при конвертации {xlsm_path}: {e}")
    finally:
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Quit()
                pythoncom.CoUninitialize()
            except:
                pass

def excel_to_pdf(excel_path: str, pdf_path: str):
    excel = None
    workbook = None
    temp_workbook = None
    try:
        remove_download_block(excel_path)
        excel = init_excel_app()
        workbook = excel.Workbooks.Open(excel_path)

        sheet_count = workbook.Sheets.Count
        filename = os.path.basename(excel_path)

        if filename.startswith("ВСИ"):
            exclude_last = 1
        elif filename.startswith("ВМ"):
            exclude_last = 2
        elif filename[0].isdigit():
            exclude_last = 0
        else:
            exclude_last = 0

        export_sheets = []
        for i in range(1, sheet_count - exclude_last + 1):
            sheet = workbook.Sheets(i)
            if sheet.Visible == -1:  # xlSheetVisible
                export_sheets.append(sheet)

        if not export_sheets:
            print(f"Нет видимых листов для экспорта в {excel_path}")
            return

        temp_workbook = excel.Workbooks.Add()
        while temp_workbook.Sheets.Count > 0:
            temp_workbook.Sheets(1).Delete()
        for sheet in reversed(export_sheets):
            sheet.Copy(Before=temp_workbook.Sheets(1))

        temp_workbook.ExportAsFixedFormat(
            Type=0,
            Filename=pdf_path,
            Quality=1,
            IncludeDocProperties=False,
            IgnorePrintAreas=False
        )

        print(f"Экспортирован в PDF (без {exclude_last} последних листов): {pdf_path}")

    except Exception as e:
        print(f"Ошибка при экспорте Excel в PDF: {e}")
    finally:
        if temp_workbook:
            try:
                temp_workbook.Close(SaveChanges=False)
            except:
                pass
        if workbook:
            try:
                workbook.Close(SaveChanges=False)
            except:
                pass
        if excel:
            try:
                excel.Quit()
                pythoncom.CoUninitialize()
            except:
                pass

def merge_all_pdfs(pdf_paths: List[str], output_path: str):
    writer = PdfWriter()
    for path in pdf_paths:
        reader = PdfReader(path)
        for page in reader.pages:
            writer.add_page(page)
    with open(output_path, 'wb') as f:
        writer.write(f)
    print(f"Создан: {output_path}")

def merge_titles(pdf_paths: List[str], output_path: str):
    writer = PdfWriter()
    for path in pdf_paths:
        reader = PdfReader(path)
        if reader.pages:
            writer.add_page(reader.pages[0])
    with open(output_path, 'wb') as f:
        writer.write(f)
    print(f"Создан: {output_path}")

def merge_without_titles(pdf_paths: List[str], output_path: str):
    writer = PdfWriter()
    for path in pdf_paths:
        reader = PdfReader(path)
        if len(reader.pages) > 1:
            for page in reader.pages[1:]:
                writer.add_page(page)
    with open(output_path, 'wb') as f:
        writer.write(f)
    print(f"Создан: {output_path}")

def process_files(directory: str):
    generated_pdfs = []

    for file in os.listdir(directory):
        full_path = os.path.join(directory, file)
        base, ext = os.path.splitext(file)

        if ext.lower() == ".xlsm":
            xlsx_path = os.path.join(directory, base + ".xlsx")
            convert_xlsm_to_xlsx(full_path, xlsx_path)
            os.remove(full_path)
            print(f"Удалён: {full_path}")
            pdf_path = os.path.join(directory, base + ".pdf")
            excel_to_pdf(xlsx_path, pdf_path)
            if os.path.exists(pdf_path):
                generated_pdfs.append(pdf_path)
            else:
                print(f"PDF не создан, файл пропущен: {pdf_path}")

        elif ext.lower() in (".xlsx", ".xls"):
            pdf_path = os.path.join(directory, base + ".pdf")
            excel_to_pdf(full_path, pdf_path)
            if os.path.exists(pdf_path):
                generated_pdfs.append(pdf_path)
            else:
                print(f"PDF не создан, файл пропущен: {pdf_path}")

    generated_pdfs.sort()
    renamed = []
    for i, path in enumerate(generated_pdfs, 1):
        filename = os.path.basename(path)
        new_name = f"{i:03d}_{filename}"
        new_path = os.path.join(directory, new_name)
        try:
            os.rename(path, new_path)
            renamed.append(new_path)
            print(f"Переименован: {path} -> {new_path}")
        except FileNotFoundError:
            print(f"Файл не найден для переименования: {path}")

    merge_all_pdfs(renamed, os.path.join(directory, "complete_merged.pdf"))
    merge_titles(renamed, os.path.join(directory, "title_merged.pdf"))
    merge_without_titles(renamed, os.path.join(directory, "no_title_merged.pdf"))

if __name__ == "__main__":
    directory = os.path.dirname(os.path.abspath(__file__))
    process_files(directory)
