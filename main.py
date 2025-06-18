import os
import win32com.client
import pythoncom
from PyPDF2 import PdfReader, PdfWriter

def convert_xlsm_to_xlsx(xlsm_file, xlsx_file):
    """
    Преобразует файл .xlsm в .xlsx.
    """
    try:
        # Инициализация COM
        pythoncom.CoInitialize()

        # Создаем объект Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Скрываем окно Excel

        # Открываем файл .xlsm
        try:
            workbook = excel.Workbooks.Open(xlsm_file)
        except Exception as e:
            print(f"Ошибка при открытии файла '{xlsm_file}': {e}")
            return

        # Сохраняем как .xlsx
        try:
            workbook.SaveAs(xlsx_file, FileFormat=51)  # 51 - это формат .xlsx
            print(f"Файл '{xlsm_file}' успешно преобразован в '{xlsx_file}'")
        except Exception as e:
            print(f"Ошибка при сохранении файла '{xlsm_file}' как '{xlsx_file}': {e}")
            return
        finally:
            # Закрываем книгу и выходим из Excel
            workbook.Close(SaveChanges=False)
            excel.Quit()

    except Exception as e:
        print(f"Ошибка при обработке файла '{xlsm_file}': {e}")
    finally:
        # Убедимся, что COM был освобожден
        pythoncom.CoUninitialize()

def excel_to_pdf(excel_file, pdf_file):
    """
    Преобразует Excel файл в PDF, с учетом префикса в названии файла для определения количества листов, которые нужно исключить.
    """
    try:
        # Инициализация COM
        pythoncom.CoInitialize()

        # Создаем объект Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Скрываем окно Excel

        # Открываем Excel файл
        try:
            workbook = excel.Workbooks.Open(excel_file)
        except Exception as e:
            print(f"Ошибка при открытии файла '{excel_file}': {e}")
            return

        # Получаем количество листов
        sheet_count = workbook.Sheets.Count

        # Определяем количество листов для исключения в зависимости от префикса файла
        filename = os.path.basename(excel_file)
        if filename.startswith("ВСИ"):
            exclude_count = 1
        elif filename.startswith("ВМ"):
            exclude_count = 2
        elif filename[0].isdigit():  # Проверяем, начинается ли имя файла с цифры
            exclude_count = 0  # Печатать все листы
        else:
            exclude_count = 0  # По умолчанию не исключаем листы

        # Определяем листы для печати
        sheets_to_print = list(range(1, sheet_count - exclude_count + 1))  # Индексы листов начинаются с 1

        # Проверяем, есть ли листы для печати
        if not sheets_to_print:
            print(f"Нет листов для печати в файле '{excel_file}'")
            workbook.Close(SaveChanges=False)
            excel.Quit()
            return

        # Сохраняем в PDF, указывая листы для печати
        try:
            workbook.ExportAsFixedFormat(
                Type=0,  # 0 соответствует PDF формату
                Filename=pdf_file,
                Quality=1,  # Стандартное качество
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

        # Закрываем книгу и выходим из Excel
        workbook.Close(SaveChanges=False)  # Не сохраняем изменения
        excel.Quit()

        print(f"Файл '{excel_file}' успешно преобразован в '{pdf_file}'")

    except Exception as e:
        print(f"Ошибка при обработке файла '{excel_file}': {e}")
    finally:
        # Убедимся, что COM был освобожден
        pythoncom.CoUninitialize()

def merge_pdfs(pdf_files, output_file):
    """
    Объединяет несколько PDF-файлов в один.
    """
    try:
        writer = PdfWriter()

        for pdf_file in pdf_files:
            reader = PdfReader(pdf_file)
            for page in reader.pages:
                writer.add_page(page)

        with open(output_file, "wb") as output_pdf:
            writer.write(output_pdf)

        print(f"Все PDF-файлы объединены в '{output_file}'")

    except Exception as e:
        print(f"Ошибка при объединении PDF-файлов: {e}")

def process_files(directory):
    """
    Обрабатывает все Excel и PDFM файлы в указанной директории.
    """
    pdf_files = []
    for filename in os.listdir(directory):
        if filename.endswith(".xlsm"):  # Преобразуем .xlsm в .xlsx
            xlsm_file = os.path.join(directory, filename)
            xlsx_file = os.path.join(directory, filename.replace(".xlsm", ".xlsx"))
            convert_xlsm_to_xlsx(xlsm_file, xlsx_file)
            # После преобразования удаляем исходный .xlsm файл
            os.remove(xlsm_file)
            print(f"Исходный файл '{xlsm_file}' удален.")
            # Обрабатываем новый .xlsx файл
            pdf_file = os.path.join(directory, filename.replace(".xlsm", ".pdf"))
            excel_to_pdf(xlsx_file, pdf_file)
        elif filename.endswith((".xlsx", ".xls")):  # Обрабатываем .xlsx и .xls
            excel_file = os.path.join(directory, filename)
            pdf_file = os.path.join(directory,
                                    filename.replace(".xlsx", ".pdf").replace(".xls", ".pdf"))
            excel_to_pdf(excel_file, pdf_file)
            pdf_files.append(pdf_file)


    pdf_files.sort()  # Сортируем файлы по имени
    for i, pdf_file in enumerate(pdf_files, start=1):
        if not os.path.exists(pdf_file):  # Проверка существования
            print(f"Файл '{pdf_file}' не найден, пропускаем...")
            continue

        new_name = os.path.join(directory, f"{i:03d}_{os.path.basename(pdf_file)}")
        os.rename(pdf_file, new_name)
        pdf_files[i - 1] = new_name
        print(f"Файл переименован: {pdf_file} -> {new_name}")


    merge_output = os.path.join(directory, "complete_merged.pdf")
    merge_pdfs(pdf_files, merge_output)

    merge_output = os.path.join(directory, "title_merged.pdf")
    merge_pdfs(pdf_files, merge_output)

    merge_output = os.path.join(directory, "no_title_merged.pdf")
    merge_pdfs(pdf_files, merge_output)


def process_scan(directory):
    """
    Обрабатывает все Excel и PDFM файлы в указанной директории.
    """
    pdf_files = []
    for filename in os.listdir(directory):
        if filename.startswith("title"):
            pdf_file = filename
            pdf_files.append(pdf_file)

    for i, pdf_file in enumerate(pdf_files, start=1):
        if not os.path.exists(pdf_file):  # Проверка существования
            print(f"Файл '{pdf_file}' не найден, пропускаем...")
            continue

        new_name = os.path.join(directory, f"{i:03d}_{os.path.basename(pdf_file)}")
        os.rename(pdf_file, new_name)
        pdf_files[i - 1] = new_name
        print(f"Файл переименован: {pdf_file} -> {new_name}")

    pdf_files.sort()  # Сортируем файлы по имени
    for i, pdf_file in enumerate(pdf_files, start=1):
        if not os.path.exists(pdf_file):  # Проверка существования
            print(f"Файл '{pdf_file}' не найден, пропускаем...")
            continue

        new_name = os.path.join(directory, f"{i:03d}_{os.path.basename(pdf_file)}")
        os.rename(pdf_file, new_name)
        pdf_files[i - 1] = new_name
        print(f"Файл переименован: {pdf_file} -> {new_name}")


    merge_output = os.path.join(directory, "complete_merged.pdf")
    merge_pdfs(pdf_files, merge_output)

    merge_output = os.path.join(directory, "title_merged.pdf")
    merge_pdfs(pdf_files, merge_output)

    merge_output = os.path.join(directory, "no_title_merged.pdf")
    merge_pdfs(pdf_files, merge_output)


# Получаем путь к директории, где находится скрипт
script_directory = os.path.dirname(os.path.abspath(__file__))

# Обрабатываем файлы в директории скрипта
process_files(script_directory)
