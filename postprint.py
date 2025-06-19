import os
import sys
from PyPDF2 import PdfReader, PdfWriter

def replace_first_page(source_pdf: str, new_first_page, output_pdf: str):
    """
    Заменяет первую страницу PDF-файла на заданную и сохраняет результат.
    """
    reader = PdfReader(source_pdf)
    writer = PdfWriter()

    writer.add_page(new_first_page)

    for page in reader.pages[1:]:
        writer.add_page(page)

    with open(output_pdf, 'wb') as f:
        writer.write(f)

def clear_output_directory(output_dir: str):
    """
    Удаляет все PDF-файлы из папки вывода.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        return

    for file in os.listdir(output_dir):
        if file.lower().endswith(".pdf"):
            os.remove(os.path.join(output_dir, file))

def process_directory_with_titles(service_dir: str, title_scan_pdf: str, output_dir: str):
    """
    Заменяет первую страницу во всех PDF-файлах в папке service_dir
    на соответствующую из title_scan_pdf, сохраняет в output_dir.
    """
    if not os.path.exists(title_scan_pdf):
        raise FileNotFoundError(f"Файл '{title_scan_pdf}' не найден. Убедитесь, что он существует в папке 'Print'.")

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
            print(f"⚠️ Нет страницы {i + 1} в {title_scan_pdf}, пропущен: {filename}")
            continue

        new_first = scanned_pages[i]
        new_name = filename[4:]  # удаляем префикс "001_"
        output_path = os.path.join(output_dir, new_name)

        replace_first_page(input_path, new_first, output_path)
        print(f"✅ Обновлён: {filename} -> {new_name}")

def run():
    """
    Точка входа при вызове как модуля.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))

    service_dir = os.path.join(base_dir, "Service")
    print_dir = os.path.join(base_dir, "Print")
    output_dir = os.path.join(base_dir, "Final")
    title_scan_pdf = os.path.join(print_dir, "title_scan.pdf")

    process_directory_with_titles(service_dir, title_scan_pdf, output_dir)
