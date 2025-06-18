import os
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

def process_directory_with_titles(directory: str, title_scan_pdf: str):
    """
    Заменяет первую страницу во всех PDF-файлах в директории
    на соответствующий лист из title_scan_pdf.
    """
    title_reader = PdfReader(title_scan_pdf)
    scanned_pages = title_reader.pages

    numbered_pdfs = sorted([
        f for f in os.listdir(directory)
        if f.lower().endswith(".pdf") and f[:3].isdigit() and "_" in f
    ])

    for i, filename in enumerate(numbered_pdfs):
        file_path = os.path.join(directory, filename)
        if i >= len(scanned_pages):
            print(f"Нет страницы {i+1} в {title_scan_pdf}, пропущен: {filename}")
            continue

        new_first = scanned_pages[i]
        new_name = filename[4:]  # удаляем префикс "001_"
        output_path = os.path.join(directory, new_name)

        replace_first_page(file_path, new_first, output_path)
        print(f"Обновлён: {filename} -> {new_name}")

if __name__ == "__main__":
    # Укажи путь к директории и скану титулов
    directory = os.path.dirname(os.path.abspath(__file__))
    title_scan_pdf = os.path.join(directory, "title_scan.pdf")

    process_directory_with_titles(directory, title_scan_pdf)
