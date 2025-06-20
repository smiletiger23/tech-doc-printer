import tkinter as tk
from tkinter import scrolledtext, messagebox
import os
import sys
import io
import shutil
import time
import pythoncom
import win32com.client
from PyPDF2 import PdfReader, PdfWriter, errors as pypdf_errors
import gc
import threading
import queue
import re

# --- CAPTURE ORIGINAL STDOUT VERY EARLY ---
_original_stdout = sys.__stdout__

# --- Глобальные переменные для GUI и очереди ---
log_queue = queue.Queue()
gui_response_queue = queue.Queue()
log_output = None
root = None


# --- КОНФИГУРАЦИЯ ---
class Config:
    """Класс для хранения путей к директориям и других глобальных настроек."""
    BASE_DIR = None  # Будет инициализирован при запуске
    EXCEL_INPUT_DIR = None
    PRINT_DIR = None
    EXPORT_DIR = None
    SERVICE_DIR = None
    FINAL_OUTPUT_DIR = None
    TITLE_SCAN_PDF = None

    @staticmethod
    def initialize_paths():
        Config.BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
            os.path.abspath(__file__))
        Config.EXCEL_INPUT_DIR = os.path.join(Config.BASE_DIR, "Excel")
        Config.PRINT_DIR = os.path.join(Config.BASE_DIR, "Print")
        Config.EXPORT_DIR = os.path.join(Config.BASE_DIR, "NotSignedExport")
        Config.SERVICE_DIR = os.path.join(Config.BASE_DIR, "Service")
        Config.FINAL_OUTPUT_DIR = os.path.join(Config.BASE_DIR, "Final")
        Config.TITLE_SCAN_PDF = os.path.join(Config.PRINT_DIR, "title_scan.pdf")


# --- 1. Настройка и утилиты (Configuration & Utilities) ---

class RedirectText(io.TextIOBase):
    """
    Класс для перенаправления вывода стандартного потока (stdout)
    в виджет ScrolledText Tkinter и файл.
    """

    def __init__(self, message_queue, original_stdout_ref):
        self.message_queue = message_queue
        self.original_stdout = original_stdout_ref

    def write(self, string):
        self.message_queue.put(string)
        if self.original_stdout:
            self.original_stdout.write(string)

        try:
            with open("app_log.txt", "a", encoding="utf-8") as f_log:
                f_log.write(string)
        except Exception as e:
            if self.original_stdout:
                self.original_stdout.write(f"[{time.strftime('%H:%M:%S')}] [ERROR] Ошибка записи в файл лога: {e}\n")

    def flush(self):
        if self.original_stdout:
            self.original_stdout.flush()


def log(msg, level="INFO"):
    """Централизованная функция для логирования сообщений."""
    timestamp = time.strftime("%H:%M:%S")
    formatted_msg = f"[{timestamp}] [{level}] {msg}\n"
    print(formatted_msg)


def ensure_and_clear_folder(folder_path, is_output_folder=False):
    """
    Проверяет существование папки, создает её при необходимости и очищает.
    Для выходных папок (Final) логирование может быть немного другим.
    """
    log_prefix = "Выходная " if is_output_folder else ""
    log(f"Очистка {log_prefix}папки: {folder_path}")

    if not os.path.exists(folder_path):
        log(f"{log_prefix}Папка не найдена: {folder_path}. Создаю.", level="WARNING")
        os.makedirs(folder_path)
        return

    for item_name in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item_name)
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path):
                os.remove(item_path)
                log(f"Удален файл: {item_name}")
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
                log(f"Удалена директория: {item_name}")
        except OSError as e:
            log(f"Ошибка при удалении '{item_path}': {e} (Код ошибки: {e.winerror})", level="ERROR")
            raise


def is_pdf_valid(file_path):
    """Проверяет, является ли PDF файл действительным и содержит ли страницы."""
    if not os.path.exists(file_path):
        log(f"PDF файл не существует: {os.path.basename(file_path)}", level="ERROR")
        return False, 0
    if os.path.getsize(file_path) == 0:
        log(f"PDF файл пуст (размер 0 байт): {os.path.basename(file_path)}", level="WARNING")
        return False, 0
    try:
        with open(file_path, "rb") as f:
            reader = PdfReader(f)
            num_pages = len(reader.pages)
            if num_pages == 0:
                log(f"PDF файл не содержит страниц: {os.path.basename(file_path)}", level="WARNING")
                return False, 0
            return True, num_pages
    except pypdf_errors.PdfReadError as e:
        log(f"Поврежденный PDF файл: {os.path.basename(file_path)} - {e}", level="ERROR")
        return False, 0
    except Exception as e:
        log(f"Непредвиденная ошибка при проверке PDF: {os.path.basename(file_path)} - {e}", level="ERROR")
        return False, 0


def is_excel_installed():
    """Проверяет, установлен ли Microsoft Excel на системе."""
    try:
        obj = win32com.client.Dispatch("Excel.Application")
        obj.Quit()
        del obj
        gc.collect()
        return True
    except pythoncom.com_error:
        return False
    except Exception:
        return False


def close_excel_app(excel_app):
    """Корректно закрывает приложение Excel и освобождает COM-объекты."""
    if excel_app:
        try:
            excel_app.DisplayAlerts = True
            excel_app.Quit()
            log("Excel приложение закрыто.")
        except Exception as e:
            log(f"Ошибка при закрытии Excel приложения: {e}", level="WARNING")
        finally:
            del excel_app
            gc.collect()
            log("COM-объект Excel освобожден и память очищена.")
    try:
        pythoncom.CoUninitialize()
        log("COM-библиотека деинициализирована.")
    except Exception as e:
        log(f"Ошибка при деинициализации COM-библиотеки: {e}", level="WARNING")


# --- 2. Функции обработки Excel (Excel Processing Functions) ---

def get_exclude_count(filename):
    """
    Определяет количество листов Excel для исключения из печати
    на основе префикса имени файла.
    """
    filename_upper = filename.upper()
    if filename_upper.startswith("ВСИ"):
        return 1
    elif filename_upper.startswith("ВМ"):
        return 2
    elif filename_upper[0].isdigit():
        return 0
    return 0


def _open_excel_workbook(excel_app, file_path):
    """Вспомогательная функция для безопасного открытия рабочей книги Excel."""
    try:
        wb = excel_app.Workbooks.Open(file_path)
        log(f"Открыта рабочая книга Excel: {os.path.basename(file_path)}")
        return wb
    except Exception as e:
        log(f"Ошибка при открытии Excel файла '{os.path.basename(file_path)}': {e}", level="ERROR")
        return None


def _close_excel_workbook(workbook):
    """Вспомогательная функция для безопасного закрытия рабочей книги Excel."""
    if workbook:
        try:
            workbook.Close(False)
            time.sleep(0.05)
            log(f"Закрыта рабочая книга Excel.")
        except Exception as e:
            log(f"Ошибка при закрытии рабочей книги Excel: {e}", level="WARNING")


def convert_xlsm_to_xlsx(excel_app, full_path, base_filename):
    """Конвертирует файл .xlsm в .xlsx и удаляет исходный .xlsm."""
    temp_xlsx_path = os.path.join(Config.EXCEL_INPUT_DIR, base_filename + ".xlsx")
    wb = None
    try:
        log(f"Попытка конвертации XLSM: {os.path.basename(full_path)}")
        wb = _open_excel_workbook(excel_app, full_path)
        if not wb:
            return None

        wb.SaveAs(temp_xlsx_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook (xlsx)
        log(f"Преобразован: {os.path.basename(full_path)} → {os.path.basename(temp_xlsx_path)}")

        try:
            os.remove(full_path)
            log(f"Исходный XLSM файл удален: {os.path.basename(full_path)}")
        except OSError as e:
            log(f"Не удалось удалить исходный XLSM файл '{os.path.basename(full_path)}': {e} (Код ошибки: {e.winerror})",
                level="WARNING")

        return temp_xlsx_path
    except Exception as e:
        log(f"❗ Общая ошибка при преобразовании {os.path.basename(full_path)} в XLSX: {e}", level="ERROR")
        return None
    finally:
        _close_excel_workbook(wb)


def export_excel_to_pdf(excel_app, full_path, base_filename, file_number):
    """
    Экспортирует листы Excel в PDF, исключая указанное количество последних листов.
    Присваивает файлу порядковый номер.
    """
    wb = None
    pdf_path = None
    expected_pages = 0
    try:
        log(f"Начинаю экспорт: {os.path.basename(full_path)}")
        wb = _open_excel_workbook(excel_app, full_path)
        if not wb:
            return None, 0

        sheet_count = wb.Sheets.Count
        exclude_count = get_exclude_count(base_filename)

        expected_pages = sheet_count - exclude_count

        if expected_pages <= 0:
            log(f"Пропущен файл '{os.path.basename(full_path)}': после исключения {exclude_count} листов, осталось {expected_pages} листов для экспорта.",
                level="WARNING")
            return None, 0

        pdf_filename = f"{file_number:03d}_{base_filename}.pdf"
        pdf_path = os.path.join(Config.SERVICE_DIR, pdf_filename)

        log(f"Экспорт листов 1-{expected_pages} из '{os.path.basename(full_path)}' в PDF: {os.path.basename(pdf_path)}")
        wb.ExportAsFixedFormat(
            Type=0,  # xlTypePDF
            Filename=pdf_path,
            Quality=0,  # xlQualityStandard
            IncludeDocProperties=False,
            IgnorePrintAreas=False,
            From=1,
            To=expected_pages
        )
        log(f"Сохранён: {os.path.basename(pdf_path)}")

        is_valid, actual_pages = is_pdf_valid(pdf_path)
        if not is_valid:
            log(f"Созданный PDF '{os.path.basename(pdf_path)}' не прошел валидацию. Возможно, он поврежден или пуст.",
                level="ERROR")
            try:
                os.remove(pdf_path)
                log(f"Удален невалидный PDF: {os.path.basename(pdf_path)}")
            except OSError as e:
                log(f"Не удалось удалить невалидный PDF '{os.path.basename(pdf_path)}': {e} (Код ошибки: {e.winerror})",
                    level="WARNING")
            return None, 0

        if actual_pages != expected_pages:
            log(f"❗ НЕСООТВЕТСТВИЕ СТРАНИЦ в '{os.path.basename(pdf_path)}': Ожидалось {expected_pages}, фактически {actual_pages}.",
                level="ERROR")
            try:
                os.remove(pdf_path)
                log(f"Удален PDF с несоответствием страниц: {os.path.basename(pdf_path)}")
            except OSError as e:
                log(f"Не удалось удалить PDF с несоответствием страниц '{os.path.basename(pdf_path)}': {e} (Код ошибки: {e.winerror})",
                    level="WARNING")
            return None, 0
        else:
            log(f"✅ Проверка страниц: Ожидалось {expected_pages}, фактически {actual_pages}. Совпадает.", level="INFO")

        return pdf_path, actual_pages
    except Exception as e:
        log(f"❗ Общая ошибка при экспорте '{os.path.basename(full_path)}' в PDF: {e}", level="ERROR")
        return None, 0
    finally:
        _close_excel_workbook(wb)


# --- 3. Функции обработки PDF (PDF Processing Functions) ---

def merge_pdfs(pdf_files, output_file, mode='full'):
    """Объединяет несколько PDF-файлов в один."""
    writer = PdfWriter()
    successful_merges = 0
    total_expected_pages = 0

    log(f"Начинаю объединение PDF файлов в режиме '{mode}' для '{os.path.basename(output_file)}'...")

    try:
        for pdf_file in pdf_files:
            is_valid, num_pages = is_pdf_valid(pdf_file)
            if not is_valid:
                log(f"Невалидный PDF файл, пропущен при объединении: {os.path.basename(pdf_file)}", level="WARNING")
                continue

            try:
                with open(pdf_file, "rb") as f:
                    reader = PdfReader(f)

                    pages_to_add = []
                    current_file_expected_pages = 0

                    if mode == 'full':
                        pages_to_add = reader.pages
                        current_file_expected_pages = num_pages
                    elif mode == 'title':
                        pages_to_add = [reader.pages[0]] if reader.pages else []
                        current_file_expected_pages = 1 if num_pages > 0 else 0
                    elif mode == 'notitle':
                        pages_to_add = reader.pages[1:]
                        current_file_expected_pages = max(0, num_pages - 1)
                    else:
                        raise ValueError(f"Неизвестный режим объединения PDF: {mode}")

                    for page in pages_to_add:
                        writer.add_page(page)

                    log(f"Добавлен файл '{os.path.basename(pdf_file)}' ({mode} режим). Ожидалось страниц: {current_file_expected_pages}")
                    total_expected_pages += current_file_expected_pages
                    successful_merges += 1
            except pypdf_errors.PdfReadError as e:
                log(f"Ошибка чтения PDF файла '{os.path.basename(pdf_file)}', пропущен: {e}", level="ERROR")
            except Exception as e:
                log(f"Непредвиденная ошибка при обработке '{os.path.basename(pdf_file)}', пропущен: {e}", level="ERROR")

        if writer.pages and successful_merges > 0:
            with open(output_file, "wb") as f:
                writer.write(f)

            is_output_valid, actual_output_pages = is_pdf_valid(output_file)
            if not is_output_valid:
                log(f"❗ Объединенный PDF '{os.path.basename(output_file)}' не прошел валидацию.", level="ERROR")
                return False

            if actual_output_pages != total_expected_pages:
                log(f"❗ НЕСООТВЕТСТВИЕ СТРАНИЦ в объединенном '{os.path.basename(output_file)}': Ожидалось {total_expected_pages}, фактически {actual_output_pages}.",
                    level="ERROR")
                return False
            else:
                log(f"✅ Проверка страниц объединенного файла: Ожидалось {total_expected_pages}, фактически {actual_output_pages}. Совпадает.",
                    level="INFO")
                log(f"✅ Успешно собран файл: {os.path.basename(output_file)}")
            return True
        else:
            log(f"Нечего объединять или все исходные PDF были невалидны. Файл '{os.path.basename(output_file)}' не создан.",
                level="WARNING")
            return False

    except Exception as e:
        log(f"❗ Общая ошибка при объединении PDF в '{output_file}': {e}", level="ERROR")
        return False


def replace_first_page(source_pdf_path, new_first_page_object, output_pdf_path, original_pages_count):
    """Заменяет первую страницу PDF-файла новой страницей."""
    try:
        is_valid, num_pages = is_pdf_valid(source_pdf_path)
        if not is_valid:
            log(f"Исходный PDF файл невалиден, пропущен при замене титульника: {os.path.basename(source_pdf_path)}",
                level="ERROR")
            return False

        reader = PdfReader(source_pdf_path)
        writer = PdfWriter()

        writer.add_page(new_first_page_object)

        if len(reader.pages) > 1:
            for i in range(1, len(reader.pages)):
                writer.add_page(reader.pages[i])

        with open(output_pdf_path, 'wb') as f:
            writer.write(f)

        is_output_valid, actual_output_pages = is_pdf_valid(output_pdf_path)
        if not is_output_valid:
            log(f"❗ Финальный PDF '{os.path.basename(output_pdf_path)}' не прошел валидацию после замены титульника.",
                level="ERROR")
            return False

        expected_output_pages = 1 + max(0, original_pages_count - 1)

        if actual_output_pages != expected_output_pages:
            log(f"❗ НЕСООТВЕТСТВИЕ СТРАНИЦ в финальном '{os.path.basename(output_pdf_path)}': Ожидалось {expected_output_pages}, фактически {actual_output_pages}.",
                level="ERROR")
            return False
        else:
            log(f"✅ Проверка страниц после замены титульника: Ожидалось {expected_output_pages}, фактически {actual_output_pages}. Совпадает.",
                level="INFO")
            log(f"✅ Обновлён файл: {os.path.basename(output_pdf_path)}")
        return True
    except pypdf_errors.PdfReadError as e:
        log(f"Ошибка чтения PDF файла '{os.path.basename(source_pdf_path)}' при замене первой страницы: {e}",
            level="ERROR")
        return False
    except Exception as e:
        log(f"❗ Общая ошибка при замене первой страницы в '{os.path.basename(source_pdf_path)}': {e}", level="ERROR")
        return False


# --- 4. Основные рабочие процессы (Core Workflow Functions) ---

def process_preprint_task():
    """Задача предварительной обработки для запуска в отдельном потоке."""
    try:
        process_preprint()
    except Exception as e:
        log(f"Критическая ошибка в процессе подготовки: {e}", level="CRITICAL")
    finally:
        log("✅ Завершено.")
        log_queue.put("PROCESS_COMPLETE")


def process_postprint_task():
    """Задача постобработки для запуска в отдельном потоке."""
    try:
        process_postprint()
    except Exception as e:
        log(f"Критическая ошибка в процессе замены титульников: {e}", level="CRITICAL")
    finally:
        log("✅ Завершено.")
        log_queue.put("PROCESS_COMPLETE")


def process_preprint():
    """
    Основная логика предварительной обработки:
    конвертация XLSM в XLSX, экспорт XLSX в PDF и объединение PDF.
    """
    # Инициализация директорий
    # Папка EXCEL_INPUT_DIR не должна очищаться
    dirs_to_create_and_clear = [
        Config.PRINT_DIR,
        Config.EXPORT_DIR,
        Config.SERVICE_DIR
    ]

    # Убедимся, что EXCEL_INPUT_DIR существует, но не очищаем её
    if not os.path.exists(Config.EXCEL_INPUT_DIR):
        try:
            log(f"Папка Excel не найдена: {Config.EXCEL_INPUT_DIR}. Создаю.", level="WARNING")
            os.makedirs(Config.EXCEL_INPUT_DIR)
        except Exception as e:
            log(f"Ошибка при создании папки Excel: {e}. Процесс остановлен.", level="CRITICAL")
            return

    for d in dirs_to_create_and_clear:
        try:
            ensure_and_clear_folder(d)
        except Exception:
            log(f"Ошибка при подготовке папки {os.path.basename(d)}. Процесс остановлен.", level="CRITICAL")
            return

    excel_app = None
    summary = {
        "excel_files_found": 0,
        "excel_files_processed_to_pdf": 0,
        "total_excel_pages_expected": 0,
        "total_pdf_exported_pages": 0,
        "pdf_export_errors": 0,
        "merge_success_complete": False,
        "merge_success_title": False,
        "merge_success_notitle": False,
        "total_merged_pages_complete": 0,
        "total_merged_pages_title": 0,
        "total_merged_pages_notitle": 0,
    }

    try:
        pythoncom.CoInitialize()

        if not is_excel_installed():
            log("Microsoft Excel не установлен или не найден. Установите Excel и попробуйте снова.", level="CRITICAL")
            return

        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.DisplayAlerts = False
        excel_app.Visible = False

        excel_files_to_process = sorted([
            f for f in os.listdir(Config.EXCEL_INPUT_DIR)
            if f.lower().endswith((".xlsx", ".xlsm", ".xls"))
        ])
        summary["excel_files_found"] = len(excel_files_to_process)

        if not excel_files_to_process:
            log("В папке 'Excel' не найдено файлов для обработки.", level="INFO")
            return

        processed_pdf_paths = []
        pdf_pages_info = {}
        file_counter = 1

        log("\n--- Начало экспорта Excel в PDF ---")
        for filename in excel_files_to_process:
            full_path = os.path.join(Config.EXCEL_INPUT_DIR, filename)
            base_filename, ext = os.path.splitext(filename)
            current_file_path = full_path

            if ext.lower() == ".xlsm":
                converted_path = convert_xlsm_to_xlsx(excel_app, full_path, base_filename)
                if converted_path:
                    current_file_path = converted_path
                    ext = ".xlsx"
                else:
                    log(f"Пропущен файл '{filename}' из-за ошибки конвертации.", level="ERROR")
                    summary["pdf_export_errors"] += 1
                    continue

            if ext.lower() in [".xlsx", ".xls"]:
                pdf_path, pages_count = export_excel_to_pdf(excel_app, current_file_path,
                                                            os.path.splitext(os.path.basename(current_file_path))[0],
                                                            file_counter)

                if pdf_path:
                    processed_pdf_paths.append(pdf_path)
                    pdf_pages_info[pdf_path] = pages_count
                    summary["total_excel_pages_expected"] += pages_count
                    summary["total_pdf_exported_pages"] += pages_count
                    summary["excel_files_processed_to_pdf"] += 1
                    file_counter += 1

                    try:
                        dest_filename = os.path.basename(pdf_path)[4:]
                        shutil.copy(pdf_path, os.path.join(Config.EXPORT_DIR, dest_filename))
                        log(f"Скопирован в NotSignedExport: {dest_filename}")
                    except Exception as e:
                        log(f"Ошибка копирования '{os.path.basename(pdf_path)}' в NotSignedExport: {e}", level="ERROR")
                else:
                    summary["pdf_export_errors"] += 1
            else:
                log(f"Пропущен файл с неподдерживаемым расширением: {filename}", level="WARNING")

        if not processed_pdf_paths:
            log("Нет успешно обработанных PDF файлов для объединения.", level="WARNING")
            return

        service_pdfs = sorted(
            [f for f in os.listdir(Config.SERVICE_DIR) if f.lower().endswith(".pdf") and f[:3].isdigit()],
            key=lambda x: int(x[:3])
        )
        service_pdfs_full_paths = [os.path.join(Config.SERVICE_DIR, f) for f in service_pdfs]

        log("\nНачинаю объединение PDF файлов...")

        total_pages_for_full_merge = sum(pdf_pages_info.values())
        total_pages_for_title_merge = len(processed_pdf_paths)
        total_pages_for_notitle_merge = total_pages_for_full_merge - total_pages_for_title_merge

        if service_pdfs_full_paths:
            log(f"\n--- Объединение: complete_merged.pdf (ожидается страниц: {total_pages_for_full_merge}) ---")
            if merge_pdfs(service_pdfs_full_paths, os.path.join(Config.PRINT_DIR, "complete_merged.pdf"), mode='full'):
                summary["merge_success_complete"] = True
                _, summary["total_merged_pages_complete"] = is_pdf_valid(
                    os.path.join(Config.PRINT_DIR, "complete_merged.pdf"))

            log(f"\n--- Объединение: title_merged.pdf (ожидается страниц: {total_pages_for_title_merge}) ---")
            if merge_pdfs(service_pdfs_full_paths, os.path.join(Config.PRINT_DIR, "title_merged.pdf"), mode='title'):
                summary["merge_success_title"] = True
                _, summary["total_merged_pages_title"] = is_pdf_valid(
                    os.path.join(Config.PRINT_DIR, "title_merged.pdf"))

            log(f"\n--- Объединение: no_title_merged.pdf (ожидается страниц: {total_pages_for_notitle_merge}) ---")
            if merge_pdfs(service_pdfs_full_paths, os.path.join(Config.PRINT_DIR, "no_title_merged.pdf"),
                          mode='notitle'):
                summary["merge_success_notitle"] = True
                _, summary["total_merged_pages_notitle"] = is_pdf_valid(
                    os.path.join(Config.PRINT_DIR, "no_title_merged.pdf"))
        else:
            log("Нет валидных PDF файлов в папке Service для объединения.", level="WARNING")

    except pythoncom.com_error as e:
        log(f"Ошибка COM-объекта (возможно, Excel не установлен или не отвечает): {e}", level="CRITICAL")
    except Exception as e:
        log(f"❗ Общая ошибка в process_preprint: {e}", level="CRITICAL")
    finally:
        close_excel_app(excel_app)

        # --- Сводка по завершении процесса ---
        log("\n--- СВОДКА ПРОЦЕССА 'ПОДГОТОВИТЬ К ПЕЧАТИ' ---")
        log(f"Найдено Excel файлов: {summary['excel_files_found']}")
        log(f"Успешно преобразовано и экспортировано в PDF: {summary['excel_files_processed_to_pdf']} файлов")
        log(f"Ошибок при экспорте Excel в PDF: {summary['pdf_export_errors']}")
        log(f"Общее ожидаемое кол-во страниц из Excel: {summary['total_excel_pages_expected']}")
        log(f"Общее кол-во страниц в экспортированных PDF (Service): {summary['total_pdf_exported_pages']}")

        log("\nРезультаты объединения PDF:")
        log(f"  complete_merged.pdf: {'✅ Успешно' if summary['merge_success_complete'] else '❌ Ошибка'}. Страниц: {summary['total_merged_pages_complete']}/{total_pages_for_full_merge if 'total_pages_for_full_merge' in locals() else 'N/A'} (фактически/ожидалось)")
        log(f"  title_merged.pdf: {'✅ Успешно' if summary['merge_success_title'] else '❌ Ошибка'}. Страниц: {summary['total_merged_pages_title']}/{total_pages_for_title_merge if 'total_pages_for_title_merge' in locals() else 'N/A'} (фактически/ожидалось)")
        log(f"  no_title_merged.pdf: {'✅ Успешно' if summary['merge_success_notitle'] else '❌ Ошибка'}. Страниц: {summary['total_merged_pages_notitle']}/{total_pages_for_notitle_merge if 'total_pages_for_notitle_merge' in locals() else 'N/A'} (фактически/ожидалось)")
        log("--- КОНЕЦ СВОДКИ ---")


def process_postprint():
    """
    Основная логика постобработки:
    замена титульных страниц в PDF-файлах отсканированными титульниками.
    """
    log("Начало этапа 'Заменить титульники'.")

    summary = {
        "title_scan_pages": 0,
        "numbered_pdfs_found": 0,
        "files_processed_successfully": 0,
        "files_skipped_no_scan_page": 0,
        "files_skipped_invalid_original": 0,
        "total_errors_during_replacement": 0
    }

    is_scan_valid, num_scanned_pages = is_pdf_valid(Config.TITLE_SCAN_PDF)
    summary["title_scan_pages"] = num_scanned_pages

    if not is_scan_valid:
        log(f"❗ Файл 'title_scan.pdf' не найден или невалиден в папке '{Config.PRINT_DIR}'. Пожалуйста, отсканируйте и положите его туда.",
            level="ERROR")
        log_queue.put(
            f"ERROR: Файл 'title_scan.pdf' не найден или невалиден в папке '{Config.PRINT_DIR}'. Пожалуйста, отсканируйте и положите его туда.")
        return

    title_scan_file_handle = None
    try:
        title_scan_file_handle = open(Config.TITLE_SCAN_PDF, "rb")
        title_reader = PdfReader(title_scan_file_handle)
        scanned_pages_objects = title_reader.pages
        log(f"Загружен 'title_scan.pdf', количество страниц: {num_scanned_pages}")

        ensure_and_clear_folder(Config.FINAL_OUTPUT_DIR, is_output_folder=True)

        numbered_pdfs = sorted([
            f for f in os.listdir(Config.SERVICE_DIR)
            if f.lower().endswith(".pdf") and f[:3].isdigit() and "_" in f
        ], key=lambda x: int(x[:3]))
        summary["numbered_pdfs_found"] = len(numbered_pdfs)

        if not numbered_pdfs:
            log("В папке 'Service' не найдено пронумерованных PDF файлов для обработки. Запустите 'Подготовить к печати' сначала.",
                level="WARNING")
            log_queue.put(
                "WARNING: В папке 'Service' не найдено пронумерованных PDF файлов для обработки. Запустите 'Подготовить к печати' сначала.")
            return

        if num_scanned_pages != len(numbered_pdfs):
            msg = (
                f"Количество отсканированных титульных страниц ({num_scanned_pages}) "
                f"не соответствует количеству файлов для замены ({len(numbered_pdfs)}).\n\n"
                "Вы хотите продолжить? (Несоответствующие файлы будут пропущены)"
            )
            log(f"ПРЕДУПРЕЖДЕНИЕ: {msg.replace('\n', ' ')}", level="WARNING")

            log_queue.put({"type": "ask_user_yn", "title": "Несоответствие количества титульников", "message": msg})

            user_response_str = gui_response_queue.get()

            if user_response_str != "ok":
                log("Пользователь отменил операцию из-за несоответствия количества.", level="INFO")
                return

        log("\n--- Начало замены титульников ---")

        for i, filename in enumerate(numbered_pdfs):
            input_path = os.path.join(Config.SERVICE_DIR, filename)

            is_original_valid, original_pages_count = is_pdf_valid(input_path)
            if not is_original_valid:
                log(f"Исходный PDF '{os.path.basename(input_path)}' невалиден, пропущен при замене титульника.",
                    level="ERROR")
                summary["files_skipped_invalid_original"] += 1
                summary["total_errors_during_replacement"] += 1
                continue

            if i >= num_scanned_pages:
                log(f"⚠️ Для файла '{filename}' отсутствует соответствующая страница в сканированном 'title_scan.pdf'. Пропущен.",
                    level="WARNING")
                summary["files_skipped_no_scan_page"] += 1
                summary["total_errors_during_replacement"] += 1
                continue

            new_name = filename[4:]
            output_path = os.path.join(Config.FINAL_OUTPUT_DIR, new_name)

            if replace_first_page(input_path, scanned_pages_objects[i], output_path, original_pages_count):
                summary["files_processed_successfully"] += 1
            else:
                summary["total_errors_during_replacement"] += 1

    except pypdf_errors.PdfReadError as e:
        log(f"Ошибка чтения файла 'title_scan.pdf': {e}", level="ERROR")
    except Exception as e:
        log(f"Непредвиденная ошибка в process_postprint: {e}", level="CRITICAL")
    finally:
        if title_scan_file_handle:
            title_scan_file_handle.close()
            log("Файл 'title_scan.pdf' закрыт.")

    log("\n--- СВОДКА ПРОЦЕССА 'ЗАМЕНИТЬ ТИТУЛЬНИКИ' ---")
    log(f"Отсканировано страниц титульников (title_scan.pdf): {summary['title_scan_pages']}")
    log(f"Найдено PDF файлов в Service для обработки: {summary['numbered_pdfs_found']}")
    log(f"Успешно заменено титульников: {summary['files_processed_successfully']}")
    log(f"Пропущено файлов (нет соответствующей сканированной страницы): {summary['files_skipped_no_scan_page']}")
    log(f"Пропущено файлов (оригинальный PDF невалиден): {summary['files_skipped_invalid_original']}")
    log(f"Всего ошибок/пропусков во время замены: {summary['total_errors_during_replacement']}")

    if summary["total_errors_during_replacement"] == 0:
        log("✅ Замена титульников завершена без ошибок.", level="INFO")
    else:
        log("❗ Замена титульников завершена с ошибками. Проверьте лог.", level="ERROR")
    log("--- КОНЕЦ СВОДКИ ---")


# --- 5. Функции GUI (GUI Callbacks & Setup) ---

def start_process_thread(process_func):
    """Запускает указанную функцию обработки в отдельном потоке."""
    if log_output:
        log_output.delete(1.0, tk.END)

    toggle_buttons_state(tk.DISABLED)

    thread = threading.Thread(target=process_func)
    thread.daemon = True
    thread.start()

    if root:
        root.after(100, update_log_display)


def handle_messagebox_response(response):
    """Колбэк функция, которая вызывается после закрытия messagebox."""
    gui_response_queue.put(response)


def update_log_display():
    """Периодически проверяет очередь сообщений и обновляет виджет лога."""
    while True:
        try:
            message = log_queue.get(block=False, timeout=0.01)

            if message == "PROCESS_COMPLETE":
                toggle_buttons_state(tk.NORMAL)
                if root:
                    root.after(100, update_log_display)
                return

            if isinstance(message, dict) and message.get("type") == "ask_user_yn":
                title = message["title"]
                msg = message["message"]

                root.after(0, lambda t=title, m=msg:
                handle_messagebox_response(messagebox.showwarning(t, m, type=messagebox.OKCANCEL)))
                continue

            if log_output:
                log_output.insert(tk.END, message)
                log_output.see(tk.END)

            # Использование f-строк для более чистого определения ошибок
            error_patterns = {
                "Ошибка Excel: Открытие файла": "Не удалось открыть файл Excel",
                "Ошибка Excel: Конвертация": "Не удалось преобразовать",
                "Ошибка Excel: Экспорт в PDF": "Не удалось экспортировать",
                "Ошибка PDF: Объединение": "Не удалось объединить PDF",
                "Ошибка PDF: Чтение": "Не удалось прочитать PDF файл",
                "Ошибка PDF: Замена титульника": "Не удалось заменить первую страницу",
                "Ошибка файловой системы: Директория": "Не удалось создать директорию",
                "Ошибка файловой системы: Удаление": "Ошибка при удалении",
                "Ошибка: Excel не найден": "Microsoft Excel не установлен",
                "Ошибка (CRITICAL)": "Критическая ошибка"
            }

            match = re.match(r'\[\d{2}:\d{2}:\d{2}\]\s\[(ERROR|CRITICAL)\]\s(.+)', message)
            if match:
                level = match.group(1)
                error_msg_content = match.group(2).strip()

                detected_title = f"Ошибка ({level})"
                for title_key, pattern in error_patterns.items():
                    if pattern in error_msg_content:
                        detected_title = title_key
                        break

                if root:
                    root.after(0, lambda t=detected_title, m=error_msg_content: messagebox.showerror(t, m))

        except queue.Empty:
            break
        except Exception as e:
            _original_stdout.write(f"[{time.strftime('%H:%M:%S')}] [CRITICAL] Ошибка в update_log_display: {e}\n")
            break

    if root:
        root.after(100, update_log_display)


def toggle_buttons_state(state):
    """Включает или выключает состояние кнопок GUI."""
    if btn_check_excel and btn_preprint and btn_postprint:
        for btn in [btn_check_excel, btn_preprint, btn_postprint]:
            btn.config(state=state)


def check_files_in_excel():
    """Проверяет наличие файлов Excel в соответствующей папке и выводит их в лог."""
    if log_output:
        log_output.delete(1.0, tk.END)

    # ensure_and_clear_folder можно использовать только для очистки, не для проверки наличия файлов.
    # Поэтому напрямую проверяем наличие папки.
    if not os.path.exists(Config.EXCEL_INPUT_DIR):
        log(f"Папка 'Excel' не найдена: {Config.EXCEL_INPUT_DIR}. Создаю.", level="INFO")
        try:
            os.makedirs(Config.EXCEL_INPUT_DIR)
        except OSError as e:
            log(f"Не удалось создать папку 'Excel': {e}", level="ERROR")
            messagebox.showerror("Ошибка создания папки", f"Не удалось создать папку 'Excel': {e}")
            return
        log("В папке 'Excel' нет файлов.", level="INFO")  # После создания, папка пуста
        return

    files = [f for f in os.listdir(Config.EXCEL_INPUT_DIR) if f.lower().endswith((".xlsx", ".xlsm", ".xls"))]
    if files:
        log("Найдены файлы Excel в папке 'Excel':")
        for f in sorted(files):
            log(f"- {f}")
    else:
        log("В папке 'Excel' не найдено файлов для обработки (.xlsx, .xlsm, .xls).", level="INFO")


def run_preprint_threaded():
    """Запускает процесс подготовки документов к печати в отдельном потоке."""
    log("▶ Запуск подготовительного этапа...")
    start_process_thread(process_preprint_task)


def run_postprint_threaded():
    """Запускает процесс замены титульных страниц в отдельном потоке."""
    log("▶ Запуск этапа 'Заменить титульники'...")
    start_process_thread(process_postprint_task)


# --- 6. Настройка и запуск GUI (GUI Setup & Execution) ---

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
button_frame.grid(row=0, column=0, sticky="n", pady=(40, 0), padx=(0, 20))

btn_check_excel = tk.Button(button_frame, text="Проверить файлы Excel", command=check_files_in_excel,
                            bg="#2196F3", fg="white", width=30, height=2,
                            font=("Arial", 10, "bold"))
btn_check_excel.grid(row=0, column=0, pady=(0, 15), sticky="w")

btn_preprint = tk.Button(button_frame, text="Подготовить к печати", command=run_preprint_threaded,
                         bg="#FF9800", fg="white", width=30, height=2,
                         font=("Arial", 10, "bold"))
btn_preprint.grid(row=1, column=0, pady=15, sticky="w")

btn_postprint = tk.Button(button_frame, text="Заменить титульники", command=run_postprint_threaded,
                          bg="#FF5722", fg="white", width=30, height=2,
                          font=("Arial", 10, "bold"))
btn_postprint.grid(row=2, column=0, pady=15, sticky="w")

instruction_frame = tk.Frame(top_frame, bg="#f0f0f0", bd=2, relief="groove")
instruction_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 10), pady=(0, 10))
instruction_label = tk.Label(instruction_frame, text="Инструкция:", bg="#f0f0f0", font=("Arial", 11, "bold"))
instruction_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
instruction_text = tk.Text(instruction_frame, height=12, wrap="word", bg="#fff", font=("Arial", 10), bd=1,
                           relief="solid")
instruction_text.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
instruction_text.insert(tk.END, """
1. Загрузите excel-файлы в папку "Excel".

2. Нажмите кнопку "Проверить файлы Excel" и убедитесь что все необходимые файлы найдены.

3. Нажмите кнопку "Подготовить к печати".
    - программа создаст папку Print, в которой будут лежать PDF
        > все файлы в одном по порядку
        > все титульники в одном
        > все файлы без титульников в одном
    - программа создаст папку NotSignedExport, в которой будут все PDF по отдельности

4. Проверьте содержание экспортированных файлов на наличие ошибок.    
    
5. Распечатайте файл "title_merged" (титульники), отправьте на подпись и отсканируйте.

    ====!ВАЖНО!========!ВАЖНО!========!ВАЖНО!====
    - при сканировании убедитесь, что листы лежат в том же порядке что и при печати
    
6. Отсканированный файл назовите "title_scan" и положите в папку Print.

    ====!ВАЖНО!========!ВАЖНО!========!ВАЖНО!====
    - Название файла со сканами СТРОГО "title_scan"
    
7. Нажмите кнопку "Заменить титульники". Программа поместит результат в папку "Final".  
""")
instruction_text.config(state='disabled')
instruction_frame.grid_rowconfigure(1, weight=1)
instruction_frame.grid_columnconfigure(0, weight=1)

log_output = scrolledtext.ScrolledText(root, wrap=tk.WORD, bg="#ffffff", fg="#000000", font=("Consolas", 10), bd=2,
                                       relief="sunken")
log_output.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=(0, 10))

sys.stdout = RedirectText(log_queue, _original_stdout)

# Инициализация путей при старте приложения
Config.initialize_paths()

# Первичная проверка и создание папки Excel, если её нет.
# Используем ensure_and_clear_folder, но только для создания, не для очистки при первом запуске.
if not os.path.exists(Config.EXCEL_INPUT_DIR):
    try:
        os.makedirs(Config.EXCEL_INPUT_DIR)
        log("Папка 'Excel' создана.", level="INFO")
    except OSError as e:
        log(f"Не удалось создать папку 'Excel': {e}", level="ERROR")
        messagebox.showerror("Ошибка создания папки", f"Не удалось создать папку 'Excel': {e}")

root.after(100, update_log_display)

root.mainloop()