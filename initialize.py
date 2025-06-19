import tkinter as tk
from tkinter import scrolledtext
import os
import sys
import io
import preprint
import postprint

# Создание папки Excel при запуске
if not os.path.exists('Excel'):
    os.mkdir('Excel')

# Перенаправление stdout
class RedirectText(io.TextIOBase):
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)

    def flush(self):
        pass

def check_files_in_excel():
    log_output.delete(1.0, tk.END)
    files = os.listdir('Excel')
    if files:
        log_output.insert(tk.END, "В папке 'Excel' найдены следующие файлы:\n")
        for file in files:
            log_output.insert(tk.END, f"{file}\n")
    else:
        log_output.insert(tk.END, "В папке 'Excel' нет файлов.\n")

def run_preprint():
    log_output.delete(1.0, tk.END)
    try:
        log_output.insert(tk.END, "Запуск подготовительного этапа...\n")
        preprint.run()
        log_output.insert(tk.END, "Завершено.\n")
    except FileNotFoundError as e:
        log_output.insert(tk.END, f"❗ Ошибка: {e}\n")
    except Exception as e:
        log_output.insert(tk.END, f"⚠️ Непредвиденная ошибка: {e}\n")

def run_postprint():
    log_output.delete(1.0, tk.END)
    try:
        log_output.insert(tk.END, "Запуск заключительного этапа...\n")
        postprint.run()
        log_output.insert(tk.END, "Завершено.\n")
    except FileNotFoundError as e:
        log_output.insert(tk.END, f"❗ Ошибка: {e}\n")
    except Exception as e:
        log_output.insert(tk.END, f"⚠️ Непредвиденная ошибка: {e}\n")

# --- GUI ---
root = tk.Tk()
root.title("Утилита для подготовки документов")
root.geometry("800x600")
root.minsize(700, 500)
root.config(bg="#f0f0f0")

# --- Основная сетка ---
root.grid_rowconfigure(1, weight=1)  # лог растягивается по вертикали
root.grid_columnconfigure(0, weight=1)  # левая часть (кнопки + инструкция)
root.grid_columnconfigure(1, weight=1)  # правая часть (инструкция)

# --- Верхний блок: кнопки + инструкция ---
top_frame = tk.Frame(root, bg="#f0f0f0")
top_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

top_frame.grid_columnconfigure(0, weight=0)
top_frame.grid_columnconfigure(1, weight=1)
top_frame.grid_rowconfigure(0, weight=1)

# --- Левая колонка: кнопки ---
button_frame = tk.Frame(top_frame, bg="#f0f0f0")
button_frame.grid(row=0, column=0, sticky="n", pady=(40, 0))

button_check_files = tk.Button(button_frame, text="Проверить файлы Excel", command=check_files_in_excel,
                               bg="#2196F3", fg="white", width=30)
button_check_files.grid(row=0, column=0, pady=(0, 15), sticky="w")

button_preprint = tk.Button(button_frame, text="Подготовить к печати", command=run_preprint,
                            bg="#FF9800", fg="white", width=30)
button_preprint.grid(row=1, column=0, pady=15, sticky="w")

button_postprint = tk.Button(button_frame, text="Заменить титульники", command=run_postprint,
                             bg="#FF5722", fg="white", width=30)
button_postprint.grid(row=2, column=0, pady=15, sticky="w")


# --- Правая колонка: инструкция ---
instruction_frame = tk.Frame(top_frame, bg="#f0f0f0")
instruction_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 10), pady=(0, 10))

instruction_label = tk.Label(instruction_frame, text="Инструкция:", bg="#f0f0f0", font=("Arial", 11, "bold"))
instruction_label.grid(row=0, column=0, sticky="w", padx=(0, 0), pady=(0, 5))

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
    
4. Распечатайте файл "titles_merged" (титульники), отправьте на подпись и отсканируйте.

    ====!ВАЖНО!========!ВАЖНО!========!ВАЖНО!====
    - при сканировании убедитесь что листы лежат в том же порядке что и при печати
    
5. Отсканированный файл назовите "titles_scan" и положить в папку Print.

    ====!ВАЖНО!========!ВАЖНО!========!ВАЖНО!====
    - Название файла со сканами СТРОГО "titles_scan"
    
6. Нажмите кнопку "Заменить титульники". Программа поместит результат в папку "Final".           
                                """)
instruction_text.config(state='disabled')
instruction_frame.grid_rowconfigure(1, weight=1)
instruction_frame.grid_columnconfigure(0, weight=1)

# --- Нижний блок: лог ---
log_output = scrolledtext.ScrolledText(root, wrap=tk.WORD,
                                       bg="#ffffff", fg="#000000",
                                       font=("Consolas", 10))
log_output.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=10, pady=(0, 10))

# Перенаправление stdout
sys.stdout = RedirectText(log_output)

root.mainloop()
