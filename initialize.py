import tkinter as tk
from tkinter import scrolledtext, messagebox
import os
import subprocess

# Функция для проверки наличия файлов в папке Excel
def check_files_in_excel():
    files = os.listdir('Excel')
    if files:
        log_output.insert(tk.END, "В папке 'Excel' найдены следующие файлы:\n")
        for file in files:
            log_output.insert(tk.END, f"{file}\n")  # Каждое имя файла на новой строке
    else:
        log_output.insert(tk.END, "В папке 'Excel' нет файлов.\n")

# Функция для запуска preprint.py
def run_preprint():
    log_output.insert(tk.END, "Запуск подготовительного этапа (preprint)...\n")
    try:
        subprocess.run(['python', 'preprint.py'], check=True)
        log_output.insert(tk.END, "Подготовительный этап завершен успешно.\n")
    except subprocess.CalledProcessError:
        log_output.insert(tk.END, "Ошибка при выполнении подготовительного этапа.\n")

# Функция для запуска postprint.py
def run_postprint():
    log_output.insert(tk.END, "Запуск заключительного этапа (postprint)...\n")
    try:
        subprocess.run(['python', 'postprint.py'], check=True)
        log_output.insert(tk.END, "Заключительный этап завершен успешно.\n")
    except subprocess.CalledProcessError:
        log_output.insert(tk.END, "Ошибка при выполнении заключительного этапа.\n")

# Автоматическое создание папки Excel при запуске программы
if not os.path.exists('Excel'):
    os.mkdir('Excel')
    print("Папка 'Excel' была создана.")

# Создание основного окна
root = tk.Tk()
root.title("Утилита для работы с файлами Excel")
root.geometry('600x400')  # Размер окна

# Улучшение внешнего вида окна
root.config(bg="#f0f0f0")

# Создание рамки для организации элементов
frame = tk.Frame(root, bg="#f0f0f0")
frame.pack(pady=20)

# Кнопка для проверки файлов в папке Excel
button_check_files = tk.Button(frame, text="Проверить файлы в Excel", command=check_files_in_excel, bg="#2196F3", fg="white", width=40)
button_check_files.grid(row=0, column=0, padx=10, pady=10)

# Кнопка для запуска preprint.py
button_preprint = tk.Button(frame, text="Запустить подготовительный этап", command=run_preprint, bg="#FF9800", fg="white", width=40)
button_preprint.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

# Кнопка для запуска postprint.py
button_postprint = tk.Button(frame, text="Запустить заключительный этап", command=run_postprint, bg="#FF5722", fg="white", width=40)
button_postprint.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

# Добавление области для отображения лога
log_output = scrolledtext.ScrolledText(root, width=70, height=10, wrap=tk.WORD, bg="#ffffff", fg="#000000", font=("Arial", 10))
log_output.pack(pady=10)

# Запуск интерфейса
root.mainloop()
