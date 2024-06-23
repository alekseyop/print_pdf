import os
import time
import glob
import win32print
import win32api
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox
import json

# Настройки
file_extension = "*.pdf"  # расширение файлов для печати
settings_file = "printer_settings.json"  # файл для сохранения настроек
sumatra_executable_name = "SumatraPDF.exe"  # имя исполняемого файла SumatraPDF
max_retries = 3  # максимальное количество попыток печати

def find_sumatra():
    for root, dirs, files in os.walk("C:\\"):
        if sumatra_executable_name in files:
            return os.path.join(root, sumatra_executable_name)
    return None

def get_installed_printers():
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    return [printer[2] for printer in printers]

def select_printer(printers):
    root = tk.Tk()
    root.title("Выбор принтера")

    tk.Label(root, text="Выберите принтер:").pack(pady=10)

    printer_var = tk.StringVar(root)
    printer_var.set(printers[0])

    printer_menu = tk.OptionMenu(root, printer_var, *printers)
    printer_menu.pack(pady=10)

    def on_select():
        root.destroy()

    tk.Button(root, text="Выбрать", command=on_select).pack(pady=10)

    root.mainloop()

    return printer_var.get()

def select_folder():
    root = tk.Tk()
    root.withdraw()  # Скрыть основное окно
    folder_selected = filedialog.askdirectory(title="Выберите папку для мониторинга")
    root.destroy()
    return folder_selected

def prompt_for_interval():
    root = tk.Tk()
    root.withdraw()  # Скрыть основное окно
    interval = simpledialog.askinteger("Период опроса", "Введите период опроса папки в секундах:", minvalue=1)
    root.destroy()
    return interval

def load_settings():
    if os.path.exists(settings_file):
        with open(settings_file, 'r') as f:
            return json.load(f)
    return {}

def save_settings(settings):
    with open(settings_file, 'w') as f:
        json.dump(settings, f)

def get_files(folder, extension):
    return glob.glob(os.path.join(folder, extension))

def print_file(sumatra_path, printer, file_path):
    for attempt in range(max_retries):
        try:
            win32api.ShellExecute(
                0,
                None,
                sumatra_path,
                f'-print-to "{printer}" "{file_path}"',
                ".",
                0
            )
            print(f"Successfully printed: {file_path}")
            return True
        except Exception as e:
            print(f"Не удалось распечатать:{file_path} при попытке {attempt + 1}. Error: {e}")
            time.sleep(1)  # Ждем 5 секунд перед следующей попыткой
    return False

def main():
    settings = load_settings()
    printer_name = settings.get("printer_name")
    folder_to_monitor = settings.get("folder_to_monitor")
    check_interval = settings.get("check_interval")
    sumatra_path = settings.get("sumatra_path")

    root = tk.Tk()
    root.withdraw()  # Скрыть основное окно
    change_printer = messagebox.askyesno("Изменить принтер", "Вы хотите изменить принтер?")
    change_folder = messagebox.askyesno("Изменить папку", "Вы хотите изменить папку для мониторинга?")
    change_interval = messagebox.askyesno("Изменить период опроса", "Вы хотите изменить период опроса папки?")
    root.destroy()

    if change_printer or not printer_name:
        printers = get_installed_printers()
        if not printers:
            messagebox.showerror("Ошибка", "Принтеры не найдены.")
            return
        printer_name = select_printer(printers)
        if printer_name:
            settings["printer_name"] = printer_name
            save_settings(settings)
        else:
            print("Принтер не выбран. Завершение программы.")
            return

    if change_folder or not folder_to_monitor:
        folder_to_monitor = select_folder()
        if folder_to_monitor:
            settings["folder_to_monitor"] = folder_to_monitor
            save_settings(settings)
        else:
            print("Папка не выбрана. Завершение программы.")
            return

    if change_interval or not check_interval:
        check_interval = prompt_for_interval()
        if check_interval:
            settings["check_interval"] = check_interval
            save_settings(settings)
        else:
            print("Период опроса не введен. Завершение программы.")
            return

    if not sumatra_path:
        sumatra_path = find_sumatra()
        if sumatra_path:
            settings["sumatra_path"] = sumatra_path
            save_settings(settings)
        else:
            print("SumatraPDF не найден. Убедитесь, что он установлен.")
            return

    while True:
        files = get_files(folder_to_monitor, file_extension)
        for file in files:
            print(f"Trying to print: {file}")
            if print_file(sumatra_path, printer_name, file):
                # Удаляем файл после успешной печати
                os.remove(file)
            else:
                print(f"Failed to print: {file} after {max_retries} attempts.")
        time.sleep(check_interval)

if __name__ == "__main__":
    main()
