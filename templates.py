import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import time

# --- СТИЛИЗАЦИЯ НЕО-КИБЕРПАНК (Бирюзовый акцент) ---
BG_COLOR = "#0f0c1a"       # Очень темно-синий/фиолетовый фон
LOG_BG = "#1a1a2e"         # Темный фон лога и полей ввода
FG_COLOR = "#00ffe1"       # Неоново-голубой (основной текст)
ACCENT_COLOR = "#00FFFF"   # <--- ИЗМЕНЕНИЕ: Ярко-бирюзовый/Неоновый Циан (замена оранжевого)
BUTTON_FG = "#00ffe1"      # Неоново-голубой текст кнопок
BUTTON_BG = LOG_BG         # Фон кнопок
ERROR_COLOR = "#ff003c"    # Неоново-красный (ошибки)
SUCCESS_COLOR = "#00ff88"  # Неоново-зеленый (успех)
INFO_COLOR = "#00ffe1"     # Неоново-голубой (информация в логе)

FONT = ("Consolas", 10)
TITLE_FONT = ("Consolas", 12, "bold")

# Глобальная переменная для текстового поля логов
log_widget = None 
root = None # Глобальная ссылка на root-окно

def log_message(message, tag="INFO"):
    """Обновляет текстовое поле логов."""
    # Убедитесь, что GUI обновляется, чтобы избежать ошибок после завершения работы.
    global root
    if not log_widget or not root.winfo_exists():
        return

    # Добавляем метку времени и тег
    timestamp = time.strftime("%H:%M:%S")
    
    # Определяем цвет текста
    if tag == "ERROR":
        color = ERROR_COLOR
        prefix = "[!] ERROR: "
    elif tag == "SUCCESS":
        color = SUCCESS_COLOR
        prefix = "[+] SUCCESS: "
    else:
        color = INFO_COLOR
        prefix = "[i] INFO: "
    
    full_message = f"[{timestamp}] {prefix}{message}\n"
    
    log_widget.configure(state='normal')
    log_widget.insert(tk.END, full_message)
    
    # Применяем цвет
    start_index = log_widget.index(tk.END + "-1c linestart")
    end_index = log_widget.index(tk.END + "-1c")
    log_widget.tag_add(tag, start_index, end_index)
    log_widget.tag_config(tag, foreground=color)
    
    log_widget.see(tk.END)
    log_widget.configure(state='disabled')
    log_widget.update_idletasks() # Принудительное обновление GUI

# --- Функции выбора файлов ---

def select_logins_file():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        logins_entry.delete(0, tk.END)
        logins_entry.insert(0, path)
        log_message(f"Выбран файл логинов: {path.split('/')[-1]}")

def select_template_file():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        template_entry.delete(0, tk.END)
        template_entry.insert(0, path)
        log_message(f"Выбран файл шаблона: {path.split('/')[-1]}")

def select_output_file():
    path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
    if path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, path)
        log_message(f"Выбран файл вывода: {path.split('/')[-1]}")

# --- Логика run_script с логгированием ---

def run_script():
    log_message("--- НАЧАЛО АНАЛИЗА ФАЙЛА ---", tag="INFO")
    
    logins_file = logins_entry.get()
    template_file = template_entry.get()
    output_file = output_entry.get()
    login_column = login_column_entry.get()
    
    logins_sheet = logins_sheet_entry.get()
    template_sheet = template_sheet_entry.get()
    
    # Конвертация листа в число или использование 0
    try:
        if logins_sheet and logins_sheet.strip().isdigit():
             logins_sheet = int(logins_sheet)
        elif not logins_sheet:
             logins_sheet = 0
        
        if template_sheet and template_sheet.strip().isdigit():
             template_sheet = int(template_sheet)
        elif not template_sheet:
             template_sheet = 0
    except Exception:
        log_message("ОШИБКА: Номер листа должен быть числом.", tag="ERROR")
        return


    if not all([logins_file, template_file, output_file, login_column]):
        log_message("ОШИБКА: Пожалуйста, заполните все поля ввода.", tag="ERROR")
        messagebox.showerror("Ошибка", "Пожалуйста, заполните все поля!")
        return

    try:
        # 1. Чтение логинов
        log_message(f"Считывание логинов из файла: {logins_file.split('/')[-1]} (Лист: {logins_sheet})")
        
        logins_df = pd.read_excel(logins_file, sheet_name=logins_sheet)
        
        if login_column not in logins_df.columns:
            logins_series = logins_df.iloc[:, 0].dropna()
            log_message("ВНИМАНИЕ: Столбец логинов не найден по имени. Используется первый столбец.", tag="INFO")
        else:
            logins_series = logins_df[login_column].dropna()
        
        logins = logins_series.astype(str).unique().tolist()
        logins = [l for l in logins if l.lower() != login_column.lower()]
        
        if not logins:
            log_message("ОШИБКА: Список логинов пуст после фильтрации.", tag="ERROR")
            messagebox.showerror("Ошибка", "Не найдены логины.")
            return

        log_message(f"Найдено уникальных логинов для обработки: {len(logins)}", tag="INFO")
        
        # 2. Чтение шаблона
        log_message(f"Считывание шаблона из файла: {template_file.split('/')[-1]} (Лист: {template_sheet})")
        template_df = pd.read_excel(template_file, sheet_name=template_sheet)
        
        if template_df.empty:
            log_message("ОШИБКА: Шаблон пуст. Проверьте лист и файл.", tag="ERROR")
            messagebox.showerror("Ошибка", "Шаблон пуст.")
            return
            
        template_rows_count = len(template_df)
        log_message(f"В шаблоне найдено строк данных: {template_rows_count}")
        
        # 3. Обработка и конкатенация
        all_user_data = []
        log_message("Начало клонирования и подстановки логинов...")
        
        # 4. Копирование и замена логина
        for i, login in enumerate(logins):
            user_df = template_df.copy()
            
            if login_column not in user_df.columns:
                 user_df.insert(1, login_column, str(login)) 
            else:
                user_df[login_column] = str(login)

            all_user_data.append(user_df)
            
            if (i + 1) % 100 == 0 or (i + 1) == len(logins):
                 log_message(f"Обработано {i + 1}/{len(logins)} логинов. Текущий: {login}")

        final_data = pd.concat(all_user_data, ignore_index=True)

        # 5. Сохранение результата
        total_rows = len(final_data)
        log_message(f"Общее количество строк для записи: {total_rows}")
        log_message(f"Запись результата в {output_file.split('/')[-1]}...")

        # 6. Сохранение (С ФОРМАТИРОВАНИЕМ ПРОЦЕНТА)
        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                # Записываем данные в Excel
                final_data.to_excel(writer, sheet_name='Worksheet', index=False, header=True)
                
                workbook  = writer.book
                worksheet = writer.sheets['Worksheet']

                # Определяем формат процента
                percent_format = workbook.add_format({'num_format': '0%'}) 

                # Ищем индекс столбца "Процент прохождения"
                try:
                    # 'Процент прохождения' должен быть именем столбца в вашем файле-шаблоне
                    percent_column_index = final_data.columns.get_loc('Процент прохождения')
                except KeyError:
                    log_message("ВНИМАНИЕ: Столбец 'Процент прохождения' не найден для форматирования. Проверьте его название.", tag="INFO")
                    percent_column_index = -1
                
                # Применяем формат к столбцу, если он найден
                if percent_column_index != -1:
                    # Применяем формат ко всему столбцу, начиная с первой строки данных (row 1, т.к. row 0 - заголовок)
                    worksheet.set_column(percent_column_index, percent_column_index, None, percent_format)
                    log_message(f"К столбцу '{final_data.columns[percent_column_index]}' применен формат '0%'.", tag="INFO")


            log_message(f"Файл импорта '{output_file.split('/')[-1]}' успешно создан.", tag="SUCCESS")
            log_message(f"Обработано {len(logins)} пользователей и {total_rows} записей.", tag="SUCCESS")

        except Exception as e:
            # Перехват ошибок, связанных с записью Excel
            log_message(f"ОШИБКА ЗАПИСИ: Не удалось записать файл Excel. {e}", tag="ERROR")
            messagebox.showerror("Ошибка записи", f"Не удалось записать файл: {e}")

    except FileNotFoundError as e:
        log_message(f"ОШИБКА ФАЙЛА: Файл не найден. Проверьте путь. {e}", tag="ERROR")
        messagebox.showerror("Ошибка", f"Файл не найден. Проверьте путь: {e}")
    except KeyError as e:
        log_message(f"ОШИБКА В ДАННЫХ: Не найден столбец {e}. Проверьте правильность его названия.", tag="ERROR")
        messagebox.showerror("Ошибка в данных", f"Не найден столбец: {e}.")
    except Exception as e:
        log_message(f"КРИТИЧЕСКАЯ ОШИБКА: {e}", tag="ERROR")
        messagebox.showerror("Критическая ошибка", str(e))


# --- GUI (Стиль КИБЕРПАНК) ---

root = tk.Tk()
root.title("МОДУЛЬ ТИРАЖИРОВАНИЯ ШАБЛОНОВ | by kucingsayabesar V0.41")
root.config(bg=BG_COLOR)
root.resizable(False, False)

# Функция для создания стилизованных Label
def create_label(parent, text, row, col, sticky="e"):
    lbl = tk.Label(parent, text=text, bg=BG_COLOR, fg=FG_COLOR, font=FONT)
    lbl.grid(row=row, column=col, sticky=sticky, padx=5, pady=5)
    return lbl

# Функция для создания стилизованных Entry
def create_entry(parent, default_text, row, col, width=50):
    entry = tk.Entry(parent, width=width, bg=LOG_BG, fg=INFO_COLOR, font=FONT, insertbackground=FG_COLOR, borderwidth=1, relief="flat")
    entry.grid(row=row, column=col, padx=5, pady=5)
    entry.insert(0, default_text)
    return entry

# Функция для создания стилизованных Button
def create_button(parent, text, command, row, col, span=1, color=BUTTON_FG):
    btn = tk.Button(parent, text=text, command=command, bg=BUTTON_BG, fg=color, 
                    font=FONT, activebackground=SUCCESS_COLOR, activeforeground=BG_COLOR, 
                    relief="flat", borderwidth=0, padx=10, pady=2)
    btn.grid(row=row, column=col, columnspan=span, padx=5, pady=5)
    return btn

# Заголовок
title_lbl = tk.Label(root, text="МОДУЛЬ ТИРАЖИРОВАНИЯ ШАБЛОНОВ⚡", bg=BG_COLOR, fg=ACCENT_COLOR, font=("Consolas", 14, "bold"))
title_lbl.grid(row=0, column=0, columnspan=3, pady=10)

# Файл с логинами
create_label(root, "Файл [LOGINS]:", 1, 0)
logins_entry = create_entry(root, "logins.xlsx", 1, 1)
create_button(root, "ВЫБРАТЬ", select_logins_file, 1, 2, color=BUTTON_FG)

# Лист с логинами
create_label(root, "Лист [LOGINS]:", 2, 0)
logins_sheet_entry = create_entry(root, "0", 2, 1)

# Файл шаблона
create_label(root, "Файл [TEMPLATE]:", 3, 0)
template_entry = create_entry(root, "template.xlsx", 3, 1)
create_button(root, "ВЫБРАТЬ", select_template_file, 3, 2, color=BUTTON_FG)

# Лист шаблона
create_label(root, "Лист [TEMPLATE]:", 4, 0)
template_sheet_entry = create_entry(root, "Worksheet", 4, 1)

# Столбец для логина
create_label(root, "СТОЛБЕЦ [TARGET]:", 5, 0)
login_column_entry = create_entry(root, "Логин пользователя", 5, 1)

# Итоговый файл
create_label(root, "Файл [OUTPUT]:", 6, 0)
output_entry = create_entry(root, "final_file.xlsx", 6, 1)
create_button(root, "СОХРАНИТЬ КАК", select_output_file, 6, 2, color=BUTTON_FG)

# Кнопка запуска
main_button = create_button(root, "СТАРТ АНАЛИЗ (ACTIVATE)", run_script, 7, 1, span=1, color=SUCCESS_COLOR) 
main_button.config(bg=LOG_BG, activebackground=SUCCESS_COLOR, activeforeground=BG_COLOR, fg=SUCCESS_COLOR) 


# --- ЛОГГИРОВАНИЕ ---
log_label = create_label(root, ":: LOG CONSOLE ::", 8, 0, sticky="w")
log_label.config(fg=SUCCESS_COLOR, font=("Consolas", 10, "bold"))

# Создание поля для логов
log_widget = scrolledtext.ScrolledText(root, height=10, width=80, bg=LOG_BG, fg=INFO_COLOR, font=FONT, 
                                       insertbackground=FG_COLOR, borderwidth=0, relief="solid", highlightthickness=1, 
                                       highlightbackground=SUCCESS_COLOR)
log_widget.grid(row=9, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
log_widget.configure(state='disabled')
log_message("СИСТЕМА ЗАПУЩЕНА. ОЖИДАНИЕ ДАННЫХ.")


root.mainloop()