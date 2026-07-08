import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def process_excel(input_file):
    """Основная логика парсинга и разноса выписки"""
    # 1. Загрузка сырых данных для поиска шапки
    raw_df = pd.read_excel(input_file, header=None)

    header_row_index = 0
    for idx, row in raw_df.iterrows():
        row_values = [str(x).strip() for x in row.values if pd.notna(x)]
        if '#' in row_values or 'Тип операции' in row_values:
            header_row_index = idx
            break

    # Перезагружаем файл с нужной строки
    df = pd.read_excel(input_file, skiprows=header_row_index)

    # Стандартизируем названия колонок
    df.columns = df.columns.astype(str).str.strip().str.replace('\n', ' ').str.replace('  ', ' ')

    # Поиск ключевых колонок
    target_col = None
    for col in df.columns:
        if 'Сумма к зачислению' in col or 'Сумма операции' in col:
            target_col = col
            break

    comm_col = None
    for col in df.columns:
        if 'Комиссия за операцию' in col or 'Комиссия' in col:
            comm_col = col
            break

    if not target_col:
        raise ValueError("Не удалось найти колонку с суммой операции в файле!")

    # Функция очистки текстовых сумм в float
    def clean_money(value):
        if pd.isna(value):
            return 0
        clean_str = str(value).replace('\xa0', '').replace(' ', '').replace(',', '.')
        try:
            return float(clean_str)
        except:
            return 0

    df[target_col] = df[target_col].apply(clean_money)
    if comm_col:
        df[comm_col] = df[comm_col].apply(clean_money)

    op_type_col = 'Тип операции' if 'Тип операции' in df.columns else None
    pay_type_col = 'Тип оплаты 2' if 'Тип оплаты 2' in df.columns else None

    if not (op_type_col and pay_type_col):
        raise ValueError("В таблице не найдены обязательные колонки 'Тип операции' или 'Тип оплаты 2'!")

    # Разделение на приход и уход
    incoming = df[df[op_type_col].astype(str).str.contains('Покупка', case=False, na=False)]
    outgoing = df[~df[op_type_col].astype(str).str.contains('Покупка', case=False, na=False)]

    payment_types = incoming[pay_type_col].dropna().unique()

    # Генерация данных для Сводки
    summary_data = []
    for p_type in payment_types:
        temp_df = incoming[incoming[pay_type_col] == p_type]
        total_sum = temp_df[target_col].sum()
        total_comm = temp_df[comm_col].sum() if comm_col else 0

        summary_data.append({
            'Группа': 'Приход (Покупка)',
            'Тип оплаты': p_type,
            'Кол-во операций': len(temp_df),
            'Общая сумма': round(total_sum, 2),
            'Общая комиссия': round(total_comm, 2)
        })

    if not outgoing.empty:
        out_sum = outgoing[target_col].sum()
        out_comm = outgoing[comm_col].sum() if comm_col else 0
        summary_data.append({
            'Группа': 'Возвраты и Списания',
            'Тип оплаты': 'Разное',
            'Кол-во операций': len(outgoing),
            'Общая сумма': round(out_sum, 2),
            'Общая комиссия': round(out_comm, 2)
        })

    summary_df = pd.DataFrame(summary_data)
    if not summary_df.empty:
        summary_df.loc['ИТОГО'] = [
            'ИТОГ', '-',
            summary_df['Кол-во операций'].sum(),
            round(summary_df['Общая сумма'].sum(), 2),
            round(summary_df['Общая комиссия'].sum(), 2)
        ]

    # Путь для сохранения итогового файла (создается в той же папке)
    dir_name = os.path.dirname(input_file)
    base_name = os.path.basename(input_file)
    output_file = os.path.join(dir_name, f"разнос_{base_name}")

    # Запись в Excel: строгая последовательность гарантирует, что СВОДКА будет первым листом
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Первым пишем лист сводки
        summary_df.to_excel(writer, sheet_name='СВОДКА', index=False)

        # Затем пишем все остальные листы
        for p_type in payment_types:
            temp_df = incoming[incoming[pay_type_col] == p_type]
            sheet_name = str(p_type)[:31]
            temp_df.to_excel(writer, sheet_name=sheet_name, index=False)

        if not outgoing.empty:
            outgoing.to_excel(writer, sheet_name='Возвраты_и_списания', index=False)

    return output_file


def select_and_process_file():
    """Обработчик нажатия на кнопку в интерфейсе"""
    file_path = filedialog.askopenfilename(
        title="Выберите файл выписки Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    label_status.config(text="Обработка файла... Пожалуйста, подождите.")
    root.update_idletasks()

    try:
        out_path = process_excel(file_path)
        label_status.config(text="Файл успешно обработан!")
        messagebox.showinfo(
            "Успех",
            f"Разнос выписок завершен!\n\nРезультат сохранен в ту же папку под именем:\n{os.path.basename(out_path)}"
        )
    except Exception as e:
        label_status.config(text="Произошла ошибка при обработке.")
        messagebox.showerror("Ошибка", f"Не удалось обработать файл.\nПричина: {str(e)}")


# Настройка графического интерфейса
root = tk.Tk()
root.title("Разнос банковских выписок")
root.geometry("450x200")
root.resizable(False, False)

label_title = tk.Label(root, text="Программа разноса банковских выписок", font=("Arial", 12, "bold"))
label_title.pack(pady=15)

btn_select = tk.Button(
    root,
    text="Выбрать файл выписки",
    command=select_and_process_file,
    font=("Arial", 10),
    bg="#2196F3",
    fg="white",
    padx=15,
    pady=5
)
btn_select.pack(pady=10)

label_status = tk.Label(root, text="Ожидание выбора файла...", font=("Arial", 9), fg="gray")
label_status.pack(pady=15)

if __name__ == "__main__":
    root.mainloop()