# Bank Statement Parser & GUI App 📊

A standalone desktop application with a graphical user interface (GUI) built in Python. It automates the parsing, categorization, and calculation of bank statements (e.g., Kaspi Pay), transforming raw data into structured Excel reports for the accounting department.

## 🚀 Project Overview
This project was designed to eliminate manual data entry for accountants. Instead of running Python scripts via the console, users interact with a simple Tkinter-based GUI. The application is compiled into a standalone `.exe` file, allowing non-technical staff to process financial data with a single click.

## ⚙️ Core Features:
* **Graphical Interface:** Easy-to-use file selection window built with `tkinter`.
* **Smart Parsing:** Automatically detects header rows, cleans messy numeric strings, and identifies transaction types.
* **Financial Logic:** Separates incoming payments (purchases) from outgoing (refunds/fees), calculates totals, and computes banking commissions.
* **Automated Reporting:** Generates a structured Excel file with a consolidated "Summary" sheet and individual sheets for each payment type.
* **Standalone Distribution:** Packaged as an `.exe` file via `PyInstaller` for deployment on any Windows machine without requiring a Python installation.

## 🛠 Tech Stack:
* **Language:** Python 3
* **GUI:** `tkinter`
* **Data Processing:** `pandas`, `openpyxl`
* **Compilation:** `PyInstaller`

## 💻 How to run from source:
1. Clone the repository.
2. Install dependencies:
   `pip install -r requirements.txt`
3. Run the app:
   `buh.py`

---
*(Russian version below)*

# Десктопное приложение: Разнос банковских выписок 📊

Оконное приложение с графическим интерфейсом (GUI) на Python для автоматического парсинга, категоризации и подсчета банковских выписок (Kaspi Pay). 

## 🚀 Описание проекта
Инструмент разработан специально для бухгалтерии. Чтобы избавить пользователей от работы с консолью, скрипт обернут в понятный графический интерфейс (Tkinter) и скомпилирован в самостоятельный `.exe` файл. Это позволяет нетехническим специалистам обрабатывать сложные финансовые выгрузки в один клик без установки Python.

## ⚙️ Основной функционал:
* **Графический интерфейс:** Удобный выбор файлов через диалоговое окно.
* **Умный парсинг:** Скрипт сам находит строку с заголовками, очищает текстовые значения сумм и определяет типы операций.
* **Финансовая логика:** Разделяет приход и уход, высчитывает общие суммы и комиссии банка по каждому типу оплат.
* **Генерация отчетов:** Создает итоговый Excel-файл со сводной таблицей на главном листе и разбивкой транзакций по отдельным листам.
* **Дистрибуция:** Сборка в `.exe` с помощью `PyInstaller` для работы на любом ПК с Windows.

## 🛠 Стек технологий:
* **Язык:** Python 3
* **GUI:** `tkinter`
* **Обработка данных:** `pandas`, `openpyxl`
* **Сборка:** `PyInstaller`

* ## 💻 Как запустить из исходного кода:
1. Клонируйте репозиторий.
2. Установите зависимости:
   `pip install -r requirements.txt`
3. Запустите приложение:
   `python buh.py`
