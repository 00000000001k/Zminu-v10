# Excel Budget Transfer Automation

Программа предназначена для автоматизации формирования файла **«Зміни.xlsx»** на основе данных из файлов **«Розрахунок»** и **«Кошторис»**.

Проект значительно сокращает время обработки финансовых документов и практически исключает ошибки при ручном переносе данных.

---

## Возможности

- 📄 Автоматический поиск файла «Розрахунок».
- 📂 Ручной выбор файлов через графический интерфейс, если они не найдены автоматически.
- 📊 Чтение данных из Excel (.xlsx, .xlsm).
- 🔍 Поиск мероприятий по номеру или названию.
- 💰 Анализ остатков по КЭКР:
  - 2210
  - 2240
  - 2250
- 📈 Расчет фактических остатков даже при использовании Excel-формул.
- 🧮 Поддержка вычисления вложенных формул и SUM().
- 📦 Автоматическое определение товаров, подлежащих переносу.
- 🛒 Отдельная обработка уже приобретённых товаров (1812).
- ⚠ Проверка соответствия сумм между документами.
- 🚨 Предупреждение о недостатке средств по выбранным КЭКР.
- 🎯 Возможность выбора:
  - какие КЭКР использовать;
  - переносить только необходимую сумму;
  - переносить весь остаток.
- 📝 Автоматическое создание и заполнение файла **«Зміни.xlsx»**.
- 🎨 Цветовое выделение строк с остатками.
- 🔄 Сохранение ранее внесённых данных в файле «Зміни».

---

## Используемые технологии

- Python 3
- OpenPyXL
- Tkinter
- pathlib
- AST
- Regular Expressions (re)

---

## Структура работы

1. Открывается файл «Розрахунок».
2. Анализируется файл «Кошторис».
3. Пользователь выбирает мероприятия.
4. Программа анализирует остатки.
5. Выполняется сравнение с потребностями сметы.
6. Пользователь выбирает режим переноса.
7. Формируется файл **«Зміни.xlsx»**.

---

## Требования

Установите зависимости:

```bash
pip install openpyxl
```

Tkinter входит в стандартную поставку Python.

---

## Запуск

```bash
python script.py
```

---

## Особенности

- Автоматически вычисляет Excel-формулы.
- Работает даже если Excel не сохранил кэш формул.
- Поддерживает объединённые ячейки.
- Проверяет корректность расчётов.
- Минимизирует ручную работу при подготовке финансовой документации.

---

## Лицензия


Проект распространяется в образовательных и служебных целях.



# Excel Budget Transfer Automation

This project automates the creation of the **"Zminy.xlsx"** file using data from **"Rozrahunok"** and **"Koshtorys"** Excel documents.

It significantly reduces manual work and minimizes errors when transferring financial data between documents.

---

## Features

- 📄 Automatically detects the **Rozrahunok** workbook.
- 📂 Allows manual file selection if files are not found.
- 📊 Reads Excel (.xlsx, .xlsm) files.
- 🔍 Finds events by number or title.
- 💰 Analyzes remaining funds for expense codes:
  - 2210
  - 2240
  - 2250
- 📈 Calculates remaining balances from Excel formulas.
- 🧮 Supports nested formulas and SUM() calculations.
- 📦 Automatically determines items that should be transferred.
- 🛒 Separately processes already purchased items (1812).
- ⚠ Validates totals between documents.
- 🚨 Warns when available funds are insufficient.
- 🎯 Lets the user choose:
  - which expense codes to use;
  - transfer only the required amount;
  - transfer the entire remaining balance.
- 📝 Automatically generates and fills **"Zminy.xlsx"**.
- 🎨 Highlights remaining balance rows.
- 🔄 Preserves existing records in the output file.

---

## Technologies

- Python 3
- OpenPyXL
- Tkinter
- pathlib
- AST
- Regular Expressions (re)

---

## Workflow

1. Open the **Rozrahunok** workbook.
2. Analyze the **Koshtorys** workbook.
3. Select events to process.
4. Analyze remaining balances.
5. Compare them with the budget requirements.
6. Choose the transfer mode.
7. Generate the **Zminy.xlsx** file automatically.

---

## Requirements

Install the required package:

```bash
pip install openpyxl
```

Tkinter is included with the standard Python installation.

---

## Run

```bash
python script.py
```

---

## Highlights

- Automatically evaluates Excel formulas.
- Works even if Excel has not stored cached formula values.
- Supports merged cells.
- Verifies calculations before exporting.
- Greatly reduces manual work in preparing financial documentation.

---

## License

This project is intended for educational and internal organizational use.
