import os
import sys
from pathlib import Path
from docx import Document
from docx_replace_ms import docx_replace
import openpyxl
from datetime import datetime
import configparser


def load_config():
    """Загружает конфигурацию из INI файла с валидацией значений."""

    config = configparser.ConfigParser()
    config_file_path = os.path.join(os.getcwd(), 'WordGenFromExcel.ini')

    if not os.path.exists(config_file_path):
        print("INI файл не найден!")
        input("Нажмите Enter для выхода ...")
        exit(1)
    try:
        config.read(config_file_path, encoding='utf-8')
        template_name = config.get('PATHS', 'template_name')
        data_file_name = config.get('PATHS', 'data_file_name')

        # валидация
        if not template_name or not isinstance(template_name, str):
            raise ValueError("Название файла шаблона договора в ini файле указано не корректно")
        if not data_file_name or not isinstance(data_file_name, str):
            raise ValueError("Название файла эксел с данными в ini файле указано не корректно")
        if not template_name.lower().endswith('.docx'):
            raise ValueError("Название файла с шаблоном договора должен оканчиваться на .docx")
        if not data_file_name.lower().endswith('.xlsx'):
            raise ValueError("Название файла с данными экселе должен оканчиваться на .xlsx")
        if not os.path.exists(os.path.join(os.getcwd(), template_name)):
            raise ValueError(f"Не найден файл шаблона договора - {template_name}")
        if not os.path.exists(os.path.join(os.getcwd(), data_file_name)):
            raise ValueError(f"Не найден файл эксел с данными - {data_file_name}")

        return template_name, data_file_name

    except Exception as e:
        print(f"Ошибка: {e}.")
        print("Проверьте ini файл на корректное заполнение или убедитесь в наличии файла шаблона и файла с данными!")
        input("Нажмите Enter для выхода ...")
        exit(1)


def main():
    # Загрузка конфигурации
    template_name, data_file_name = load_config()
    # Конфигурация путей
    exe_dir = os.getcwd()

    template_path = os.path.join(exe_dir, template_name)
    xlsx_path = os.path.join(exe_dir, data_file_name)

    try:
        # Работа с Excel-данными
        wb = openpyxl.load_workbook(xlsx_path)
        sheet = wb.active
        rows = list(sheet.iter_rows(values_only=True))

        if not rows:
            raise ValueError("Файл Excel не содержит данных")

        headers = []
        for cell in rows[0]:
            if cell is not None and str(cell).strip() != "":
                headers.append(str(cell))
            else:
                break
        columns_count = len(headers)
        if columns_count == 0:
            raise ValueError("Все ячейки верхней строки Excel файла должны быть заполнены!")

        for row_idx, row in enumerate(rows[1:], 1):
            # Нормализация данных
            row_data = []
            for cell in row[:columns_count]:
                if cell is None:
                    row_data.append("")
                elif isinstance(cell, datetime):
                    row_data.append(cell.strftime("%d.%m.%Y"))
                else:
                    row_data.append(str(cell).strip() or "-")

            # Дополнение данных до количества столбцов
            row_data += [""] * (columns_count - len(row_data))

            # Извлечение имени документа
            doc_name = row_data[0].strip() or f"row_{row_idx}"
            if not doc_name:
                print(f"Предупреждение: Пустое имя в строке {row_idx}, пропуск")
                continue

            # Загрузка шаблона
            doc = Document(template_path)

            # Выполнение замен
            for col_idx in range(1, columns_count):
                placeholder = headers[col_idx]
                value = row_data[col_idx]
                docx_replace(doc, placeholder=value)
                print(f"Замена: {placeholder} → {value}")

            # Сохранение результата
            output_file = f"{doc_name}{Path(template_name).suffix}"
            output_path = os.path.join(exe_dir, output_file)
            doc.save(output_path)
            print(f"Создан документ: {output_path}\n")

        wb.close()

    except Exception as e:
        print(f"КРИТИЧЕСКАЯ ОШИБКА: {str(e)}", file=sys.stderr)
        input("Нажмите Enter для выхода ...")
        sys.exit(1)

if __name__ == "__main__":
    main()