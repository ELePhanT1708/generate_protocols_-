import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from tkinter import Tk, filedialog, simpledialog, messagebox
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn
from copy import deepcopy
import re
import os
from collections import defaultdict
import tempfile
import zipfile
import traceback

from mapping import mapping, mapping_protocol_name
from parse_name import extract_app_info
from replacing_substring import replace_text_with_formatting


# Настройка логирования
def setup_logging():
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / "desktop_app.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[
            RotatingFileHandler(log_file, maxBytes=5 * 1024 * 1024, backupCount=3),  # 5 MB на файл, храним 3 версии
            logging.StreamHandler()  # Вывод в консоль
        ]
    )


setup_logging()
logger = logging.getLogger(__name__)

# Путь к исходной заявке
SOURCE_DOC = "636. ООО КХ г. Дятьково.docx"
OOO_TITLE = "Общество с ограниченной ответственностью «Коммунальное хозяйство г. Дятьково»"

# Шаблоны для 5 программ
TEMPLATES = {
    '1': "templates/one_row/00. ПП Шаблон 1.docx",
    '2': "templates/one_row/00. СИЗ ШАБЛОН.docx",
    '3': "templates/one_row/00. А ШАБЛОН.docx",
    '4': "templates/one_row/00.Б ШАБЛОН.docx",
    '5': "templates/one_row/00. В ШАБЛОН.docx",
}

TEMPLATE_LIST_ATTENDANCE = "templates/one_row/00. УП пустой.docx"


def check_tables_in_file(doc):
    for idx, table in enumerate(doc.tables):
        print(f"\n=== Таблица {idx} ===")
        for row_idx, row in enumerate(table.rows):
            row_data = [cell.text.strip() for cell in row.cells]
            print(f"  Строка {row_idx}: {row_data}")


# 1. Извлечь данные из исходного файла
def parse_applications(path):
    doc = Document(path)
    rows = []
    for table in doc.tables:
        for row in table.rows[1:]:  # Пропускаем заголовок
            cells = row.cells
            try:
                fio = cells[1].text.strip()
                snils = cells[2].text.strip()
                role = cells[3].text.strip()
                programs = re.split(r'[,\s]+', cells[4].text.strip())
                if fio:  # Пропускаем пустые строки
                    rows.append({
                        'fio': fio,
                        'snils': snils,
                        'role': role,
                        'programs': programs
                    })
            except IndexError:
                continue
    return rows


# 2. Сгруппировать сотрудников по номерам программ
def group_by_program(rows):
    program_dict = defaultdict(list)
    for row in rows:
        for program in row['programs']:
            program_dict[program].append(row)
    return program_dict


def clone_row(table, row_idx: int, i: int):
    """
    Клонирует строку с индексом row_idx и возвращает ссылку на ячейки новой строки.
    """
    tbl = table._tbl
    tr = table.rows[row_idx]._tr
    new_tr = deepcopy(tr)
    if i == 1:
        tbl.remove(tr)
    tbl.append(new_tr)
    return table.rows[-1].cells


def fill_cell(cell, value, alignment='left'):
    """
    Чистая вставка текста без лишних переносов, с сохранением форматирования
    """
    cell.text = ""  # Очищаем
    p = cell.paragraphs[0]
    p.clear()

    run = p.add_run(str(value).strip())

    # Устанавливаем шрифт Times New Roman 10
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10)
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    if alignment == 'left':
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if alignment == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT


# Импортируй здесь свои функции


def generate_protocols_from_file(file_path: object, organization_name: object) -> object:
    try:
        logger.info(f"Начало обработки файла: {file_path}")
        save_path = file_path
        # Получаем директорию, где лежит исходный Excel
        base_dir = os.path.dirname(file_path)

        # Создаем временную папку для генерации DOCX файлов
        output_dir = tempfile.mkdtemp()



        # Парсим данные
        rows = parse_applications(save_path)
        logger.info(f"Найдено {len(rows)} записей")
        grouped_data = group_by_program(rows)
        logger.info(f"Данные сгруппированы по программам: {list(grouped_data.keys())}")
        app_number, org_name = extract_app_info(save_path)
        logger.info(f"Номер заявки: {app_number}, Организация: {org_name}")
        logger.info(f"Выходная директория: {output_dir}")
        generated_files = []

        for program, people in grouped_data.items():
            logger.info(f"Обработка программы {program} ({len(people)} человек)")
            try:
                doc = Document(TEMPLATES.get(program))

                if int(program) > 5:
                    doc = Document(TEMPLATES['5'])
                    OLD_THEME = "«Обучение по безопасным методам и ..."  # Сократил
                    NEW_THEME = mapping[program]
                    old_protocol_substring = mapping_protocol_name.get('5')
                    new_line = old_protocol_substring.replace("_____", app_number)
                    replace_text_with_formatting(doc, old_protocol_substring, new_line, highlight_substring=new_line)
                    replace_text_with_formatting(doc, OLD_THEME, NEW_THEME)
                else:
                    old_protocol_substring = mapping_protocol_name.get(program)
                    new_line = old_protocol_substring.replace("_____", app_number)
                    replace_text_with_formatting(doc, old_protocol_substring, new_line, highlight_substring=new_line)

                table = doc.tables[1]
                template_row_idx = 1

                for i, person in enumerate(people, start=1):
                    cells = clone_row(table, template_row_idx, i)
                    values = [
                        str(f"{i}."), person['fio'], person['role'],
                        organization_name, "удовлетворительно",
                        "Согласно Приложению № 1 к настоящему протоколу", "", ""
                    ]
                    for cell, value in zip(cells, values):
                        fill_cell(cell, value, 'right')

                safe_org = re.sub(r'[^\w\s-]', '', org_name).strip().replace(' ', '_')
                output_path = os.path.join(output_dir, f"Протокол_{app_number}_{safe_org}_Программа_{program}.docx")
                doc.save(output_path)
                generated_files.append(output_path)
                logger.info(f"✅ Файл сохранен: {output_path}")

                ## Лист посещений
                list_attendance_template = Document(TEMPLATE_LIST_ATTENDANCE)
                new_line = f"Группа {app_number}_{safe_org}_Программа_{program}"
                replace_text_with_formatting(list_attendance_template, "Группа ______________________", new_line,
                                             highlight_substring=new_line)

                list_attendance_table = list_attendance_template.tables[0]
                template_row_idx_attendance = 3

                for i, person in enumerate(people, start=1):
                    cells = clone_row(list_attendance_table, template_row_idx_attendance, i)
                    values = [str(f"{i}."), person['fio']]
                    for cell, value in zip(cells, values):
                        fill_cell(cell, value, 'left')

                output_path = os.path.join(output_dir,
                                           f"Лист_Посещении_{app_number}_{safe_org}_Программа_{program}.docx")
                list_attendance_template.save(output_path)
                generated_files.append(output_path)
                logger.info(f"✅ Лист посещения сохранен: {output_path}")

            except Exception as e:
                logger.error(f"Ошибка при обработке программы {program}: {str(e)}\n{traceback.format_exc()}")

        # Архивируем
        # Указываем финальный путь архива
        zip_filename = "protocols.zip"
        zip_path = os.path.join(base_dir, zip_filename)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
            for file_path in generated_files:
                zipf.write(file_path, os.path.basename(file_path))

        logger.info(f"Архив создан: {zip_path}")
        messagebox.showinfo("Успех", f"Протоколы сгенерированы:\n{zip_path}")

    except Exception as e:
        logger.critical(f"Критическая ошибка: {str(e)}\n{traceback.format_exc()}")
        messagebox.showerror("Ошибка", f"Ошибка генерации: {str(e)}")


def main_gui():
    root = Tk()
    root.withdraw()  # Не показывать основное окно

    file_path = filedialog.askopenfilename(
        title="Выберите файл для загрузки",
        filetypes=[("Word", "*.docx")]
    )

    if not file_path:
        messagebox.showwarning("Внимание", "Файл не выбран.")
        return

    org_name = simpledialog.askstring("Организация", "Введите название организации:")
    if not org_name:
        messagebox.showwarning("Внимание", "Название организации не указано.")
        return

    generate_protocols_from_file(file_path, org_name)


if __name__ == "__main__":
    main_gui()
