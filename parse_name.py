import os
import re

SOURCE_DOC = "636. ООО КХ г. Дятьково.docx"


def extract_app_info(filename):
    """
    Извлекает номер заявки и название организации из имени файла.
    Пример: "636. ООО КХ г. Дятьково.docx" → ("636", "ООО КХ г. Дятьково")
    """
    base = os.path.basename(filename)
    name_part = os.path.splitext(base)[0]

    match = re.match(r"(\d+)\.\s*(.+)", name_part)
    if match:
        number, org_name = match.groups()
        return number.strip(), org_name.strip()
    else:
        return "Номер", "Организация"
