import sys
import os


def resource_path(relative_path):
    """Получает абсолютный путь к ресурсу (работает и для PyInstaller и для IDE)."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


if __name__ == '__main__':
    template_path = resource_path("templates/one_row/00. В ШАБЛОН.docx")
