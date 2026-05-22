# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Что делает проект

Инструмент для автоматической генерации пакета документов по охране труда из заявки на обучение (`.docx`). На вход — файл заявки, на выход — ZIP-архив с протоколами по программам, листом посещаемости и согласием.

Есть два варианта запуска:
- **Desktop-приложение** (`desktop_app.py`) — Tkinter GUI, собирается в `.exe` через PyInstaller
- **FastAPI-сервер** (`app.py`) — REST API, принимает файл + organization_name через multipart/form-data

## Запуск

```powershell
# Установить зависимости
.\venv\Scripts\python.exe -m pip install -r requirements.txt

# Запустить FastAPI-сервер
.\venv\Scripts\python.exe app.py
# Swagger UI: http://localhost:8000/docs

# Запустить desktop GUI
.\venv\Scripts\python.exe desktop_app.py

# Собрать .exe
.\venv\Scripts\pyinstaller.exe desktop_app.spec
```

## Архитектура

### Основной поток данных

1. **Парсинг заявки** (`parse_applications`) — читает `.docx`, извлекает из таблицы: ФИО (col 1), СНИЛС (col 2), должность (col 3), номера программ (col 4). Первая строка таблицы — заголовок, пропускается.

2. **Группировка по программам** (`group_by_program`) — разбивает сотрудников по номерам программ 1–5. Программы >5 мэппятся в программу 5 (специфика бизнес-логики, уникальность по fio+snils).

3. **Генерация документов** — для каждой программы:
   - Берётся шаблон из `templates/one_row/`
   - Через `replace_text_with_formatting` вставляется номер договора в заголовок протокола
   - Строки в таблицу добавляются через `clone_row` (клонирование строки-образца, индекс 1)
   - `fill_cell` вставляет текст с Times New Roman 10pt без лишних переносов

4. **Дополнительные документы** — лист посещаемости (уникальные сотрудники по fio+snils) и согласие

5. **Архивирование** — результат пакуется в ZIP

### Ключевые модули

| Файл | Роль |
|---|---|
| `app.py` | FastAPI-обёртка: принимает файл, возвращает ZIP в ответе |
| `desktop_app.py` | Tkinter GUI: диалог выбора файла, вызывает ту же логику |
| `mapping.py` | Словари: `mapping` (программа >5 → название темы), `mapping_protocol_name` (программа → шаблон номера протокола) |
| `parse_name.py` | `extract_app_info` — парсит номер заявки и организацию из имени файла вида `636. ООО КХ г. Дятьково.docx` |
| `replacing_substring.py` | `replace_text_with_formatting` — замена текста в DOCX с сохранением форматирования runs |
| `replacing_theme.py` | Устаревшая версия `replace_text_with_formatting` без поддержки bold-подсветки, не используется в продакшне |
| `main.py` | Прототип/sandbox, не используется |
| `clone_row_main.py` | Экспериментальный файл, не используется |

### Шаблоны

`templates/one_row/` — актуальные шаблоны (используются в `app.py` и `desktop_app.py`):
- `00. ПП Шаблон.docx` / `00. ПП Шаблон 1.docx` — программа 1
- `00. СИЗ ШАБЛОН.docx` — программа 2
- `00. А ШАБЛОН.docx` — программа 3
- `00.Б ШАБЛОН.docx` — программа 4
- `00. В ШАБЛОН.docx` — программа 5 (и все программы >5)
- `00. УП пустой.docx` — шаблон листа посещаемости
- `00. Шаблон.docx` — шаблон согласия

`templates/empty/` — старые шаблоны, не используются.

### Особенности работы с DOCX

- Замена текста (`replace_text_with_formatting`) работает только с `doc.paragraphs`, не обходит таблицы рекурсивно — учитывай при правках шаблонов
- `clone_row` клонирует строку через deepcopy XML-элемента, при первом вызове (`i==1`) удаляет строку-образец
- Номер договора вставляется в Desktop-версии через `simpledialog`, в API-версии извлекается из имени файла через `parse_name.extract_app_info`

### PyInstaller

Spec-файл `desktop_app.spec` включает папку `templates` в бандл (`datas=[('templates', 'templates')]`). `resource_path()` в `desktop_app.py` корректно резолвит пути как в `.py`, так и в `.exe` через `sys._MEIPASS`.

Актуальный spec — `desktop_app.spec`. Остальные `.spec`-файлы в корне — исторические версии сборки.
