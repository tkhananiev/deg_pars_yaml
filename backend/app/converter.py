"""Преобразование YAML в книгу XLSX с сохранением всех уровней вложенности.

Документ обходится рекурсивно; каждое конечное (скалярное) значение становится
строкой листа. Колонки «Уровень 1..N» содержат путь до значения: ключи словарей
и порядковые номера элементов списков в виде «[1]», «[2]», …  Такое представление
не теряет ни одного уровня исходной структуры и остаётся удобным для
автофильтра Excel.
"""

import io
import re
from datetime import datetime, timezone

import yaml
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

_MIN_COLUMN_WIDTH = 12
_MAX_COLUMN_WIDTH = 70

_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
_REPEAT_FONT = Font(color="A6A6A6")  # повтор родительского уровня — приглушённый серый
_VALUE_FONT = Font(bold=False)
_LABEL_FONT = Font(bold=True)
_THIN_SIDE = Side(style="thin", color="D9D9D9")
_THIN_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)

_SHEET_TITLE_FORBIDDEN_RE = re.compile(r"[\[\]:*?/\\]")


class YamlConversionError(Exception):
    """Ошибка разбора YAML или построения XLSX."""


def _walk(node, path, rows):
    """Рекурсивно собирает пары (путь, значение) для всех конечных значений."""
    if isinstance(node, dict):
        if not node:
            rows.append((path, "{}"))
        for key, value in node.items():
            _walk(value, path + (str(key),), rows)
    elif isinstance(node, (list, tuple, set)):
        if not node:
            rows.append((path, "[]"))
        for index, item in enumerate(node, start=1):
            _walk(item, path + (f"[{index}]",), rows)
    else:
        rows.append((path, node))


def _load_all(yaml_text: str) -> list:
    return [doc for doc in yaml.safe_load_all(yaml_text) if doc is not None]


def _parse_documents(yaml_text: str) -> tuple[list, list[str]]:
    """Разбирает YAML; возвращает (документы, примечания о применённых исправлениях).

    Спецификация YAML запрещает символы табуляции, но в реальных файлах они
    встречаются внутри значений. Если строгий разбор не удался, пробуем ещё раз,
    заменив табуляции на пробелы.
    """
    notes = []
    try:
        documents = _load_all(yaml_text)
    except yaml.YAMLError as first_error:
        if "\t" not in yaml_text:
            raise YamlConversionError(f"Не удалось разобрать YAML: {first_error}") from first_error
        try:
            documents = _load_all(yaml_text.replace("\t", " "))
        except yaml.YAMLError:
            raise YamlConversionError(f"Не удалось разобрать YAML: {first_error}") from first_error
        notes.append(
            "В файле найдены символы табуляции — при разборе они заменены на пробелы."
        )
    if not documents:
        raise YamlConversionError("Файл не содержит данных YAML.")
    return documents, notes


def _write_value_cell(cell, value):
    """Записывает значение с сохранением типа, безопасно для Excel."""
    if value is None:
        cell.value = "null"
    elif isinstance(value, bool) or isinstance(value, (int, float)):
        cell.value = value
    elif isinstance(value, datetime) and value.tzinfo is not None:
        cell.value = value.isoformat()  # Excel не поддерживает даты с таймзоной
    else:
        text = str(value)
        cell.value = text
        if text.startswith("="):
            cell.data_type = "s"  # не даём Excel принять строку за формулу


def _sheet_title(file_name: str) -> str:
    base = re.sub(r"\.(ya?ml)$", "", file_name, flags=re.IGNORECASE)
    title = _SHEET_TITLE_FORBIDDEN_RE.sub("_", base).strip()
    return title[:31] or "Данные"


def _fill_data_sheet(sheet, rows, depth):
    headers = [f"Уровень {level}" for level in range(1, depth + 1)] + ["Значение"]
    total_columns = len(headers)
    column_widths = [len(header) for header in headers]

    for column, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=column, value=header)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = _HEADER_ALIGNMENT
        cell.border = _THIN_BORDER

    previous_path = ()
    for row_index, (path, value) in enumerate(rows, start=2):
        same_prefix = True
        for level in range(depth):
            text = path[level] if level < len(path) else ""
            cell = sheet.cell(row=row_index, column=level + 1, value=text)
            cell.border = _THIN_BORDER
            is_repeat = (
                same_prefix
                and level < len(previous_path)
                and level < len(path)
                and path[level] == previous_path[level]
            )
            if is_repeat:
                cell.font = _REPEAT_FONT
            else:
                same_prefix = False
            if text:
                column_widths[level] = max(column_widths[level], len(text))

        value_cell = sheet.cell(row=row_index, column=total_columns)
        _write_value_cell(value_cell, value)
        value_cell.font = _VALUE_FONT
        value_cell.border = _THIN_BORDER
        column_widths[-1] = max(column_widths[-1], len(str(value_cell.value or "")))
        previous_path = path

    for column, width in enumerate(column_widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = max(
            _MIN_COLUMN_WIDTH, min(_MAX_COLUMN_WIDTH, width + 2)
        )

    last_row = len(rows) + 1
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(total_columns)}{last_row}"


def _fill_info_sheet(sheet, *, source_url, file_name, row_count, depth, notes):
    info_rows = [
        ("Источник (URL)", source_url),
        ("Имя файла", file_name),
        ("Дата выгрузки (UTC)", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")),
        ("Строк данных", row_count),
        ("Уровней вложенности", depth),
    ]
    info_rows.extend(("Примечание", note) for note in notes)
    for row_index, (label, value) in enumerate(info_rows, start=1):
        label_cell = sheet.cell(row=row_index, column=1, value=label)
        label_cell.font = _LABEL_FONT
        sheet.cell(row=row_index, column=2, value=value)
    sheet.column_dimensions["A"].width = 24
    sheet.column_dimensions["B"].width = 90


def convert_yaml_to_xlsx(yaml_text: str, *, source_url: str, file_name: str) -> tuple[bytes, int]:
    """Строит книгу XLSX и возвращает (байты файла, число строк данных)."""
    documents, notes = _parse_documents(yaml_text)

    rows = []
    if len(documents) == 1:
        _walk(documents[0], (), rows)
    else:
        for doc_index, document in enumerate(documents, start=1):
            _walk(document, (f"Документ {doc_index}",), rows)

    if not rows:
        raise YamlConversionError("В файле не найдено ни одного значения.")

    depth = max(len(path) for path, _ in rows)

    workbook = Workbook()
    data_sheet = workbook.active
    data_sheet.title = _sheet_title(file_name)
    _fill_data_sheet(data_sheet, rows, depth)

    info_sheet = workbook.create_sheet("Инфо")
    _fill_info_sheet(
        info_sheet,
        source_url=source_url,
        file_name=file_name,
        row_count=len(rows),
        depth=depth,
        notes=notes,
    )

    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue(), len(rows)
