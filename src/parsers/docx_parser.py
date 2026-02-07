"""
Парсер таблиц из Word документов (.docx)
"""

import pandas as pd
from typing import Union, List, Dict
from io import BytesIO
import re

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


def clean_text(text: str) -> str:
    """Очистка текста от лишних пробелов"""
    if not text:
        return ""
    # Убираем множественные пробелы и переводы строк
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def extract_tables_from_docx(file: Union[str, BytesIO]) -> List[List[List[str]]]:
    """
    Извлекает все таблицы из Word документа

    Args:
        file: Путь к файлу или BytesIO объект

    Returns:
        Список таблиц, где каждая таблица - это список строк,
        а каждая строка - список ячеек (текстов)
    """
    if not HAS_DOCX:
        print("ОШИБКА: python-docx не установлен")
        return []

    try:
        doc = Document(file)
    except Exception as e:
        print(f"ОШИБКА: Не удалось открыть Word документ: {e}")
        return []

    tables_data = []

    for table_idx, table in enumerate(doc.tables):
        table_data = []

        for row_idx, row in enumerate(table.rows):
            row_data = []
            for cell in row.cells:
                cell_text = clean_text(cell.text)
                row_data.append(cell_text)
            table_data.append(row_data)

        if table_data:  # Только если таблица не пустая
            tables_data.append(table_data)
            print(f"  Таблица {table_idx + 1}: {len(table_data)} строк, {len(table_data[0]) if table_data else 0} колонок")

    return tables_data


def table_to_dataframe(table_data: List[List[str]], has_header: bool = True) -> pd.DataFrame:
    """
    Преобразует таблицу в DataFrame

    Args:
        table_data: Список строк таблицы
        has_header: Первая строка - заголовок

    Returns:
        DataFrame
    """
    if not table_data:
        return pd.DataFrame()

    if has_header and len(table_data) > 1:
        # Первая строка - заголовки
        headers = table_data[0]
        data_rows = table_data[1:]

        # Создаём DataFrame
        df = pd.DataFrame(data_rows, columns=headers)
    else:
        # Нет заголовков - генерируем имена колонок
        df = pd.DataFrame(table_data)

    return df


def find_table_by_keywords(tables: List[List[List[str]]], keywords: List[str]) -> int:
    """
    Находит таблицу, содержащую ключевые слова в заголовке

    Args:
        tables: Список таблиц
        keywords: Ключевые слова для поиска

    Returns:
        Индекс найденной таблицы или -1
    """
    for idx, table in enumerate(tables):
        if not table:
            continue

        # Проверяем первую строку (заголовок)
        header_text = ' '.join(table[0]).lower()

        # Если хотя бы одно ключевое слово найдено
        if any(kw.lower() in header_text for kw in keywords):
            return idx

    return -1


def parse_docx_to_dataframes(file: Union[str, BytesIO], keywords: List[str] = None) -> List[pd.DataFrame]:
    """
    Парсит Word документ и возвращает список DataFrame для каждой таблицы

    Args:
        file: Путь к файлу или BytesIO объект
        keywords: Ключевые слова для фильтрации нужных таблиц (опционально)

    Returns:
        Список DataFrame
    """
    print("Парсинг Word документа...")

    tables = extract_tables_from_docx(file)

    if not tables:
        print("ОШИБКА: Таблицы не найдены в документе")
        return []

    print(f"  Найдено таблиц: {len(tables)}")

    # Если указаны ключевые слова - ищем нужную таблицу
    if keywords:
        table_idx = find_table_by_keywords(tables, keywords)
        if table_idx >= 0:
            print(f"  Используем таблицу {table_idx + 1} (найдены ключевые слова)")
            df = table_to_dataframe(tables[table_idx], has_header=True)
            return [df]
        else:
            print(f"  ВНИМАНИЕ: Таблица с ключевыми словами {keywords} не найдена")

    # Преобразуем все таблицы в DataFrame
    dataframes = []
    for idx, table in enumerate(tables):
        df = table_to_dataframe(table, has_header=True)
        if len(df) > 0:
            dataframes.append(df)
            print(f"  Таблица {idx + 1} → DataFrame: {len(df)} строк, {len(df.columns)} колонок")

    return dataframes
