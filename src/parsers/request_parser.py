"""
Парсер запросов КП (Word .docx)
Извлекает названия товаров, количества и описания из таблиц Word документов
"""

import pandas as pd
import re
from typing import Union
from io import BytesIO

try:
    from src.parsers.docx_parser import parse_docx_to_dataframes
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


def clean_number(value) -> float:
    """Преобразование значения в число"""
    if pd.isna(value) or value is None:
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    s = re.sub(r'[\[\]\(\)]', '', s)
    s = s.replace(' ', '').replace('\xa0', '')
    s = s.replace(',', '.')

    clean = re.sub(r'[^\d.]', '', s)

    try:
        result = float(clean)
        if result > 10000000:
            return 0.0
        return result
    except ValueError:
        return 0.0


def clean_product_name(text: str) -> str:
    """Очищает название продукта от лишнего"""
    text = re.sub(r'ГОСТ\s*[РР]?\s*[\d\-\.]+', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\d{2}\.\d{2}\.\d{2}', '', text)
    text = re.sub(r'Соответств\w+', '', text, flags=re.IGNORECASE)
    text = re.sub(r'требован\w+', '', text, flags=re.IGNORECASE)
    text = re.sub(r'Технич\w+', '', text, flags=re.IGNORECASE)
    text = re.sub(r'условия\w*', '', text, flags=re.IGNORECASE)
    text = ' '.join(text.split())
    return text.strip()


def calculate_request_confidence(name: str, qty: float) -> tuple:
    """
    Рассчитывает уверенность в корректности данных запроса

    Returns:
        (confidence_score, issues_list)
    """
    issues = []
    score = 100

    garbage_chars = sum(1 for c in name if c in '[]{}|\\<>~`')
    if garbage_chars > 2:
        issues.append("⚠️ Мусор в названии")
        score -= 25

    if qty > 100000:
        issues.append("⚠️ Большое кол-во")
        score -= 20

    if len(name) < 5:
        issues.append("⚠️ Короткое название")
        score -= 20

    tech_words = ['соответств', 'требован', 'технич', 'условия', 'допуска', 'школа', 'детский сад', 'гбоу', 'гбдоу']
    if any(w in name.lower() for w in tech_words):
        issues.append("❌ Похоже на описание")
        score -= 40

    address_words = ['область', 'район', 'улица', 'ул.', 'пгт', 'село', 'с.']
    if any(w in name.lower() for w in address_words):
        issues.append("❌ Похоже на адрес")
        score -= 50

    return max(0, score), issues


def parse_request_docx(file: Union[str, BytesIO]) -> pd.DataFrame:
    """
    Парсинг запроса на КП из Word документа

    Args:
        file: Путь к файлу или BytesIO объект

    Returns:
        DataFrame с колонками: №, Наименование, Описание, Ед.изм., Кол-во
    """
    print("Парсинг запроса КП из Word...")

    if not HAS_DOCX:
        print("ОШИБКА: python-docx модуль недоступен")
        return pd.DataFrame(columns=['№', 'Наименование', 'Описание', 'Ед.изм.', 'Кол-во'])

    dataframes = parse_docx_to_dataframes(file)

    if not dataframes:
        print("ОШИБКА: Таблицы не найдены в документе")
        return pd.DataFrame(columns=['№', 'Наименование', 'Описание', 'Ед.изм.', 'Кол-во'])

    main_table = max(dataframes, key=len)
    print(f"  Выбрана таблица с {len(main_table)} строками")
    print(f"  Колонки: {list(main_table.columns)}")

    data = []
    found_products = set()

    for idx, row in main_table.iterrows():
        product_name = ""
        description = ""
        qty = 0
        unit = "кг"

        # Ищем наименование товара
        for col in main_table.columns:
            col_lower = str(col).lower()
            if 'наименован' in col_lower or 'товар' in col_lower or 'продукт' in col_lower:
                product_name = str(row[col]).strip()
                break

        # Ищем описание (ГОСТ, характеристики, требования)
        for col in main_table.columns:
            col_lower = str(col).lower()
            if any(word in col_lower for word in ['описан', 'гост', 'характерист', 'техн', 'требован', 'соответств', 'поставля']):
                desc_value = str(row[col]).strip()
                if desc_value and desc_value != 'nan':
                    description += " " + desc_value

        # Если описания в отдельной колонке нет, ищем в самом названии
        if not description and product_name:
            parts = product_name.split('ГОСТ')
            if len(parts) > 1:
                product_name = parts[0].strip()
                description = 'ГОСТ' + parts[1].strip()

        # Ищем количество
        for col in main_table.columns:
            col_lower = str(col).lower()
            if 'кол' in col_lower or 'количеств' in col_lower:
                qty = clean_number(row[col])
                break

        # Ищем единицу измерения
        for col in main_table.columns:
            col_lower = str(col).lower()
            if 'ед' in col_lower and 'изм' in col_lower:
                unit_val = str(row[col]).lower().strip()
                if unit_val and unit_val != 'nan':
                    unit = unit_val
                break

        # Валидация
        if not product_name or product_name == 'nan':
            continue

        if qty <= 0:
            # Ищем количество только в колонках, похожих на "кол-во", НЕ в ценовых
            for col in main_table.columns:
                col_lower = str(col).lower()
                if 'цена' in col_lower or 'стоимость' in col_lower or 'сумма' in col_lower:
                    continue
                val = clean_number(row[col])
                if 10 < val < 100000:
                    qty = val
                    break

        if qty <= 0:
            print(f"    ⚠️ Кол-во <= 0 для «{product_name[:50]}», строка {idx} — ставим 0")
            qty = 0

        # Проверка на дубликаты — по названию как есть
        name_key = product_name.lower().strip()
        if name_key in found_products:
            print(f"    ⚠️ Дубликат: «{product_name[:50]}» — пропускаем")
            continue
        found_products.add(name_key)

        # Рассчитываем уверенность
        confidence, issues = calculate_request_confidence(product_name, qty)

        if confidence >= 80:
            check_status = "✅"
        elif confidence >= 50:
            check_status = "⚠️"
        else:
            check_status = "❌"

        data.append({
            '№': len(data) + 1,
            '⚡': check_status,
            'Наименование': product_name,
            'Описание': description.strip(),
            'Ед.изм.': unit,
            'Кол-во': qty,
            'Уверенность': confidence,
            'Проблемы': ', '.join(issues) if issues else ''
        })

        print(f"    ✓ Найден: {product_name[:40]} | {qty} {unit} | {description[:30] if description else '(нет описания)'}")

    result = pd.DataFrame(data)

    print(f"  ✅ Извлечено {len(result)} позиций из Word")

    return result


def parse_request_file(file: Union[str, BytesIO], filename: str = None) -> pd.DataFrame:
    """
    Парсер запроса КП (.docx)

    Args:
        file: Путь к файлу или BytesIO объект
        filename: Имя файла (для проверки формата)

    Returns:
        DataFrame с данными запроса
    """
    if filename and not filename.lower().endswith('.docx'):
        raise ValueError(f"Неподдерживаемый формат файла: {filename}. Используйте .docx")
    return parse_request_docx(file)
