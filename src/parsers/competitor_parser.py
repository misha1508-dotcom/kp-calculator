"""
Парсер КП конкурента (Word .docx)
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

    # Убираем скобки
    s = re.sub(r'[\[\]\(\)]', '', s)

    # Паттерн: "27 500 00" -> 27500.00
    match = re.match(r'^(\d{1,3})\s+(\d{3})\s+(\d{2})$', s)
    if match:
        return float(f"{match.group(1)}{match.group(2)}.{match.group(3)}")

    # Убираем пробелы (разделители тысяч)
    if ' ' in s:
        parts = s.split()
        if all(p.replace(',', '').replace('.', '').isdigit() for p in parts if p):
            s = ''.join(parts)

    s = s.replace(' ', '').replace('\xa0', '')

    # Запятая как десятичный разделитель
    if ',' in s and '.' not in s:
        if re.search(r',\d{2}$', s):
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')

    clean = re.sub(r'[^\d.]', '', s)

    # Если несколько точек, оставляем только последнюю
    if clean.count('.') > 1:
        parts = clean.split('.')
        clean = ''.join(parts[:-1]) + '.' + parts[-1]

    try:
        result = float(clean)
        if result > 10000000:
            return 0.0
        return result
    except ValueError:
        return 0.0


def calculate_confidence(name: str, qty: float, price: float, total: float) -> tuple:
    """
    Рассчитывает уверенность в корректности данных

    Returns:
        (confidence_score, issues_list)
    """
    issues = []
    score = 100

    if total > 0:
        calculated = qty * price
        diff_percent = abs(calculated - total) / total * 100
        if diff_percent > 30:
            issues.append("❌ Сумма не сходится")
            score -= 40
        elif diff_percent > 10:
            issues.append("⚠️ Сумма примерно")
            score -= 15

    garbage_chars = sum(1 for c in name if c in '[]{}|\\<>~`')
    if garbage_chars > 2:
        issues.append("⚠️ Мусор в названии")
        score -= 20

    if qty > 100000:
        issues.append("⚠️ Большое кол-во")
        score -= 15
    if price > 10000:
        issues.append("⚠️ Высокая цена")
        score -= 10

    if len(name) < 5:
        issues.append("⚠️ Короткое название")
        score -= 15

    tech_words = ['соответств', 'требован', 'технич', 'условия', 'гост', 'допуска']
    if any(w in name.lower() for w in tech_words):
        issues.append("⚠️ Похоже на описание")
        score -= 10

    return max(0, score), issues


def parse_competitor_docx(file: Union[str, BytesIO]) -> pd.DataFrame:
    """
    Парсинг КП конкурента из Word документа

    Args:
        file: Путь к файлу или BytesIO объект

    Returns:
        DataFrame с колонками: №, Наименование, Кол-во, Ед.изм., Цена, Сумма
    """
    print("Парсинг КП конкурента из Word...")

    if not HAS_DOCX:
        print("ОШИБКА: python-docx модуль недоступен")
        return pd.DataFrame(columns=['№', 'Наименование', 'Кол-во', 'Ед.изм.', 'Цена', 'Сумма'])

    dataframes = parse_docx_to_dataframes(file)

    if not dataframes:
        print("ОШИБКА: Таблицы не найдены в документе")
        return pd.DataFrame(columns=['№', 'Наименование', 'Кол-во', 'Ед.изм.', 'Цена', 'Сумма'])

    main_table = max(dataframes, key=len)
    print(f"  Выбрана таблица с {len(main_table)} строками")
    print(f"  Колонки: {list(main_table.columns)}")

    data = []

    for idx, row in main_table.iterrows():
        product_name = ""
        qty = 0
        unit = ""
        price = 0
        total = 0

        # Ищем наименование
        for col in main_table.columns:
            col_lower = str(col).lower()
            if 'наименован' in col_lower or 'товар' in col_lower or 'продукт' in col_lower:
                product_name = str(row[col]).strip()
                break

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

        # Ищем цену
        for col in main_table.columns:
            col_lower = str(col).lower()
            if 'цена' in col_lower and 'ед' in col_lower:
                price = clean_number(row[col])
                break

        # Если не нашли "цена за ед", ищем просто "цена"
        if price == 0:
            for col in main_table.columns:
                col_lower = str(col).lower()
                if 'цена' in col_lower and 'конкурент' not in col_lower:
                    price = clean_number(row[col])
                    break

        # Ищем сумму
        for col in main_table.columns:
            col_lower = str(col).lower()
            if 'сумма' in col_lower or 'итого' in col_lower or 'стоимость' in col_lower:
                total = clean_number(row[col])
                break

        # Валидация
        if not product_name or product_name == 'nan':
            continue

        if qty <= 0 or price <= 0:
            continue

        if total == 0:
            total = qty * price

        # Убираем ГОСТ и коды из названия
        product_name = re.sub(r'ГОСТ\s*[РР]?\s*[\d\-\.]+', '', product_name, flags=re.IGNORECASE)
        product_name = re.sub(r'\d{2}\.\d{2}\.\d{2}', '', product_name)
        product_name = ' '.join(product_name.split()).strip()

        if len(product_name) < 3:
            continue

        # Рассчитываем уверенность
        confidence, issues = calculate_confidence(product_name, qty, price, total)

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
            'Кол-во': qty,
            'Ед.изм.': unit,
            'Цена': price,
            'Сумма': total,
            'Уверенность': confidence,
            'Проблемы': ', '.join(issues) if issues else ''
        })

        print(f"    ✓ Найден: {product_name[:40]} | {qty} {unit} × {price} = {total}")

    result = pd.DataFrame(data)

    if len(result) == 0:
        print("ВНИМАНИЕ: Не удалось извлечь данные из Word документа")
    else:
        high_conf = len([d for d in data if d['Уверенность'] >= 80])
        low_conf = len([d for d in data if d['Уверенность'] < 50])
        print(f"  ✅ Извлечено {len(result)} позиций")
        print(f"     ├─ Уверены: {high_conf}")
        print(f"     └─ Проверить: {low_conf}")
        total_sum = result['Сумма'].sum()
        print(f"  Итого: {total_sum:,.0f} руб.")

    return result


def parse_competitor_file(file: Union[str, BytesIO], filename: str = None) -> pd.DataFrame:
    """
    Парсер КП конкурента (.docx)

    Args:
        file: Путь к файлу или BytesIO объект
        filename: Имя файла (для проверки формата)

    Returns:
        DataFrame с данными КП конкурента
    """
    if filename and not filename.lower().endswith('.docx'):
        raise ValueError(f"Неподдерживаемый формат файла: {filename}. Используйте .docx")
    return parse_competitor_docx(file)
