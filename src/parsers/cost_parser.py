"""
Парсер файла себестоимости (Excel)

Извлекает цены от поставщиков, выбирает минимальную.
Сохраняет информацию о таре из ценовых ячеек и названий.
"""

import pandas as pd
import re
from typing import Union, Optional, Tuple
from io import BytesIO


def clean_price(value) -> float:
    """Очистка и преобразование цены в число"""
    if pd.isna(value) or value is None:
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    if s in ['-', ' -', '  -', '', 'неактуал.']:
        return 0.0

    match = re.search(r'(\d+[,.]?\d*)', s.replace(' ', ''))
    if match:
        price_str = match.group(1).replace(',', '.')
        return float(price_str)

    return 0.0


def extract_packaging_from_cell(value) -> str:
    """
    Извлекает тару из ценовой ячейки.
    Примеры: "62,5(400г)" → "400г", "50,5 (800г)" → "800г"
    """
    if pd.isna(value) or value is None:
        return ''

    if isinstance(value, (int, float)):
        return ''

    s = str(value).strip()

    # Ищем тару в скобках: (400г), (1кг), (0.5л), (800г)
    pkg_match = re.search(r'\(\s*(\d+[,.]?\d*\s*(?:гр|г|кг|мл|л|шт)\.?)\s*\)', s, re.IGNORECASE)
    if pkg_match:
        return pkg_match.group(1).strip()

    return ''


def extract_packaging_from_name(name: str) -> str:
    """
    Извлекает тару из названия товара.
    Работает и с текстом в скобках, и без.

    Примеры:
    - "Молоко сгущенное 270гр." → "270гр"
    - "Мука пшеничная 1кг." → "1кг"
    - "Масло подсолнечное 5л" → "5л"
    - "Кефир 0,9л" → "0,9л"
    - "Гречневая крупа" → "" (нет тары)
    """
    if not name:
        return ''

    s = str(name).strip()

    # Ищем вес/объём: 270гр, 1кг, 400г, 5л, 200мл, 0.5л, 0,9л, 1.5г
    match = re.search(r'(\d+[,.]?\d*)\s*(гр|г|кг|мл|л|шт)\.?', s, re.IGNORECASE)
    if match:
        return match.group(0).rstrip('.')

    return ''


def parse_cost_file(file: Union[str, BytesIO]) -> pd.DataFrame:
    """
    Парсинг файла себестоимости

    Структура файла:
    - Колонка 0: номер
    - Колонка 1: наименование (с тарой в названии)
    - Колонка 2: цена поставщика 1 (может быть "62,5(400г)")
    - Колонка 3: цена поставщика 2 (может быть "50,5(800г)")
    - Колонка 4: цена поставщика 3 (обычно число)
    """
    try:
        df = pd.read_excel(file, header=None)
    except Exception:
        df = pd.read_excel(file, header=None, engine='xlrd')

    # Ищем строку заголовка
    header_row = None
    for idx, row in df.iterrows():
        row_str = ' '.join(str(x) for x in row if pd.notna(x)).lower()
        if 'наименование' in row_str or 'спецификация' in row_str:
            header_row = idx
            break

    if header_row is None:
        header_row = 0

    data_rows = []
    for idx in range(header_row + 1, len(df)):
        row = df.iloc[idx]

        if pd.isna(row.iloc[1]) or str(row.iloc[1]).strip() == '':
            continue

        name = str(row.iloc[1]).strip()
        if not name:
            continue

        # Получаем цены и тару из ценовых ячеек
        raw1 = row.iloc[2] if len(row) > 2 else None
        raw2 = row.iloc[3] if len(row) > 3 else None
        raw3 = row.iloc[4] if len(row) > 4 else None

        price1 = clean_price(raw1)
        price2 = clean_price(raw2)
        price3 = clean_price(raw3)

        pkg1 = extract_packaging_from_cell(raw1)
        pkg2 = extract_packaging_from_cell(raw2)
        pkg3 = extract_packaging_from_cell(raw3)

        # Тара из названия (fallback)
        name_pkg = extract_packaging_from_name(name)

        # Собираем пары (цена, тара) для каждого поставщика
        price_pkg_pairs = []
        if price1 > 0:
            price_pkg_pairs.append((price1, pkg1 or name_pkg))
        if price2 > 0:
            price_pkg_pairs.append((price2, pkg2 or name_pkg))
        if price3 > 0:
            price_pkg_pairs.append((price3, pkg3 or name_pkg))

        if not price_pkg_pairs:
            continue

        # Минимальная цена и тара той цены
        min_pair = min(price_pkg_pairs, key=lambda x: x[0])
        cost = min_pair[0]
        packaging = min_pair[1]

        data_rows.append({
            'Наименование': name,
            'Себестоимость': round(cost, 2),
            'Тара': packaging,
            'Цена1': price1,
            'Цена2': price2,
            'Цена3': price3
        })

    result = pd.DataFrame(data_rows)

    # Убираем строки без себестоимости
    result = result[result['Себестоимость'] > 0].reset_index(drop=True)

    # Статистика
    with_pkg = len(result[result['Тара'] != '']) if 'Тара' in result.columns else 0
    print(f"  Себестоимость: {len(result)} позиций, с тарой: {with_pkg}")
    # Показать тару для отладки
    for _, r in result.iterrows():
        if r['Тара']:
            print(f"    {r['Наименование'][:40]} → себес {r['Себестоимость']} за [{r['Тара']}]")

    return result


def normalize_product_name(name: str) -> str:
    """Нормализация названия товара для матчинга"""
    s = name.lower()
    s = re.sub(r'\([^)]*\)', '', s)
    s = re.sub(r'\d+\s*(гр?|кг|мл|л|шт)\.?', '', s)
    s = ' '.join(s.split())
    return s.strip()
