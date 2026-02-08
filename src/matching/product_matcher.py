"""
Матчинг товаров — строгий режим

Только точные совпадения названий + проверка тары.
Если тара не совпадает — пересчитываем себестоимость пропорционально и помечаем.
Если не уверен — цену конкурента не ставим, наша цена = себестоимость + 30%.
"""

import pandas as pd
import re
from typing import Optional, Tuple, Dict
from fuzzywuzzy import fuzz


def extract_packaging(name: str) -> Dict:
    """
    Извлекает информацию о таре/упаковке из названия.

    Возвращает dict с ключами:
    - weight_g: вес в граммах (400г, 1кг → 1000г)
    - volume_ml: объём в мл (1л → 1000мл, 200мл)
    - fat_pct: жирность (2.5%, 3,2%)
    - count: штуки (10шт)
    """
    if not name:
        return {}
    s = str(name).lower()
    result = {}

    # Вес: 400г, 400 г, 1кг, 1 кг, 270гр, 0.5кг
    weight_matches = re.findall(r'(\d+[,.]?\d*)\s*(гр|г|кг)\b\.?', s)
    if weight_matches:
        val_str, unit = weight_matches[0]
        val = float(val_str.replace(',', '.'))
        if unit == 'кг':
            val *= 1000
        result['weight_g'] = val

    # Объём: 1л, 0.5л, 200мл, 1 л
    volume_matches = re.findall(r'(\d+[,.]?\d*)\s*(мл|л)\b\.?', s)
    if volume_matches:
        val_str, unit = volume_matches[0]
        val = float(val_str.replace(',', '.'))
        if unit == 'л':
            val *= 1000
        result['volume_ml'] = val

    # Жирность: 2.5%, 3,2%
    fat_matches = re.findall(r'(\d+[,.]?\d*)\s*%', s)
    if fat_matches:
        result['fat_pct'] = float(fat_matches[0].replace(',', '.'))

    # Штуки: 10шт, 10 шт
    count_matches = re.findall(r'(\d+)\s*шт\.?', s)
    if count_matches:
        result['count'] = int(count_matches[0])

    return result


def packaging_compatible(pkg1: Dict, pkg2: Dict) -> bool:
    """
    Проверяет совместимость тары двух товаров.
    Если у обоих есть параметр и он различается — несовместимы.
    Если одна тара есть, а другой нет — несовместимы.
    """
    # Если обе пусты — совместимы (нет данных для сравнения)
    # Если одна пуста, другая нет — несовместимы
    has1 = bool(pkg1) and ('weight_g' in pkg1 or 'volume_ml' in pkg1)
    has2 = bool(pkg2) and ('weight_g' in pkg2 or 'volume_ml' in pkg2)
    if has1 != has2:
        return False

    if 'weight_g' in pkg1 and 'weight_g' in pkg2:
        if abs(pkg1['weight_g'] - pkg2['weight_g']) > 1:
            return False

    if 'volume_ml' in pkg1 and 'volume_ml' in pkg2:
        if abs(pkg1['volume_ml'] - pkg2['volume_ml']) > 1:
            return False

    if 'fat_pct' in pkg1 and 'fat_pct' in pkg2:
        if abs(pkg1['fat_pct'] - pkg2['fat_pct']) > 0.1:
            return False

    if 'count' in pkg1 and 'count' in pkg2:
        if pkg1['count'] != pkg2['count']:
            return False

    return True


def calc_packaging_ratio(target_pkg: Dict, source_pkg: Dict) -> Optional[float]:
    """
    Считает коэффициент пересчёта цены по таре.

    Например: запрос 0.5л, себестоимость за 1л → ratio = 0.5
    Запрос 1кг, себестоимость за 400г → ratio = 2.5

    Returns: коэффициент или None если нельзя пересчитать
    """
    # По весу
    if 'weight_g' in target_pkg and 'weight_g' in source_pkg:
        if source_pkg['weight_g'] > 0:
            return target_pkg['weight_g'] / source_pkg['weight_g']

    # По объёму
    if 'volume_ml' in target_pkg and 'volume_ml' in source_pkg:
        if source_pkg['volume_ml'] > 0:
            return target_pkg['volume_ml'] / source_pkg['volume_ml']

    return None


def unit_to_base(unit_str: str) -> Optional[Dict]:
    """
    Переводит единицу измерения запроса в базовые единицы.

    'кг' → {'weight_g': 1000}
    'г'  → {'weight_g': 1}
    'л'  → {'volume_ml': 1000}
    'мл' → {'volume_ml': 1}
    'шт' → None (поштучно, не пересчитываем)
    """
    u = unit_str.lower().strip().rstrip('.')
    if u in ('кг',):
        return {'weight_g': 1000}
    if u in ('г', 'гр'):
        return {'weight_g': 1}
    if u in ('л',):
        return {'volume_ml': 1000}
    if u in ('мл',):
        return {'volume_ml': 1}
    return None


def adjust_cost_for_unit(cost_price: float, cost_tara: str, request_unit: str) -> Tuple[float, str]:
    """
    Пересчёт себестоимости из тары в единицу измерения запроса.

    Примеры:
    - себес 50.5 за [800г], ед.изм. 'кг' → 50.5 * (1000/800) = 63.13 за кг
    - себес 609 за [5л], ед.изм. 'л' → 609 * (1000/5000) = 121.80 за л
    - себес 104.31 за [270гр], ед.изм. 'шт' → без пересчёта (штуки = упаковки)
    - себес 66 за [1л], ед.изм. 'л' → без пересчёта (уже совпадает)
    """
    if not cost_tara or not request_unit:
        return cost_price, ''

    tara_pkg = extract_packaging(cost_tara)
    if not tara_pkg:
        return cost_price, ''

    target_base = unit_to_base(request_unit)
    if not target_base:
        # шт, уп, бут — нет пересчёта
        return cost_price, ''

    # Считаем коэффициент
    ratio = None
    if 'weight_g' in target_base and 'weight_g' in tara_pkg:
        if tara_pkg['weight_g'] > 0:
            ratio = target_base['weight_g'] / tara_pkg['weight_g']
    elif 'volume_ml' in target_base and 'volume_ml' in tara_pkg:
        if tara_pkg['volume_ml'] > 0:
            ratio = target_base['volume_ml'] / tara_pkg['volume_ml']

    if ratio is None:
        # Несовпадение типов: вес↔объём
        tara_type = 'вес' if 'weight_g' in tara_pkg else 'объём' if 'volume_ml' in tara_pkg else '?'
        unit_type = 'вес' if 'weight_g' in target_base else 'объём' if 'volume_ml' in target_base else '?'
        if tara_type != unit_type:
            note = f"Несовпадение: тара [{cost_tara}] ({tara_type}) vs ед.изм. [{request_unit}] ({unit_type})"
            return cost_price, note
        return cost_price, ''

    if abs(ratio - 1.0) < 0.001:
        # Уже совпадает — без пересчёта
        return cost_price, ''

    adjusted = round(cost_price * ratio, 2)
    note = f"{cost_price} за [{cost_tara}] → {adjusted} за [{request_unit}] (×{ratio:.3f})"
    return adjusted, note


def format_packaging(pkg: Dict) -> str:
    """Форматирует информацию о таре для отображения"""
    parts = []
    if 'weight_g' in pkg:
        w = pkg['weight_g']
        if w >= 1000:
            if w % 1000 != 0:
                parts.append(f"{w/1000:.1f}".rstrip('0').rstrip('.') + 'кг')
            else:
                parts.append(f"{int(w/1000)}кг")
        else:
            parts.append(f"{int(w)}г" if w == int(w) else f"{w}г")
    if 'volume_ml' in pkg:
        v = pkg['volume_ml']
        if v >= 1000:
            if v % 1000 != 0:
                parts.append(f"{v/1000:.1f}".rstrip('0').rstrip('.') + 'л')
            else:
                parts.append(f"{int(v/1000)}л")
        else:
            parts.append(f"{int(v)}мл" if v == int(v) else f"{v}мл")
    if 'fat_pct' in pkg:
        parts.append(f"{pkg['fat_pct']}%")
    if 'count' in pkg:
        parts.append(f"{pkg['count']}шт")
    return ', '.join(parts)


def normalize_name(name: str) -> str:
    """Нормализация: убираем вес, объём, жирность, скобки, лишнее"""
    if not name:
        return ""
    s = str(name).lower().strip()
    # Убираем вес/объём: 270гр, 1кг, 400г, 5л, 200мл
    s = re.sub(r'\d+[,.]?\d*\s*(гр|г|кг|мл|л|шт)\.?', '', s)
    # Убираем жирность: 2.5%, 3,2%
    s = re.sub(r'\d+[,.]?\d*\s*%', '', s)
    # Убираем скобки и содержимое
    s = re.sub(r'[(\[«"][^)\]»"]*[)\]»"]', '', s)
    # Убираем кавычки
    s = re.sub(r'[«»""\']', '', s)
    # Убираем числа-номера в начале
    s = re.sub(r'^\d+\s*\.?\s*', '', s)
    # Множественные пробелы
    s = ' '.join(s.split())
    return s.strip()


def get_row_packaging(row, name_col: str = 'Наименование') -> Dict:
    """
    Извлекает тару из строки DataFrame.
    Сначала смотрит колонку 'Тара' (из парсера себестоимости),
    потом извлекает из названия.
    """
    # Колонка "Тара" из cost_parser (например "400г")
    tara_col = str(row.get('Тара', '') or '')
    if tara_col:
        pkg = extract_packaging(tara_col)
        if pkg:
            return pkg

    # Извлекаем из названия
    candidate_name = str(row.get(name_col, ''))
    return extract_packaging(candidate_name)


def find_best_match(target_name: str, candidates: pd.DataFrame,
                    name_col: str = 'Наименование') -> Tuple[Optional[pd.Series], int, str]:
    """
    Строгий матчинг — название + предпочтение совпадения тары.

    1. Fuzzy-сравнение нормализованных названий (порог 90)
    2. Предпочитаем совместимые по таре
    3. Если совместимых нет — всё равно возвращаем лучший (тару пересчитаем)
    """
    if candidates is None or len(candidates) == 0:
        return None, 0, ''

    target_norm = normalize_name(target_name)
    if not target_norm:
        return None, 0, ''

    target_pkg = extract_packaging(target_name)

    matches = []

    for idx, row in candidates.iterrows():
        candidate_name = str(row.get(name_col, ''))
        candidate_norm = normalize_name(candidate_name)

        if not candidate_norm:
            continue

        ratio = fuzz.ratio(target_norm, candidate_norm)
        token_sort = fuzz.token_sort_ratio(target_norm, candidate_norm)
        token_set = fuzz.token_set_ratio(target_norm, candidate_norm)

        score = max(ratio, token_sort, token_set)

        if score >= 90:
            # Тара: сначала из колонки "Тара", потом из названия
            candidate_pkg = get_row_packaging(row, name_col)
            is_compatible = packaging_compatible(target_pkg, candidate_pkg)

            matches.append({
                'row': row,
                'name': candidate_name,
                'score': score,
                'pkg': candidate_pkg,
                'compatible': is_compatible,
            })

    if not matches:
        # Нет кандидатов выше порога
        best_score = 0
        best_name = ''
        for idx, row in candidates.iterrows():
            candidate_name = str(row.get(name_col, ''))
            candidate_norm = normalize_name(candidate_name)
            if not candidate_norm:
                continue
            score = max(
                fuzz.ratio(target_norm, candidate_norm),
                fuzz.token_sort_ratio(target_norm, candidate_norm),
                fuzz.token_set_ratio(target_norm, candidate_norm),
            )
            if score > best_score:
                best_score = score
                best_name = candidate_name
        return None, best_score, best_name

    # Предпочитаем совместимые по таре
    compatible_matches = [m for m in matches if m['compatible']]

    if compatible_matches:
        best = max(compatible_matches, key=lambda m: m['score'])
        return best['row'], best['score'], best['name']

    # Совместимых нет — берём лучший несовместимый (тару пересчитаем в match_products)
    best = max(matches, key=lambda m: m['score'])
    return best['row'], best['score'], best['name']


def match_products(
    request_df: pd.DataFrame,
    cost_df: pd.DataFrame,
    competitor_df: pd.DataFrame
) -> pd.DataFrame:
    """
    Строгий матчинг: точные совпадения + проверка тары.
    Если тара не совпадает — пересчитываем себестоимость пропорционально.
    Если нет матча конкурента — наша цена = себестоимость + 30%.
    """
    result_data = []

    print(f"  Входных позиций: {len(request_df)}")

    for idx, request_row in request_df.iterrows():
        product_name = str(request_row.get('Наименование', '') or '')
        qty = float(request_row.get('Кол-во', 0) or 0)
        description = str(request_row.get('Описание', '') or '')
        unit = str(request_row.get('Ед.изм.', 'кг') or 'кг')

        if not product_name or product_name == 'nan':
            print(f"  ⚠️ Пустое название, строка {idx}")
            product_name = f"(позиция {idx + 1} — без названия)"

        qty_note = ''
        if qty > 100000:
            old_qty = qty
            qty = qty / 1000 if qty > 1000000 else qty / 100
            qty_note = f"⚠️ Кол-во: {old_qty:.0f}→{qty:.0f}"
            print(f"  ⚠️ Кол-во скорректировано: {old_qty} → {qty} для «{product_name[:50]}»")

        # Ищем себестоимость
        cost_match, cost_score, cost_name = find_best_match(product_name, cost_df, 'Наименование')
        cost_price = 0
        tara_note = ''

        if cost_match is not None:
            cost_price = float(cost_match.get('Себестоимость', 0) or 0)
            cost_tara_str = str(cost_match.get('Тара', '') or '')

            # Пересчитываем себестоимость в единицу измерения запроса
            if cost_tara_str:
                adjusted_price, adjust_note = adjust_cost_for_unit(cost_price, cost_tara_str, unit)
                if adjust_note:
                    # Был пересчёт
                    tara_note = f"⚠️ {adjust_note}"
                    print(f"  ⚠️ Пересчёт: «{cost_name[:40]}» {adjust_note}")
                    cost_price = adjusted_price
                else:
                    # Совпадает или нет данных для пересчёта
                    tara_note = f"Себес за [{cost_tara_str}]"

        if cost_price <= 0:
            print(f"  ⚠️ Нет себестоимости: {product_name[:50]} (лучший: {cost_name[:40]}, скор: {cost_score})")
            tara_note = f"❌ Себес не найдена (лучший: {cost_name[:40]}, скор: {cost_score})"

        # Ищем цену конкурента — строго
        comp_match, comp_score, comp_name = find_best_match(product_name, competitor_df, 'Наименование')
        competitor_price = 0
        has_competitor = False

        if comp_match is not None:
            competitor_price = float(comp_match.get('Цена', 0) or 0)
            if competitor_price > 0 and competitor_price < 100000:
                # Проверяем тару конкурента тоже
                target_pkg = extract_packaging(product_name)
                comp_pkg = extract_packaging(comp_name)
                comp_pkg_ok = packaging_compatible(target_pkg, comp_pkg)

                if not comp_pkg_ok:
                    ratio = calc_packaging_ratio(target_pkg, comp_pkg)
                    if ratio is not None and ratio > 0:
                        original_comp = competitor_price
                        competitor_price = round(competitor_price * ratio, 2)
                        comp_tara = (
                            f" | Конк.тара: [{format_packaging(comp_pkg)}]→×{ratio:.2f} ({original_comp}→{competitor_price})"
                        )
                        tara_note = (tara_note + comp_tara) if tara_note else f"⚠️{comp_tara.lstrip(' |')}"
                    else:
                        # Не можем пересчитать цену конкурента — не используем её
                        competitor_price = 0
                        t_info = format_packaging(target_pkg) if target_pkg else '?'
                        c_info = format_packaging(comp_pkg) if comp_pkg else '?'
                        extra = f" | Конк: тара не совпадает [{t_info}] vs [{c_info}]"
                        tara_note = (tara_note + extra) if tara_note else f"⚠️{extra.lstrip(' |')}"

                if competitor_price > 0:
                    has_competitor = True
            else:
                competitor_price = 0

        # Формируем инфо о матче — только если есть конкурент
        if has_competitor:
            match_info = f"Себес: {cost_name[:30]}({cost_score}) | Конк: {comp_name[:30]}({comp_score})"
        else:
            match_info = ''

        # Название всегда из запроса КП
        clean_name = product_name

        result_data.append({
            '№': len(result_data) + 1,
            'Наименование': clean_name,
            'Описание': description,
            'Ед.изм.': unit,
            'Кол-во': qty,
            'Себестоимость': cost_price,
            'Цена конкурента': competitor_price if has_competitor else 0,
            'Есть конкурент': has_competitor,
            'Матч': match_info,
            'Тара': (qty_note + ' | ' + tara_note).strip(' |') if qty_note else tara_note,
        })

    result = pd.DataFrame(result_data)

    # Статистика
    with_comp = len([r for r in result_data if r['Есть конкурент']])
    without_comp = len([r for r in result_data if not r['Есть конкурент']])
    tara_issues = len([r for r in result_data if r.get('Тара', '')])
    print(f"\nМатчинг завершён:")
    print(f"  Позиций в запросе: {len(request_df)}")
    print(f"  С ценой конкурента: {with_comp}")
    print(f"  Без конкурента (себес+30%): {without_comp}")
    print(f"  Вопросы по таре: {tara_issues}")
    print(f"  Итого: {len(result)}")

    return result
