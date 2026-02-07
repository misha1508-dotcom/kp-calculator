"""
Расчёт цен для КП

Стратегия: КОНТРАКТНАЯ ОПТИМИЗАЦИЯ
- Общая сумма контракта ровно 0.1% ниже конкурента
- Скидка достигается за счёт снижения цен на самые низкомаржинальные товары
- Максимизация общей маржи
- Если маржа минусовая — показываем как есть
"""

import pandas as pd


def calculate_prices(
    df: pd.DataFrame,
    target_discount_percent: float = 0.1
) -> pd.DataFrame:
    """
    Расчёт наших цен с контрактной оптимизацией.

    Алгоритм:
    1. Начальная цена = цена конкурента - 0.01 (строго ниже)
    2. Считаем сколько нужно скинуть от общей суммы (0.1%)
    3. Снижаем цены на самых низкомаржинальных товарах (у которых есть запас)
    4. Маржа может быть минусовой — это нормально

    Args:
        df: DataFrame с колонками: Наименование, Кол-во, Себестоимость, Цена конкурента
        target_discount_percent: Целевая скидка от суммы конкурента (по умолчанию 0.1%)

    Returns:
        DataFrame с добавленными колонками: Наша цена, Сумма, Маржа, Маржа %
    """
    result = df.copy()

    if 'Цена конкурента' not in result.columns:
        result['Цена конкурента'] = 0
    if 'Себестоимость' not in result.columns:
        result['Себестоимость'] = 0

    # Шаг 1: Начальная цена
    # Если есть конкурент — цена конкурента - 0.01
    # Если нет конкурента — себестоимость + 30%
    for idx in result.index:
        competitor_price = float(result.at[idx, 'Цена конкурента'] or 0)
        cost = float(result.at[idx, 'Себестоимость'] or 0)

        if competitor_price <= 0:
            # Нет конкурента — себестоимость + 30%
            result.at[idx, 'Наша цена'] = round(cost * 1.30, 2) if cost > 0 else 0
        else:
            result.at[idx, 'Наша цена'] = round(competitor_price - 0.01, 2)

    # Шаг 2: Скидка 0.1% — только по позициям с ценой конкурента
    has_comp = result['Цена конкурента'] > 0
    comp_rows = result[has_comp]

    if len(comp_rows) > 0:
        competitor_total = (comp_rows['Цена конкурента'] * comp_rows['Кол-во']).sum()
        target_total = competitor_total * (1 - target_discount_percent / 100)
        our_comp_total = (comp_rows['Наша цена'] * comp_rows['Кол-во']).sum()
        delta = our_comp_total - target_total

        if delta > 0:
            # Шаг 3: Маржинальность только позиций с конкурентом
            margins = []
            for idx in comp_rows.index:
                our_price = float(result.at[idx, 'Наша цена'])
                cost = float(result.at[idx, 'Себестоимость'] or 0)
                qty = float(result.at[idx, 'Кол-во'] or 0)

                if our_price > 0 and cost > 0 and qty > 0:
                    margin_pct = (our_price - cost) / our_price
                    max_reduction_per_unit = max(0, our_price - cost)
                    max_reduction_total = max_reduction_per_unit * qty
                else:
                    margin_pct = -999
                    max_reduction_total = 0

                margins.append({
                    'idx': idx,
                    'margin_pct': margin_pct,
                    'qty': qty,
                    'max_reduction_total': max_reduction_total
                })

            margins.sort(key=lambda x: x['margin_pct'])

            # Шаг 4: Распределяем дельту по низкомаржинальным
            remaining_delta = delta
            for item in margins:
                if remaining_delta <= 0.01:
                    break
                if item['max_reduction_total'] <= 0:
                    continue

                reduction = min(remaining_delta, item['max_reduction_total'])
                price_reduction = reduction / item['qty'] if item['qty'] > 0 else 0

                idx = item['idx']
                current_price = float(result.at[idx, 'Наша цена'])
                cost = float(result.at[idx, 'Себестоимость'] or 0)
                new_price = max(current_price - price_reduction, cost)
                new_price = round(new_price, 2)
                result.at[idx, 'Наша цена'] = new_price

                actual_reduction = (current_price - new_price) * item['qty']
                remaining_delta -= actual_reduction

            if remaining_delta > 0.5:
                print(f"  ⚠️ Не удалось полностью достичь целевой скидки, остаток дельты: {remaining_delta:.2f} руб")

    # Шаг 5: Финальная проверка — наша цена строго < конкурента (где есть конкурент)
    # Если себестоимость > конкурента — маржа минусовая, но цена ВСЕГДА ниже конкурента
    for idx in result.index:
        competitor_price = float(result.at[idx, 'Цена конкурента'] or 0)
        if competitor_price > 0:
            our_price = float(result.at[idx, 'Наша цена'])
            if our_price >= competitor_price:
                result.at[idx, 'Наша цена'] = round(competitor_price - 0.01, 2)

    # Шаг 6: Финальная коррекция скидки (компенсация ошибок округления)
    if len(comp_rows) > 0:
        remaining_correction = (result.loc[has_comp, 'Наша цена'] * result.loc[has_comp, 'Кол-во']).sum() - target_total

        if abs(remaining_correction) > 0.5:
            # Сортируем позиции по маржинальности (от макс к мин) для коррекции
            correction_candidates = []
            for idx in result.loc[has_comp].index:
                p = float(result.at[idx, 'Наша цена'])
                c = float(result.at[idx, 'Себестоимость'] or 0)
                q = float(result.at[idx, 'Кол-во'] or 0)
                comp_p = float(result.at[idx, 'Цена конкурента'] or 0)
                if q > 0 and p > c:
                    correction_candidates.append({
                        'idx': idx, 'margin': (p - c) / p, 'qty': q,
                        'price': p, 'cost': c, 'comp': comp_p
                    })
            correction_candidates.sort(key=lambda x: x['margin'], reverse=True)

            for item in correction_candidates:
                if abs(remaining_correction) <= 0.01:
                    break
                q = item['qty']
                per_unit = remaining_correction / q
                old_p = item['price']
                new_p = round(old_p - per_unit, 2)
                # Не ниже себестоимости и не выше конкурента
                new_p = max(new_p, item['cost'])
                if item['comp'] > 0:
                    new_p = min(new_p, round(item['comp'] - 0.01, 2))
                result.at[item['idx'], 'Наша цена'] = new_p
                actual = (old_p - new_p) * q
                remaining_correction -= actual

    # Рассчитываем итоговые показатели (маржа может быть минусовой)
    result['Сумма'] = (result['Наша цена'] * result['Кол-во']).round(2)
    result['Маржа'] = (result['Наша цена'] - result['Себестоимость']).round(2)
    result['Маржа %'] = (result['Маржа'] / result['Наша цена'] * 100).replace([float('inf'), float('-inf')], 0).fillna(0).round(1)

    if '№' in result.columns:
        result = result.sort_values('№').reset_index(drop=True)

    return result


def recalculate_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Пересчёт итогов после ручного редактирования"""
    result = df.copy()
    result['Сумма'] = (result['Наша цена'] * result['Кол-во']).round(2)
    result['Маржа'] = (result['Наша цена'] - result['Себестоимость']).round(2)
    result['Маржа %'] = (result['Маржа'] / result['Наша цена'] * 100).replace([float('inf'), float('-inf')], 0).fillna(0).round(1)
    return result
