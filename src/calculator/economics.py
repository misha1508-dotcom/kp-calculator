"""
Расчёт экономики поставки
"""

import pandas as pd
from typing import Dict


def calculate_economics(df: pd.DataFrame) -> Dict:
    """
    Расчёт общей экономики поставки

    Args:
        df: DataFrame с колонками: Кол-во, Себестоимость, Наша цена, Цена конкурента

    Returns:
        Словарь с показателями:
        - competitor_total: Сумма конкурента
        - our_total: Наша сумма
        - cost_total: Общая себестоимость
        - profit: Прибыль
        - margin_percent: Средняя маржинальность
        - discount_percent: Скидка от конкурента
    """
    # Общая сумма контракта (все товары)
    contract_total = (df['Наша цена'] * df['Кол-во']).sum()
    cost_total = (df['Себестоимость'] * df['Кол-во']).sum()

    # Прибыль и маржа — полностью с контракта
    profit = contract_total - cost_total
    margin_percent = (profit / contract_total * 100) if contract_total > 0 else 0

    # Позиции с конкурентом — для расчёта скидки
    has_comp = df['Цена конкурента'] > 0
    comp_competitor_total = (df.loc[has_comp, 'Цена конкурента'] * df.loc[has_comp, 'Кол-во']).sum()
    comp_our_total = (df.loc[has_comp, 'Наша цена'] * df.loc[has_comp, 'Кол-во']).sum()

    if comp_competitor_total > 0:
        discount_percent = ((comp_competitor_total - comp_our_total) / comp_competitor_total) * 100
    else:
        discount_percent = 0

    # Маржа конкурента (его цена минус себестоимость по его позициям)
    comp_cost_total = (df.loc[has_comp, 'Себестоимость'] * df.loc[has_comp, 'Кол-во']).sum()
    competitor_margin = comp_competitor_total - comp_cost_total
    competitor_margin_percent = (competitor_margin / comp_competitor_total * 100) if comp_competitor_total > 0 else 0

    # Статистика по позициям
    total_positions = len(df)
    positions_with_comp = int(has_comp.sum())
    positions_without_comp = total_positions - positions_with_comp
    margin_per_item = (df['Наша цена'] - df['Себестоимость'])
    loss_positions = int((margin_per_item < 0).sum())
    loss_per_position = margin_per_item * df['Кол-во']
    loss_total_rub = float(loss_per_position[margin_per_item < 0].sum()) if loss_positions > 0 else 0
    median_loss = float(loss_per_position[margin_per_item < 0].median()) if loss_positions > 0 else 0

    return {
        'contract_total': contract_total,       # Общая сумма контракта
        'competitor_total': comp_competitor_total, # Сумма конкурента по его позициям
        'our_comp_total': comp_our_total,        # Наша цена по позициям конкурента
        'cost_total': cost_total,
        'profit': profit,                        # Товарная маржа полностью с контракта
        'margin_percent': margin_percent,        # Маржа полностью с контракта
        'discount_percent': discount_percent,    # Скидка только по конкурентным позициям
        'competitor_margin': competitor_margin,   # Маржа конкурента в рублях
        'competitor_margin_percent': competitor_margin_percent,  # Маржа конкурента в %
        'total_positions': total_positions,
        'positions_with_comp': positions_with_comp,
        'positions_without_comp': positions_without_comp,
        'loss_positions': loss_positions,
        'loss_total_rub': loss_total_rub,
        'median_loss': median_loss,
    }


def get_economics_details(df: pd.DataFrame) -> pd.DataFrame:
    """
    Детальный расчёт экономики по каждой позиции

    Args:
        df: DataFrame с данными КП

    Returns:
        DataFrame с колонками для экспорта в Excel
    """
    result = df.copy()

    # Добавляем колонки если их нет
    if 'Сумма' not in result.columns:
        result['Сумма'] = result['Наша цена'] * result['Кол-во']

    if 'Маржа' not in result.columns:
        result['Маржа'] = result['Наша цена'] - result['Себестоимость']

    if 'Маржа %' not in result.columns:
        result['Маржа %'] = (result['Маржа'] / result['Наша цена'] * 100).replace([float('inf'), float('-inf')], 0).fillna(0)

    # Прибыль по позиции
    result['Прибыль'] = result['Маржа'] * result['Кол-во']

    # Сумма конкурента по позиции
    result['Сумма конкурента'] = result['Цена конкурента'] * result['Кол-во']

    # Скидка от конкурента
    result['Скидка %'] = ((result['Цена конкурента'] - result['Наша цена']) /
                          result['Цена конкурента'] * 100).replace([float('inf'), float('-inf')], 0).fillna(0)

    # Выбираем нужные колонки
    columns = [
        '№', 'Наименование', 'Ед.изм.', 'Кол-во',
        'Себестоимость', 'Наша цена', 'Цена конкурента',
        'Сумма', 'Сумма конкурента',
        'Маржа', 'Маржа %', 'Прибыль', 'Скидка %'
    ]

    # Оставляем только существующие колонки
    columns = [c for c in columns if c in result.columns]

    return result[columns]
