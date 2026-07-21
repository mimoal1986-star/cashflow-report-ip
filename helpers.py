import pandas as pd
from datetime import datetime
from typing import Tuple
import io
from models import BalanceReport

# ============================================
# РАБОТА С ДАТАМИ
# ============================================

def get_date_range(ip_ops: pd.DataFrame, phys_ops: pd.DataFrame) -> Tuple[datetime, datetime]:
    """Определяет минимальную и максимальную дату из всех операций"""
    if ip_ops.empty and phys_ops.empty:
        return datetime.now(), datetime.now()
    
    all_dates = pd.concat([
        ip_ops["date"] if not ip_ops.empty else pd.Series(dtype='datetime64[ns]'),
        phys_ops["date"] if not phys_ops.empty else pd.Series(dtype='datetime64[ns]')
    ])
    
    if all_dates.empty:
        return datetime.now(), datetime.now()
    
    return all_dates.min(), all_dates.max()


def format_date(date: datetime) -> str:
    """Форматирует дату для отображения"""
    return date.strftime("%d.%m.%Y")


def format_currency(amount: float) -> str:
    """Форматирует сумму в валютном формате"""
    return f"{amount:,.2f} ₽"


# ============================================
# ЭКСПОРТ В EXCEL
# ============================================

def create_excel_report(
    ip_report: BalanceReport,
    phys_report: BalanceReport,
    ip_operations: pd.DataFrame,
    phys_operations: pd.DataFrame
) -> io.BytesIO:
    """Создает Excel-файл с отчетом"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Сводка по ИП
        ip_summary = pd.DataFrame({
            "Показатель": ["Начальный остаток", "Конечный остаток", "Изменение"],
            "Значение": [
                ip_report.start_balance,
                ip_report.end_balance,
                ip_report.end_balance - ip_report.start_balance
            ]
        })
        ip_summary.to_excel(writer, sheet_name="ИП_Сводка", index=False)
        
        # Динамика ИП
        if not ip_report.monthly_dynamics.empty:
            ip_report.monthly_dynamics.to_excel(
                writer, sheet_name="ИП_Динамика", index=False
            )
        
        # Сводка по физлицу
        phys_summary = pd.DataFrame({
            "Показатель": ["Начальный остаток", "Конечный остаток", "Изменение"],
            "Значение": [
                phys_report.start_balance,
                phys_report.end_balance,
                phys_report.end_balance - phys_report.start_balance
            ]
        })
        phys_summary.to_excel(writer, sheet_name="Физлицо_Сводка", index=False)
        
        # Динамика физлица
        if not phys_report.monthly_dynamics.empty:
            phys_report.monthly_dynamics.to_excel(
                writer, sheet_name="Физлицо_Динамика", index=False
            )
        
        # Детальные операции
        if not ip_operations.empty:
            ip_operations.to_excel(writer, sheet_name="ИП_Операции", index=False)
        if not phys_operations.empty:
            phys_operations.to_excel(writer, sheet_name="Физлицо_Операции", index=False)
        
        # Сводный отчет
        combined = pd.DataFrame({
            "Показатель": [
                "Начальный остаток ИП",
                "Конечный остаток ИП",
                "Изменение ИП",
                "Начальный остаток Физлицо",
                "Конечный остаток Физлицо",
                "Изменение Физлицо",
                "Общий начальный остаток",
                "Общий конечный остаток",
                "Общее изменение"
            ],
            "Значение": [
                ip_report.start_balance,
                ip_report.end_balance,
                ip_report.end_balance - ip_report.start_balance,
                phys_report.start_balance,
                phys_report.end_balance,
                phys_report.end_balance - phys_report.start_balance,
                ip_report.start_balance + phys_report.start_balance,
                ip_report.end_balance + phys_report.end_balance,
                (ip_report.end_balance + phys_report.end_balance) - 
                (ip_report.start_balance + phys_report.start_balance)
            ]
        })
        combined.to_excel(writer, sheet_name="Сводный_отчет", index=False)
    
    output.seek(0)
    return output


# ============================================
# РАБОТА С ДЕПОЗИТАМИ (helper-функции)
# ============================================

def get_deposit_operations(df: pd.DataFrame) -> pd.DataFrame:
    """
    Возвращает DataFrame с депозитными операциями
    (размещение, возврат, проценты)
    
    Это HELPER, потому что:
    1. Он просто извлекает данные по меткам
    2. Не содержит бизнес-логики расчета
    3. Может использоваться в разных местах
    """
    if df.empty:
        return pd.DataFrame()
    
    # Проверяем, есть ли метки депозитов
    if "is_deposit_operation" not in df.columns:
        return pd.DataFrame()
    
    # Возвращаем все депозитные операции
    deposits = df[
        df["is_deposit_operation"] | 
        df["is_deposit_interest"]
    ].copy()
    
    return deposits


def get_non_deposit_operations(df: pd.DataFrame) -> pd.DataFrame:
    """
    Возвращает DataFrame с операциями, исключая депозитные
    (размещение и возврат исключены, проценты остаются)
    
    Это HELPER, потому что:
    1. Простая фильтрация по меткам
    2. Не содержит бизнес-логики
    3. Может использоваться в разных местах
    """
    if df.empty:
        return pd.DataFrame()
    
    if "is_deposit_operation" not in df.columns:
        return df.copy()
    
    # Возвращаем только НЕ депозитные операции
    # (проценты остаются, т.к. они не marked as deposit_operation)
    non_deposits = df[~df["is_deposit_operation"]].copy()
    
    return non_deposits


def get_deposit_summary(df: pd.DataFrame) -> dict:
    """
    Возвращает сводку по депозитам
    
    Это HELPER, потому что:
    1. Агрегирует данные
    2. Не содержит бизнес-логики
    3. Используется для отображения
    """
    if df.empty:
        return {
            "total_placed": 0.0,
            "total_returned": 0.0,
            "total_interest": 0.0,
            "net_deposit": 0.0
        }
    
    if "is_deposit_operation" not in df.columns:
        return {
            "total_placed": 0.0,
            "total_returned": 0.0,
            "total_interest": 0.0,
            "net_deposit": 0.0
        }
    
    total_placed = df[df["is_deposit_placement"]]["amount"].sum()
    total_returned = df[df["is_deposit_return"]]["amount"].abs().sum()
    total_interest = df[df["is_deposit_interest"]]["amount"].sum()
    net_deposit = total_interest - total_placed + total_returned
    
    return {
        "total_placed": total_placed,
        "total_returned": total_returned,
        "total_interest": total_interest,
        "net_deposit": net_deposit
    }
