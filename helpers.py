import pandas as pd
from datetime import datetime
from typing import Tuple
import io
from models import BalanceReport

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