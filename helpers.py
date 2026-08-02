import pandas as pd
from datetime import datetime
import io
from models import BalanceReport
from deposit_report import DepositReportGenerator

# ============================================
# ЭКСПОРТ В EXCEL
# ============================================

def create_excel_report(
    ip_report: BalanceReport,
    ip_operations: pd.DataFrame
) -> io.BytesIO:
    """Создает Excel-файл с отчетом"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # ============================================
        # Лист 1: ИП_Динамика
        # ============================================
        if not ip_report.monthly_dynamics.empty:
            df_dynamics = ip_report.monthly_dynamics.copy()
            
            # Получаем начальные остатки для каждого месяца
            balances = []
            for i in range(len(df_dynamics)):
                if i == 0:
                    start_bal = ip_report.start_balance
                else:
                    start_bal = df_dynamics.iloc[i-1]["balance"]
                balances.append(start_bal)
            
            df_dynamics["Баланс начало месяца"] = balances
            df_dynamics["Баланс конец месяца"] = df_dynamics["balance"]
            
            # Переименовываем колонку month
            df_dynamics = df_dynamics.rename(columns={"month": "Месяц"})
            
            # Выбираем нужные колонки
            result_df = df_dynamics[["Месяц", "Баланс начало месяца", "Баланс конец месяца"]].copy()
            
            # Форматируем суммы
            result_df["Баланс начало месяца"] = result_df["Баланс начало месяца"].apply(lambda x: f"{x:,.2f}")
            result_df["Баланс конец месяца"] = result_df["Баланс конец месяца"].apply(lambda x: f"{x:,.2f}")
            
            result_df.to_excel(writer, sheet_name="ИП_Динамика", index=False)
        
        # ============================================
        # Лист 2: ИП_Операции
        # ============================================
        if not ip_operations.empty:
            ops_df = ip_operations.copy()
            
            # Формируем нужные колонки
            ops_df["Дата"] = ops_df["date"].dt.strftime("%d.%m.%Y")
            ops_df["Дебет"] = ops_df["debit"].apply(lambda x: f"{x:,.2f}" if x != 0 else "")
            ops_df["Кредит"] = ops_df["credit"].apply(lambda x: f"{x:,.2f}" if x != 0 else "")
            ops_df["Итого"] = ops_df["amount"].apply(lambda x: f"{x:+,.2f}")
            ops_df["Описание"] = ops_df["description"]
            
            # Выбираем нужные колонки
            result_ops = ops_df[["Дата", "Дебет", "Кредит", "Итого", "Описание"]].copy()
            
            result_ops.to_excel(writer, sheet_name="ИП_Операции", index=False)
    
    output.seek(0)
    return output

# ============================================
# РАБОТА С ДЕПОЗИТАМИ
# ============================================

def export_deposit_report_to_excel(report_df: pd.DataFrame, ip_operations: pd.DataFrame) -> io.BytesIO:
    """
    Экспортирует депозитный отчет в Excel.
    """
    deposit_ops = ip_operations.attrs.get("deposits", pd.DataFrame())
    return DepositReportGenerator.export_to_excel(report_df, deposit_ops)
