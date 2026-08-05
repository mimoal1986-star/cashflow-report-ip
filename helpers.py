import pandas as pd
from datetime import datetime
import io
from models import BalanceReport, BalanceReportFL
from deposit_report import DepositReportGenerator

# ============================================
# ЭКСПОРТ В EXCEL
# ============================================

def create_excel_report(
    ip_report: BalanceReport,
    ip_operations: pd.DataFrame,
    fl_report: BalanceReportFL = None,
    fl_operations: pd.DataFrame = None
) -> io.BytesIO:
    """Создает единый Excel-файл с отчетами"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # ============================================
        # ИП_Динамика
        # ============================================
        if ip_report is not None and not ip_report.monthly_dynamics.empty:
            df_dynamics = ip_report.monthly_dynamics.copy()
            
            balances = []
            for i in range(len(df_dynamics)):
                if i == 0:
                    start_bal = ip_report.start_balance
                else:
                    start_bal = df_dynamics.iloc[i-1]["balance"]
                balances.append(start_bal)
            
            df_dynamics["Баланс начало месяца"] = balances
            df_dynamics["Баланс конец месяца"] = df_dynamics["balance"]
            df_dynamics = df_dynamics.rename(columns={"month": "Месяц"
