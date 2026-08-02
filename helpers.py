import pandas as pd
from datetime import datetime
import io
from models import BalanceReport
from deposit_report import DepositReportGenerator

# ============================================
# ЭКСПОРТ В EXCEL (единый файл)
# ============================================

def create_excel_report(
    ip_report: BalanceReport,
    ip_operations: pd.DataFrame
) -> io.BytesIO:
    """
    Создает единый Excel-файл с отчетами:
    - ИП_Динамика
    - ИП_Операции
    - Депозиты_Динамика
    - Депозиты_Операции
    """
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
        
        # ============================================
        # Лист 3: Депозиты_Динамика
        # ============================================
        deposit_ops = ip_operations.attrs.get("deposits", pd.DataFrame())
        
        if not deposit_ops.empty:
            deposit_report = DepositReportGenerator.generate_report(deposit_ops)
            
            if not deposit_report.empty:
                # Форматируем даты в ДД.ММ.ГГГГ
                deposit_report_copy = deposit_report.copy()
                
                if "Дата начала" in deposit_report_copy.columns:
                    deposit_report_copy["Дата начала"] = deposit_report_copy["Дата начала"].apply(
                        lambda x: x.strftime("%d.%m.%Y") if pd.notna(x) else ""
                    )
                
                if "Дата завершения" in deposit_report_copy.columns:
                    deposit_report_copy["Дата завершения"] = deposit_report_copy["Дата завершения"].apply(
                        lambda x: x.strftime("%d.%m.%Y") if pd.notna(x) else ""
                    )
                
                # Переименовываем колонки для единообразия
                deposit_report_copy = deposit_report_copy.rename(columns={
                    "Номер сделки": "Номер сделки",
                    "Дата начала": "Дата начала",
                    "Дата завершения": "Дата завершения",
                    "Сумма депозита (руб)": "Сумма депозита (руб)",
                    "Процент депозита (руб)": "Процент депозита (руб)",
                    "Дней": "Дней"
                })
                
                deposit_report_copy.to_excel(writer, sheet_name="Депозиты_Динамика", index=False)
        
        # ============================================
        # Лист 4: Депозиты_Операции
        # ============================================
        if not deposit_ops.empty:
            detail_df = deposit_ops.copy()
            
            # Извлекаем номера сделок
            from deposit_report import DepositReportGenerator as DRG
            detail_df["Номер сделки"] = detail_df["description"].apply(DRG.extract_deal_number)
            
            # Форматируем дату в ДД.ММ.ГГГГ
            detail_df["Дата"] = detail_df["date"].dt.strftime("%d.%m.%Y")
            
            # Форматируем сумму
            detail_df["Сумма"] = detail_df["amount"].apply(lambda x: f"{x:,.2f}")
            
            # Берем описание
            detail_df["Назначение платежа"] = detail_df["description"]
            
            # Выбираем нужные колонки
            result_detail = detail_df[["Номер сделки", "Дата", "Сумма", "Назначение платежа"]].copy()
            
            # Удаляем строки без номера сделки
            result_detail = result_detail.dropna(subset=["Номер сделки"])
            
            result_detail.to_excel(writer, sheet_name="Депозиты_Операции", index=False)
    
    output.seek(0)
    return output

# ============================================
# РАБОТА С ДЕПОЗИТАМИ (экспорт отдельного файла - больше не нужен)
# ============================================

# Функция export_deposit_report_to_excel больше не нужна,
# так как депозиты включены в основной Excel-файл.
# Оставляем её для обратной совместимости, но она не используется.

def export_deposit_report_to_excel(report_df: pd.DataFrame, ip_operations: pd.DataFrame) -> io.BytesIO:
    """
    Устаревшая функция. Депозиты теперь включены в основной Excel-файл.
    Оставлена для обратной совместимости.
    """
    deposit_ops = ip_operations.attrs.get("deposits", pd.DataFrame())
    return DepositReportGenerator.export_to_excel(report_df, deposit_ops)
