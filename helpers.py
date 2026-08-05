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
            df_dynamics = df_dynamics.rename(columns={"month": "Месяц"})
            
            result_df = df_dynamics[["Месяц", "Баланс начало месяца", "Баланс конец месяца"]].copy()
            result_df["Баланс начало месяца"] = result_df["Баланс начало месяца"].apply(lambda x: f"{x:,.2f}")
            result_df["Баланс конец месяца"] = result_df["Баланс конец месяца"].apply(lambda x: f"{x:,.2f}")
            
            result_df.to_excel(writer, sheet_name="ИП_Динамика", index=False)
        
        # ============================================
        # ИП_Операции
        # ============================================
        if ip_operations is not None and not ip_operations.empty:
            ops_df = ip_operations.copy()
            
            ops_df["Дата"] = ops_df["date"].dt.strftime("%d.%m.%Y")
            ops_df["Дебет"] = ops_df["debit"].apply(lambda x: f"{x:,.2f}" if x != 0 else "")
            ops_df["Кредит"] = ops_df["credit"].apply(lambda x: f"{x:,.2f}" if x != 0 else "")
            ops_df["Итого"] = ops_df["amount"].apply(lambda x: f"{x:+,.2f}")
            ops_df["Описание"] = ops_df["description"]
            
            result_ops = ops_df[["Дата", "Дебет", "Кредит", "Итого", "Описание"]].copy()
            result_ops.to_excel(writer, sheet_name="ИП_Операции", index=False)
        
        # ============================================
        # Депозиты_Динамика
        # ============================================
        if ip_operations is not None and not ip_operations.empty:
            deposit_ops = ip_operations.attrs.get("deposits", pd.DataFrame())
            
            if not deposit_ops.empty:
                deposit_report = DepositReportGenerator.generate_report(deposit_ops)
                
                if not deposit_report.empty:
                    deposit_report_copy = deposit_report.copy()
                    
                    if "Дата начала" in deposit_report_copy.columns:
                        deposit_report_copy["Дата начала"] = deposit_report_copy["Дата начала"].apply(
                            lambda x: x.strftime("%d.%m.%Y") if pd.notna(x) else ""
                        )
                    
                    if "Дата завершения" in deposit_report_copy.columns:
                        deposit_report_copy["Дата завершения"] = deposit_report_copy["Дата завершения"].apply(
                            lambda x: x.strftime("%d.%m.%Y") if pd.notna(x) else ""
                        )
                    
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
        # Депозиты_Операции
        # ============================================
        if ip_operations is not None and not ip_operations.empty:
            deposit_ops = ip_operations.attrs.get("deposits", pd.DataFrame())
            
            if not deposit_ops.empty:
                detail_df = deposit_ops.copy()
                
                from deposit_report import DepositReportGenerator as DRG
                detail_df["Номер сделки"] = detail_df["description"].apply(DRG.extract_deal_number)
                detail_df["Дата"] = detail_df["date"].dt.strftime("%d.%m.%Y")
                detail_df["Сумма"] = detail_df["amount"].apply(lambda x: f"{x:,.2f}")
                detail_df["Назначение платежа"] = detail_df["description"]
                
                result_detail = detail_df[["Номер сделки", "Дата", "Сумма", "Назначение платежа"]].copy()
                result_detail = result_detail.dropna(subset=["Номер сделки"])
                
                result_detail.to_excel(writer, sheet_name="Депозиты_Операции", index=False)
        
        # ============================================
        # ФЛ_Динамика (заглушка)
        # ============================================
        if fl_report is not None and not fl_report.monthly_dynamics.empty:
            fl_report.monthly_dynamics.to_excel(writer, sheet_name="ФЛ_Динамика", index=False)
        else:
            # Пустой лист с сообщением
            empty_df = pd.DataFrame({"Сообщение": ["Расчет ФЛ в разработке"]})
            empty_df.to_excel(writer, sheet_name="ФЛ_Динамика", index=False)
        
        # ============================================
        # ФЛ_Операции (заглушка)
        # ============================================
        if fl_operations is not None and not fl_operations.empty:
            fl_ops = fl_operations.copy()
            fl_ops["Дата"] = fl_ops["date"].dt.strftime("%d.%m.%Y")
            fl_ops["Описание"] = fl_ops["description"]
            fl_ops["Сумма"] = fl_ops["amount"].apply(lambda x: f"{x:,.2f}")
            fl_ops["Счет"] = fl_ops["account_name"] if "account_name" in fl_ops.columns else ""
            
            result_fl = fl_ops[["Дата", "Описание", "Сумма", "Счет"]].copy()
            result_fl.to_excel(writer, sheet_name="ФЛ_Операции", index=False)
        else:
            empty_df = pd.DataFrame({"Сообщение": ["Расчет ФЛ в разработке"]})
            empty_df.to_excel(writer, sheet_name="ФЛ_Операции", index=False)
    
    output.seek(0)
    return output


def export_deposit_report_to_excel(report_df: pd.DataFrame, ip_operations: pd.DataFrame) -> io.BytesIO:
    """
    Устаревшая функция. Оставлена для обратной совместимости.
    """
    deposit_ops = ip_operations.attrs.get("deposits", pd.DataFrame())
    return DepositReportGenerator.export_to_excel(report_df, deposit_ops)
