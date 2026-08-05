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
            # Формат: 7269017,22 (без разделителей тысяч, запятая как разделитель)
            result_df["Баланс начало месяца"] = result_df["Баланс начало месяца"].apply(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            result_df["Баланс конец месяца"] = result_df["Баланс конец месяца"].apply(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            
            result_df.to_excel(writer, sheet_name="ИП_Динамика", index=False)
        
        # ============================================
        # ИП_Операции
        # ============================================
        if ip_operations is not None and not ip_operations.empty:
            ops_df = ip_operations.copy()
            
            # Дата в числовом формате Excel
            ops_df["Дата"] = pd.to_datetime(ops_df["date"]).apply(lambda x: x.toordinal() + 693594)
            
            # Формат чисел: 7269017,22
            ops_df["Дебет"] = ops_df["debit"].apply(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if x != 0 else ""
            )
            ops_df["Кредит"] = ops_df["credit"].apply(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if x != 0 else ""
            )
            ops_df["Итого"] = ops_df["amount"].apply(
                lambda x: f"{x:+,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
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
                    
                    # Даты в числовом формате Excel
                    if "Дата начала" in deposit_report_copy.columns:
                        deposit_report_copy["Дата начала"] = pd.to_datetime(
                            deposit_report_copy["Дата начала"], errors="coerce"
                        ).apply(lambda x: x.toordinal() + 693594 if pd.notna(x) else "")
                    
                    if "Дата завершения" in deposit_report_copy.columns:
                        deposit_report_copy["Дата завершения"] = pd.to_datetime(
                            deposit_report_copy["Дата завершения"], errors="coerce"
                        ).apply(lambda x: x.toordinal() + 693594 if pd.notna(x) else "")
                    
                    # Формат чисел
                    deposit_report_copy["Сумма депозита (руб)"] = deposit_report_copy["Сумма депозита (руб)"].apply(
                        lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if x else ""
                    )
                    deposit_report_copy["Процент депозита (руб)"] = deposit_report_copy["Процент депозита (руб)"].apply(
                        lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if x else ""
                    )
                    
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
                
                # Дата в числовом формате Excel
                detail_df["Дата"] = pd.to_datetime(detail_df["date"]).apply(lambda x: x.toordinal() + 693594)
                
                # Формат чисел
                detail_df["Сумма"] = detail_df["amount"].apply(
                    lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                )
                detail_df["Назначение платежа"] = detail_df["description"]
                
                result_detail = detail_df[["Номер сделки", "Дата", "Сумма", "Назначение платежа"]].copy()
                result_detail = result_detail.dropna(subset=["Номер сделки"])
                
                result_detail.to_excel(writer, sheet_name="Депозиты_Операции", index=False)
        
        # ============================================
        # ФЛ_Динамика
        # ============================================
        if fl_report is not None and not fl_report.monthly_dynamics.empty:
            df_fl_dynamics = fl_report.monthly_dynamics.copy()
            
            balances = []
            for i in range(len(df_fl_dynamics)):
                if i == 0:
                    start_bal = fl_report.start_balance
                else:
                    start_bal = df_fl_dynamics.iloc[i-1]["balance"]
                balances.append(start_bal)
            
            df_fl_dynamics["Баланс начало месяца"] = balances
            df_fl_dynamics["Баланс конец месяца"] = df_fl_dynamics["balance"]
            df_fl_dynamics = df_fl_dynamics.rename(columns={"month": "Месяц"})
            
            result_df = df_fl_dynamics[["Месяц", "Баланс начало месяца", "Баланс конец месяца"]].copy()
            result_df["Баланс начало месяца"] = result_df["Баланс начало месяца"].apply(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            result_df["Баланс конец месяца"] = result_df["Баланс конец месяца"].apply(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            
            result_df.to_excel(writer, sheet_name="ФЛ_Динамика", index=False)
        else:
            empty_df = pd.DataFrame({"Сообщение": ["Нет данных для динамики ФЛ"]})
            empty_df.to_excel(writer, sheet_name="ФЛ_Динамика", index=False)
        
        # ============================================
        # ФЛ_Вклады
        # ============================================
        if fl_report is not None and fl_report.deposits_data:
            all_deposits = []
            for deposit_item in fl_report.deposits_data:
                df_dep = deposit_item["data"].copy()
                df_dep["Название счета"] = deposit_item["account_name"]
                all_deposits.append(df_dep)
            
            if all_deposits:
                combined_df = pd.concat(all_deposits, ignore_index=True)
                combined_df = combined_df[["Название счета", "month", "balance", "interest"]]
                combined_df = combined_df.rename(columns={
                    "month": "Месяц",
                    "balance": "Остаток на конец месяца",
                    "interest": "Выплата процентов"
                })
                combined_df["Остаток на конец месяца"] = combined_df["Остаток на конец месяца"].apply(
                    lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                )
                combined_df["Выплата процентов"] = combined_df["Выплата процентов"].apply(
                    lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                )
                combined_df.to_excel(writer, sheet_name="ФЛ_Вклады", index=False)
            else:
                empty_df = pd.DataFrame({"Сообщение": ["Нет данных по вкладам ФЛ"]})
                empty_df.to_excel(writer, sheet_name="ФЛ_Вклады", index=False)
        else:
            empty_df = pd.DataFrame({"Сообщение": ["Нет данных по вкладам ФЛ"]})
            empty_df.to_excel(writer, sheet_name="ФЛ_Вклады", index=False)
        
        # ============================================
        # ФЛ_Операции (НОВЫЙ ЛИСТ)
        # ============================================
        if fl_operations is not None and not fl_operations.empty:
            fl_ops = fl_operations.copy()
            
            # Дата в числовом формате Excel
            fl_ops["Дата"] = pd.to_datetime(fl_ops["date"]).apply(lambda x: x.toordinal() + 693594)
            fl_ops["Описание"] = fl_ops["description"]
            fl_ops["Сумма"] = fl_ops["amount"].apply(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            fl_ops["Счет"] = fl_ops["account_name"] if "account_name" in fl_ops.columns else ""
            fl_ops["Тип"] = fl_ops["type"] if "type" in fl_ops.columns else ""
            fl_ops["Категория"] = fl_ops["category"] if "category" in fl_ops.columns else ""
            
            result_fl = fl_ops[["Дата", "Описание", "Сумма", "Счет", "Тип", "Категория"]].copy()
            result_fl.to_excel(writer, sheet_name="ФЛ_Операции", index=False)
        else:
            empty_df = pd.DataFrame({"Сообщение": ["Нет данных по операциям ФЛ"]})
            empty_df.to_excel(writer, sheet_name="ФЛ_Операции", index=False)
        
        # ============================================
        # ФЛ_Вклады_Операции (НОВЫЙ ЛИСТ)
        # ============================================
        if fl_operations is not None and not fl_operations.empty:
            # Фильтруем только операции по вкладам
            deposit_names = ["Альфа-Счёт на минимальный остаток", "Альфа-Счёт на ежедневный остаток"]
            fl_deposit_ops = fl_operations[fl_operations["account_name"].isin(deposit_names)].copy()
            
            if not fl_deposit_ops.empty:
                fl_deposit_ops["Дата"] = pd.to_datetime(fl_deposit_ops["date"]).apply(lambda x: x.toordinal() + 693594)
                fl_deposit_ops["Описание"] = fl_deposit_ops["description"]
                fl_deposit_ops["Сумма"] = fl_deposit_ops["amount"].apply(
                    lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                )
                fl_deposit_ops["Счет"] = fl_deposit_ops["account_name"]
                fl_deposit_ops["Тип"] = fl_deposit_ops["type"] if "type" in fl_deposit_ops.columns else ""
                
                result_deposit_ops = fl_deposit_ops[["Дата", "Описание", "Сумма", "Счет", "Тип"]].copy()
                result_deposit_ops.to_excel(writer, sheet_name="ФЛ_Вклады_Операции", index=False)
            else:
                empty_df = pd.DataFrame({"Сообщение": ["Нет операций по вкладам ФЛ"]})
                empty_df.to_excel(writer, sheet_name="ФЛ_Вклады_Операции", index=False)
        else:
            empty_df = pd.DataFrame({"Сообщение": ["Нет данных по вкладам ФЛ"]})
            empty_df.to_excel(writer, sheet_name="ФЛ_Вклады_Операции", index=False)
    
    output.seek(0)
    return output


def export_deposit_report_to_excel(report_df: pd.DataFrame, ip_operations: pd.DataFrame) -> io.BytesIO:
    """
    Устаревшая функция. Оставлена для обратной совместимости.
    """
    deposit_ops = ip_operations.attrs.get("deposits", pd.DataFrame())
    return DepositReportGenerator.export_to_excel(report_df, deposit_ops)
