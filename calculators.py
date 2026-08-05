import pandas as pd
from datetime import datetime
from typing import Dict, Tuple, List
from models import BalanceReport, BalanceReportFL

class BalanceCalculator:
    """Калькулятор остатков для ИП"""
    
    @staticmethod
    def calculate(
        ip_operations: pd.DataFrame,
        start_date: datetime,
        end_date: datetime
    ) -> Dict[str, BalanceReport]:
        if ip_operations.empty:
            raise ValueError("Нет данных для расчета")
        
        ip_ops = ip_operations.copy()
        
        start_month = start_date.replace(day=1)
        ip_before = ip_ops[ip_ops["date"] < start_month]
        start_balance_ip = ip_before["amount"].sum() if not ip_before.empty else 0.0
        
        ip_period = ip_ops[
            (ip_ops["date"] >= start_date) & 
            (ip_ops["date"] <= end_date)
        ]
        
        end_balance_ip = start_balance_ip + (ip_period["amount"].sum() if not ip_period.empty else 0.0)
        
        dynamics = BalanceCalculator._calculate_monthly_dynamics(
            ip_ops, start_date, end_date, start_balance_ip
        )
        
        return {
            "ip": BalanceReport(start_balance_ip, end_balance_ip, dynamics)
        }
    
    @staticmethod
    def _calculate_monthly_dynamics(
        ip_ops: pd.DataFrame,
        start_date: datetime,
        end_date: datetime,
        start_balance_ip: float
    ) -> pd.DataFrame:
        start_month = start_date.replace(day=1)
        months = pd.date_range(start=start_month, end=end_date, freq="MS")
        
        if len(months) == 0:
            return pd.DataFrame(columns=["month", "balance"])
        
        dynamics = []
        current_balance = start_balance_ip
        
        for month_start in months:
            month_end = month_start + pd.offsets.MonthEnd(1)
            
            month_ops = ip_ops[
                (ip_ops["date"] >= month_start) & 
                (ip_ops["date"] <= min(month_end, end_date))
            ] if not ip_ops.empty else pd.DataFrame()
            
            current_balance += month_ops["amount"].sum() if not month_ops.empty else 0.0
            
            dynamics.append({
                "month": month_start.strftime("%B %Y"),
                "balance": round(current_balance, 2)
            })
        
        return pd.DataFrame(dynamics)


# ============================================
# КАЛЬКУЛЯТОР ДЛЯ ФЛ
# ============================================

class BalanceCalculatorFL:
    """Калькулятор остатков для физлица"""
    
    # Названия счетов
    MAIN_ACCOUNT = "Текущий счёт"
    DEPOSIT_ACCOUNTS = ["Альфа-Счёт на минимальный остаток", "Альфа-Счёт на ежедневный остаток"]
    INTEREST_KEYWORDS = ["выплата процентов", "проценты"]
    
    @staticmethod
    def calculate(
        fl_operations: pd.DataFrame,
        start_date: datetime,
        end_date: datetime
    ) -> BalanceReportFL:
        """
        Рассчитывает остатки для физлица
        """
        if fl_operations.empty:
            return BalanceReportFL(0.0, 0.0, pd.DataFrame(), [], 0.0)
        
        df = fl_operations.copy()
        
        # ============================================
        # 1. ОПРЕДЕЛЯЕМ ТИПЫ ОПЕРАЦИЙ
        # ============================================
        
        # Основной счет (Текущий счёт)
        is_main = df["account_name"] == BalanceCalculatorFL.MAIN_ACCOUNT
        
        # Вклады
        is_deposit = df["account_name"].isin(BalanceCalculatorFL.DEPOSIT_ACCOUNTS)
        
        # Проценты (по описанию)
        is_interest = df["description"].str.lower().apply(
            lambda x: any(kw in x for kw in BalanceCalculatorFL.INTEREST_KEYWORDS)
        )
        
        # ============================================
        # 2. СПАРИВАНИЕ ОПЕРАЦИЙ "МЕЖДУ СВОИМИ СЧЕТАМИ"
        # ============================================
        
        # Фильтруем операции "Между своими счетами"
        between_own = df[df["description"].str.lower() == "между своими счетами"].copy()
        
        # Группируем по дате для поиска пар
        pairs = []
        used_indices = set()
        
        for date in between_own["date"].unique():
            day_ops = between_own[between_own["date"] == date].copy()
            
            # Отделяем пополнения и списания
            inflows = day_ops[day_ops["type"] == "Пополнение"]
            outflows = day_ops[day_ops["type"] == "Списание"]
            
            # Ищем пары по сумме
            for _, out_row in outflows.iterrows():
                if out_row.name in used_indices:
                    continue
                
                # Ищем пополнение с той же суммой
                for _, in_row in inflows.iterrows():
                    if in_row.name in used_indices:
                        continue
                    
                    if abs(out_row["amount"] + in_row["amount"]) < 0.01:  # сумма равна (одна +, другая -)
                        # Проверяем: один счет - текущий, другой - вклад
                        out_is_main = out_row["account_name"] == BalanceCalculatorFL.MAIN_ACCOUNT
                        in_is_main = in_row["account_name"] == BalanceCalculatorFL.MAIN_ACCOUNT
                        out_is_deposit = out_row["account_name"] in BalanceCalculatorFL.DEPOSIT_ACCOUNTS
                        in_is_deposit = in_row["account_name"] in BalanceCalculatorFL.DEPOSIT_ACCOUNTS
                        
                        # Одна операция должна быть по Текущему счету, другая по вкладу
                        if (out_is_main and in_is_deposit) or (out_is_deposit and in_is_main):
                            pairs.append({
                                "date": date,
                                "from_account": out_row["account_name"],
                                "to_account": in_row["account_name"],
                                "amount": abs(out_row["amount"]),
                                "from_index": out_row.name,
                                "to_index": in_row.name
                            })
                            used_indices.add(out_row.name)
                            used_indices.add(in_row.name)
                            break
        
        # Создаем DataFrame с парами
        pairs_df = pd.DataFrame(pairs) if pairs else pd.DataFrame()
        
        # ============================================
        # 3. КОРРЕКТИРОВКА ОПЕРАЦИЙ
        # ============================================
        
        # Создаем копию для расчета
        df_calc = df.copy()
        
        # Для пар "Между своими счетами" меняем описание
        if not pairs_df.empty:
            for _, pair in pairs_df.iterrows():
                # Для операции списания (откуда ушло)
                df_calc.loc[pair["from_index"], "description"] = f"Перевод на {pair['to_account']}"
                # Для операции пополнения (куда пришло)
                df_calc.loc[pair["to_index"], "description"] = f"Перевод с {pair['from_account']}"
        
        # ============================================
        # 4. РАСЧЕТ ДИНАМИКИ ПО ТЕКУЩЕМУ СЧЕТУ
        # ============================================
        
        # Берем только операции по Текущему счету
        main_ops = df_calc[df_calc["account_name"] == BalanceCalculatorFL.MAIN_ACCOUNT].copy()
        
        if main_ops.empty:
            return BalanceReportFL(0.0, 0.0, pd.DataFrame(), [], 0.0)
        
        # Сортируем по дате
        main_ops = main_ops.sort_values("date").reset_index(drop=True)
        
        # Начальный остаток = сумма всех операций до start_date
        main_before = main_ops[main_ops["date"] < pd.Timestamp(start_date)]
        start_balance = main_before["amount"].sum() if not main_before.empty else 0.0
        
        # Помесячная динамика
        start_month = start_date.replace(day=1)
        months = pd.date_range(start=start_month, end=end_date, freq="MS")
        
        dynamics = []
        current_balance = start_balance
        
        for month_start in months:
            month_end = month_start + pd.offsets.MonthEnd(1)
            
            month_ops = main_ops[
                (main_ops["date"] >= month_start) & 
                (main_ops["date"] <= min(month_end, end_date))
            ]
            
            current_balance += month_ops["amount"].sum() if not month_ops.empty else 0.0
            
            dynamics.append({
                "month": month_start.strftime("%B %Y"),
                "balance": round(current_balance, 2)
            })
        
        dynamics_df = pd.DataFrame(dynamics)
        
        # ============================================
        # 5. РАСЧЕТ ПО ВКЛАДАМ
        # ============================================
        
        deposits_data = []
        
        for deposit_name in BalanceCalculatorFL.DEPOSIT_ACCOUNTS:
            # Операции по вкладу
            deposit_ops = df_calc[df_calc["account_name"] == deposit_name].copy()
            
            if deposit_ops.empty:
                continue
            
            # Сортируем по дате
            deposit_ops = deposit_ops.sort_values("date").reset_index(drop=True)
            
            # Остаток на конец каждого месяца
            monthly_data = []
            
            for month_start in months:
                month_end = month_start + pd.offsets.MonthEnd(1)
                
                # Операции за месяц по этому вкладу
                month_ops = deposit_ops[
                    (deposit_ops["date"] >= month_start) & 
                    (deposit_ops["date"] <= min(month_end, end_date))
                ]
                
                # Проценты за месяц (суммируем все выплаты процентов)
                interest_ops = month_ops[
                    month_ops["description"].str.lower().apply(
                        lambda x: any(kw in x for kw in BalanceCalculatorFL.INTEREST_KEYWORDS)
                    )
                ]
                interest_sum = interest_ops["amount"].sum() if not interest_ops.empty else 0.0
                
                # Остаток на конец месяца (сумма всех операций по вкладу до конца месяца)
                ops_until_end = deposit_ops[deposit_ops["date"] <= min(month_end, end_date)]
                balance_end = ops_until_end["amount"].sum() if not ops_until_end.empty else 0.0
                
                monthly_data.append({
                    "month": month_start.strftime("%B %Y"),
                    "balance": round(balance_end, 2),
                    "interest": round(interest_sum, 2)
                })
            
            # Создаем DataFrame по вкладу
            deposit_df = pd.DataFrame(monthly_data)
            
            if not deposit_df.empty:
                deposits_data.append({
                    "account_name": deposit_name,
                    "data": deposit_df
                })
        
        # ============================================
        # 6. КОНЕЧНЫЙ ОСТАТОК ФЛ = СУММА ВСЕХ СЧЕТОВ
        # ============================================
        
        # Сумма остатков на вкладах на конец периода
        deposits_on_end = 0.0
        for deposit_item in deposits_data:
            last_row = deposit_item["data"].iloc[-1] if not deposit_item["data"].empty else None
            if last_row is not None:
                deposits_on_end += last_row["balance"]
        
        # Остаток на Текущем счете на конец периода
        main_period = main_ops[
            (main_ops["date"] >= pd.Timestamp(start_date)) & 
            (main_ops["date"] <= pd.Timestamp(end_date))
        ]
        main_end_balance = start_balance + (main_period["amount"].sum() if not main_period.empty else 0.0)
        
        # Итоговый конечный остаток = текущий счет + все вклады
        end_balance = main_end_balance + deposits_on_end
        
        return BalanceReportFL(
            start_balance=start_balance,
            end_balance=end_balance,
            monthly_dynamics=dynamics_df,
            deposits_data=deposits_data,
            deposits_on_end=deposits_on_end
        )
