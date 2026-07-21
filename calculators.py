import pandas as pd
from datetime import datetime
from typing import Dict
from models import BalanceReport

class BalanceCalculator:
    """Калькулятор остатков"""
    
    @staticmethod
    def calculate(
        ip_operations: pd.DataFrame,
        phys_operations: pd.DataFrame,
        start_date: datetime,
        end_date: datetime,
        use_zero_start: bool = True
    ) -> Dict[str, BalanceReport]:
        """
        Рассчитывает остатки для ИП и физлица
        """
        # Проверяем входные данные
        if ip_operations.empty and phys_operations.empty:
            raise ValueError("Нет данных для расчета")
        
        # Создаем копии данных
        ip_ops = ip_operations.copy() if not ip_operations.empty else pd.DataFrame()
        phys_ops = phys_operations.copy() if not phys_operations.empty else pd.DataFrame()
        
        # ============================================
        # 1. НАЧАЛЬНЫЙ ОСТАТОК (на 1-е число месяца начала периода)
        # ============================================
        start_month = start_date.replace(day=1)
        
        # Все операции ДО 1-го числа месяца начала периода
        ip_before = ip_ops[ip_ops["date"] < start_month] if not ip_ops.empty else pd.DataFrame()
        phys_before = phys_ops[phys_ops["date"] < start_month] if not phys_ops.empty else pd.DataFrame()
        
        # Начальный остаток ИП
        start_balance_ip = ip_before["amount"].sum() if not ip_before.empty else 0.0
        
        # Начальный остаток физлица
        if use_zero_start:
            start_balance_phys = 0.0
        else:
            start_balance_phys = phys_before["amount"].sum() if not phys_before.empty else 0.0
        
        # ============================================
        # 2. ФИЛЬТРУЕМ ОПЕРАЦИИ ЗА ПЕРИОД (для отображения)
        # ============================================
        ip_period = ip_ops[
            (ip_ops["date"] >= start_date) & 
            (ip_ops["date"] <= end_date)
        ] if not ip_ops.empty else pd.DataFrame()
        
        phys_period = phys_ops[
            (phys_ops["date"] >= start_date) & 
            (phys_ops["date"] <= end_date)
        ] if not phys_ops.empty else pd.DataFrame()
        
        # ============================================
        # 3. КОНЕЧНЫЙ ОСТАТОК
        # ============================================
        # Суммируем ВСЕ операции ДО end_date (включая операции до периода)
        ip_until_end = ip_ops[ip_ops["date"] <= end_date] if not ip_ops.empty else pd.DataFrame()
        end_balance_ip = ip_until_end["amount"].sum() if not ip_until_end.empty else 0.0
        
        if use_zero_start:
            phys_until_end = phys_ops[phys_ops["date"] <= end_date] if not phys_ops.empty else pd.DataFrame()
            end_balance_phys = phys_until_end["amount"].sum() if not phys_until_end.empty else 0.0
        else:
            phys_until_end = phys_ops[phys_ops["date"] <= end_date] if not phys_ops.empty else pd.DataFrame()
            end_balance_phys = phys_until_end["amount"].sum() if not phys_until_end.empty else 0.0
        
        # ============================================
        # 4. ПОМЕСЯЧНАЯ ДИНАМИКА (на 1-е число каждого месяца)
        # ============================================
        dynamics = BalanceCalculator._calculate_monthly_dynamics(
            ip_ops, phys_ops, start_date, end_date,
            start_balance_ip, start_balance_phys,
            use_zero_start
        )
        
        return {
            "ip": BalanceReport(start_balance_ip, end_balance_ip, dynamics["ip"]),
            "phys": BalanceReport(start_balance_phys, end_balance_phys, dynamics["phys"])
        }
    
    @staticmethod
    def _calculate_monthly_dynamics(
        ip_ops: pd.DataFrame,
        phys_ops: pd.DataFrame,
        start_date: datetime,
        end_date: datetime,
        start_balance_ip: float,
        start_balance_phys: float,
        use_zero_start: bool = True
    ) -> Dict[str, pd.DataFrame]:
        """
        Рассчитывает помесячную динамику остатков на 1-е число каждого месяца
        """
        # Генерируем первое число каждого месяца в периоде
        start_month = start_date.replace(day=1)
        months = pd.date_range(start=start_month, end=end_date, freq="MS")
        
        if len(months) == 0:
            return {
                "ip": pd.DataFrame(columns=["month", "balance"]),
                "phys": pd.DataFrame(columns=["month", "balance"])
            }
        
        dynamics_ip = []
        dynamics_phys = []
        
        # Текущий остаток = начальный остаток
        current_balance_ip = start_balance_ip
        current_balance_phys = start_balance_phys
        
        # === ПЕРВЫЙ МЕСЯЦ ===
        first_month = months[0]
        dynamics_ip.append({
            "month": first_month.strftime("%B %Y"),
            "balance": round(current_balance_ip, 2)
        })
        dynamics_phys.append({
            "month": first_month.strftime("%B %Y"),
            "balance": round(current_balance_phys, 2)
        })
        
        # === ОСТАЛЬНЫЕ МЕСЯЦА ===
        for i in range(1, len(months)):
            current_month = months[i]
            prev_month = months[i-1]
            
            # Конец предыдущего месяца
            prev_month_end = prev_month + pd.offsets.MonthEnd(1)
            
            # Операции за ПРЕДЫДУЩИЙ месяц (с prev_month по prev_month_end)
            month_ops_ip = ip_ops[
                (ip_ops["date"] >= prev_month) & 
                (ip_ops["date"] <= prev_month_end)
            ] if not ip_ops.empty else pd.DataFrame()
            
            month_ops_phys = phys_ops[
                (phys_ops["date"] >= prev_month) & 
                (phys_ops["date"] <= prev_month_end)
            ] if not phys_ops.empty else pd.DataFrame()
            
            # Добавляем операции ТОЛЬКО за предыдущий месяц
            current_balance_ip += month_ops_ip["amount"].sum() if not month_ops_ip.empty else 0.0
            
            if use_zero_start:
                current_balance_phys += month_ops_phys["amount"].sum() if not month_ops_phys.empty else 0.0
            else:
                current_balance_phys += month_ops_phys["amount"].sum() if not month_ops_phys.empty else 0.0
            
            dynamics_ip.append({
                "month": current_month.strftime("%B %Y"),
                "balance": round(current_balance_ip, 2)
            })
            dynamics_phys.append({
                "month": current_month.strftime("%B %Y"),
                "balance": round(current_balance_phys, 2)
            })
        
        return {
            "ip": pd.DataFrame(dynamics_ip),
            "phys": pd.DataFrame(dynamics_phys)
        }
