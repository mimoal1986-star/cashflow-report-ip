import pandas as pd
from datetime import datetime
from typing import Dict, Tuple
from models import BalanceReport

class BalanceCalculator:
    """Калькулятор остатков"""
    
    @staticmethod
    def calculate(
        ip_operations: pd.DataFrame,
        phys_operations: pd.DataFrame,
        start_date: datetime,
        end_date: datetime,
        use_zero_start: bool = True  # ← ДОБАВИТЬ ЭТОТ ПАРАМЕТР
    ) -> Dict[str, BalanceReport]:
        """
        Рассчитывает остатки для ИП и физлица
        
        Args:
            ip_operations: DataFrame с операциями ИП
            phys_operations: DataFrame с операциями физлица
            start_date: дата начала периода
            end_date: дата окончания периода
            use_zero_start: если True - начальный остаток физлица = 0
        
        Returns:
            Dict с ключами "ip" и "phys" - объекты BalanceReport
        """
        # Проверяем входные данные
        if ip_operations.empty and phys_operations.empty:
            raise ValueError("Нет данных для расчета")
        
        # Создаем копии данных, чтобы не изменять оригиналы
        ip_ops = ip_operations.copy() if not ip_operations.empty else pd.DataFrame()
        phys_ops = phys_operations.copy() if not phys_operations.empty else pd.DataFrame()
        
        # Фильтруем операции по периоду
        ip_period = ip_ops[
            (ip_ops["date"] >= start_date) & 
            (ip_ops["date"] <= end_date)
        ] if not ip_ops.empty else pd.DataFrame()
        
        phys_period = phys_ops[
            (phys_ops["date"] >= start_date) & 
            (phys_ops["date"] <= end_date)
        ] if not phys_ops.empty else pd.DataFrame()
        
        # Операции до начала периода
        ip_before = ip_ops[ip_ops["date"] < start_date] if not ip_ops.empty else pd.DataFrame()
        phys_before = phys_ops[phys_ops["date"] < start_date] if not phys_ops.empty else pd.DataFrame()
        
        # ============================================
        # РАСЧЕТ ДЛЯ ИП
        # ============================================
        start_balance_ip = ip_before["amount"].sum() if not ip_before.empty else 0.0
        end_balance_ip = start_balance_ip + (ip_period["amount"].sum() if not ip_period.empty else 0.0)
        
        # ============================================
        # РАСЧЕТ ДЛЯ ФИЗЛИЦА
        # ============================================
        if use_zero_start:
            # По условию: начальный остаток = 0
            start_balance_phys = 0.0
            end_balance_phys = phys_period["amount"].sum() if not phys_period.empty else 0.0
        else:
            # Реальный остаток на начало периода
            start_balance_phys = phys_before["amount"].sum() if not phys_before.empty else 0.0
            end_balance_phys = start_balance_phys + (phys_period["amount"].sum() if not phys_period.empty else 0.0)
        
        # ============================================
        # ПОМЕСЯЧНАЯ ДИНАМИКА
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
        """Рассчитывает помесячную динамику остатков"""
        
        # Генерируем первое число каждого месяца в периоде
        months = pd.date_range(
            start=start_date.replace(day=1),
            end=end_date,
            freq="MS"
        )
        
        if len(months) == 0:
            return {
                "ip": pd.DataFrame(columns=["month", "balance"]),
                "phys": pd.DataFrame(columns=["month", "balance"])
            }
        
        dynamics_ip = []
        dynamics_phys = []
        
        # Текущие остатки
        current_balance_ip = start_balance_ip
        current_balance_phys = start_balance_phys
        
        for month_start in months:
            # Конец месяца
            month_end = month_start + pd.offsets.MonthEnd(1)
            
            # Операции за месяц (только в пределах периода)
            month_ops_ip = ip_ops[
                (ip_ops["date"] >= month_start) & 
                (ip_ops["date"] <= min(month_end, end_date))
            ] if not ip_ops.empty else pd.DataFrame()
            
            month_ops_phys = phys_ops[
                (phys_ops["date"] >= month_start) & 
                (phys_ops["date"] <= min(month_end, end_date))
            ] if not phys_ops.empty else pd.DataFrame()
            
            # Обновляем остатки
            current_balance_ip += month_ops_ip["amount"].sum() if not month_ops_ip.empty else 0.0
            
            if use_zero_start:
                # Для физлица считаем только операции в периоде
                current_balance_phys += month_ops_phys["amount"].sum() if not month_ops_phys.empty else 0.0
            else:
                # Для физлица учитываем все операции
                current_balance_phys += month_ops_phys["amount"].sum() if not month_ops_phys.empty else 0.0
            
            dynamics_ip.append({
                "month": month_start.strftime("%B %Y"),
                "balance": round(current_balance_ip, 2)
            })
            dynamics_phys.append({
                "month": month_start.strftime("%B %Y"),
                "balance": round(current_balance_phys, 2)
            })
        
        return {
            "ip": pd.DataFrame(dynamics_ip),
            "phys": pd.DataFrame(dynamics_phys)
        }
