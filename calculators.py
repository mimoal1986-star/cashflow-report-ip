import pandas as pd
from datetime import datetime
from typing import Dict
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
# НОВЫЙ КАЛЬКУЛЯТОР ДЛЯ ФЛ
# ============================================

class BalanceCalculatorFL:
    """Калькулятор остатков для физлица"""
    
    @staticmethod
    def calculate(
        fl_operations: pd.DataFrame,
        start_date: datetime,
        end_date: datetime
    ) -> BalanceReportFL:
        """
        Рассчитывает остатки для физлица.
        Пока возвращает пустой отчет.
        """
        if fl_operations.empty:
            return BalanceReportFL(0.0, 0.0, pd.DataFrame(), pd.DataFrame())
        
        # ============================================
        # ВРЕМЕННО: возвращаем пустой отчет
        # TODO: реализовать расчет
        # ============================================
        
        return BalanceReportFL(0.0, 0.0, pd.DataFrame(), pd.DataFrame())
