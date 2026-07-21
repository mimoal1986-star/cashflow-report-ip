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
        end_date: datetime
    ) -> Dict[str, BalanceReport]:
        """
        Рассчитывает остатки для ИП и физлица
        """
        # Проверяем входные данные
        if ip_operations.empty and phys_operations.empty:
            raise ValueError("Нет данных для расчета")
        
        # Создаем копии данных, чтобы не изменять оригиналы
        ip_ops = ip_operations.copy()
        phys_ops = phys_operations.copy()
        
        # Фильтруем операции по периоду
        ip_period = ip_ops[
            (ip_ops["date"] >= start_date) & 
            (ip_ops["date"] <= end_date)
        ]
        phys_period = phys_ops[
            (phys_ops["date"] >= start_date) & 
            (phys_ops["date"] <= end_date)
        ]
        
        # Операции до начала периода
        ip_before = ip_ops[ip_ops["date"] < start_date]
        phys_before = phys_ops[phys_ops["date"] < start_date]
        
        # Расчет для ИП
        start_balance_ip = ip_before["amount"].sum()
        end_balance_ip = start_balance_ip + ip_period["amount"].sum()
        
        # Расчет для физлица
        # ВАЖНО: Если нужно, чтобы на 1 июля было 0, игнорируем операции до этой даты
        # Иначе считаем реальный остаток
        use_zero_start = True  # По условию задачи
        
        if use_zero_start:
            # Начальный остаток = 0, игнорируем все операции до start_date
            start_balance_phys = 0.0
            # Пересчитываем операции периода с нулевого старта
            end_balance_phys = phys_period["amount"].sum()
        else:
            # Реальный остаток на начало периода
            start_balance_phys = phys_before["amount"].sum()
            end_balance_phys = start_balance_phys + phys_period["amount"].sum()
        
        # Помесячная динамика
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
        
        # Если используем нулевой старт для физлица, обнуляем накопления до start_date
        if use_zero_start:
            # Начинаем с нуля, все операции до start_date игнорируются
            pass
        
        for month_start in months:
            # Конец месяца
            month_end = month_start + pd.offsets.MonthEnd(1)
            
            # Операции за месяц (только в пределах периода)
            month_ops_ip = ip_ops[
                (ip_ops["date"] >= month_start) & 
                (ip_ops["date"] <= min(month_end, end_date))
            ]
            month_ops_phys = phys_ops[
                (phys_ops["date"] >= month_start) & 
                (phys_ops["date"] <= min(month_end, end_date))
            ]
            
            # Обновляем остатки
            current_balance_ip += month_ops_ip["amount"].sum()
            
            if use_zero_start:
                # Для физлица считаем только операции в периоде
                current_balance_phys += month_ops_phys["amount"].sum()
            else:
                # Для физлица учитываем все операции
                current_balance_phys += month_ops_phys["amount"].sum()
            
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