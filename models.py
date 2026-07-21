from dataclasses import dataclass
from datetime import datetime
import pandas as pd

@dataclass
class Operation:
    """Модель операции"""
    date: datetime
    amount: float
    description: str
    source: str

@dataclass
class BalanceReport:
    """Результат расчета остатков"""
    start_balance: float
    end_balance: float
    monthly_dynamics: pd.DataFrame
    
    def to_dict(self):
        return {
            "start_balance": self.start_balance,
            "end_balance": self.end_balance,
            "monthly_dynamics": self.monthly_dynamics
        }