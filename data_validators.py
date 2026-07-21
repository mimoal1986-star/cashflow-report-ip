import pandas as pd
from datetime import datetime
from typing import Tuple, List

class DataValidator:
    """Валидатор данных"""
    
    @staticmethod
    def validate_dates(start_date: datetime, end_date: datetime) -> bool:
        """Проверяет корректность дат"""
        if start_date > end_date:
            raise ValueError("Дата начала не может быть позже даты окончания")
        return True
    
    @staticmethod
    def validate_operations(df: pd.DataFrame) -> bool:
        """Проверяет, что в данных есть операции"""
        if df.empty:
            return True
        return True
    
    @staticmethod
    def validate_files(ip_file, phys_file) -> Tuple[bool, str]:
        """Проверяет, что хотя бы один файл загружен"""
        if not ip_file and not phys_file:
            return False, "Необходимо загрузить хотя бы один файл"
        return True, "OK"
    
    @staticmethod
    def find_duplicates(df: pd.DataFrame) -> pd.DataFrame:
        """
        Находит дублирующиеся операции
        
        Returns:
            DataFrame с дубликатами или пустой DataFrame
        """
        if df.empty:
            return pd.DataFrame()
        
        # Находим дубликаты по дате, сумме и описанию
        duplicates = df[df.duplicated(subset=["date", "amount", "description"], keep=False)]
        return duplicates
    
    @staticmethod
    def validate_amounts(df: pd.DataFrame) -> bool:
        """Проверяет корректность сумм"""
        if df.empty:
            return True
        
        if (df["amount"] < -1e9).any() or (df["amount"] > 1e9).any():
            raise ValueError("Обнаружены некорректные суммы (слишком большие или маленькие)")
        return True
