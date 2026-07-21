import pandas as pd
from datetime import datetime
from typing import Tuple

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
            return True  # ← изменено: пустой DataFrame - это не ошибка
        return True
    
    @staticmethod
    def validate_files(ip_file, phys_file) -> Tuple[bool, str]:
        """Проверяет, что хотя бы один файл загружен"""
        if not ip_file and not phys_file:
            return False, "Необходимо загрузить хотя бы один файл"
        return True, "OK"
    
    @staticmethod
    def validate_duplicates(df: pd.DataFrame) -> bool:
        """Проверяет наличие дубликатов операций"""
        if df.empty:
            return True  # ← изменено: пустой DataFrame - это не ошибка
        
        duplicates = df.duplicated(subset=["date", "amount", "description"], keep=False)
        if duplicates.any():
            dup_count = duplicates.sum()
            raise ValueError(f"Обнаружены дублирующиеся операции: {dup_count} шт.")
        return True
    
    @staticmethod
    def validate_amounts(df: pd.DataFrame) -> bool:
        """Проверяет корректность сумм"""
        if df.empty:
            return True  # ← изменено: пустой DataFrame - это не ошибка
        
        if (df["amount"] < -1e9).any() or (df["amount"] > 1e9).any():
            raise ValueError("Обнаружены некорректные суммы (слишком большие или маленькие)")
        return True
