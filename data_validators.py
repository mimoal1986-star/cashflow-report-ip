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
            raise ValueError("Нет операций в выбранном периоде")
        return True
    
    @staticmethod
    def validate_files(ip_file, phys_file) -> Tuple[bool, str]:
        """Проверяет, что файлы загружены"""
        if not ip_file or not phys_file:
            return False, "Необходимо загрузить оба файла"
        return True, "OK"
    
    @staticmethod
    def validate_duplicates(df: pd.DataFrame) -> bool:
        """Проверяет наличие дубликатов операций"""
        # Проверяем дубликаты по дате и сумме
        duplicates = df.duplicated(subset=["date", "amount", "description"], keep=False)
        if duplicates.any():
            dup_count = duplicates.sum()
            raise ValueError(f"Обнаружены дублирующиеся операции: {dup_count} шт.")
        return True
    
    @staticmethod
    def validate_amounts(df: pd.DataFrame) -> bool:
        """Проверяет корректность сумм"""
        if (df["amount"] < -1e9).any() or (df["amount"] > 1e9).any():
            raise ValueError("Обнаружены некорректные суммы (слишком большие или маленькие)")
        return True
