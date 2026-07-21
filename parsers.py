import pandas as pd
from typing import Optional, List
import re

class ParserError(Exception):
    """Ошибка парсинга файла"""
    pass

class BaseParser:
    """Базовый класс для парсеров"""
    
    @staticmethod
    def clean_amount(value) -> float:
        """Очищает строку с суммой от лишних символов"""
        if pd.isna(value):
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        
        # Убираем пробелы, заменяем запятую на точку
        cleaned = str(value).replace(" ", "").replace(",", ".")
        # Оставляем только цифры, точку и минус
        cleaned = re.sub(r"[^\d.-]", "", cleaned)
        
        try:
            return float(cleaned) if cleaned else 0.0
        except ValueError:
            return 0.0
    
    @staticmethod
    def find_column(df: pd.DataFrame, possible_names: List[str]) -> Optional[str]:
        """Ищет колонку по списку возможных названий"""
        df_cols_lower = {col.lower(): col for col in df.columns}
        
        for name in possible_names:
            name_lower = name.lower()
            if name_lower in df_cols_lower:
                return df_cols_lower[name_lower]
        
        # Поиск по частичному совпадению
        for name in possible_names:
            name_lower = name.lower()
            for col in df.columns:
                if name_lower in col.lower():
                    return col
        
        return None

class IPParser(BaseParser):
    """Парсер выписки ИП"""
    
    @staticmethod
    def parse(file) -> pd.DataFrame:
        """Парсит файл выписки ИП"""
        try:
            df = pd.read_excel(file, engine="openpyxl")
            
            # Ищем колонки
            date_col = BaseParser.find_column(df, ["дата"])
            debit_col = BaseParser.find_column(df, ["дебет"])
            credit_col = BaseParser.find_column(df, ["кредит"])
            purpose_col = BaseParser.find_column(df, ["назначение платежа", "назнач"])
            
            if not all([date_col, debit_col, credit_col, purpose_col]):
                missing = []
                if not date_col: missing.append("Дата")
                if not debit_col: missing.append("Дебет")
                if not credit_col: missing.append("Кредит")
                if not purpose_col: missing.append("Назначение платежа")
                raise ParserError(f"Отсутствуют колонки: {', '.join(missing)}")
            
            # Очищаем данные
            debit_values = df[debit_col].apply(BaseParser.clean_amount)
            credit_values = df[credit_col].apply(BaseParser.clean_amount)
            
            # Создаем результат
            result = pd.DataFrame()
            result["date"] = pd.to_datetime(df[date_col], errors="coerce")
            result["debit"] = debit_values
            result["credit"] = credit_values
            result["amount"] = credit_values - debit_values  # Приход - расход
            result["description"] = df[purpose_col].fillna("").astype(str)
            result["source"] = "ip"
            
            # Удаляем строки с пустыми датами
            result = result.dropna(subset=["date"])
            
            # Удаляем строки с нулевой суммой (опционально)
            result = result[abs(result["amount"]) > 0.001]
            
            # Сортируем по дате
            result = result.sort_values("date").reset_index(drop=True)
            
            return result
            
        except Exception as e:
            raise ParserError(f"Ошибка при парсинге файла ИП: {str(e)}")

class PhysParser(BaseParser):
    """Парсер выписки физлица"""
    
    @staticmethod
    def parse(file) -> pd.DataFrame:
        """Парсит файл выписки физлица"""
        try:
            df = pd.read_excel(file, engine="openpyxl")
            
            # Ищем колонки
            date_col = BaseParser.find_column(df, ["дата операции", "дата опер"])
            desc_col = BaseParser.find_column(df, ["описание"])
            amount_col = BaseParser.find_column(df, ["сумма в валюте счета", "сумма"])
            
            if not all([date_col, desc_col, amount_col]):
                missing = []
                if not date_col: missing.append("Дата операции")
                if not desc_col: missing.append("Описание")
                if not amount_col: missing.append("Сумма")
                raise ParserError(f"Отсутствуют колонки: {', '.join(missing)}")
            
            result = pd.DataFrame()
            result["date"] = pd.to_datetime(df[date_col], errors="coerce")
            result["description"] = df[desc_col].fillna("").astype(str)
            result["amount"] = df[amount_col].apply(BaseParser.clean_amount)
            result["source"] = "phys"
            
            # Удаляем строки с пустыми датами
            result = result.dropna(subset=["date"])
            
            # Удаляем строки с нулевой суммой
            result = result[abs(result["amount"]) > 0.001]
            
            # Сортируем по дате
            result = result.sort_values("date").reset_index(drop=True)
            
            return result
            
        except Exception as e:
            raise ParserError(f"Ошибка при парсинге файла физлица: {str(e)}")