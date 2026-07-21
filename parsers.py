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
        
        cleaned = str(value).replace(" ", "").replace(",", ".")
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
        
        for name in possible_names:
            name_lower = name.lower()
            for col in df.columns:
                if name_lower in col.lower():
                    return col
        
        return None
    
    @staticmethod
    def is_deposit_placement(text: str) -> bool:
        """Проверяет, является ли операция размещением депозита"""
        if not text:
            return False
        text = str(text).lower()
        keywords = [
            "размещение денежных средств во вклад",
            "размещение дс во вклад",
            "размещение во вклад",
            "открытие депозита"
        ]
        return any(keyword in text for keyword in keywords)
    
    @staticmethod
    def is_deposit_return(text: str) -> bool:
        """Проверяет, является ли операция возвратом депозита"""
        if not text:
            return False
        text = str(text).lower()
        keywords = [
            "возврат депозита",
            "возврат вклада",
            "закрытие депозита",
            "возврат суммы депозита"
        ]
        return any(keyword in text for keyword in keywords)
    
    @staticmethod
    def is_deposit_interest(text: str) -> bool:
        """Проверяет, является ли операция уплатой процентов по депозиту"""
        if not text:
            return False
        text = str(text).lower()
        keywords = [
            "уплата процентов депозит",
            "проценты по депозиту",
            "проценты по вкладу",
            "выплата процентов",
            "начисление процентов"
        ]
        return any(keyword in text for keyword in keywords)
    
    @staticmethod
    def is_deposit_operation(text: str) -> bool:
        """Проверяет, является ли операция депозитной (размещение или возврат)"""
        return BaseParser.is_deposit_placement(text) or BaseParser.is_deposit_return(text)

class IPParser(BaseParser):
    """Парсер выписки ИП"""
    
    @staticmethod
    def parse(file) -> pd.DataFrame:
        """Парсит файл выписки ИП"""
        try:
            df = pd.read_excel(file, engine="openpyxl")
            
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
            
            # Парсим даты
            try:
                result_date = pd.to_datetime(df[date_col], format="%d.%m.%Y", errors="coerce")
            except:
                result_date = pd.to_datetime(df[date_col], errors="coerce")
            
            if result_date.isna().all():
                date_str = df[date_col].astype(str)
                result_date = pd.to_datetime(date_str, format="%d.%m.%Y", errors="coerce")
            
            # Очищаем суммы
            debit_values = df[debit_col].apply(BaseParser.clean_amount)
            credit_values = df[credit_col].apply(BaseParser.clean_amount)
            
            purpose_text = df[purpose_col].fillna("").astype(str)
            
            # Определяем типы операций
            is_deposit_placement = purpose_text.apply(BaseParser.is_deposit_placement)
            is_deposit_return = purpose_text.apply(BaseParser.is_deposit_return)
            is_deposit_interest = purpose_text.apply(BaseParser.is_deposit_interest)
            is_deposit_operation = is_deposit_placement | is_deposit_return
            
            # Создаем результат
            result = pd.DataFrame()
            result["date"] = result_date
            result["debit"] = debit_values
            result["credit"] = credit_values
            result["amount"] = credit_values - debit_values
            result["description"] = purpose_text
            result["source"] = "ip"
            
            # Метки депозитов
            result["is_deposit_placement"] = is_deposit_placement
            result["is_deposit_return"] = is_deposit_return
            result["is_deposit_interest"] = is_deposit_interest
            result["is_deposit_operation"] = is_deposit_operation
            
            # Фильтруем для основного отчета
            result_main = result[~is_deposit_operation].copy()
            
            # Сохраняем депозиты отдельно
            result_deposits = result[is_deposit_operation | is_deposit_interest].copy()
            
            # Удаляем пустые даты
            result_main = result_main.dropna(subset=["date"])
            result_deposits = result_deposits.dropna(subset=["date"])
            
            # Удаляем нулевые суммы
            result_main = result_main[abs(result_main["amount"]) > 0.001]
            result_deposits = result_deposits[abs(result_deposits["amount"]) > 0.001]
            
            # Сортируем
            result_main = result_main.sort_values("date").reset_index(drop=True)
            result_deposits = result_deposits.sort_values("date").reset_index(drop=True)
            
            # Сохраняем депозиты в атрибутах
            result_main.attrs["deposits"] = result_deposits
            
            return result_main
            
        except Exception as e:
            raise ParserError(f"Ошибка при парсинге файла ИП: {str(e)}")


class PhysParser(BaseParser):
    """Парсер выписки физлица"""
    
    @staticmethod
    def parse(file) -> pd.DataFrame:
        """Парсит файл выписки физлица"""
        try:
            df = pd.read_excel(file, engine="openpyxl")
            
            date_col = BaseParser.find_column(df, ["дата операции", "дата опер"])
            desc_col = BaseParser.find_column(df, ["описание"])
            amount_col = BaseParser.find_column(df, ["сумма в валюте счета", "сумма"])
            
            if not all([date_col, desc_col, amount_col]):
                missing = []
                if not date_col: missing.append("Дата операции")
                if not desc_col: missing.append("Описание")
                if not amount_col: missing.append("Сумма")
                raise ParserError(f"Отсутствуют колонки: {', '.join(missing)}")
            
            # Парсим даты
            try:
                result_date = pd.to_datetime(df[date_col], format="%d.%m.%Y", errors="coerce")
            except:
                result_date = pd.to_datetime(df[date_col], errors="coerce")
            
            if result_date.isna().all():
                date_str = df[date_col].astype(str)
                result_date = pd.to_datetime(date_str, format="%d.%m.%Y", errors="coerce")
            
            result = pd.DataFrame()
            result["date"] = result_date
            result["description"] = df[desc_col].fillna("").astype(str)
            result["amount"] = df[amount_col].apply(BaseParser.clean_amount)
            result["source"] = "phys"
            
            # Удаляем пустые даты
            result = result.dropna(subset=["date"])
            
            # Удаляем нулевые суммы
            result = result[abs(result["amount"]) > 0.001]
            
            # Сортируем
            result = result.sort_values("date").reset_index(drop=True)
            
            return result
            
        except Exception as e:
            raise ParserError(f"Ошибка при парсинге файла физлица: {str(e)}")


# ============================================
# ЯВНЫЙ ЭКСПОРТ КЛАССОВ
# ============================================
__all__ = ['IPParser', 'PhysParser', 'ParserError']
