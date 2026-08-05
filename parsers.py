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
    
    # ============================================
    # Методы для депозитов ИП
    # ============================================
    @staticmethod
    def is_deposit_placement(text: str) -> bool:
        if not text:
            return False
        text = str(text)
        return "Размещение денежных средств во Вклад" in text
    
    @staticmethod
    def is_deposit_return(text: str) -> bool:
        if not text:
            return False
        text = str(text)
        return "ВОЗВРАТ ДЕПОЗИТ С ИП" in text
    
    @staticmethod
    def is_deposit_interest(text: str) -> bool:
        if not text:
            return False
        text = str(text)
        return "УПЛАТА ПРОЦЕНТОВ ДЕПОЗИТ С ИП" in text
    
    @staticmethod
    def is_deposit_operation(text: str) -> bool:
        return BaseParser.is_deposit_placement(text) or BaseParser.is_deposit_return(text)

class IPParser(BaseParser):
    """Парсер выписки ИП"""
    
    @staticmethod
    def parse(file) -> pd.DataFrame:
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
            
            try:
                result_date = pd.to_datetime(df[date_col], format="%d.%m.%Y", errors="coerce")
            except:
                result_date = pd.to_datetime(df[date_col], errors="coerce")
            
            if result_date.isna().all():
                date_str = df[date_col].astype(str)
                result_date = pd.to_datetime(date_str, format="%d.%m.%Y", errors="coerce")
            
            debit_values = df[debit_col].apply(BaseParser.clean_amount)
            credit_values = df[credit_col].apply(BaseParser.clean_amount)
            
            purpose_text = df[purpose_col].fillna("").astype(str)
            
            is_deposit_placement = purpose_text.apply(BaseParser.is_deposit_placement)
            is_deposit_return = purpose_text.apply(BaseParser.is_deposit_return)
            is_deposit_interest = purpose_text.apply(BaseParser.is_deposit_interest)
            is_deposit_operation = is_deposit_placement | is_deposit_return
            
            result = pd.DataFrame()
            result["date"] = result_date
            result["debit"] = debit_values
            result["credit"] = credit_values
            result["amount"] = credit_values - debit_values
            result["description"] = purpose_text
            result["source"] = "ip"
            
            result["is_deposit_placement"] = is_deposit_placement
            result["is_deposit_return"] = is_deposit_return
            result["is_deposit_interest"] = is_deposit_interest
            result["is_deposit_operation"] = is_deposit_operation
            
            result_main = result[~is_deposit_operation].copy()
            result_deposits = result[is_deposit_operation | is_deposit_interest].copy()
            
            result_main = result_main.dropna(subset=["date"])
            result_deposits = result_deposits.dropna(subset=["date"])
            
            result_main = result_main[abs(result_main["amount"]) > 0.001]
            result_deposits = result_deposits[abs(result_deposits["amount"]) > 0.001]
            
            result_main = result_main.sort_values("date").reset_index(drop=True)
            result_deposits = result_deposits.sort_values("date").reset_index(drop=True)
            
            result_main.attrs["deposits"] = result_deposits
            
            return result_main
            
        except Exception as e:
            raise ParserError(f"Ошибка при парсинге файла ИП: {str(e)}")


# ============================================
# НОВЫЙ ПАРСЕР ДЛЯ ФЛ
# ============================================

class FLParser(BaseParser):
    """Парсер выписки физлица"""
    
    REQUIRED_COLUMNS = [
        "Дата операции",
        "Название счета",
        "Номер счета",
        "Описание операции",
        "Сумма",
        "Статус",
        "Категория",
        "Тип",
        "Комментарий"
    ]
    
    @staticmethod
    def parse(file) -> pd.DataFrame:
        """
        Парсит файл выписки физлица.
        Возвращает DataFrame с колонками: date, description, amount, source
        """
        try:
            df = pd.read_excel(file, engine="openpyxl")
            
            # Проверяем наличие всех обязательных колонок
            missing_cols = []
            for col in FLParser.REQUIRED_COLUMNS:
                if col not in df.columns:
                    missing_cols.append(col)
            
            if missing_cols:
                raise ParserError(
                    f"Отсутствуют колонки: {', '.join(missing_cols)}. "
                    f"Расчет ФЛ не будет произведен."
                )
            
            # Парсим даты
            try:
                result_date = pd.to_datetime(df["Дата операции"], format="%d.%m.%Y", errors="coerce")
            except:
                result_date = pd.to_datetime(df["Дата операции"], errors="coerce")
            
            if result_date.isna().all():
                date_str = df["Дата операции"].astype(str)
                result_date = pd.to_datetime(date_str, format="%d.%m.%Y", errors="coerce")
            
            # Очищаем сумму (убираем пробелы, запятые)
            amount_values = df["Сумма"].apply(BaseParser.clean_amount)
            
            # Создаем результат
            result = pd.DataFrame()
            result["date"] = result_date
            result["description"] = df["Описание операции"].fillna("").astype(str)
            result["amount"] = amount_values
            result["source"] = "fl"
            
            # Сохраняем дополнительные колонки для информации
            result["account_name"] = df["Название счета"].fillna("").astype(str)
            result["account_number"] = df["Номер счета"].fillna("").astype(str)
            result["status"] = df["Статус"].fillna("").astype(str)
            result["category"] = df["Категория"].fillna("").astype(str)
            result["type"] = df["Тип"].fillna("").astype(str)
            result["comment"] = df["Комментарий"].fillna("").astype(str)
            
            # Удаляем строки с пустыми датами
            result = result.dropna(subset=["date"])
            
            # Удаляем строки с нулевой суммой
            result = result[abs(result["amount"]) > 0.001]
            
            # Сортируем по дате
            result = result.sort_values("date").reset_index(drop=True)
            
            return result
            
        except ParserError:
            raise
        except Exception as e:
            raise ParserError(f"Ошибка при парсинге файла ФЛ: {str(e)}")


__all__ = ['IPParser', 'FLParser', 'ParserError']
