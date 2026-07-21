import pandas as pd
import re
from typing import Dict, Optional, List
import io

class DepositReportGenerator:
    """Генератор отчета по депозитам"""
    
    @staticmethod
    def extract_deal_number(text: str) -> Optional[str]:
        """
        Извлекает номер сделки из текста назначения платежа.
        Формат: CB + 14-16 цифр (всего 16-18 символов)
        """
        if not text or not isinstance(text, str):
            return None
        
        # Ищем паттерн: CB + от 14 до 16 цифр
        # Регулярное выражение ищет CB, за которым следуют 14-16 цифр
        pattern = r'CB\d{14,16}'
        match = re.search(pattern, text)
        
        if match:
            return match.group(0)
        return None
    
    @staticmethod
    def generate_report(deposit_operations: pd.DataFrame) -> pd.DataFrame:
        """
        Генерирует отчет по депозитам на основе всех депозитных операций.
        
        Args:
            deposit_operations: DataFrame с депозитными операциями (из attrs["deposits"])
            
        Returns:
            DataFrame с отчетом по депозитам:
            - Номер сделки
            - Дата начала (размещение)
            - Дата завершения (возврат)
            - Сумма депозита (руб)
            - Процент депозита (руб)
            - Дней между началом и завершением
        """
        if deposit_operations.empty:
            return pd.DataFrame(columns=[
                "Номер сделки",
                "Дата начала",
                "Дата завершения",
                "Сумма депозита (руб)",
                "Процент депозита (руб)",
                "Дней"
            ])
        
        # Копируем данные, чтобы не изменять оригинал
        df = deposit_operations.copy()
        
        # Извлекаем номер сделки из описания
        df["deal_number"] = df["description"].apply(DepositReportGenerator.extract_deal_number)
        
        # Удаляем операции без номера сделки
        df = df.dropna(subset=["deal_number"])
        
        if df.empty:
            return pd.DataFrame(columns=[
                "Номер сделки",
                "Дата начала",
                "Дата завершения",
                "Сумма депозита (руб)",
                "Процент депозита (руб)",
                "Дней"
            ])
        
        # Группируем по номеру сделки
        deals = {}
        
        for deal_number in df["deal_number"].unique():
            deal_ops = df[df["deal_number"] == deal_number]
            
            # Ищем операции по типам
            placement = deal_ops[deal_ops["is_deposit_placement"]]
            returns = deal_ops[deal_ops["is_deposit_return"]]
            interests = deal_ops[deal_ops["is_deposit_interest"]]
            
            # Дата начала (из размещения)
            start_date = placement["date"].min() if not placement.empty else None
            
            # Дата завершения (из возврата) - берем максимальную дату, если несколько возвратов
            end_date = returns["date"].max() if not returns.empty else None
            
            # Сумма депозита (из размещения) - берем максимальную сумму, если несколько
            deposit_amount = placement["credit"].max() if not placement.empty and "credit" in placement.columns else 0.0
            # Если нет колонки credit, используем amount (сумма должна быть положительной)
            if deposit_amount == 0.0 and not placement.empty:
                deposit_amount = placement["amount"].max()
            
            # Процент депозита (сумма всех процентов)
            interest_amount = interests["amount"].sum() if not interests.empty else 0.0
            
            # Количество дней между началом и завершением
            days = None
            if start_date and end_date:
                days = (end_date - start_date).days
            
            deals[deal_number] = {
                "Номер сделки": deal_number,
                "Дата начала": start_date,
                "Дата завершения": end_date,
                "Сумма депозита (руб)": round(deposit_amount, 2) if deposit_amount else 0.0,
                "Процент депозита (руб)": round(interest_amount, 2) if interest_amount else 0.0,
                "Дней": days
            }
        
        # Создаем DataFrame
        result_df = pd.DataFrame(list(deals.values()))
        
        # Сортируем по дате начала (если есть)
        if not result_df.empty and "Дата начала" in result_df.columns:
            result_df = result_df.sort_values("Дата начала", na_position="last").reset_index(drop=True)
        
        return result_df
    
    @staticmethod
    def export_to_excel(report_df: pd.DataFrame, deposit_operations: pd.DataFrame) -> io.BytesIO:
        """
        Экспортирует депозитный отчет в Excel с детализацией операций.
        """
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Основной отчет по депозитам
            report_df.to_excel(writer, sheet_name="Депозитный_отчет", index=False)
            
            # Детальные операции по депозитам
            if not deposit_operations.empty:
                # Копируем и извлекаем номера сделок
                detail_df = deposit_operations.copy()
                detail_df["deal_number"] = detail_df["description"].apply(
                    DepositReportGenerator.extract_deal_number
                )
                
                # Форматируем для вывода
                detail_df["date"] = detail_df["date"].dt.strftime("%d.%m.%Y")
                detail_df["amount"] = detail_df["amount"].round(2)
                
                # Переименовываем колонки для понятности
                detail_df = detail_df.rename(columns={
                    "date": "Дата",
                    "amount": "Сумма",
                    "description": "Назначение платежа",
                    "deal_number": "Номер сделки"
                })
                
                # Выбираем нужные колонки
                cols = ["Номер сделки", "Дата", "Сумма", "Назначение платежа"]
                detail_df = detail_df[[c for c in cols if c in detail_df.columns]]
                
                detail_df.to_excel(writer, sheet_name="Детальные_операции", index=False)
        
        output.seek(0)
        return output