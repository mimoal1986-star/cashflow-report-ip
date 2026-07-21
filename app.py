import streamlit as st
import pandas as pd
from datetime import datetime
import logging

from parsers import IPParser, PhysParser, ParserError
from calculators import BalanceCalculator
from data_validators import DataValidator
from helpers import get_date_range, create_excel_report, export_deposit_report_to_excel
from deposit_report import DepositReportGenerator

# Настройка страницы
st.set_page_config(
    page_title="Отчет по ДДС ИП",
    page_icon="💰",
    layout="wide"
)

st.title("💰 Автоматический отчет по движению денежных средств ИП")

# Инициализация сессии
if "ip_operations" not in st.session_state:
    st.session_state.ip_operations = None
if "phys_operations" not in st.session_state:
    st.session_state.phys_operations = None
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False
if "report_ready" not in st.session_state:
    st.session_state.report_ready = False
if "ip_report" not in st.session_state:
    st.session_state.ip_report = None
if "phys_report" not in st.session_state:
    st.session_state.phys_report = None
if "excel_data" not in st.session_state:
    st.session_state.excel_data = None
if "report_filename" not in st.session_state:
    st.session_state.report_filename = None

# -------------------------------
# Блок загрузки файлов
# -------------------------------
st.header("📁 Загрузка выписок банков")

col1, col2 = st.columns(2)

with col1:
    st.subheader("🏢 Выписка ИП")
    file_ip = st.file_uploader(
        "Загрузите Excel-файл выписки ИП",
        type=["xlsx", "xls"],
        key="ip_upload"
    )

with col2:
    st.subheader("👤 Выписка физлица")
    file_phys = st.file_uploader(
        "Загрузите Excel-файл выписки физлица",
        type=["xlsx", "xls"],
        key="phys_upload"
    )

# Кнопка обработки
if st.button("🔄 Обработать файлы", type="primary"):
    if file_ip or file_phys:
        try:
            with st.spinner("Обработка файлов..."):
                if file_ip:
                    st.session_state.ip_operations = IPParser.parse(file_ip)
                else:
                    st.session_state.ip_operations = pd.DataFrame()
                
                if file_phys:
                    st.session_state.phys_operations = PhysParser.parse(file_phys)
                else:
                    st.session_state.phys_operations = pd.DataFrame()
                
                if not st.session_state.ip_operations.empty:
                    duplicates_ip = DataValidator.find_duplicates(st.session_state.ip_operations)
                    if not duplicates_ip.empty:
                        st.warning(f"⚠️ Обнаружены дублирующиеся операции в выписке ИП: {len(duplicates_ip)} шт.")
                        with st.expander("📋 Показать дубликаты (ИП)"):
                            st.dataframe(
                                duplicates_ip[["date", "amount", "description"]],
                                use_container_width=True,
                                hide_index=True
                            )
                
                if not st.session_state.phys_operations.empty:
                    duplicates_phys = DataValidator.find_duplicates(st.session_state.phys_operations)
                    if not duplicates_phys.empty:
                        st.warning(f"⚠️ Обнаружены дублирующиеся операции в выписке физлица: {len(duplicates_phys)} шт.")
                        with st.expander("📋 Показать дубликаты (Физлицо)"):
                            st.dataframe(
                                duplicates_phys[["date", "amount", "description"]],
                                use_container_width=True,
                                hide_index=True
                            )
                
                if not st.session_state.ip_operations.empty:
                    DataValidator.validate_amounts(st.session_state.ip_operations)
                
                if not st.session_state.phys_operations.empty:
                    DataValidator.validate_amounts(st.session_state.phys_operations)
                
                st.session_state.data_loaded = True
                st.session_state.report_ready = False
                
                st.success("✅ Файлы успешно обработаны и проверены!")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    count_ip = len(st.session_state.ip_operations) if not st.session_state.ip_operations.empty else 0
                    st.metric("Операций ИП", count_ip)
                with col2:
                    count_phys = len(st.session_state.phys_operations) if not st.session_state.phys_operations.empty else 0
                    st.metric("Операций физлица", count_phys)
                with col3:
                    total = count_ip + count_phys
                    st.metric("Всего операций", total)
                
                if file_ip and not file_phys:
                    st.info("ℹ️ Загружена только выписка ИП. Отчет будет сформирован только по ИП.")
                elif file_phys and not file_ip:
                    st.info("ℹ️ Загружена только выписка физлица. Отчет будет сформирован только по физлицу.")
                
        except ParserError as e:
            st.error(f"❌ Ошибка при обработке: {str(e)}")
            st.session_state.data_loaded = False
        except ValueError as e:
            st.error(f"❌ Ошибка валидации: {str(e)}")
            st.session_state.data_loaded = False
        except Exception as e:
            st.error(f"❌ Непредвиденная ошибка: {str(e)}")
            st.session_state.data_loaded = False
    else:
        st.warning("⚠️ Загрузите хотя бы один файл для обработки")

# -------------------------------
# Основной функционал
# -------------------------------
if st.session_state.data_loaded and \
   (st.session_state.ip_operations is not None or \
    st.session_state.phys_operations is not None):
    
    ip_empty = st.session_state.ip_operations.empty if st.session_state.ip_operations is not None else True
    phys_empty = st.session_state.phys_operations.empty if st.session_state.phys_operations is not None else True
    
    if ip_empty and phys_empty:
        st.warning("⚠️ Нет данных для формирования отчета")
        st.stop()
    
    try:
        all_dates = []
        if not ip_empty:
            all_dates.extend(st.session_state.ip_operations["date"].tolist())
        if not phys_empty:
            all_dates.extend(st.session_state.phys_operations["date"].tolist())
        
        if all_dates:
            min_date = min(all_dates)
            max_date = max(all_dates)
        else:
            st.warning("⚠️ Нет данных с датами")
            st.stop()
            
    except Exception as e:
        st.error(f"❌ Ошибка определения диапазона дат: {str(e)}")
        st.stop()
    
    st.subheader("📅 Настройка периода отчета")
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input(
            "📆 Дата начала периода",
            value=min_date,
            min_value=min_date,
            max_value=max_date,
            key="start_date"
        )
    with col2:
        end_date = st.date_input(
            "📆 Дата окончания периода",
            value=max_date,
            min_value=min_date,
            max_value=max_date,
            key="end_date"
        )
    
    use_zero_start = st.checkbox(
        "Начальный остаток физлица = 0 (согласно условию)",
        value=True,
        help="Если включено, все операции физлица до даты начала игнорируются"
    )
    
    # ============================================
    # КНОПКА РАСЧЕТА
    # ============================================
    if st.button("📊 Сформировать отчет", type="primary"):
        try:
            DataValidator.validate_dates(start_date, end_date)
            
            with st.spinner("Расчет отчета..."):
                reports = BalanceCalculator.calculate(
                    st.session_state.ip_operations,
                    st.session_state.phys_operations,
                    pd.Timestamp(start_date),
                    pd.Timestamp(end_date),
                    use_zero_start
                )
                
                st.session_state.ip_report = reports["ip"]
                st.session_state.phys_report = reports["phys"]
                
                excel_file = create_excel_report(
                    st.session_state.ip_report,
                    st.session_state.phys_report,
                    st.session_state.ip_operations,
                    st.session_state.phys_operations
                )
                
                st.session_state.excel_data = excel_file.getvalue()
                st.session_state.report_filename = f"Отчет_ДДС_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
                st.session_state.report_ready = True
                
                st.success("✅ Отчет сформирован!")
                
        except ValueError as e:
            st.error(f"❌ Ошибка валидации: {str(e)}")
        except Exception as e:
            st.error(f"❌ Ошибка при формировании отчета: {str(e)}")
    
    # ============================================
    # ОТОБРАЖЕНИЕ РЕЗУЛЬТАТОВ
    # ============================================
    if st.session_state.report_ready and st.session_state.ip_report is not None:
        ip_report = st.session_state.ip_report
        phys_report = st.session_state.phys_report
        
        st.header("📈 Отчет по движению денежных средств")
        
        # ============================================
        # РАСЧЕТ "Из них на депозите" на конец периода
        # ============================================
        
        deposit_ops_all = st.session_state.ip_operations.attrs.get("deposits", pd.DataFrame()) if st.session_state.ip_operations is not None else pd.DataFrame()
        
        if not deposit_ops_all.empty:
            deposit_report_full = DepositReportGenerator.generate_report(deposit_ops_all)
            if not deposit_report_full.empty:
                end_ts = pd.Timestamp(end_date)
                
                # Активные на конец периода:
                # - депозит начался ДО или В период (дата начала <= end_date)
                # - И (нет даты завершения ИЛИ завершение ПОСЛЕ end_date)
                active_on_end = deposit_report_full[
                    (deposit_report_full["Дата начала"] <= end_ts) &
                    (
                        (deposit_report_full["Дата завершения"].isna()) |
                        (deposit_report_full["Дата завершения"] > end_ts)
                    )
                ]
                
                ip_on_deposit = active_on_end["Сумма депозита (руб)"].sum() if not active_on_end.empty else 0.0
            else:
                ip_on_deposit = 0.0
        else:
            ip_on_deposit = 0.0
        
        # Для физлица пока 0
        phys_on_deposit = 0.0

        
        # ============================================
        # ОТОБРАЖЕНИЕ МЕТРИК (6 колонок)
        # ============================================
        
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        
        with col1:
            st.metric(
                "🏢 Начальный остаток ИП",
                f"{ip_report.start_balance:,.2f} ₽"
            )
        
        with col2:
            st.metric(
                "🏢 Конечный остаток ИП",
                f"{ip_report.end_balance:,.2f} ₽",
                delta=f"{ip_report.end_balance - ip_report.start_balance:,.2f} ₽"
            )
        
        with col3:
            st.metric(
                "🏦 Их них на депозите",
                f"{ip_on_deposit:,.2f} ₽"
            )
        
        with col4:
            st.metric(
                "👤 Начальный остаток физлица",
                f"{phys_report.start_balance:,.2f} ₽"
            )
        
        with col5:
            st.metric(
                "👤 Конечный остаток физлица",
                f"{phys_report.end_balance:,.2f} ₽",
                delta=f"{phys_report.end_balance - phys_report.start_balance:,.2f} ₽"
            )
        
        with col6:
            st.metric(
                "🏦 Из них на вкладе",
                f"{phys_on_deposit:,.2f} ₽"
            )
        
        total_start = ip_report.start_balance + phys_report.start_balance
        total_end = ip_report.end_balance + phys_report.end_balance
        st.info(f"💰 **Общий остаток:** {total_start:,.2f} ₽ → {total_end:,.2f} ₽ (изменение: {total_end - total_start:,.2f} ₽)")
        
        # ✅ Все три вкладки создаются вместе
        tab1, tab2, tab3 = st.tabs(["📊 Динамика ИП", "📊 Динамика физлица", "🏦 Депозиты"])
        
        with tab1:
            st.subheader("Динамика остатка ИП помесячно")
            
            if not ip_report.monthly_dynamics.empty:
                # Берем данные динамики
                df_dynamics = ip_report.monthly_dynamics.copy()
                
                # Преобразуем месяц в краткий формат: Янв'26, Фев'26, ...
                df_dynamics["month_short"] = pd.to_datetime(
                    df_dynamics["month"], format="%B %Y"
                ).dt.strftime("%b'%y")
                
                # ============================================
                # РАСЧЕТ ПОКАЗАТЕЛЕЙ
                # ============================================
                
                # Получаем начальный остаток (первая строка динамики)
                start_balance = df_dynamics["balance"].iloc[0] if not df_dynamics.empty else 0
                
                # Динамика = изменение остатка
                df_dynamics["dynamics"] = df_dynamics["balance"].diff().fillna(0)
                
                # Для поступлений и списаний за каждый месяц
                ip_ops = st.session_state.ip_operations
                
                if not ip_ops.empty:
                    # Копируем и создаем колонку с месяцем
                    ops = ip_ops.copy()
                    ops["month_period"] = ops["date"].dt.to_period("M").dt.strftime("%b'%y")
                    
                    # Группируем по месяцу: поступления (amount > 0) и списания (amount < 0)
                    monthly_income = ops[ops["amount"] > 0].groupby("month_period")["amount"].sum()
                    monthly_expense = ops[ops["amount"] < 0].groupby("month_period")["amount"].sum()
                    
                    # Создаем таблицу с месяцами по горизонтали
                    months = df_dynamics["month_short"].tolist()
                    
                    # Формируем данные для таблицы
                    table_data = {
                        "Показатель": [
                            "Начальный остаток, млн ₽",
                            "Конечный остаток, млн ₽",
                            "Динамика, млн ₽",
                            "───────────",  # ← разделительная черта
                            "Поступления, млн ₽",
                            "Списания, млн ₽"
                        ]
                    }
                    
                    for month in months:
                        # Находим данные для этого месяца
                        row = df_dynamics[df_dynamics["month_short"] == month]
                        
                        if not row.empty:
                            balance = row["balance"].iloc[0] / 1_000_000  # переводим в млн
                            dynamics = row["dynamics"].iloc[0] / 1_000_000
                            
                            # Поступления и списания за этот месяц
                            income = monthly_income.get(month, 0) / 1_000_000
                            expense = monthly_expense.get(month, 0) / 1_000_000
                        else:
                            balance = 0
                            dynamics = 0
                            income = 0
                            expense = 0
                        
                        # Начальный остаток для этого месяца
                        # Для первого месяца берем start_balance, для остальных - остаток на начало месяца
                        if month == months[0]:
                            start_bal = start_balance / 1_000_000
                        else:
                            # Берем остаток из предыдущего месяца
                            prev_row = df_dynamics[df_dynamics["month_short"] == month].index
                            if len(prev_row) > 0:
                                idx = prev_row[0]
                                if idx > 0:
                                    start_bal = df_dynamics.iloc[idx - 1]["balance"] / 1_000_000
                                else:
                                    start_bal = start_balance / 1_000_000
                            else:
                                start_bal = 0
                        
                        # Форматируем значения с одним знаком после запятой
                        table_data[month] = [
                            f"{start_bal:.2f}",
                            f"{balance:.2f}",
                            f"{dynamics:+.2f}",
                            "─",  # ← разделительная черта
                            f"{income:+.2f}",
                            f"{expense:+.2f}"
                        ]
                    
                    # Создаем DataFrame для отображения
                    df_table = pd.DataFrame(table_data)
                    
                    # Убираем индекс и показываем
                    st.dataframe(
                        df_table,
                        use_container_width=True,
                        hide_index=True
                    )
                else:
                    st.info("Нет данных для отображения динамики ИП")
            else:
                st.info("Нет данных для отображения динамики ИП")
        
        with tab2:
            st.subheader("Динамика остатка физлица помесячно")
            if not phys_report.monthly_dynamics.empty:
                st.line_chart(
                    phys_report.monthly_dynamics.set_index("month")["balance"]
                )
                st.dataframe(
                    phys_report.monthly_dynamics,
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("Нет данных для отображения динамики физлица")
        
        # Вкладка с депозитами
        with tab3:
            st.header("🏦 Отчет по депозитам")
            
            if st.session_state.ip_operations is not None and not st.session_state.ip_operations.empty:
                # Получаем все депозитные операции
                deposit_ops_all = st.session_state.ip_operations.attrs.get("deposits", pd.DataFrame())
                
                if deposit_ops_all.empty:
                    st.info("ℹ️ Нет депозитных операций в выписке ИП")
                else:
                    # Генерируем полный отчет по депозитам
                    deposit_report_full = DepositReportGenerator.generate_report(deposit_ops_all)
                    
                    if deposit_report_full.empty:
                        st.info("ℹ️ Не найдены депозитные операции с номерами сделок")
                    else:
                        # ============================================
                        # ФИЛЬТРАЦИЯ: только депозиты, начавшиеся в периоде
                        # ============================================
                        start_ts = pd.Timestamp(start_date)
                        end_ts = pd.Timestamp(end_date)
                        
                        # Оставляем только те депозиты, у которых дата начала в периоде
                        deposit_report = deposit_report_full[
                            (deposit_report_full["Дата начала"] >= start_ts) & 
                            (deposit_report_full["Дата начала"] <= end_ts)
                        ].copy()
                        
                        if deposit_report.empty:
                            st.info(f"ℹ️ Нет депозитов, начавшихся в период {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}")
                        else:
                            # Активные депозиты (без даты завершения)
                            active_deposits = deposit_report[deposit_report["Дата завершения"].isna()]
                            active_count = len(active_deposits)
                            active_amount = active_deposits["Сумма депозита (руб)"].sum() if not active_deposits.empty else 0.0
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                st.metric("📌 Кол-во депозитов активно (шт)", active_count)
                            with col2:
                                st.metric("💰 Общая сумма рублей на активных депозитах (руб)", f"{active_amount:,.2f} ₽")
                            
                            # Таблица
                            st.dataframe(
                                deposit_report,
                                use_container_width=True,
                                hide_index=True
                            )
                            
                            # Кнопка скачивания
                            excel_file = export_deposit_report_to_excel(deposit_report, st.session_state.ip_operations)
                            st.download_button(
                                label="📥 Скачать депозитный отчет Excel",
                                data=excel_file,
                                file_name=f"Депозитный_отчет_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_deposits"
                            )
            else:
                st.info("ℹ️ Нет данных ИП для формирования депозитного отчета")
        
        # ============================================
        # КНОПКА СКАЧИВАНИЯ ОСНОВНОГО ОТЧЕТА
        # ============================================
        if st.session_state.excel_data is not None:
            st.download_button(
                label="📥 Скачать отчет Excel",
                data=st.session_state.excel_data,
                file_name=st.session_state.report_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="download_excel"
            )

else:
    if not st.session_state.data_loaded:
        st.info("👆 Загрузите файлы и нажмите 'Обработать файлы'")
    else:
        st.warning("⚠️ Данные не загружены. Попробуйте перезагрузить файлы.")

# -------------------------------
# Информация о проекте
# -------------------------------
with st.expander("ℹ️ Информация о проекте"):
    st.markdown("""
    ### Как работает сервис:
    1. Загрузите два Excel-файла: выписку ИП и выписку физлица
    2. Нажмите "Обработать файлы" - данные будут проверены
    3. Выберите период отчета
    4. Нажмите "Сформировать отчет"
    5. Скачайте готовый отчет в Excel
    
    ### Формат файлов:
    **Выписка ИП:** колонки Дата, Дебет, Кредит, Назначение платежа
    
    **Выписка физлица:** колонки Дата операции, Описание, Сумма в валюте счета
    
    ### Важные замечания:
    - Начальный остаток физлица можно настроить (по умолчанию = 0)
    - Все суммы округляются до 2 знаков после запятой
    - Автоматическая проверка дубликатов и корректности данных
    - Визуализация динамики остатков в виде графиков
    """)
