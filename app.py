import streamlit as st
import pandas as pd
from datetime import datetime
import logging

from parsers import IPParser, PhysParser, ParserError
from calculators import BalanceCalculator
from data_validators import DataValidator
from helpers import get_date_range, create_excel_report

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
                # Парсим только загруженные файлы
                if file_ip:
                    st.session_state.ip_operations = IPParser.parse(file_ip)
                else:
                    st.session_state.ip_operations = pd.DataFrame()
                
                if file_phys:
                    st.session_state.phys_operations = PhysParser.parse(file_phys)
                else:
                    st.session_state.phys_operations = pd.DataFrame()
                
                # ============================================
                # ВАЛИДАЦИЯ ДАННЫХ (с показом дубликатов)
                # ============================================
                
                # Проверка дубликатов в ИП
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
                        # НЕ останавливаем работу!
                
                # Проверка дубликатов в физлице
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
                        # НЕ останавливаем работу!
                
                # Проверка сумм
                if not st.session_state.ip_operations.empty:
                    DataValidator.validate_amounts(st.session_state.ip_operations)
                
                if not st.session_state.phys_operations.empty:
                    DataValidator.validate_amounts(st.session_state.phys_operations)
                
                st.session_state.data_loaded = True
                
                st.success("✅ Файлы успешно обработаны и проверены!")
                
                # Показываем статистику
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
                
                # Предупреждение если загружен только 1 файл
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
    
    # Проверяем, что есть данные хотя бы в одном файле
    ip_empty = st.session_state.ip_operations.empty if st.session_state.ip_operations is not None else True
    phys_empty = st.session_state.phys_operations.empty if st.session_state.phys_operations is not None else True
    
    if ip_empty and phys_empty:
        st.warning("⚠️ Нет данных для формирования отчета")
        st.stop()
    
    # Определяем диапазон дат (из доступных данных)
    try:
        # Собираем все доступные даты
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
    
    # Опция: учитывать ли операции физлица до start_date
    use_zero_start = st.checkbox(
        "Начальный остаток физлица = 0 (согласно условию)",
        value=True,
        help="Если включено, все операции физлица до даты начала игнорируются"
    )
    
    # Кнопка расчета
    if st.button("📊 Сформировать отчет", type="primary"):
        try:
            # Валидация
            DataValidator.validate_dates(start_date, end_date)
            
            with st.spinner("Расчет отчета..."):
                # Расчет
                reports = BalanceCalculator.calculate(
                    st.session_state.ip_operations,
                    st.session_state.phys_operations,
                    pd.Timestamp(start_date),
                    pd.Timestamp(end_date)
                )
                
                ip_report = reports["ip"]
                phys_report = reports["phys"]
                
                # Вывод результатов
                st.header("📈 Отчет по движению денежных средств")
                
                # Карточки с остатками
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric(
                        "🏢 Начальный остаток ИП",
                        f"{ip_report.start_balance:,.2f} ₽",
                        delta=f"{ip_report.end_balance - ip_report.start_balance:,.2f} ₽"
                    )
                with col2:
                    st.metric(
                        "🏢 Конечный остаток ИП",
                        f"{ip_report.end_balance:,.2f} ₽"
                    )
                with col3:
                    st.metric(
                        "👤 Начальный остаток физлица",
                        f"{phys_report.start_balance:,.2f} ₽",
                        delta=f"{phys_report.end_balance - phys_report.start_balance:,.2f} ₽"
                    )
                with col4:
                    st.metric(
                        "👤 Конечный остаток физлица",
                        f"{phys_report.end_balance:,.2f} ₽"
                    )
                
                # Общий остаток
                total_start = ip_report.start_balance + phys_report.start_balance
                total_end = ip_report.end_balance + phys_report.end_balance
                st.info(f"💰 **Общий остаток:** {total_start:,.2f} ₽ → {total_end:,.2f} ₽ (изменение: {total_end - total_start:,.2f} ₽)")
                
                # Таблицы динамики
                tab1, tab2 = st.tabs(["📊 Динамика ИП", "📊 Динамика физлица"])
                
                with tab1:
                    st.subheader("Динамика остатка ИП помесячно")
                    if not ip_report.monthly_dynamics.empty:
                        # Добавляем визуализацию
                        st.line_chart(
                            ip_report.monthly_dynamics.set_index("month")["balance"]
                        )
                        st.dataframe(
                            ip_report.monthly_dynamics,
                            use_container_width=True,
                            hide_index=True
                        )
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
                
                # Кнопка скачивания
                try:
                    excel_file = create_excel_report(
                        ip_report,
                        phys_report,
                        st.session_state.ip_operations,
                        st.session_state.phys_operations
                    )
                    
                    st.download_button(
                        label="📥 Скачать отчет Excel",
                        data=excel_file,
                        file_name=f"Отчет_ДДС_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                except Exception as e:
                    st.error(f"❌ Ошибка создания Excel-файла: {str(e)}")
                
        except ValueError as e:
            st.error(f"❌ Ошибка валидации: {str(e)}")
        except Exception as e:
            st.error(f"❌ Ошибка при формировании отчета: {str(e)}")
else:
    if not st.session_state.data_loaded:
        st.info("👆 Загрузитехотя бы один файл и нажмите 'Обработать файлы'")
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
