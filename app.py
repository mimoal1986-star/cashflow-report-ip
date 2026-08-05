import streamlit as st
import pandas as pd
from datetime import datetime

from parsers import IPParser, FLParser, ParserError
from calculators import BalanceCalculator, BalanceCalculatorFL
from data_validators import DataValidator
from helpers import create_excel_report
from deposit_report import DepositReportGenerator

# Настройка страницы
st.set_page_config(
    page_title="Отчет по ДДС ИП и ФЛ",
    page_icon="💰",
    layout="wide"
)

st.title("💰 Автоматический отчет по движению денежных средств")

# Инициализация сессии
if "ip_operations" not in st.session_state:
    st.session_state.ip_operations = None
if "fl_operations" not in st.session_state:
    st.session_state.fl_operations = None
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = False
if "report_ready" not in st.session_state:
    st.session_state.report_ready = False
if "ip_report" not in st.session_state:
    st.session_state.ip_report = None
if "fl_report" not in st.session_state:
    st.session_state.fl_report = None
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
    file_fl = st.file_uploader(
        "Загрузите Excel-файл выписки физлица (необязательно)",
        type=["xlsx", "xls"],
        key="fl_upload"
    )

# Кнопка обработки
if st.button("🔄 Обработать файлы", type="primary"):
    if file_ip or file_fl:
        try:
            with st.spinner("Обработка файлов..."):
                fl_error = False
                
                # ============================================
                # Обработка ИП
                # ============================================
                if file_ip:
                    st.session_state.ip_operations = IPParser.parse(file_ip)
                    if not st.session_state.ip_operations.empty:
                        duplicates_ip = DataValidator.find_duplicates(st.session_state.ip_operations)
                        if not duplicates_ip.empty:
                            st.warning(f"⚠️ Обнаружены дублирующиеся операции в ИП: {len(duplicates_ip)} шт.")
                            with st.expander("📋 Показать дубликаты (ИП)"):
                                st.dataframe(
                                    duplicates_ip[["date", "amount", "description"]],
                                    use_container_width=True,
                                    hide_index=True
                                )
                        DataValidator.validate_amounts(st.session_state.ip_operations)
                else:
                    st.session_state.ip_operations = pd.DataFrame()
                
                # ============================================
                # Обработка ФЛ
                # ============================================
                if file_fl:
                    try:
                        st.session_state.fl_operations = FLParser.parse(file_fl)
                        if not st.session_state.fl_operations.empty:
                            duplicates_fl = DataValidator.find_duplicates(st.session_state.fl_operations)
                            if not duplicates_fl.empty:
                                st.warning(f"⚠️ Обнаружены дублирующиеся операции в ФЛ: {len(duplicates_fl)} шт.")
                                with st.expander("📋 Показать дубликаты (ФЛ)"):
                                    st.dataframe(
                                        duplicates_fl[["date", "amount", "description"]],
                                        use_container_width=True,
                                        hide_index=True
                                    )
                            DataValidator.validate_amounts(st.session_state.fl_operations)
                    except ParserError as e:
                        st.warning(f"⚠️ {str(e)}")
                        st.session_state.fl_operations = pd.DataFrame()
                        fl_error = True
                else:
                    st.session_state.fl_operations = pd.DataFrame()
                
                st.session_state.data_loaded = True
                st.session_state.report_ready = False
                
                st.success("✅ Файлы успешно обработаны!")
                
            # ============================================
            # СТАТИСТИКА (вне with spinner)
            # ============================================
            col1, col2 = st.columns(2)
            with col1:
                count_ip = len(st.session_state.ip_operations) if not st.session_state.ip_operations.empty else 0
                st.metric("📊 Операций ИП", count_ip)
            with col2:
                count_fl = len(st.session_state.fl_operations) if not st.session_state.fl_operations.empty else 0
                st.metric("📊 Операций ФЛ", count_fl)
            
            if fl_error:
                st.info("ℹ️ Расчет ФЛ не будет произведен из-за отсутствия колонок.")
                
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
if st.session_state.data_loaded:
    
    has_ip = st.session_state.ip_operations is not None and not st.session_state.ip_operations.empty
    has_fl = st.session_state.fl_operations is not None and not st.session_state.fl_operations.empty
    
    if not has_ip and not has_fl:
        st.warning("⚠️ Нет данных для формирования отчета")
        st.stop()
    
    # Определяем диапазон дат
    try:
        all_dates = []
        if has_ip:
            all_dates.extend(st.session_state.ip_operations["date"].tolist())
        if has_fl:
            all_dates.extend(st.session_state.fl_operations["date"].tolist())
        
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
    

    # ============================================
    # КНОПКА РАСЧЕТА
    # ============================================
    if st.button("📊 Сформировать отчет", type="primary"):
        try:
            DataValidator.validate_dates(start_date, end_date)
            
            with st.spinner("Расчет отчета..."):
                # Расчет ИП
                if has_ip:
                    ip_reports = BalanceCalculator.calculate(
                        st.session_state.ip_operations,
                        pd.Timestamp(start_date),
                        pd.Timestamp(end_date)
                    )
                    st.session_state.ip_report = ip_reports["ip"]
                else:
                    st.session_state.ip_report = None
                
                # Расчет ФЛ (проверяем, что нет ошибки)
                if has_fl:
                    # Проверяем, что fl_operations не пустой и нет ошибки
                    if not st.session_state.fl_operations.empty:
                        st.session_state.fl_report = BalanceCalculatorFL.calculate(
                            st.session_state.fl_operations,
                            pd.Timestamp(start_date),
                            pd.Timestamp(end_date)
                        )
                    else:
                        st.session_state.fl_report = None
                else:
                    st.session_state.fl_report = None
                
                # Создаем Excel файл
                excel_file = create_excel_report(
                    st.session_state.ip_report,
                    st.session_state.ip_operations,
                    st.session_state.fl_report,
                    st.session_state.fl_operations
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
    if st.session_state.report_ready:
        
        st.header("📈 Отчет по движению денежных средств")
        
        # ============================================
        # БЛОК ИП
        # ============================================
        if st.session_state.ip_report is not None:
            ip_report = st.session_state.ip_report
            
            st.subheader("🏢 Отчет по ИП")
            
            # Расчет "Из них на депозите"
            deposit_ops_all = st.session_state.ip_operations.attrs.get("deposits", pd.DataFrame()) if st.session_state.ip_operations is not None else pd.DataFrame()
            
            if not deposit_ops_all.empty:
                deposit_report_full = DepositReportGenerator.generate_report(deposit_ops_all)
                if not deposit_report_full.empty:
                    end_ts = pd.Timestamp(end_date)
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
            
            # Метрики ИП
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(
                    "🏢 Начальный остаток ИП",
                    f"{ip_report.start_balance / 1_000_000:.2f} млн ₽"
                )
            with col2:
                st.metric(
                    "🏢 Конечный остаток ИП",
                    f"{ip_report.end_balance / 1_000_000:.2f} млн ₽",
                    delta=f"{(ip_report.end_balance - ip_report.start_balance) / 1_000_000:+.2f} млн ₽"
                )
            with col3:
                st.metric(
                    "🏦 Из них на депозите",
                    f"{ip_on_deposit / 1_000_000:.2f} млн ₽"
                )
            
            # Вкладки ИП
            tab1_ip, tab2_ip = st.tabs(["📊 Динамика ИП", "🏦 Депозиты ИП"])
            
            with tab1_ip:
                st.subheader("Динамика остатка ИП помесячно")
                if not ip_report.monthly_dynamics.empty:
                    df_dynamics = ip_report.monthly_dynamics.copy()
                    df_dynamics["month_short"] = pd.to_datetime(
                        df_dynamics["month"], format="%B %Y"
                    ).dt.strftime("%b'%y")
                    
                    start_balance = df_dynamics["balance"].iloc[0] if not df_dynamics.empty else 0
                    df_dynamics["dynamics"] = df_dynamics["balance"].diff().fillna(0)
                    
                    ip_ops = st.session_state.ip_operations
                    if not ip_ops.empty:
                        ops = ip_ops.copy()
                        ops["month_period"] = ops["date"].dt.to_period("M").dt.strftime("%b'%y")
                        
                        monthly_income = ops[ops["amount"] > 0].groupby("month_period")["amount"].sum()
                        monthly_expense = ops[ops["amount"] < 0].groupby("month_period")["amount"].sum()
                        
                        months = df_dynamics["month_short"].tolist()
                        
                        table_data = {
                            "Показатель": [
                                "Начальный остаток, млн ₽",
                                "Конечный остаток, млн ₽",
                                "Динамика, млн ₽",
                                "Поступления, млн ₽",
                                "Списания, млн ₽"
                            ]
                        }
                        
                        for month in months:
                            row = df_dynamics[df_dynamics["month_short"] == month]
                            if not row.empty:
                                balance = row["balance"].iloc[0] / 1_000_000
                                dynamics = row["dynamics"].iloc[0] / 1_000_000
                                income = monthly_income.get(month, 0) / 1_000_000
                                expense = monthly_expense.get(month, 0) / 1_000_000
                            else:
                                balance = 0
                                dynamics = 0
                                income = 0
                                expense = 0
                            
                            if month == months[0]:
                                start_bal = start_balance / 1_000_000
                            else:
                                prev_idx = df_dynamics[df_dynamics["month_short"] == month].index
                                if len(prev_idx) > 0:
                                    idx = prev_idx[0]
                                    if idx > 0:
                                        start_bal = df_dynamics.iloc[idx - 1]["balance"] / 1_000_000
                                    else:
                                        start_bal = start_balance / 1_000_000
                                else:
                                    start_bal = 0
                            
                            table_data[month] = [
                                f"{start_bal:.2f}",
                                f"{balance:.2f}",
                                f"{dynamics:+.2f}",
                                f"{income:+.2f}",
                                f"{expense:+.2f}"
                            ]
                        
                        df_table = pd.DataFrame(table_data)
                        styled_df = df_table.style.hide(axis="index").set_properties(
                            **{'border-bottom': '2px solid #cccccc'}, 
                            subset=pd.IndexSlice[2, :]
                        )
                        st.dataframe(styled_df, use_container_width=True)
                    else:
                        st.info("Нет данных для отображения динамики ИП")
                else:
                    st.info("Нет данных для отображения динамики ИП")
            
            with tab2_ip:
                st.subheader("🏦 Отчет по депозитам ИП")
                if st.session_state.ip_operations is not None and not st.session_state.ip_operations.empty:
                    deposit_ops_all = st.session_state.ip_operations.attrs.get("deposits", pd.DataFrame())
                    if deposit_ops_all.empty:
                        st.info("ℹ️ Нет депозитных операций в выписке ИП")
                    else:
                        deposit_report_full = DepositReportGenerator.generate_report(deposit_ops_all)
                        if deposit_report_full.empty:
                            st.info("ℹ️ Не найдены депозитные операции с номерами сделок")
                        else:
                            start_ts = pd.Timestamp(start_date)
                            end_ts = pd.Timestamp(end_date)
                            deposit_report = deposit_report_full[
                                (deposit_report_full["Дата начала"] >= start_ts) & 
                                (deposit_report_full["Дата начала"] <= end_ts)
                            ].copy()
                            if deposit_report.empty:
                                st.info(f"ℹ️ Нет депозитов, начавшихся в период {start_date.strftime('%d.%m.%Y')} - {end_date.strftime('%d.%m.%Y')}")
                            else:
                                active_deposits = deposit_report[deposit_report["Дата завершения"].isna()]
                                active_count = len(active_deposits)
                                active_amount = active_deposits["Сумма депозита (руб)"].sum() if not active_deposits.empty else 0.0
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.metric("📌 Кол-во депозитов активно (шт)", active_count)
                                with col2:
                                    st.metric("💰 Общая сумма рублей на активных депозитах (руб)", f"{active_amount:,.2f} ₽")
                                st.dataframe(deposit_report, use_container_width=True, hide_index=True)
                else:
                    st.info("ℹ️ Нет данных ИП для формирования депозитного отчета")
        
        # ============================================
        # БЛОК ФЛ
        # ============================================
        if st.session_state.fl_report is not None:
            fl_report = st.session_state.fl_report
            
            st.subheader("👤 Отчет по физлицу")
            
            # Метрики ФЛ
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric(
                    "👤 Начальный остаток ФЛ",
                    f"{fl_report.start_balance / 1_000_000:.2f} млн ₽"
                )
            with col2:
                st.metric(
                    "👤 Конечный остаток ФЛ",
                    f"{fl_report.end_balance / 1_000_000:.2f} млн ₽",
                    delta=f"{(fl_report.end_balance - fl_report.start_balance) / 1_000_000:+.2f} млн ₽"
                )
            with col3:
                st.metric(
                    "🏦 Из них на вкладе",
                    f"{fl_report.deposits_on_end / 1_000_000:.2f} млн ₽"
                )
            
            # Вкладки ФЛ
            tab1_fl, tab2_fl = st.tabs(["📊 Динамика ФЛ", "🏦 Вклады"])
            
            with tab1_fl:
                st.subheader("Динамика остатка ФЛ помесячно")
                if not fl_report.monthly_dynamics.empty:
                    df_dynamics = fl_report.monthly_dynamics.copy()
                    df_dynamics["month_short"] = pd.to_datetime(
                        df_dynamics["month"], format="%B %Y"
                    ).dt.strftime("%b'%y")
                    
                    start_balance = df_dynamics["balance"].iloc[0] if not df_dynamics.empty else 0
                    df_dynamics["dynamics"] = df_dynamics["balance"].diff().fillna(0)
                    
                    # Для поступлений и списаний по ФЛ (только Текущий счет)
                    fl_ops = st.session_state.fl_operations
                    if not fl_ops.empty:
                        ops = fl_ops[fl_ops["account_name"] == "Текущий счёт"].copy()
                        ops["month_period"] = ops["date"].dt.to_period("M").dt.strftime("%b'%y")
                        
                        monthly_income = ops[ops["amount"] > 0].groupby("month_period")["amount"].sum()
                        monthly_expense = ops[ops["amount"] < 0].groupby("month_period")["amount"].sum()
                        
                        months = df_dynamics["month_short"].tolist()
                        
                        table_data = {
                            "Показатель": [
                                "Начальный остаток, млн ₽",
                                "Конечный остаток, млн ₽",
                                "Динамика, млн ₽",
                                "Поступления, млн ₽",
                                "Списания, млн ₽"
                            ]
                        }
                        
                        for month in months:
                            row = df_dynamics[df_dynamics["month_short"] == month]
                            if not row.empty:
                                balance = row["balance"].iloc[0] / 1_000_000
                                dynamics = row["dynamics"].iloc[0] / 1_000_000
                                income = monthly_income.get(month, 0) / 1_000_000
                                expense = monthly_expense.get(month, 0) / 1_000_000
                            else:
                                balance = 0
                                dynamics = 0
                                income = 0
                                expense = 0
                            
                            if month == months[0]:
                                start_bal = start_balance / 1_000_000
                            else:
                                prev_idx = df_dynamics[df_dynamics["month_short"] == month].index
                                if len(prev_idx) > 0:
                                    idx = prev_idx[0]
                                    if idx > 0:
                                        start_bal = df_dynamics.iloc[idx - 1]["balance"] / 1_000_000
                                    else:
                                        start_bal = start_balance / 1_000_000
                                else:
                                    start_bal = 0
                            
                            table_data[month] = [
                                f"{start_bal:.2f}",
                                f"{balance:.2f}",
                                f"{dynamics:+.2f}",
                                f"{income:+.2f}",
                                f"{expense:+.2f}"
                            ]
                        
                        df_table = pd.DataFrame(table_data)
                        styled_df = df_table.style.hide(axis="index").set_properties(
                            **{'border-bottom': '2px solid #cccccc'}, 
                            subset=pd.IndexSlice[2, :]
                        )
                        st.dataframe(styled_df, use_container_width=True)
                    else:
                        st.info("Нет данных для отображения динамики ФЛ")
                else:
                    st.info("Нет данных для отображения динамики ФЛ")
            
            with tab2_fl:
                st.subheader("🏦 Отчет по вкладам ФЛ")
                
                if fl_report.deposits_data:
                    for deposit_item in fl_report.deposits_data:
                        df_dep = deposit_item["data"].copy()
                        if not df_dep.empty:
                            df_dep["month"] = pd.to_datetime(df_dep["month"], format="%B %Y").dt.strftime("%b'%y")
                            df_dep = df_dep.rename(columns={
                                "month": "Месяц",
                                "balance": "Остаток на конец месяца",
                                "interest": "Выплата процентов"
                            })
                            st.subheader(f"📊 {deposit_item['account_name']}")
                            st.dataframe(
                                df_dep,
                                use_container_width=True,
                                hide_index=True
                            )
                else:
                    st.info("ℹ️ Нет данных по вкладам ФЛ")
        
        # ============================================
        # КНОПКА СКАЧИВАНИЯ
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
    1. Загрузите Excel-файл выписки ИП (обязательно)
    2. Загрузите Excel-файл выписки физлица (опционально)
    3. Нажмите "Обработать файлы"
    4. Выберите период отчета
    5. Нажмите "Сформировать отчет"
    6. Скачайте готовый отчет в Excel
    
    ### Формат файлов:
    **Выписка ИП:** колонки Дата, Дебет, Кредит, Назначение платежа
    
    **Выписка физлица:** колонки Дата операции, Название счета, Номер счета, 
    Описание операции, Сумма, Статус, Категория, Тип, Комментарий
    
    ### Важные замечания:
    - Все суммы отображаются в млн ₽
    - По ФЛ: Текущий счёт — основной, Альфа-Счёт... — вклады
    """)
