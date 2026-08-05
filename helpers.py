        # ============================================
        # ФЛ_Динамика
        # ============================================
        if fl_report is not None and not fl_report.monthly_dynamics.empty:
            df_fl_dynamics = fl_report.monthly_dynamics.copy()
            
            # Получаем начальные остатки для каждого месяца
            balances = []
            for i in range(len(df_fl_dynamics)):
                if i == 0:
                    start_bal = fl_report.start_balance
                else:
                    start_bal = df_fl_dynamics.iloc[i-1]["balance"]
                balances.append(start_bal)
            
            df_fl_dynamics["Баланс начало месяца"] = balances
            df_fl_dynamics["Баланс конец месяца"] = df_fl_dynamics["balance"]
            df_fl_dynamics = df_fl_dynamics.rename(columns={"month": "Месяц"})
            
            result_df = df_fl_dynamics[["Месяц", "Баланс начало месяца", "Баланс конец месяца"]].copy()
            result_df["Баланс начало месяца"] = result_df["Баланс начало месяца"].apply(lambda x: f"{x:,.2f}")
            result_df["Баланс конец месяца"] = result_df["Баланс конец месяца"].apply(lambda x: f"{x:,.2f}")
            
            result_df.to_excel(writer, sheet_name="ФЛ_Динамика", index=False)
        else:
            empty_df = pd.DataFrame({"Сообщение": ["Нет данных для динамики ФЛ"]})
            empty_df.to_excel(writer, sheet_name="ФЛ_Динамика", index=False)
        
        # ============================================
        # ФЛ_Вклады
        # ============================================
        if fl_report is not None and fl_report.deposits_data:
            # Объединяем данные по всем вкладам в один лист
            all_deposits = []
            for deposit_item in fl_report.deposits_data:
                df_dep = deposit_item["data"].copy()
                df_dep["Название счета"] = deposit_item["account_name"]
                all_deposits.append(df_dep)
            
            if all_deposits:
                combined_df = pd.concat(all_deposits, ignore_index=True)
                combined_df = combined_df[["Название счета", "month", "balance", "interest"]]
                combined_df = combined_df.rename(columns={
                    "month": "Месяц",
                    "balance": "Остаток на конец месяца",
                    "interest": "Выплата процентов"
                })
                combined_df["Месяц"] = pd.to_datetime(combined_df["Месяц"], format="%B %Y").dt.strftime("%B %Y")
                combined_df["Остаток на конец месяца"] = combined_df["Остаток на конец месяца"].apply(lambda x: f"{x:,.2f}")
                combined_df["Выплата процентов"] = combined_df["Выплата процентов"].apply(lambda x: f"{x:,.2f}")
                combined_df.to_excel(writer, sheet_name="ФЛ_Вклады", index=False)
            else:
                empty_df = pd.DataFrame({"Сообщение": ["Нет данных по вкладам ФЛ"]})
                empty_df.to_excel(writer, sheet_name="ФЛ_Вклады", index=False)
        else:
            empty_df = pd.DataFrame({"Сообщение": ["Нет данных по вкладам ФЛ"]})
            empty_df.to_excel(writer, sheet_name="ФЛ_Вклады", index=False)
