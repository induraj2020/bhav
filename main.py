from typing import final
import streamlit as st
import pandas as pd
import io
import ast
import warnings
from datetime import datetime
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.filterwarnings("ignore")
pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 1000)
import sys
import re
from datetime import datetime, timedelta
from pytz import timezone, all_timezones
import time
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
from openpyxl import load_workbook
from openpyxl.styles import Font
pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 1000)
from openpyxl import load_workbook
from utils import *
from io import BytesIO

# Title
st.title("Bhav Analyzer")

# Integer inputs for Nifty and Bank spot values
nifty_spot = st.number_input("Enter Nifty Spot value:", min_value=0, value=10000)
bank_spot = st.number_input("Enter Bank Spot value:", min_value=0, value=25000)
fin_spot = st.number_input("Enter Fin Spot value:", min_value=0, value=25000)
own_file_main = st.file_uploader("Upload Main file", type=["csv", "xlsx"])
bhav_copy_file = st.file_uploader("Upload bhav copy from nse", type=["csv", "xlsx"])
enter_nifty_weekly = st.date_input('select nifty weekly')
enter_nifty_monthly = st.date_input('select nifty monthly')
enter_bank_weekly = st.date_input('select bank weekly')
enter_bank_monthly = st.date_input('select bank monthly')
enter_fin_weekly = st.date_input('select fin weekly')
enter_fin_monthly = st.date_input('select fin monthly')

cols = ['Date', 'Expiry', 'Spot', 
        'CE_EOD', 'CE_CHANGE',  'ITM_CE_EOD', 'ITM_CE_CHANGE', 'CE_%CHANGE', 'ITM_CE_%EOD', 'ITM_CE_%CHANGE', 
        'PE_EOD', 'PE_CHANGE', 'ITM_PE_EOD', 'ITM_PE_CHANGE', 'PE_%CHANGE', 'ITM_PE_%EOD', 'ITM_PE_%CHANGE']

if own_file_main and bhav_copy_file and enter_nifty_weekly and enter_nifty_monthly and enter_bank_weekly and enter_bank_monthly and enter_fin_weekly and enter_fin_monthly:
    try:
        date_match = re.search(r'(\d{8})', bhav_copy_file.name)
        if date_match:
            date_str = date_match.group(1)
            date_obj = datetime.strptime(date_str, '%Y%m%d')
            formatted_date = date_obj.strftime('%d-%m-%y')
            new_filename = '{}.xlsx'.format(formatted_date)

        df = pd.read_csv(bhav_copy_file)
        df['TradDt'] = pd.to_datetime(df['TradDt'])
        df['BizDt'] = pd.to_datetime(df['BizDt'])
        df['XpryDt'] = pd.to_datetime(df['XpryDt'])
        df_org = df.copy()

        df = df[df['FinInstrmTp']=='IDO'].reset_index(drop=True)
        col_to_keep = ['TradDt', 'TckrSymb', 'XpryDt',  'StrkPric','OpnIntrst','ChngInOpnIntrst', 'OptnTp', 'ClsPric']
        df = df[col_to_keep]

        nifty_df = df[df['TckrSymb']=='NIFTY'].reset_index(drop=True)
        bank_df = df[df['TckrSymb']=='BANKNIFTY'].reset_index(drop=True)
        finnifty_df = df[df['TckrSymb']=='FINNIFTY'].reset_index(drop=True)

        final_df_nifty_weekly = pd.DataFrame(columns=cols)
        final_df_nifty_monthly = pd.DataFrame(columns=cols)
        final_df_bank_weekly = pd.DataFrame(columns=cols)
        final_df_bank_monthly = pd.DataFrame(columns=cols)
        final_df_fin_weekly = pd.DataFrame(columns=cols)
        final_df_fin_monthly = pd.DataFrame(columns=cols)

        nifty_weekly_expirty_CE_df, nifty_weekly_expirty_PE_df = filter_by_ce_pe(nifty_df, enter_nifty_weekly)
        nifty_monthly_expirty_CE_df, nifty_monthly_expirty_PE_df = filter_by_ce_pe(nifty_df, enter_nifty_monthly)

        bank_weekly_expirty_CE_df, bank_weekly_expirty_PE_df = filter_by_ce_pe(bank_df, enter_bank_weekly)
        bank_monthly_expirty_CE_df, bank_monthly_expirty_PE_df = filter_by_ce_pe(bank_df, enter_bank_monthly)

        fin_weekly_expirty_CE_df, fin_weekly_expirty_PE_df = filter_by_ce_pe(finnifty_df, enter_fin_weekly)
        fin_monthly_expirty_CE_df, fin_monthly_expirty_PE_df = filter_by_ce_pe(finnifty_df, enter_fin_monthly)

        #--------------
        nifty_weekly_changes_df = find_changes(nifty_weekly_expirty_CE_df, nifty_weekly_expirty_PE_df, nifty_spot)
        nifty_monthly_changes_df = find_changes(nifty_monthly_expirty_CE_df, nifty_monthly_expirty_PE_df,nifty_spot)

        bank_weekly_changes_df = find_changes(bank_weekly_expirty_CE_df, bank_weekly_expirty_PE_df, bank_spot)
        bank_monthly_changes_df = find_changes(bank_monthly_expirty_CE_df, bank_monthly_expirty_PE_df, bank_spot)

        fin_weekly_changes_df = find_changes(fin_weekly_expirty_CE_df, fin_weekly_expirty_PE_df, fin_spot)
        fin_monthly_changes_df = find_changes(fin_monthly_expirty_CE_df, fin_monthly_expirty_PE_df, fin_spot)

        #----------------
        final_df_nifty_weekly = pd.concat([final_df_nifty_weekly,nifty_weekly_changes_df],ignore_index=True)
        final_df_nifty_monthly = pd.concat([final_df_nifty_monthly,nifty_monthly_changes_df],ignore_index=True)
        final_df_bank_weekly = pd.concat([final_df_bank_weekly,bank_weekly_changes_df],ignore_index=True)
        final_df_bank_monthly = pd.concat([final_df_bank_monthly,bank_monthly_changes_df],ignore_index=True)
        final_df_fin_weekly = pd.concat([final_df_fin_weekly,fin_weekly_changes_df],ignore_index=True)
        final_df_fin_monthly = pd.concat([final_df_fin_monthly,fin_monthly_changes_df],ignore_index=True)

        final_df1 = save_df_as_excel(own_file_main,final_df_nifty_weekly,'Nifty-W')
        final_df2 = save_df_as_excel(own_file_main,final_df_nifty_monthly,'Nifty-M')
        final_df3 = save_df_as_excel(own_file_main,final_df_bank_weekly,'Bank-W')
        final_df4 = save_df_as_excel(own_file_main,final_df_bank_monthly,'Bank-M')
        final_df5 = save_df_as_excel(own_file_main,final_df_fin_weekly,'Fin-W')
        final_df6 = save_df_as_excel(own_file_main,final_df_fin_monthly,'Fin-M')


        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df1.to_excel(writer, sheet_name='Nifty-W', index=False)
            final_df2.to_excel(writer, sheet_name='Nifty-M', index=False)
            final_df3.to_excel(writer, sheet_name='Bank-W', index=False)
            final_df4.to_excel(writer, sheet_name='Bank-M', index=False)
            final_df5.to_excel(writer, sheet_name='Fin-W', index=False)
            final_df6.to_excel(writer, sheet_name='Fin-M', index=False)
        output.seek(0)
        st.download_button( label="Download bhav copy File", data=output, file_name="bhav_copy.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        output1 = io.BytesIO()
        with pd.ExcelWriter(output1, engine='openpyxl') as writer:
            nifty_weekly_expirty_CE_df.to_excel(writer, sheet_name='Nifty-CE-W', index=False)
            nifty_monthly_expirty_CE_df.to_excel(writer, sheet_name='Nifty-CE-M', index=False)
            nifty_weekly_expirty_PE_df.to_excel(writer, sheet_name='Nifty-PE-W', index=False)
            nifty_monthly_expirty_PE_df.to_excel(writer, sheet_name='Nifty-PE-M', index=False)

            bank_weekly_expirty_CE_df.to_excel(writer, sheet_name='Bank-CE-W', index=False)
            bank_monthly_expirty_CE_df.to_excel(writer, sheet_name='Bank-CE-M', index=False)
            bank_weekly_expirty_PE_df.to_excel(writer, sheet_name='Bank-PE-W', index=False)
            bank_monthly_expirty_PE_df.to_excel(writer, sheet_name='Bank-PE-M', index=False)

            fin_weekly_expirty_CE_df.to_excel(writer, sheet_name='Fin-CE-W', index=False)
            fin_monthly_expirty_CE_df.to_excel(writer, sheet_name='Fin-CE-M', index=False)
            fin_weekly_expirty_PE_df.to_excel(writer, sheet_name='Fin-PE-W', index=False)
            fin_monthly_expirty_PE_df.to_excel(writer, sheet_name='Fin-PE-M', index=False)
        output1.seek(0)
        st.download_button( label="Download daily File", data=output1, file_name=f"{new_filename}", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        print(e)
        st.write('Check ur input data. There is some mistake in the input')