import pandas as pd
import input_database
import constants
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook


def hints_month_excel():
    end_of_month_excel = constants.LONGSHORT_DIR + f"/10304_일자별기준가격 조회_{datetime.today().month}월.xlsx"
    eom_df = pd.read_excel(end_of_month_excel)
    df_cleaned = eom_df.dropna(subset=['No', '통화코드'])
    fund_dict = {
        'ipo_1': 'DM12001', 'ipo_2': 'DM12002', 'ipo_f': 'DM12003', 'ipo_a': 'DM12007',
        'kv_1': 'DM13001', 'kv_2': 'DM13020', 'kv_3': 'DM13030',
        'multi_1': 'DM14001', 'multi_2': 'DM14021', 'ipo_b': 'DM16001', 'prelude': 'DM15012'
    }
    eom_pnl_dict = {}
    for key, fund_code in fund_dict.items():
        try:
            eom_pnl_dict[key] = df_cleaned.loc[df_cleaned['펀드코드'] == fund_code, 'Unnamed: 9'].tolist()[0]
        except IndexError:
            eom_pnl_dict[key] = None  # 값이 없을 경우 None으로 설정

    return eom_pnl_dict


def hints_yesterday_excel():
    yesterday = (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")
    yesterday_excel = constants.LONGSHORT_DIR + f"/10304_일자별기준가격 조회_{yesterday}.xlsx"
    eom_df = pd.read_excel(yesterday_excel)
    df_cleaned = eom_df.dropna(subset=['No', '통화코드'])
    fund_dict = {
        'ipo_1': 'DM12001', 'ipo_2': 'DM12002', 'ipo_f': 'DM12003', 'ipo_a': 'DM12007',
        'kv_1': 'DM13001', 'kv_2': 'DM13020', 'kv_3': 'DM13030',
        'multi_1': 'DM14001', 'multi_2': 'DM14021', 'ipo_b': 'DM16001', 'prelude': 'DM15012'
    }
    ex_pnl_dict = {}
    for key, fund_code in fund_dict.items():
        try:
            ex_pnl_dict[key] = df_cleaned.loc[df_cleaned['펀드코드'] == fund_code, 'Unnamed: 9'].tolist()[0]
        except IndexError:
            ex_pnl_dict[key] = None  # 값이 없을 경우 None으로 설정

    return ex_pnl_dict


def hints_1ytd_excel():
    end_of_year_excel = constants.LONGSHORT_DIR + f"/10304_일자별기준가격 조회_2024.xlsx"
    eoy_df = pd.read_excel(end_of_year_excel)
    df_cleaned = eoy_df.dropna(subset=['통화코드'])
    fund_dict = {
        'ipo_1': 'DM12001', 'ipo_2': 'DM12002', 'ipo_f': 'DM12003', 'ipo_a': 'DM12007',
        'kv_1': 'DM13001', 'kv_2': 'DM13020', 'kv_3': 'DM13030',
        'multi_1': 'DM14001', 'multi_2': 'DM14021', 'ipo_b': 'DM16001', 'prelude': 'DM15012'
    }
    one_year_ago = datetime.today() - relativedelta(years=1)
    df_cleaned.loc[:, '날짜'] = pd.to_datetime(df_cleaned['기준가산출일'], format='%Y-%m-%d')
    eom_pnl_dict = {}
    for key, fund_code in fund_dict.items():
        try:
            condition = (df_cleaned['펀드코드'] == fund_code) & (df_cleaned['날짜'] >= one_year_ago)
            eom_pnl_dict[key] = df_cleaned.loc[condition, 'Unnamed: 9'].tolist()[0]
        except IndexError:
            if fund_code == "15012":
                eom_pnl_dict[key] = 10
            else:
                eom_pnl_dict[key] = 1000

    return eom_pnl_dict


if __name__ == "__main__":
    print(hints_month_excel())