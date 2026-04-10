import pandas as pd
from tabulate import tabulate
import FinanceDataReader as fdr
from datetime import datetime, timedelta



def read_oms_futures_file(target_date):

    today = target_date.strftime("%Y-%m-%d")

    # 1. 엑셀 파일 불러오기
    file_path = f'Z:/02.펀드/019. 일간매매내역/파생/{today}_futures.xlsx'  # 파일 경로를 적절히 수정하세요.
    df = pd.read_excel(file_path)

    # 2. 필요한 칼럼만 남기기
    columns_to_keep = ['펀드명', '매매\n구분', '종목코드', '종목명', '체결단가', '체결\n계약수', '약정금액']
    df = df[columns_to_keep]

    # 칼럼 이름 정리
    df.columns = ['펀드명', '매매구분', '종목코드', '종목명', '체결단가', '체결계약수', '약정금액']

    # 3. 펀드명 변환
    fund_name_mapping = {
        "두나미스 공모주멀티 일반사모투자신탁": "공모주1호",
        "두나미스 공모주 일반사모투자신탁 제2호": "공모주2호",
        "두나미스 공모주 포커스 일반사모투자신탁": "포커스",
        "두나미스 공모주 알파 일반사모투자신탁 운": "알파",
        "두나미스 코스닥벤처 일반사모투자신탁": "코스닥벤처1호",
        "두나미스 코스닥벤처 일반사모투자신탁 2호": "코스닥벤처2호",
        "두나미스 코스닥벤처 일반사모투자신탁 3호": "코스닥벤처3호",
        "두나미스 멀티전략 일반사모(운)": "멀티1호",
        "두나미스 멀티전략 일반사모투자신탁 2호": "멀티2호",
        "두나미스 블록딜공모주 일반사모투자신탁 1호(운용)": "블록딜",
        "DUNAMIS_PRELUDE 일임 (USD)": "Prelude"
    }
    df['펀드명'] = df['펀드명'].replace(fund_name_mapping)

    # 4. 종목코드 변환 (앞 3자리와 뒤 5자리만 남기기)
    df['종목코드'] = df['종목코드'].str[3:8]

    # 단축코드 6자리로 수정
    df['종목코드'] = df['종목코드'].astype(str).str.zfill(6)

    # 5. 칼럼 이름 변경
    column_rename_mapping = {
        "종목코드": "단축코드",
        "체결계약수": "체결수량",
        "약정금액": "체결금액"
    }
    df.rename(columns=column_rename_mapping, inplace=True)

    # 6. 가공된 데이터프레임 저장
    oms_futures_df = df

    # 결과 확인
    # print(tabulate(oms_futures_df, headers = 'keys', tablefmt = 'pretty'))

    return oms_futures_df


def read_oms_stock_file(target_date):

    today = target_date.strftime("%Y-%m-%d")

    # 1. 엑셀 파일 불러오기
    file_path = f'Z:/02.펀드/019. 일간매매내역/주식/{today}_stock.xlsx'  # 파일 경로를 적절히 수정하세요.
    df = pd.read_excel(file_path)

    # 2. 필요한 칼럼만 남기기
    columns_to_keep = ['펀드명', '매매구분', '종목명', '체결가격', '체결수량', '매매금액']
    df = df[columns_to_keep]

    # 3. 펀드명 변환
    fund_name_mapping = {
        "두나미스 공모주멀티 일반사모투자신탁": "공모주1호",
        "두나미스 공모주 일반사모투자신탁 제2호": "공모주2호",
        "두나미스 공모주 포커스 일반사모투자신탁": "포커스",
        "두나미스 공모주 알파 일반사모투자신탁 운용": "알파",
        "두나미스 코스닥벤처 일반사모투자신탁": "코스닥벤처1호",
        "두나미스 코스닥벤처 일반사모투자신탁 2호": "코스닥벤처2호",
        "두나미스 코스닥벤처 일반사모투자신탁 3호": "코스닥벤처3호",
        "두나미스 멀티전략 일반사모(운)": "멀티1호",
        "두나미스 멀티전략 일반사모투자신탁 2호": "멀티2호",
        "두나미스 블록딜공모주 일반사모투자신탁 1호(운용)": "블록딜",
        "DUNAMIS_PRELUDE 일임 (USD)": "Prelude"
    }
    df['펀드명'] = df['펀드명'].replace(fund_name_mapping)

    # 종목명 클렌징
    df['종목명'] = df['종목명'].str.strip()

    # 4. FinanceDataReader를 사용해 종목명으로 단축코드 생성
    # pykrx의 API 통신 에러(KeyError: '시장') 발생으로 인해 FinanceDataReader로 전환
    krx_df = fdr.StockListing('KRX') # 코스피, 코스닥 종목 포함
    etf_df = fdr.StockListing('ETF/KR') # 국내 ETF 종목 포함
    
    krx_mapping = dict(zip(krx_df['Name'], krx_df['Code']))
    etf_mapping = dict(zip(etf_df['Name'], etf_df['Symbol']))

    # 두 시장의 매핑 통합
    stock_code_mapping = {**krx_mapping, **etf_mapping}

    # 종목명으로 단축코드 매핑
    df['단축코드'] = df['종목명'].map(stock_code_mapping)
    
    # 매핑되지 않은 종목명 확인
    unmatched = df[df['단축코드'].isna()]['종목명'].unique()
    if unmatched.size > 0:
        print("매핑되지 않은 종목명:", unmatched)

    # 단축코드 칼럼을 종목명 앞에 배치
    df = df[['펀드명', '매매구분', '단축코드', '종목명', '체결가격', '체결수량', '매매금액']]

    # 5. 칼럼 이름 변경
    df.rename(columns={
        "체결가격": "체결단가",
        "매매금액": "체결금액"
    }, inplace=True)

    # 6. 가공된 데이터프레임 저장
    oms_stock_df = df

    # 결과 확인
    # print(tabulate(oms_stock_df, headers = 'keys', tablefmt = 'pretty'))

    return oms_stock_df

def read_prelude_stock_trade_history(target_date):

    today = target_date.strftime("%m%d%y")

    file_path = f"Z:/02.펀드/003.매매보고서 대사/PRELUDE_RECAP/Korea Stocks - {today}.xls"
    df = pd.read_excel(file_path)

    # 2. 필요한 칼럼만 남기기
    columns_to_keep = [
        'Client Account Name', 'Buy/Sell', 'Ric', 'Stock Description', 'Price (Gross) Listing Ccy', 'Stock Quantity'
    ]
    df = df[columns_to_keep]

    # 3. 칼럼 이름 변경
    df.columns = ['펀드명', '매매구분', '단축코드', '종목명', '체결단가', '체결수량']

    # 4. 종목코드 앞의 여섯 자리만 남기기
    df['단축코드'] = df['단축코드'].astype(str).str[:6]

    # 5. 매매금액 칼럼 추가 (체결가격 * 체결수량)
    df['체결금액'] = df['체결단가'] * df['체결수량']

    # 6. 펀드명 데이터를 'Prelude'로 변경
    df['펀드명'] = 'Prelude'

    # 7. 매매구분 값을 변환 (Buy -> 매수, Sell -> 매도)
    df['매매구분'] = df['매매구분'].replace({'Buy': '매수', 'Sell': '매도', 'Sell Short': '매도', 'Buy Cover': '매수', 'Short sell': '매도', 'Buy cover': '매수'})

    prelude_stock_df = df

    # print(tabulate(df, headers = 'keys', tablefmt = 'pretty'))

    return prelude_stock_df


def read_trader_file(target_date):

    today = f"{target_date.month}월{target_date.day}일"

    # 1. 엑셀 파일 불러오기
    file_path = f'Z:/02.펀드/019. 일간매매내역/{today} 전체.xlsx'  # 파일 경로를 적절히 수정하세요.
    df = pd.read_excel(file_path)

    # 2. 필요한 칼럼만 남기기
    columns_to_keep = ['펀드', '매매구분', '단축코드', '종목명', '체결단가', '체결수량', '체결금액']
    df = df[columns_to_keep]

    # 3. 단축코드 6자리로 수정
    df['단축코드'] = df['단축코드'].astype(str).str.zfill(6)

    # 4. 칼럼 이름 변경
    df.rename(columns={
        "펀드": "펀드명"
    }, inplace=True)

    # 5. 매매구분 변환
    direction_mapping = {
        "Buy": "매수",
        "Sell": "매도",
        "Sell Short": "매도",
        "Buy Cover": "매수",
        "Short sell": "매도", 
        'Buy cover': '매수'
    }
    df['매매구분'] = df['매매구분'].replace(direction_mapping)

    # 데이터프레임 저장
    trader_df = df

    # print(tabulate(trader_df, headers = 'keys', tablefmt = 'pretty'))

    return trader_df


def aggregate_by_key(df):
    # 펀드명, 매매구분, 단축코드로 그룹화
    aggregated_df = df.groupby(['펀드명', '매매구분', '단축코드'], as_index=False).agg({
        '종목명': 'first',  # 그룹에서 첫 번째 종목명을 사용
        '체결수량': 'sum',
        '체결금액': 'sum',
        '체결단가': lambda x: (x * df.loc[x.index, '체결수량']).sum() / df.loc[x.index, '체결수량'].sum()
    })
    return aggregated_df


def merge_and_reconcile(output_path, target_date):
    # OMS Futures와 Stock 데이터를 읽기
    oms_futures_df = read_oms_futures_file(target_date)
    oms_stock_df = read_oms_stock_file(target_date)
    prelude_stock_df = read_prelude_stock_trade_history(target_date)

    # 1. oms_futures_df를 oms_stock_df 아래에 병합
    trade_history_combined = pd.concat([oms_stock_df, oms_futures_df, prelude_stock_df], ignore_index=True)
    # trade_history_combined = pd.concat([oms_stock_df, oms_futures_df], ignore_index=True)
    aggregated_oms = aggregate_by_key(trade_history_combined)

    # Trader 데이터를 읽기
    trader_df = read_trader_file(target_date)
    aggregated_trader = aggregate_by_key(trader_df)

    # 2. 대사 작업: 펀드명, 매매구분, 단축코드가 일치하는 행에서 체결단가, 체결수량, 체결금액 비교
    reconciliation_df = aggregated_oms.merge(
        aggregated_trader,
        on=['펀드명', '매매구분', '단축코드'],
        suffixes=('_oms', '_trader'),
        how='outer',
        indicator=True
    )

    # 체결단가, 체결수량, 체결금액의 차이 계산
    reconciliation_df['체결단가_차이'] = reconciliation_df['체결단가_oms'] - reconciliation_df['체결단가_trader']
    reconciliation_df['체결수량_차이'] = reconciliation_df['체결수량_oms'] - reconciliation_df['체결수량_trader']
    reconciliation_df['체결금액_차이'] = reconciliation_df['체결금액_oms'] - reconciliation_df['체결금액_trader']

    # 종목명 유지: Trader의 종목명을 우선 사용, 없으면 OMS의 종목명을 사용
    # reconciliation_df['종목명'] = reconciliation_df['종목명_trader'].combine_first(reconciliation_df['종목명_oms'])

    # 필요 없는 종목명 관련 컬럼 제거
    # reconciliation_df.drop(columns=['종목명_trader', '종목명_oms'], inplace=True, errors='ignore')

    # 결과 출력
    print("대사 결과:")
    # print(tabulate(reconciliation_df[['펀드명', '매매구분', '종목명', '체결단가_차이', '체결수량_차이', '체결금액_차이']], headers = 'keys', tablefmt = 'pretty'))
    print(tabulate(reconciliation_df[['펀드명', '매매구분', '단축코드', '체결단가_차이', '체결수량_차이', '체결금액_차이']], headers = 'keys', tablefmt = 'pretty'))

    # 결과를 파일로 저장 (옵션)
    reconciliation_df.to_excel(output_path, index=False)

    return reconciliation_df

def main():
    # 기준 날짜 설정: 특정 날짜로 실행하려면 아래 변수를 수정하세요.
    # 예: target_date = datetime(2025, 4, 25)
    # target_date = datetime.now()
    target_date = datetime(2026, 4, 9)

    today_str = target_date.strftime("%Y-%m-%d")
    # read_oms_futures_file()
    # read_oms_stock_file()
    # read_trader_file()
    output_path = f'Z:/02.펀드/019. 일간매매내역/recon_result/{today_str}_recon result.xlsx'  # 저장 경로 지정
    merge_and_reconcile(output_path, target_date)

if __name__ == "__main__":
    main()