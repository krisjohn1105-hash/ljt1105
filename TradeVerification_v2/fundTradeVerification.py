import pandas as pd
from tabulate import tabulate
import FinanceDataReader as fdr
from datetime import datetime, timedelta
from pathlib import Path
import os

# Z 드라이브 경로 상수 정의 (규칙 7)
Z_DRIVE = Path("Z:/")

# 검토 필요 대상 행들을 수집하기 위한 전역 리스트 (규칙 6)
REVIEW_LIST = []

def check_and_filter_nan(df, file_label, required_cols):
    """ 데이터프레임 내 필수 칼럼들의 결측치(NaN)를 검사하여
        결측치가 존재하는 행은 '검토 필요' 리스트로 분류하고 제외한 정상 데이터만 반환합니다.
        (규칙 6 준수)
    """
    if df.empty:
        return df
    
    # 필수 칼럼이 없는 경우 있는 것만 체크
    cols_to_check = [c for c in required_cols if c in df.columns]
    nan_rows = df[df[cols_to_check].isna().any(axis=1)]
    if not nan_rows.empty:
        print(f"[경고] {file_label}에서 결측치(NaN)가 발견되었습니다. {len(nan_rows)}행을 '검토 필요'로 분류합니다.")
        for idx, row in nan_rows.iterrows():
            row_dict = row.to_dict()
            row_dict['출처파일'] = file_label
            row_dict['검토사유'] = "필수 항목 누락(NaN)"
            REVIEW_LIST.append(row_dict)
        # 결측치가 없는 정상 행만 필터링
        df = df[df[cols_to_check].notna().all(axis=1)]
    return df

def extract_mapping_from_recent_files():
    """ FinanceDataReader 실패 시 우회하기 위해, Z 드라이브의 최근 전체 매매내역(* 전체.xlsx) 파일들로부터
        종목명-단축코드 매핑 딕셔너리를 추출합니다. (규칙 11 우회 로직)
    """
    mapping = {}
    parent_dir = Z_DRIVE / '02.펀드' / '019. 일간매매내역'
    if not parent_dir.exists():
        return mapping
    
    # 최근 수정된 '* 전체.xlsx' 파일들 최대 5개 탐색
    files = sorted(
        parent_dir.glob('* 전체.xlsx'),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )[:5]
    
    for file in files:
        try:
            df = pd.read_excel(file)
            if '종목명' in df.columns and '단축코드' in df.columns:
                temp_df = df[['종목명', '단축코드']].dropna()
                for _, row in temp_df.iterrows():
                    name = str(row['종목명']).strip()
                    code = str(row['단축코드']).strip().zfill(6) # 규칙 9
                    mapping[name] = code
        except Exception:
            continue
            
    return mapping

def get_stock_code_mapping():
    """ FinanceDataReader를 통해 KRX 및 ETF 종목코드 매핑 정보를 가져옵니다.
        실패 시 Z 드라이브 최근 거래내역 파일에서 추출하는 우회 로직을 작동합니다. (규칙 11)
    """
    print("[작업 시작] 종목코드 매핑 정보 로드 시작")
    mapping = {}
    try:
        krx_df = fdr.StockListing('KRX')
        etf_df = fdr.StockListing('ETF/KR')
        krx_mapping = dict(zip(krx_df['Name'], krx_df['Code']))
        etf_mapping = dict(zip(etf_df['Name'], etf_df['Symbol']))
        mapping = {**krx_mapping, **etf_mapping}
        print("[작업 완료] FinanceDataReader를 통한 종목코드 매핑 완료")
    except Exception as e:
        print(f"[경고] FinanceDataReader 호출 실패: {str(e)}. 우회 로직(최근 매매내역 기반 매핑 추출)을 실행합니다.")
        try:
            mapping = extract_mapping_from_recent_files()
            print(f"[작업 완료] 최근 파일 기반 종목코드 매핑 완료 (추출된 매핑 수: {len(mapping)}개)")
        except Exception as bypass_err:
            print(f"[에러] 종목코드 매핑 우회 로직도 실패하였습니다: {str(bypass_err)}")
    
    return mapping

def get_latest_available_date(start_date):
    """ start_date부터 역순으로 최대 10일까지 탐색하여,
        실제 데이터 파일(주식 파일 및 전체 파일)이 존재하는 가장 최근 영업일을 반환합니다. (규칙 10)
    """
    print("[작업 시작] 영업일 데이터 탐색 시작")
    for i in range(10):
        current_date = start_date - timedelta(days=i)
        date_str = current_date.strftime("%Y-%m-%d")
        korean_date_str = f"{current_date.month}월{current_date.day}일"
        
        # 파일 존재 검사
        stock_file = Z_DRIVE / '02.펀드' / '019. 일간매매내역' / '주식' / f'{date_str}_stock.xlsx'
        trader_file = Z_DRIVE / '02.펀드' / '019. 일간매매내역' / f'{korean_date_str} 전체.xlsx'
        
        if stock_file.exists() and trader_file.exists():
            print(f"[작업 완료] 사용 가능한 영업일 발견: {date_str}")
            return current_date
            
    print(f"[경고] 최근 10일 이내에 사용 가능한 영업일 데이터를 찾지 못했습니다. 입력 기준일({start_date.strftime('%Y-%m-%d')})로 진행합니다.")
    return start_date


def read_oms_futures_file(target_date):
    """ OMS 파생상품 거래 내역 파일을 읽어 가공한 뒤 반환합니다. """
    today = target_date.strftime("%Y-%m-%d")
    print(f"[작업 시작] OMS 파생상품 파일 로드 시작 (날짜: {today})")
    
    file_path = Z_DRIVE / '02.펀드' / '019. 일간매매내역' / '파생' / f'{today}_futures.xlsx'
    
    if not file_path.exists():
        print(f"[경고] OMS 파생상품 파일이 존재하지 않습니다: {file_path}")
        print("[작업 완료] 빈 OMS 파생상품 DataFrame 반환")
        return pd.DataFrame(columns=['펀드명', '매매구분', '단축코드', '종목명', '체결단가', '체결수량', '체결금액'])

    df = pd.read_excel(file_path)

    # 필요한 칼럼만 남기기
    columns_to_keep = ['펀드명', '매매\n구분', '종목코드', '종목명', '체결단가', '체결\n계약수', '약정금액']
    missing_cols = [c for c in columns_to_keep if c not in df.columns]
    if missing_cols:
        raise ValueError(f"OMS 파생상품 파일에 필수 칼럼이 누락되었습니다: {missing_cols}")

    df = df[columns_to_keep]
    df.columns = ['펀드명', '매매구분', '종목코드', '종목명', '체결단가', '체결계약수', '약정금액']

    # 결측치 NaN 검사 및 검토 필요 분류 (규칙 6)
    df = check_and_filter_nan(df, file_path.name, ['펀드명', '매매구분', '종목코드', '체결단가', '체결계약수', '약정금액'])

    # 펀드명 변환
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

    # 종목코드 변환 (앞 3자리와 뒤 5자리만 남기기)
    df['종목코드'] = df['종목코드'].str[3:8]

    # 단축코드 6자리로 수정 (규칙 9)
    df['종목코드'] = df['종목코드'].astype(str).str.zfill(6)

    # 칼럼 이름 변경
    column_rename_mapping = {
        "종목코드": "단축코드",
        "체결계약수": "체결수량",
        "약정금액": "체결금액"
    }
    df.rename(columns=column_rename_mapping, inplace=True)

    # 소수점 2째자리 반올림 (규칙 3)
    df['체결단가'] = df['체결단가'].astype(float).round(2)
    df['체결수량'] = df['체결수량'].astype(float).round(2)
    df['체결금액'] = df['체결금액'].astype(float).round(2)

    print(f"[작업 완료] OMS 파생상품 파일 로드 완료 (데이터 수: {len(df)}건)")
    return df


def read_oms_stock_file(target_date):
    """ OMS 국내주식 거래 내역 파일을 읽어 가공한 뒤 반환합니다. """
    today = target_date.strftime("%Y-%m-%d")
    print(f"[작업 시작] OMS 주식 파일 로드 시작 (날짜: {today})")

    file_path = Z_DRIVE / '02.펀드' / '019. 일간매매내역' / '주식' / f'{today}_stock.xlsx'
    
    if not file_path.exists():
        print(f"[경고] OMS 주식 파일이 존재하지 않습니다: {file_path}")
        print("[작업 완료] 빈 OMS 주식 DataFrame 반환")
        return pd.DataFrame(columns=['펀드명', '매매구분', '단축코드', '종목명', '체결단가', '체결수량', '체결금액'])

    df = pd.read_excel(file_path)

    # 필요한 칼럼만 남기기
    columns_to_keep = ['펀드명', '매매구분', '종목명', '체결가격', '체결수량', '매매금액']
    missing_cols = [c for c in columns_to_keep if c not in df.columns]
    if missing_cols:
        raise ValueError(f"OMS 주식 파일에 필수 칼럼이 누락되었습니다: {missing_cols}")

    df = df[columns_to_keep]

    # 결측치 NaN 검사 및 검토 필요 분류 (규칙 6)
    df = check_and_filter_nan(df, file_path.name, ['펀드명', '매매구분', '종목명', '체결가격', '체결수량', '매매금액'])

    # 펀드명 변환
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

    # FinanceDataReader를 사용해 종목명으로 단축코드 생성
    stock_code_mapping = get_stock_code_mapping()

    # 종목명으로 단축코드 매핑
    df['단축코드'] = df['종목명'].map(stock_code_mapping)
    
    # 매핑되지 않은 종목명 확인
    unmatched = df[df['단축코드'].isna()]['종목명'].unique()
    if unmatched.size > 0:
        print("[경고] 매핑되지 않은 종목명:", unmatched)

    # 단축코드 6자리 보장 (규칙 9)
    df['단축코드'] = df['단축코드'].apply(lambda x: str(x).strip().zfill(6) if pd.notna(x) else x)

    # 단축코드 칼럼을 종목명 앞에 배치
    df = df[['펀드명', '매매구분', '단축코드', '종목명', '체결가격', '체결수량', '매매금액']]

    # 칼럼 이름 변경
    df.rename(columns={
        "체결가격": "체결단가",
        "매매금액": "체결금액"
    }, inplace=True)

    # 소수점 2째자리 반올림 (규칙 3)
    df['체결단가'] = df['체결단가'].astype(float).round(2)
    df['체결수량'] = df['체결수량'].astype(float).round(2)
    df['체결금액'] = df['체결금액'].astype(float).round(2)

    print(f"[작업 완료] OMS 주식 파일 로드 완료 (데이터 수: {len(df)}건)")
    return df


def read_prelude_stock_trade_history(target_date):
    """ Prelude 주식 거래 내역 파일을 읽어 가공한 뒤 반환합니다. """
    today = target_date.strftime("%m%d%y")
    print(f"[작업 시작] Prelude 주식 파일 로드 시작 (날짜: {today})")

    file_dir = Z_DRIVE / '02.펀드' / '003.매매보고서 대사' / 'PRELUDE_RECAP'
    
    preferred_path = file_dir / f"Korea Stocks - {today}.xls"
    fallback_path = file_dir / f"Pre allocation Korea Stocks - {today}.xls"
    
    selected_path = None
    if preferred_path.exists():
        selected_path = preferred_path
    elif fallback_path.exists():
        print(f"[정보] {preferred_path.name} 파일이 없어 {fallback_path.name} 파일로 대체합니다.")
        selected_path = fallback_path
    else:
        # 두 파일 모두 없는 경우, 폴더 내 가장 최근 수정된 파일 검색
        candidates = sorted(
            list(file_dir.glob("Korea Stocks - *.xls")) + list(file_dir.glob("Pre allocation Korea Stocks - *.xls")),
            key=lambda p: p.stat().st_mtime
        )
        if candidates:
            selected_path = candidates[-1]
            print(f"[경고] 당일 Prelude 매매보고서 파일이 존재하지 않아, 가장 최근의 파일인 {selected_path.name}을 사용합니다.")
        else:
            print("[경고] Prelude 매매보고서 파일이 존재하지 않으며, 폴더 내 대체 파일도 없습니다.")
            print("[작업 완료] 빈 Prelude DataFrame 반환")
            return pd.DataFrame(columns=['펀드명', '매매구분', '단축코드', '종목명', '체결단가', '체결수량', '체결금액'])

    df = pd.read_excel(selected_path)

    # 필요한 칼럼만 남기기
    columns_to_keep = [
        'Client Account Name', 'Buy/Sell', 'Ric', 'Stock Description', 'Price (Gross) Listing Ccy', 'Stock Quantity'
    ]
    missing_cols = [c for c in columns_to_keep if c not in df.columns]
    if missing_cols:
        raise ValueError(f"Prelude 파일에 필수 칼럼이 누락되었습니다: {missing_cols}")

    df = df[columns_to_keep]
    df.columns = ['펀드명', '매매구분', '단축코드', '종목명', '체결단가', '체결수량']

    # 결측치 NaN 검사 및 검토 필요 분류 (규칙 6)
    df = check_and_filter_nan(df, selected_path.name, ['펀드명', '매매구분', '단축코드', '체결단가', '체결수량'])

    # 종목코드 앞의 여섯 자리만 남기고 zfill(6) 보장 (규칙 9)
    df['단축코드'] = df['단축코드'].astype(str).str[:6].str.zfill(6)

    # 매매금액 칼럼 추가 (체결가격 * 체결수량)
    df['체결금액'] = df['체결단가'] * df['체결수량']

    # 펀드명 데이터를 'Prelude'로 변경
    df['펀드명'] = 'Prelude'

    # 매매구분 값 변환
    df['매매구분'] = df['매매구분'].replace({'Buy': '매수', 'Sell': '매도', 'Sell Short': '매도', 'Buy Cover': '매수', 'Short sell': '매도', 'Buy cover': '매수'})

    # 소수점 2째자리 반올림 (규칙 3)
    df['체결단가'] = df['체결단가'].astype(float).round(2)
    df['체결수량'] = df['체결수량'].astype(float).round(2)
    df['체결금액'] = df['체결금액'].astype(float).round(2)

    print(f"[작업 완료] Prelude 주식 파일 로드 완료 (데이터 수: {len(df)}건)")
    return df


def read_trader_file(target_date):
    """ Trader 거래내역 전체 파일을 읽어 가공한 뒤 반환합니다. """
    today = f"{target_date.month}월{target_date.day}일"
    print(f"[작업 시작] Trader 전체 파일 로드 시작 (날짜: {today})")

    file_path = Z_DRIVE / '02.펀드' / '019. 일간매매내역' / f'{today} 전체.xlsx'
    
    if not file_path.exists():
        print(f"[경고] Trader 전체 파일이 존재하지 않습니다: {file_path}")
        print("[작업 완료] 빈 Trader DataFrame 반환")
        return pd.DataFrame(columns=['펀드명', '매매구분', '단축코드', '종목명', '체결단가', '체결수량', '체결금액'])

    df = pd.read_excel(file_path)

    # 필요한 칼럼만 남기기
    columns_to_keep = ['펀드', '매매구분', '단축코드', '종목명', '체결단가', '체결수량', '체결금액']
    missing_cols = [c for c in columns_to_keep if c not in df.columns]
    if missing_cols:
        raise ValueError(f"Trader 전체 파일에 필수 칼럼이 누락되었습니다: {missing_cols}")

    df = df[columns_to_keep]

    # 결측치 NaN 검사 및 검토 필요 분류 (규칙 6)
    df = check_and_filter_nan(df, file_path.name, ['펀드', '매매구분', '단축코드', '체결단가', '체결수량', '체결금액'])

    # 단축코드 6자리로 수정 (규칙 9)
    df['단축코드'] = df['단축코드'].astype(str).str.zfill(6)

    # 칼럼 이름 변경
    df.rename(columns={
        "펀드": "펀드명"
    }, inplace=True)

    # 매매구분 변환
    direction_mapping = {
        "Buy": "매수",
        "Sell": "매도",
        "Sell Short": "매도",
        "Buy Cover": "매수",
        "Short sell": "매도", 
        'Buy cover': '매수'
    }
    df['매매구분'] = df['매매구분'].replace(direction_mapping)

    # 소수점 2째자리 반올림 (규칙 3)
    df['체결단가'] = df['체결단가'].astype(float).round(2)
    df['체결수량'] = df['체결수량'].astype(float).round(2)
    df['체결금액'] = df['체결금액'].astype(float).round(2)

    print(f"[작업 완료] Trader 전체 파일 로드 완료 (데이터 수: {len(df)}건)")
    return df


def aggregate_by_key(df):
    """ 펀드명, 매매구분, 단축코드를 기준으로 데이터를 그룹화하고 합계 및 가중평균단가를 계산합니다. """
    print("[작업 시작] 키 기준 데이터 그룹화 및 합산 진행")
    if df.empty:
        print("[정보] 데이터프레임이 비어 있어 그룹화를 진행하지 않고 반환합니다.")
        return pd.DataFrame(columns=['펀드명', '매매구분', '단축코드', '종목명', '체결수량', '체결금액', '체결단가'])
        
    aggregated_df = df.groupby(['펀드명', '매매구분', '단축코드'], as_index=False).agg({
        '종목명': 'first',  # 그룹에서 첫 번째 종목명을 사용
        '체결수량': 'sum',
        '체결금액': 'sum',
        # 소수점 3째자리에서 반올림하여 2째자리로 표기 (규칙 3)
        '체결단가': lambda x: round((x * df.loc[x.index, '체결수량']).sum() / df.loc[x.index, '체결수량'].sum(), 2)
    })
    
    # 각 합산값도 소수점 2자리 반올림 (규칙 3)
    aggregated_df['체결수량'] = aggregated_df['체결수량'].round(2)
    aggregated_df['체결금액'] = aggregated_df['체결금액'].round(2)
    
    print("[작업 완료] 키 기준 데이터 그룹화 완료")
    return aggregated_df


def merge_and_reconcile(output_path, target_date):
    """ OMS와 Trader 거래 내역을 병합하고 대사 작업을 진행합니다. """
    print(f"[작업 시작] 대사(Reconciliation) 프로세스 시작 (기준일: {target_date.strftime('%Y-%m-%d')})")
    
    # OMS Futures와 Stock 데이터를 읽기
    oms_futures_df = read_oms_futures_file(target_date)
    oms_stock_df = read_oms_stock_file(target_date)
    prelude_stock_df = read_prelude_stock_trade_history(target_date)

    # 1. OMS 주식, 파생, Prelude 데이터를 하나로 병합
    trade_history_combined = pd.concat([oms_stock_df, oms_futures_df, prelude_stock_df], ignore_index=True)
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

    # 체결단가, 체결수량, 체결금액의 차이 계산 (소수점 2째자리로 반올림 - 규칙 3)
    reconciliation_df['체결단가_차이'] = (reconciliation_df['체결단가_oms'].fillna(0) - reconciliation_df['체결단가_trader'].fillna(0)).round(2)
    reconciliation_df['체결수량_차이'] = (reconciliation_df['체결수량_oms'].fillna(0) - reconciliation_df['체결수량_trader'].fillna(0)).round(2)
    reconciliation_df['체결금액_차이'] = (reconciliation_df['체결금액_oms'].fillna(0) - reconciliation_df['체결금액_trader'].fillna(0)).round(2)

    # 결과 출력
    print("대사 결과:")
    print(tabulate(reconciliation_df[['펀드명', '매매구분', '단축코드', '체결단가_차이', '체결수량_차이', '체결금액_차이']], headers = 'keys', tablefmt = 'pretty'))

    # 대사 결과 차이 분석 및 로깅 (규칙 5)
    # 단가 차이, 수량 차이, 금액 차이가 존재하거나 한쪽 데이터가 누락된 경우
    diff_condition = (
        (reconciliation_df['체결단가_차이'].abs() > 0.001) |
        (reconciliation_df['체결수량_차이'].abs() > 0.001) |
        (reconciliation_df['체결금액_차이'].abs() > 0.001) |
        (reconciliation_df['_merge'] != 'both')
    )
    diff_df = reconciliation_df[diff_condition].copy()
    
    if not diff_df.empty:
        print("\n[!] 대사 결과 차이가 발생하였습니다. 임의 수정하지 않고 차이 내역을 로그에 기록합니다.")
        
        # 불일치 상세 내역을 표로 만들기 위한 리스트 구축
        table_rows = []
        for idx, row in diff_df.iterrows():
            # 종목명 가져오기 (Trader 우선, 없으면 OMS)
            stock_name = row['종목명_trader'] if pd.notna(row['종목명_trader']) else (row['종목명_oms'] if pd.notna(row['종목명_oms']) else '')
            
            # 수량 표시 포맷
            qty_oms = f"{row['체결수량_oms']:.2f}" if pd.notna(row['체결수량_oms']) else "N/A"
            qty_trader = f"{row['체결수량_trader']:.2f}" if pd.notna(row['체결수량_trader']) else "N/A"
            qty_diff = f"{row['체결수량_차이']:.2f}"
            
            # 단가 표시 포맷
            price_oms = f"{row['체결단가_oms']:.2f}" if pd.notna(row['체결단가_oms']) else "N/A"
            price_trader = f"{row['체결단가_trader']:.2f}" if pd.notna(row['체결단가_trader']) else "N/A"
            price_diff = f"{row['체결단가_차이']:.2f}"

            # 금액 표시 포맷
            amt_oms = f"{row['체결금액_oms']:.2f}" if pd.notna(row['체결금액_oms']) else "N/A"
            amt_trader = f"{row['체결금액_trader']:.2f}" if pd.notna(row['체결금액_trader']) else "N/A"
            amt_diff = f"{row['체결금액_차이']:.2f}"
            
            # 불일치 사유
            if row['_merge'] == 'left_only':
                reason = "OMS에만 존재 (Trader 누락)"
            elif row['_merge'] == 'right_only':
                reason = "Trader에만 존재 (OMS 누락)"
            else:
                diffs = []
                if abs(row['체결단가_차이']) > 0.001:
                    diffs.append("단가 불일치")
                if abs(row['체결수량_차이']) > 0.001:
                    diffs.append("수량 불일치")
                if abs(row['체결금액_차이']) > 0.001:
                    diffs.append("금액 불일치")
                reason = ", ".join(diffs)
                
            table_rows.append({
                "펀드명": row['펀드명'],
                "구분": row['매매구분'],
                "단축코드": row['단축코드'],
                "종목명": stock_name,
                "OMS수량": qty_oms,
                "Trader수량": qty_trader,
                "수량차이": qty_diff,
                "OMS단가": price_oms,
                "Trader단가": price_trader,
                "단가차이": price_diff,
                "OMS금액": amt_oms,
                "Trader금액": amt_trader,
                "금액차이": amt_diff,
                "사유": reason
            })
            
        summary_df = pd.DataFrame(table_rows)
        print("\n======== [대사 불일치 요약 표] ========")
        print(tabulate(summary_df, headers='keys', tablefmt='pretty', showindex=False))
        print("========================================")
    else:
        print("\n[+] 대사 결과: 모든 내역이 상호 일치합니다.")

    # 결측치(NaN) 검토 필요 리스트가 있을 경우 출력 (규칙 6)
    if REVIEW_LIST:
        print("\n======== [검토 필요 리스트 (결측치 존재)] ========")
        review_df = pd.DataFrame(REVIEW_LIST)
        cols_to_show = ['출처파일', '검토사유'] + [c for c in review_df.columns if c not in ['출처파일', '검토사유']]
        review_df = review_df[cols_to_show]
        print(tabulate(review_df, headers='keys', tablefmt='pretty', showindex=False))
        print("====================================================")

    # output_path 디렉토리가 없으면 생성 (pathlib 활용)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    # 결과를 파일로 저장 (원본 파일은 수정하거나 삭제하지 않음 - 규칙 8)
    reconciliation_df.to_excel(output_path, index=False)
    print(f"[작업 완료] 대사 결과 파일 저장 완료: {output_path}")

    return reconciliation_df


def main():
    print(f"[{datetime.now()}] 주식/파생 대사 검증 프로그램 시작")
    
    # 1. 타겟 날짜 설정: 기본값은 현재 날짜
    target_date = datetime.now()
    
    # 2. 영업일 기준 날짜 조정 (규칙 10)
    adjusted_date = get_latest_available_date(target_date)
    
    today_str = adjusted_date.strftime("%Y-%m-%d")
    output_path = Z_DRIVE / '02.펀드' / '019. 일간매매내역' / 'recon_result' / f'{today_str}_recon result.xlsx'
    
    # 대사 진행
    merge_and_reconcile(output_path, adjusted_date)
    
    print(f"[{datetime.now()}] 주식/파생 대사 검증 프로그램 종료")


if __name__ == "__main__":
    main()