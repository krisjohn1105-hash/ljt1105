import sys
import json
import pandas as pd
from datetime import datetime
from pathlib import Path

# 모듈 경로 추가 (어느 위치에서든 상대 경로 로드가 가능하게 함)
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from api import fetch_eod_report

# ENF 로그인 정보
ENF_USERNAME = 'ops@dunamiscap.com'
ENF_PASSWORD = 'Kingduna10!'

# 결측치(NaN)가 있는 행을 관리하기 위한 '검토 필요' 리스트 (규칙 6 준수)
REVIEW_LIST = []


def check_and_filter_nan(df, file_label, required_cols):
    """
    데이터프레임 내 필수 칼럼들의 결측치(NaN)를 검사하여
    결측치가 존재하는 행은 '검토 필요' 리스트로 분류하고 제외한 정상 데이터만 반환합니다.
    (규칙 6 준수)
    """
    if df.empty:
        return df
    
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


def load_pre_trade_excel(directory, prefix, date_str):
    """
    지정된 날짜의 pre-trade 엑셀 파일을 로드합니다.
    파일이 존재하지 않을 경우 가장 최근 영업일의 파일을 백업으로 활용합니다. (규칙 10 준수)
    """
    preferred = directory / f'{prefix} - {date_str}.xls'
    if preferred.exists():
        return pd.read_excel(preferred)

    candidates = sorted(
        directory.glob(f'{prefix} - *.xls'),
        key=lambda p: p.stat().st_mtime,
    )
    if candidates:
        print(f"[WARN] {preferred.name} 파일이 존재하지 않아 최신 수정 파일인 {candidates[-1].name}을 로드합니다.")
        return pd.read_excel(candidates[-1])

    print(f"[WARN] {directory} 내에 '{prefix} - *.xls' 패턴의 파일이 존재하지 않습니다. 빈 데이터프레임을 반환합니다.")
    return pd.DataFrame()


def load_data(data_dir, eod_trade_path, today_str):
    """
    대사에 필요한 pre-trade 엑셀 파일들과 EOD Trade JSON 파일을 로드합니다.
    (규칙 2 준수 - 진행 로그 출력)
    """
    print(f"[작업 시작] 데이터 로드 시작 (기준일: {today_str})")
    
    # 1. ENF Trade JSON 파일 로드
    if not eod_trade_path.exists():
        raise FileNotFoundError(
            f"{eod_trade_path} 파일이 존재하지 않습니다. api.py를 먼저 실행해 주세요."
        )
    with open(eod_trade_path, "r") as f:
        data = json.load(f)
    enf_df = pd.json_normalize(data['rows'])
    
    # 2. Pre allocation Korea Stocks 엑셀 로드
    ms_df = load_pre_trade_excel(data_dir, 'Pre allocation Korea Stocks', today_str)
    
    # 3. Pre allocation Korea Futures 엑셀 로드
    ms_futures = load_pre_trade_excel(data_dir, 'Pre allocation Korea Futures', today_str)
    
    print("[작업 완료] 데이터 로드 완료")
    return enf_df, ms_df, ms_futures


def preprocess_data(enf_df, ms_df, ms_futures, today_str):
    """
    ENF 데이터와 MS 데이터를 대사 가능한 형태로 전처리하고 날짜 필터링을 수행합니다.
    (규칙 2, 규칙 3, 규칙 6, 규칙 9 준수)
    """
    print("[작업 시작] 데이터 전처리 시작")
    
    # 1. MS 데이터에서 실제 대사 대상 날짜(Trade Date) 추출 및 ENF 날짜 필터링 (규칙 10 관련)
    target_date_str = None
    if not ms_df.empty and 'Trade Date' in ms_df.columns:
        try:
            # 첫 번째 행의 Trade Date를 파싱하여 YYYY-MM-DD 형식으로 변환
            target_date_str = pd.to_datetime(ms_df['Trade Date']).dt.strftime('%Y-%m-%d').iloc[0]
        except Exception as e:
            print(f"[경고] MS Stocks의 Trade Date 파싱 실패: {e}")
            
    if not target_date_str:
        try:
            target_date_str = datetime.strptime(today_str, '%m%d%y').strftime('%Y-%m-%d')
        except Exception:
            target_date_str = datetime.today().strftime('%Y-%m-%d')
            
    print(f"[정보] 대사 대상 일자 설정: {target_date_str}")
    
    # ENF 데이터의 날짜 필터링 (가장 핵심적인 불일치 수정 포인트)
    if 'tradeDate.value' in enf_df.columns:
        enf_df['tradeDate_ymd'] = pd.to_datetime(enf_df['tradeDate.value']).dt.strftime('%Y-%m-%d')
        enf_df = enf_df[enf_df['tradeDate_ymd'] == target_date_str].copy()
        enf_df = enf_df.drop(columns=['tradeDate_ymd'])
    
    # 2. ENF 필요한 컬럼만 추출 및 정규화
    required_enf_cols = [
        'bBYellow.value', 'rIC.value', 'tradeTransactionType.value', 'notionalQuantity.value', 
        'tradePrice.value', 'tradePrice.currency', 'instrumentDescription.value', 
        'tradePrice.quoteType', 'tradePrice.quotationFormat', 'instrumentSubtype.value', 
        'tradeDate.value', 'settleDate.value', 'tradeCurrency.value', 'settleCurrency.value', 
        'actualTradeToSettleFXRate.value', 'grossTradeCommisions.value', 
        'grossTradeCommisions.currency', 'grossFees.value', 'grossTaxes.value', 
        'netTradePrice.value', 'netTradePrice.currency', 'tradingNotionalNetProceeds.value'
    ]
    
    # 필수 칼럼이 없을 경우 경고 메시지 출력
    missing_enf_cols = [c for c in required_enf_cols if c not in enf_df.columns]
    if missing_enf_cols:
         print(f"[경고] ENF 데이터에 일부 필수 칼럼이 누락되었습니다: {missing_enf_cols}")
    
    # ENF 결측치 검사 및 '검토 필요' 리스트 분류 (규칙 6 준수)
    enf_df = check_and_filter_nan(
        enf_df, 
        'EOD_Trade.json', 
        ['rIC.value', 'tradeTransactionType.value', 'notionalQuantity.value', 'tradingNotionalNetProceeds.value']
    )
    
    # 칼럼명 변경: '.value' 접미사를 지우고, 나머지 속성(currency 등)은 고유 이름 유지
    enf_df.columns = [col.replace('.value', '') for col in enf_df.columns]
    
    # 3. MS 데이터 전처리 및 컬럼 매핑
    ms_dict = {
        'Buy/Sell': 'tradeTransactionType', 
        'Stock Quantity': 'notionalQuantity',
        'Net Notional in Swap Ccy': 'tradingNotionalNetProceeds', 
        'Ric': 'rIC', 
        'Swap Ccy': 'settleCurrency',
        'Futures Contracts Quantity': 'notionalQuantity'
    }
    
    # 결측치 검사 및 '검토 필요' 리스트 분류 (규칙 6 준수)
    if not ms_df.empty:
        ms_df = check_and_filter_nan(
            ms_df, 
            'Pre allocation Korea Stocks', 
            ['Ric', 'Buy/Sell', 'Stock Quantity', 'Net Notional in Swap Ccy']
        )
        ms_df = ms_df.rename(columns=ms_dict)
        
    if not ms_futures.empty:
        ms_futures = check_and_filter_nan(
            ms_futures, 
            'Pre allocation Korea Futures', 
            ['Ric', 'Buy/Sell', 'Futures Contracts Quantity', 'Net Notional in Swap Ccy']
        )
        ms_futures = ms_futures.rename(columns=ms_dict)
        
    # 두 데이터프레임 병합
    ms_combined = pd.concat([ms_df, ms_futures], ignore_index=True)
    
    # 4. 종목 코드(rIC/Ric)의 6자리 zfill(6) 처리 및 정규화 (규칙 9 준수)
    def normalize_ric(ric_str):
        if pd.isna(ric_str):
            return ""
        ric_str = str(ric_str).strip()
        if '.' in ric_str:
            parts = ric_str.split('.')
            code = parts[0]
            suffix = parts[1]
            if code.isdigit():
                return f"{code.zfill(6)}.{suffix}"
        else:
            if ric_str.isdigit():
                return ric_str.zfill(6)
        return ric_str

    if not enf_df.empty and 'rIC' in enf_df.columns:
        enf_df['rIC'] = enf_df['rIC'].apply(normalize_ric)
    if not ms_combined.empty and 'rIC' in ms_combined.columns:
        ms_combined['rIC'] = ms_combined['rIC'].apply(normalize_ric)
        
    # 5. 비교 대상 키 값들 형식 정규화 (strip 및 문자열 변환)
    key_enf = ['rIC', 'tradeTransactionType', 'settleCurrency']
    for col in key_enf:
        if col in enf_df.columns:
            enf_df[col] = enf_df[col].astype(str).str.strip()
        if col in ms_combined.columns:
            ms_combined[col] = ms_combined[col].astype(str).str.strip()
            
    # 6. 절대값 처리 (수량 및 금액)
    if not enf_df.empty:
        enf_df['notionalQuantity'] = pd.to_numeric(enf_df['notionalQuantity'], errors='coerce').abs()
        enf_df['tradingNotionalNetProceeds'] = pd.to_numeric(enf_df['tradingNotionalNetProceeds'], errors='coerce').abs()
    if not ms_combined.empty:
        ms_combined['notionalQuantity'] = pd.to_numeric(ms_combined['notionalQuantity'], errors='coerce').abs()
        ms_combined['tradingNotionalNetProceeds'] = pd.to_numeric(ms_combined['tradingNotionalNetProceeds'], errors='coerce').abs()
        
    # 7. 거래 유형 변환 (ENF의 Buy to Cover -> Buy, Sell Short -> Sell)
    if not enf_df.empty and 'tradeTransactionType' in enf_df.columns:
        enf_df['tradeTransactionType'] = enf_df['tradeTransactionType'].replace({'Buy to Cover': 'Buy', 'Sell Short': 'Sell'})
        
    print("[작업 완료] 데이터 전처리 완료")
    return enf_df, ms_combined


def reconcile_trades(enf_df, ms_df):
    """
    ENF와 MS 데이터를 대조하여 누락된 거래와 차이가 발생하는 거래를 검출합니다.
    (규칙 3, 규칙 5 준수)
    """
    print("[작업 시작] 대사(Reconciliation) 수행 시작")
    
    key_cols = ['rIC', 'tradeTransactionType', 'settleCurrency']
    agg_map = {
        'notionalQuantity': 'sum',
        'tradingNotionalNetProceeds': 'sum'
    }
    
    # groupby로 합산 (동일 종목의 중복 체결 등 합산 처리)
    enf_grouped = enf_df.groupby(key_cols, as_index=False).agg(agg_map)
    ms_grouped = ms_df.groupby(key_cols, as_index=False).agg(agg_map)
    
    # 1. 복합 키 생성 및 비교
    enf_grouped["_key"] = list(zip(*[enf_grouped[col] for col in key_cols]))
    ms_grouped["_key"] = list(zip(*[ms_grouped[col] for col in key_cols]))
    
    keys_enf = set(enf_grouped["_key"])
    keys_ms = set(ms_grouped["_key"])
    
    # 2. 누락 데이터 검출 (Missing report)
    missing_in_ms = enf_grouped[enf_grouped["_key"].isin(keys_enf - keys_ms)][key_cols].copy()
    missing_in_ms["issue"] = "In enf_df, missing in ms_df"
    
    missing_in_enf = ms_grouped[ms_grouped["_key"].isin(keys_ms - keys_enf)][key_cols].copy()
    missing_in_enf["issue"] = "In ms_df, missing in enf_df"
    
    missing_report = pd.concat([missing_in_ms, missing_in_enf], ignore_index=True)
    
    # 3. 공통 키에 대한 값 차이 비교 (Difference report)
    enf_idx = enf_grouped.set_index(key_cols)
    ms_idx = ms_grouped.set_index(key_cols)
    common_keys = keys_enf & keys_ms
    
    diff_rows = []
    compare_cols = ['notionalQuantity', 'tradingNotionalNetProceeds']
    
    for key in common_keys:
        for col in compare_cols:
            val_1 = enf_idx.loc[key, col] if col in enf_idx.columns else 0.0
            val_2 = ms_idx.loc[key, col] if col in ms_idx.columns else 0.0
            
            # 소수점 3째 자리에서 반올림하여 2째자리까지 표기 규칙 적용 (규칙 3 준수)
            try:
                v1_rounded = round(float(val_1), 2)
                v2_rounded = round(float(val_2), 2)
                diff = round(v1_rounded - v2_rounded, 2)
            except (TypeError, ValueError):
                v1_rounded = val_1
                v2_rounded = val_2
                diff = "N/A"
                
            if v1_rounded != v2_rounded:
                diff_rows.append({
                    "key": key,
                    "variable": col,
                    "val_enf": v1_rounded,
                    "val_ms": v2_rounded,
                    "Difference": diff
                })
                
    diff_report = pd.DataFrame(
        diff_rows,
        columns=["key", "variable", "val_enf", "val_ms", "Difference"]
    )
    
    # 오차 범위 적용 (절대값 1.0 이상인 경우만 차이로 분류)
    tolerance = 1.0
    diff_report['Difference'] = pd.to_numeric(diff_report['Difference'], errors='coerce')
    diff_report = diff_report[(diff_report["Difference"].isna()) | (diff_report["Difference"].abs() >= tolerance)]
    
    print("[작업 완료] 대사 수행 완료")
    return missing_report, diff_report


def main():
    """
    대사 프로그램의 메인 제어 흐름입니다.
    """
    print("="*60)
    print("대사(Reconciliation) 프로그램 실행을 시작합니다.")
    print("="*60)
    
    # 1. EOD 리포트 실시간 페치 시도 (규칙 11 - 예외 처리 및 우회 로직 준수)
    try:
        fetch_eod_report(ENF_USERNAME, ENF_PASSWORD)
    except Exception as exc:
        print(f"[WARN] EOD 리포트 페치 실패 ({exc}). 기존 EOD_Trade.json 파일을 사용합니다.")
        
    # 2. 경로 설정 및 파일 로드
    DATA_DIR = BASE_DIR.parent / 'data' / 'input' / 'pre-trade'
    EOD_TRADE_PATH = BASE_DIR / "EOD_Trade.json"
    today_str = datetime.today().strftime('%m%d%y')
    
    try:
        enf_df, ms_df, ms_futures = load_data(DATA_DIR, EOD_TRADE_PATH, today_str)
    except Exception as e:
        print(f"[에러] 데이터 로드 실패: {e}")
        sys.exit(1)
        
    # 3. 데이터 전처리 및 날짜 필터링
    enf_df, ms_df = preprocess_data(enf_df, ms_df, ms_futures, today_str)
    
    # 4. 대사 수행
    missing_report, diff_report = reconcile_trades(enf_df, ms_df)
    
    # 5. 결과 리포트 출력 및 로그 기록 (규칙 5 준수)
    print("\n" + "="*20 + " 대사 결과 리포트 " + "="*20)
    
    if missing_report.empty:
        print("\n[누락 내역] 누락된 거래가 없습니다.")
    else:
        print("\n[누락 내역] 다음과 같은 거래 누락이 발견되었습니다:")
        print(missing_report.to_string(index=False))
        
    if diff_report.empty:
        print("\n[차이 내역] 수량 및 금액의 불일치가 없습니다.")
    else:
        print("\n[차이 내역] [경고] 수량/금액 불일치가 발생하였습니다:")
        print(diff_report.to_string(index=False))
        
    # 검토 필요 리스트 출력 (규칙 6 준수)
    if REVIEW_LIST:
        print("\n[검토 필요] 일부 소스 파일에서 NaN 결측치가 발견되어 해당 행들을 분류했습니다:")
        for idx, item in enumerate(REVIEW_LIST):
            print(f"  {idx+1}. 파일: {item.get('출처파일')}, 사유: {item.get('검토사유')}")
            # 결측 정보 디테일 출력
            clean_item = {k: v for k, v in item.items() if k not in ['출처파일', '검토사유']}
            print(f"     내용: {clean_item}")
            
    print("="*60)
    print("대사 프로그램 실행이 완료되었습니다.")
    print("="*60)


if __name__ == "__main__":
    main()
