import sys
import re
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

# GS PBS 거래를 ENF 전체 블로터에서 걸러내기 위한 counterPartyName 식별 키워드
GS_COUNTERPARTY_KEYWORDS = ['goldman', 'gsil']

SWAP_DIR = BASE_DIR.parent / 'data' / 'input' / 'QRT' / 'Swap'
CASH_DIR = BASE_DIR.parent / 'data' / 'input' / 'QRT' / 'Cash'
EOD_TRADE_PATH = BASE_DIR / 'EOD_Trade.json'

# 결측치(NaN)가 있는 행을 관리하기 위한 '검토 필요' 리스트
REVIEW_LIST = []

# B/S, SIDE 코드 -> Buy/Sell 변환 (QRT/ENF와 동일한 컨벤션 사용)
SIDE_MAP = {'B': 'Buy', 'S': 'Sell', 'SS': 'Sell', 'BC': 'Buy'}


def normalize_ric(ric_str):
    """종목 코드(RIC)의 숫자 부분을 6자리로 zfill 하여 정규화합니다."""
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


def check_and_filter_nan(df, file_label, required_cols):
    """필수 칼럼에 결측치가 있는 행을 '검토 필요' 리스트로 분류하고 제외합니다."""
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
        df = df[df[cols_to_check].notna().all(axis=1)]
    return df


def extract_trailing_date_token(stem):
    """
    파일명(확장자 제외)에 포함된 마지막 연속 숫자 구간의 끝 6자리를 추출합니다.
    브로커별로 뒤에 ' - Excel' 등의 접미사가 붙어도 날짜(YYMMDD) 뒤의 텍스트일 뿐이므로,
    '파일명 내 마지막 숫자 구간'을 찾아 그 끝 6자리를 취하면 접미사 유무와 무관하게 동작합니다.
    예) 'GSETASIARECAP20260706' -> 마지막 숫자 구간 '20260706' -> 끝 6자리 '260706'
        'KR - Execution Report from GS_260706 - Excel' -> 숫자 구간 '260706' -> '260706'
    """
    digit_runs = re.findall(r'\d+', stem)
    if not digit_runs:
        return None
    last_run = digit_runs[-1]
    if len(last_run) < 6:
        return None
    return last_run[-6:]


def find_todays_file(directory, today_dt):
    """
    directory 내 엑셀 파일들 중 파일명에 오늘 날짜(YYMMDD)가 포함된 파일을 찾습니다.
    없으면 가장 최근에 수정된 파일을 백업으로 사용합니다.
    """
    target = today_dt.strftime('%y%m%d')
    candidates = sorted(directory.glob('*.xls*'))

    for path in candidates:
        if extract_trailing_date_token(path.stem) == target:
            return path

    if candidates:
        latest = max(candidates, key=lambda p: p.stat().st_mtime)
        print(f"[WARN] {directory}에서 오늘({target}) 날짜의 파일을 찾지 못해 최신 수정 파일인 {latest.name}을 로드합니다.")
        return latest

    return None


def load_swap_file(today_dt):
    """QRT/Swap 폴더에서 오늘 날짜(YYMMDD)의 GS 스왑 리캡 엑셀을 로드합니다."""
    if not SWAP_DIR.exists():
        print(f"[WARN] {SWAP_DIR} 폴더가 존재하지 않습니다.")
        return pd.DataFrame()

    swap_path = find_todays_file(SWAP_DIR, today_dt)
    if swap_path is None:
        print(f"[WARN] {SWAP_DIR} 내에 스왑 파일이 없습니다. 빈 데이터프레임을 반환합니다.")
        return pd.DataFrame()

    print(f"[정보] 스왑 파일 로드: {swap_path.name}")
    return pd.read_excel(swap_path)


def parse_cash_execution_report(path):
    """
    GS Cash Execution Report(.xls)는 Buy/Sell 별로
    'FUND RIC BBERG ISIN NAME SIDE QUANTITY ... PRICE CCY DATE DATE' 형태의
    헤더-데이터-Totals 블록이 반복되는 포맷입니다.
    'RIC' 헤더 셀을 기준으로 열 위치를 잡고, 그 아래 데이터 행을 블록별로 추출합니다.
    """
    raw = pd.read_excel(path, header=None)
    n_rows, n_cols = raw.shape
    records = []

    for i in range(n_rows):
        row_str = raw.iloc[i].astype(str).str.strip()
        ric_cols = row_str[row_str == 'RIC'].index.tolist()

        for ric_col in ric_cols:
            qty_col = ric_col + 5
            # 'RIC'에서 5칸 뒤에 'QUANTITY'가 있는 행만 실제 헤더 행으로 인정 (오탐 방지)
            if qty_col >= n_cols or str(raw.iat[i, qty_col]).strip() != 'QUANTITY':
                continue

            side_col = ric_col + 4
            net_principal_col = ric_col + 10
            settle_ccy_col = ric_col + 12

            j = i + 1
            while j < n_rows:
                ric_val = raw.iat[j, ric_col]
                if pd.isna(ric_val) or str(ric_val).strip() == '':
                    break
                records.append({
                    'RIC': ric_val,
                    'SIDE': raw.iat[j, side_col],
                    'QUANTITY': raw.iat[j, qty_col],
                    'NET_PRINCIPAL': raw.iat[j, net_principal_col],
                    'SETTLE_CCY': raw.iat[j, settle_ccy_col],
                })
                j += 1

    return pd.DataFrame(records)


def load_cash_file(today_dt):
    """QRT/Cash 폴더에서 오늘 날짜(YYMMDD)의 GS Execution Report를 로드합니다."""
    if not CASH_DIR.exists():
        print(f"[WARN] {CASH_DIR} 폴더가 존재하지 않습니다.")
        return pd.DataFrame()

    cash_path = find_todays_file(CASH_DIR, today_dt)
    if cash_path is None:
        print(f"[WARN] {CASH_DIR} 내에 Cash 파일이 없습니다. 빈 데이터프레임을 반환합니다.")
        return pd.DataFrame()

    print(f"[정보] Cash 파일 로드: {cash_path.name}")
    return parse_cash_execution_report(cash_path)


def load_data(today_dt):
    """대사에 필요한 ENF EOD Trade JSON과 GS PBS(Swap/Cash) 파일들을 로드합니다."""
    print(f"[작업 시작] 데이터 로드 시작 (기준일: {today_dt.strftime('%Y-%m-%d')})")

    if not EOD_TRADE_PATH.exists():
        raise FileNotFoundError(
            f"{EOD_TRADE_PATH} 파일이 존재하지 않습니다. api.py를 먼저 실행해 주세요."
        )
    with open(EOD_TRADE_PATH, "r") as f:
        data = json.load(f)
    enf_df = pd.json_normalize(data['rows'])

    swap_df = load_swap_file(today_dt)
    cash_df = load_cash_file(today_dt)

    print("[작업 완료] 데이터 로드 완료")
    return enf_df, swap_df, cash_df


def preprocess_enf(enf_df, today_dt):
    """ENF 데이터에서 GS PBS 거래만 필터링하고 대사 가능한 형태로 전처리합니다."""
    if enf_df.empty:
        return enf_df

    # 1. GS PBS 거래만 필터링 (counterPartyName에 'Goldman' 또는 'GSIL' 포함)
    if 'counterPartyName.value' in enf_df.columns:
        cp_lower = enf_df['counterPartyName.value'].astype(str).str.lower()
        gs_mask = cp_lower.apply(lambda v: any(k in v for k in GS_COUNTERPARTY_KEYWORDS))
        enf_df = enf_df[gs_mask].copy()
    else:
        print("[경고] counterPartyName.value 칼럼이 없어 GS 거래 필터링을 수행하지 못했습니다.")

    # 2. 거래일(Trade Date) 필터링
    target_date_str = today_dt.strftime('%Y-%m-%d')
    if 'tradeDate.value' in enf_df.columns:
        enf_df['tradeDate_ymd'] = pd.to_datetime(enf_df['tradeDate.value']).dt.strftime('%Y-%m-%d')
        enf_df = enf_df[enf_df['tradeDate_ymd'] == target_date_str].copy()
        enf_df = enf_df.drop(columns=['tradeDate_ymd'])

    # 3. 대사에 필요한 컬럼만 추출
    required_cols = ['rIC.value', 'tradeTransactionType.value', 'settleCurrency.value',
                      'notionalQuantity.value', 'tradingNotionalNetProceeds.value']
    missing_cols = [c for c in required_cols if c not in enf_df.columns]
    if missing_cols:
        print(f"[경고] ENF 데이터에 일부 필수 칼럼이 누락되었습니다: {missing_cols}")

    enf_df = check_and_filter_nan(
        enf_df,
        'EOD_Trade.json (GS PBS)',
        ['rIC.value', 'tradeTransactionType.value', 'notionalQuantity.value', 'tradingNotionalNetProceeds.value']
    )

    enf_df.columns = [col.replace('.value', '') for col in enf_df.columns]

    if not enf_df.empty:
        enf_df['rIC'] = enf_df['rIC'].apply(normalize_ric)
        enf_df['tradeTransactionType'] = enf_df['tradeTransactionType'].replace(
            {'Buy to Cover': 'Buy', 'Sell Short': 'Sell'}
        )
        for col in ['rIC', 'tradeTransactionType', 'settleCurrency']:
            enf_df[col] = enf_df[col].astype(str).str.strip()
        enf_df['notionalQuantity'] = pd.to_numeric(enf_df['notionalQuantity'], errors='coerce').abs()
        enf_df['tradingNotionalNetProceeds'] = pd.to_numeric(enf_df['tradingNotionalNetProceeds'], errors='coerce').abs()

    return enf_df


def preprocess_swap(swap_df):
    """GS 스왑 리캡 데이터를 대사 가능한 공통 스키마로 정규화합니다."""
    if swap_df.empty:
        return pd.DataFrame(columns=['rIC', 'tradeTransactionType', 'settleCurrency',
                                      'notionalQuantity', 'tradingNotionalNetProceeds'])

    swap_df = check_and_filter_nan(swap_df, 'GS Swap Recap', ['RIC', 'B/S', 'QTY', 'Net Px'])
    if swap_df.empty:
        return pd.DataFrame(columns=['rIC', 'tradeTransactionType', 'settleCurrency',
                                      'notionalQuantity', 'tradingNotionalNetProceeds'])

    out = pd.DataFrame()
    out['rIC'] = swap_df['RIC'].apply(normalize_ric)
    out['tradeTransactionType'] = swap_df['B/S'].astype(str).str.strip().replace(SIDE_MAP)
    # 스왑 거래는 Net Px(현지 통화 가격을 FX Rate로 나눈 USD 환산 가격)를 사용하여 USD로 결제됩니다.
    out['settleCurrency'] = 'USD'
    out['notionalQuantity'] = pd.to_numeric(swap_df['QTY'], errors='coerce').abs()
    net_px = pd.to_numeric(swap_df['Net Px'], errors='coerce')
    out['tradingNotionalNetProceeds'] = (out['notionalQuantity'] * net_px).abs()
    out['rIC'] = out['rIC'].astype(str).str.strip()
    out['tradeTransactionType'] = out['tradeTransactionType'].astype(str).str.strip()

    return out


def preprocess_cash(cash_df):
    """GS Cash Execution Report 데이터를 대사 가능한 공통 스키마로 정규화합니다."""
    if cash_df.empty:
        return pd.DataFrame(columns=['rIC', 'tradeTransactionType', 'settleCurrency',
                                      'notionalQuantity', 'tradingNotionalNetProceeds'])

    cash_df = check_and_filter_nan(cash_df, 'GS Cash Execution Report',
                                    ['RIC', 'SIDE', 'QUANTITY', 'NET_PRINCIPAL', 'SETTLE_CCY'])
    if cash_df.empty:
        return pd.DataFrame(columns=['rIC', 'tradeTransactionType', 'settleCurrency',
                                      'notionalQuantity', 'tradingNotionalNetProceeds'])

    out = pd.DataFrame()
    out['rIC'] = cash_df['RIC'].apply(normalize_ric)
    out['tradeTransactionType'] = cash_df['SIDE'].astype(str).str.strip().replace(SIDE_MAP)
    out['settleCurrency'] = cash_df['SETTLE_CCY'].astype(str).str.strip()
    out['notionalQuantity'] = pd.to_numeric(cash_df['QUANTITY'], errors='coerce').abs()
    out['tradingNotionalNetProceeds'] = pd.to_numeric(cash_df['NET_PRINCIPAL'], errors='coerce').abs()
    out['rIC'] = out['rIC'].astype(str).str.strip()
    out['tradeTransactionType'] = out['tradeTransactionType'].astype(str).str.strip()

    return out


def reconcile_trades(enf_df, gs_df):
    """ENF(GS PBS 필터링분)와 GS PBS(Swap+Cash) 데이터를 대조합니다."""
    print("[작업 시작] 대사(Reconciliation) 수행 시작")

    key_cols = ['rIC', 'tradeTransactionType', 'settleCurrency']
    agg_map = {'notionalQuantity': 'sum', 'tradingNotionalNetProceeds': 'sum'}

    enf_grouped = enf_df.groupby(key_cols, as_index=False).agg(agg_map) if not enf_df.empty else pd.DataFrame(columns=key_cols + list(agg_map))
    gs_grouped = gs_df.groupby(key_cols, as_index=False).agg(agg_map) if not gs_df.empty else pd.DataFrame(columns=key_cols + list(agg_map))

    enf_grouped["_key"] = list(zip(*[enf_grouped[col] for col in key_cols])) if not enf_grouped.empty else []
    gs_grouped["_key"] = list(zip(*[gs_grouped[col] for col in key_cols])) if not gs_grouped.empty else []

    keys_enf = set(enf_grouped["_key"])
    keys_gs = set(gs_grouped["_key"])

    missing_in_gs = enf_grouped[enf_grouped["_key"].isin(keys_enf - keys_gs)][key_cols].copy()
    missing_in_gs["issue"] = "In enf_df, missing in gs_df"

    missing_in_enf = gs_grouped[gs_grouped["_key"].isin(keys_gs - keys_enf)][key_cols].copy()
    missing_in_enf["issue"] = "In gs_df, missing in enf_df"

    missing_report = pd.concat([missing_in_gs, missing_in_enf], ignore_index=True)

    enf_idx = enf_grouped.set_index(key_cols)
    gs_idx = gs_grouped.set_index(key_cols)
    common_keys = keys_enf & keys_gs

    diff_rows = []
    compare_cols = ['notionalQuantity', 'tradingNotionalNetProceeds']

    for key in common_keys:
        for col in compare_cols:
            val_1 = enf_idx.loc[key, col] if col in enf_idx.columns else 0.0
            val_2 = gs_idx.loc[key, col] if col in gs_idx.columns else 0.0

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
                    "val_gs": v2_rounded,
                    "Difference": diff
                })

    diff_report = pd.DataFrame(diff_rows, columns=["key", "variable", "val_enf", "val_gs", "Difference"])

    tolerance = 1.0
    diff_report['Difference'] = pd.to_numeric(diff_report['Difference'], errors='coerce')
    diff_report = diff_report[(diff_report["Difference"].isna()) | (diff_report["Difference"].abs() >= tolerance)]

    print("[작업 완료] 대사 수행 완료")
    return missing_report, diff_report


def main():
    print("=" * 60)
    print("GS PBS 대사(Reconciliation) 프로그램 실행을 시작합니다.")
    print("=" * 60)

    try:
        fetch_eod_report(ENF_USERNAME, ENF_PASSWORD)
    except Exception as exc:
        print(f"[WARN] EOD 리포트 페치 실패 ({exc}). 기존 EOD_Trade.json 파일을 사용합니다.")

    today_dt = datetime.today()

    try:
        enf_df, swap_df, cash_df = load_data(today_dt)
    except Exception as e:
        print(f"[에러] 데이터 로드 실패: {e}")
        sys.exit(1)

    enf_df = preprocess_enf(enf_df, today_dt)
    swap_norm = preprocess_swap(swap_df)
    cash_norm = preprocess_cash(cash_df)
    gs_df = pd.concat([swap_norm, cash_norm], ignore_index=True)

    print(f"[정보] GS PBS 필터링 후 ENF 거래 건수: {len(enf_df)}")
    print(f"[정보] Swap 거래 건수: {len(swap_norm)}, Cash 거래 건수: {len(cash_norm)}")

    missing_report, diff_report = reconcile_trades(enf_df, gs_df)

    print("\n" + "=" * 20 + " 대사 결과 리포트 " + "=" * 20)

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

    if REVIEW_LIST:
        print("\n[검토 필요] 일부 소스 파일에서 NaN 결측치가 발견되어 해당 행들을 분류했습니다:")
        for idx, item in enumerate(REVIEW_LIST):
            print(f"  {idx + 1}. 파일: {item.get('출처파일')}, 사유: {item.get('검토사유')}")
            clean_item = {k: v for k, v in item.items() if k not in ['출처파일', '검토사유']}
            print(f"     내용: {clean_item}")

    print("=" * 60)
    print("GS PBS 대사 프로그램 실행이 완료되었습니다.")
    print("=" * 60)


if __name__ == "__main__":
    main()
