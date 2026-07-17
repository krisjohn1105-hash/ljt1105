import sys
import re
import json
import pandas as pd
from tabulate import tabulate
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

# 경로 설정
PRELUDE_DATA_DIR = BASE_DIR.parent / 'data' / 'input' / 'pre-trade'
SWAP_DIR = BASE_DIR.parent / 'data' / 'input' / 'QRT' / 'Swap'
CASH_DIR = BASE_DIR.parent / 'data' / 'input' / 'QRT' / 'Cash'
EOD_TRADE_PATH = BASE_DIR / 'EOD_Trade.json'
RESULT_DIR = Path('Z:/02.펀드/019. 일간매매내역/SMA_recon_result')

# 계좌 구분 기준 (소문자 비교, 정규식)
# 1순위: fundName — 'Prelude SMA' -> Prelude, 'QSMA' -> Qube
# 2순위(폴백): fund가 아직 배정 전(예: 'Preallocation')이면 counterPartyName으로 구분
#   - Prelude: 'MS Equity Algo', 'MS Future Algo', 'MS Equity HT', 'Morgan Stanley ...' 등
#   - Qube(GS PBS): 'GS Equity Algo', 'GS Equity HT', 'Goldman ...', 'GSIL ...' 등
ACCOUNTS = {
    'Prelude': {'fund': r'prelude', 'counterparty': r'^ms\s|morgan stanley'},
    'Qube': {'fund': r'qsma', 'counterparty': r'^gs\s|goldman|gsil'},
}
ALL_FUND_PATTERN = '|'.join(a['fund'] for a in ACCOUNTS.values())

# 대사 키/비교 변수 (모든 소스 공통 스키마)
KEY_COLS = ['rIC', 'tradeTransactionType', 'settleCurrency']
COMPARE_COLS = ['notionalQuantity', 'tradingNotionalNetProceeds']

# 거래 유형 정규화: 브로커 약어(B/S/SS/BC)와 ENF 표기(Buy to Cover 등)를 Buy/Sell로 통일
SIDE_MAP = {
    'B': 'Buy', 'S': 'Sell', 'SS': 'Sell', 'BC': 'Buy',
    'Buy to Cover': 'Buy', 'Sell Short': 'Sell',
}

# 오차 허용 범위 (절대값 1.0 이상인 경우만 차이로 분류)
TOLERANCE = 1.0

# 결측치(NaN)가 있는 행을 관리하기 위한 '검토 필요' 리스트 (Prelude/GS PBS 공통)
REVIEW_LIST = []


# ==================================================================================
# 공통 유틸리티
# ==================================================================================

def normalize_ric(ric_str):
    """종목 코드(RIC)의 숫자 부분을 6자리로 zfill 하여 정규화합니다."""
    if pd.isna(ric_str):
        return ""
    ric_str = str(ric_str).strip()
    code, _, suffix = ric_str.partition('.')
    if code.isdigit():
        return f"{code.zfill(6)}.{suffix}" if suffix else code.zfill(6)
    return ric_str


def check_and_filter_nan(df, file_label, required_cols):
    """필수 칼럼에 결측치가 있는 행을 '검토 필요' 리스트로 분류하고 제외합니다."""
    if df.empty:
        return df

    cols_to_check = [c for c in required_cols if c in df.columns]
    nan_mask = df[cols_to_check].isna().any(axis=1)
    if nan_mask.any():
        print(f"[경고] {file_label}에서 결측치(NaN)가 발견되었습니다. {int(nan_mask.sum())}행을 '검토 필요'로 분류합니다.")
        for _, row in df[nan_mask].iterrows():
            REVIEW_LIST.append({**row.to_dict(), '출처파일': file_label, '검토사유': "필수 항목 누락(NaN)"})
        df = df[~nan_mask]
    return df


def finalize_match_frame(df):
    """
    소스별 데이터프레임을 대사 공통 스키마(KEY_COLS + COMPARE_COLS)로 정리합니다.
    - RIC 6자리 정규화, 거래 유형 Buy/Sell 통일, 키 공백 제거
    - 수량/금액은 절대값 처리 (방향 정보는 키의 tradeTransactionType이 담당)
    """
    if df.empty:
        return pd.DataFrame(columns=KEY_COLS + COMPARE_COLS)

    df = df.reindex(columns=KEY_COLS + COMPARE_COLS)
    df['rIC'] = df['rIC'].apply(normalize_ric)
    df['tradeTransactionType'] = df['tradeTransactionType'].astype(str).str.strip().replace(SIDE_MAP)
    for col in KEY_COLS:
        df[col] = df[col].astype(str).str.strip()
    for col in COMPARE_COLS:
        df[col] = pd.to_numeric(df[col], errors='coerce').abs()
    return df


def account_membership_mask(df, account):
    """
    ENF 블로터에서 해당 계좌(account) 소속 거래를 판별하는 불리언 마스크를 반환합니다.
    1) fundName이 계좌의 fund 패턴에 매칭되면 해당 계좌
    2) fundName이 어느 계좌의 fund 패턴에도 매칭되지 않으면(배정 전 등)
       counterPartyName 패턴으로 폴백하여 구분
    """
    idx = df.index
    fund = df['fundName.value'].astype(str).str.lower() if 'fundName.value' in df.columns else pd.Series('', index=idx)
    cp = df['counterPartyName.value'].astype(str).str.lower() if 'counterPartyName.value' in df.columns else pd.Series('', index=idx)

    fund_known = fund.str.contains(ALL_FUND_PATTERN, regex=True)
    fund_match = fund.str.contains(ACCOUNTS[account]['fund'], regex=True)
    cp_match = cp.str.contains(ACCOUNTS[account]['counterparty'], regex=True)

    return fund_match | (~fund_known & cp_match)


def preprocess_enf(enf_df, target_date_str, source_label, account=None):
    """
    ENF EOD 블로터를 대사 공통 스키마로 전처리합니다.
    account('Prelude'/'Qube')가 주어지면 해당 계좌 소속 거래만 남깁니다.
    """
    if enf_df.empty:
        return finalize_match_frame(enf_df)

    df = enf_df.copy()

    if account:
        df = df[account_membership_mask(df, account)]

    if 'tradeDate.value' in df.columns:
        trade_dates = pd.to_datetime(df['tradeDate.value']).dt.strftime('%Y-%m-%d')
        df = df[trade_dates == target_date_str]

    required_cols = [f'{c}.value' for c in KEY_COLS + COMPARE_COLS]
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        print(f"[경고] ENF 데이터에 일부 필수 칼럼이 누락되었습니다: {missing_cols}")

    df = check_and_filter_nan(
        df, f'EOD_Trade.json ({source_label})',
        ['rIC.value', 'tradeTransactionType.value', 'notionalQuantity.value', 'tradingNotionalNetProceeds.value']
    )

    df.columns = [col.replace('.value', '') for col in df.columns]
    return finalize_match_frame(df)


# 비교 변수의 리포트용 한글 이름
COMPARE_COL_KOR = {'notionalQuantity': '체결수량', 'tradingNotionalNetProceeds': '체결금액'}

# 키 컬럼의 리포트용 한글 이름
KEY_COL_KOR = {'rIC': '종목코드', 'tradeTransactionType': '매매구분', 'settleCurrency': '결제통화'}


# 통합 대사 키: 계좌(Prelude/Qube) + 종목/매매구분/결제통화
RECON_KEYS = ['계좌'] + KEY_COLS


def reconcile_trades(enf_df, broker_df):
    """
    ENF 전체 거래와 브로커(Prelude MS recap + Qube GS PBS) 전체 거래를
    계좌 + 종목/매매구분/결제통화 기준으로 통합 대사합니다.

    각 키에 대해 양측의 체결수량/체결금액과 차이를 모두 표기하며,
    한쪽에만 존재하는 거래는 없는 쪽을 NaN으로 두고 '사유'에 누락 사실을 기록합니다.
    (차이 계산 시에는 누락 측을 0으로 간주합니다.)
    """
    print("[작업 시작] 통합 대사(Reconciliation) 수행 시작")

    enf_grouped = enf_df.groupby(RECON_KEYS, as_index=False)[COMPARE_COLS].sum()
    broker_grouped = broker_df.groupby(RECON_KEYS, as_index=False)[COMPARE_COLS].sum()

    merged = enf_grouped.merge(
        broker_grouped, on=RECON_KEYS, how='outer',
        suffixes=('_enf', '_oth'), indicator=True, sort=True,
    )

    report = merged[RECON_KEYS].rename(columns=KEY_COL_KOR)
    report['매매구분'] = report['매매구분'].replace({'Buy': '매수', 'Sell': '매도'})

    for col, kor in COMPARE_COL_KOR.items():
        # 소수점 3째 자리에서 반올림하여 2째자리까지 비교
        # (한쪽이 빈 데이터프레임이면 merge 결과가 object dtype이 되므로 숫자형으로 강제 변환)
        v_enf = pd.to_numeric(merged[f'{col}_enf'], errors='coerce').round(2)
        v_oth = pd.to_numeric(merged[f'{col}_oth'], errors='coerce').round(2)
        report[f'{kor}_ENF'] = v_enf
        report[f'{kor}_브로커'] = v_oth
        # 차이 계산 시 누락 측은 0으로 간주
        report[f'{kor}_차이'] = (v_enf.fillna(0) - v_oth.fillna(0)).round(2)

    # 사유 판정: 누락 > 오차 범위(TOLERANCE) 초과 차이 > 일치
    def build_reason(row_merge, qty_diff, amt_diff):
        if row_merge == 'left_only':
            return "ENF에만 존재 (브로커 누락)"
        if row_merge == 'right_only':
            return "브로커에만 존재 (ENF 누락)"
        diffs = []
        if abs(qty_diff) >= TOLERANCE:
            diffs.append("수량 불일치")
        if abs(amt_diff) >= TOLERANCE:
            diffs.append("금액 불일치")
        return ", ".join(diffs) if diffs else "일치"

    report['사유'] = [
        build_reason(m, q, a)
        for m, q, a in zip(merged['_merge'], report['체결수량_차이'], report['체결금액_차이'])
    ]

    # 미분류 계좌 거래는 비교할 브로커 파일이 없으므로 사유를 별도 표기
    report.loc[report['계좌'].astype(str).str.startswith('미분류'), '사유'] = '계좌 미분류 (확인 필요)'

    print("[작업 완료] 통합 대사 수행 완료")
    return report


def print_reconciliation_report(report):
    print("\n" + "=" * 20 + " SMA 통합 대사 결과 " + "=" * 20)

    if report.empty:
        print("\n대사 대상 거래가 없습니다.")
        return

    # 전체 대사 결과 (Prelude + Qube 모든 거래의 차이 내역)
    print("대사 결과:")
    print(tabulate(
        report[['계좌', '종목코드', '매매구분', '결제통화', '체결수량_차이', '체결금액_차이']],
        headers='keys', tablefmt='pretty'
    ))

    # 대사 결과 차이 분석 및 로깅
    diff_df = report[report['사유'] != '일치']

    if not diff_df.empty:
        print("\n[!] 대사 결과 차이가 발생하였습니다. 임의 수정하지 않고 차이 내역을 로그에 기록합니다.")

        def fmt(v):
            return f"{v:.2f}" if pd.notna(v) else "N/A"

        table_rows = [{
            "계좌": row['계좌'],
            "종목코드": row['종목코드'],
            "구분": row['매매구분'],
            "결제통화": row['결제통화'],
            "ENF수량": fmt(row['체결수량_ENF']),
            "브로커수량": fmt(row['체결수량_브로커']),
            "수량차이": fmt(row['체결수량_차이']),
            "ENF금액": fmt(row['체결금액_ENF']),
            "브로커금액": fmt(row['체결금액_브로커']),
            "금액차이": fmt(row['체결금액_차이']),
            "사유": row['사유'],
        } for _, row in diff_df.iterrows()]

        summary_df = pd.DataFrame(table_rows)
        print("\n======== [대사 불일치 요약 표] ========")
        print(tabulate(summary_df, headers='keys', tablefmt='pretty', showindex=False))
        print("========================================")
    else:
        print("\n[+] 대사 결과: 모든 내역이 상호 일치합니다.")


def save_results_to_excel(report_sections, today_dt):
    """
    대사 결과를 엑셀 파일(SMA_recon_result_YYYYMMDD.xlsx)로 저장합니다.
    report_sections: {시트명: 데이터프레임} — 섹션별 누락/차이 리포트.
    검토 필요(NaN) 리스트가 있으면 별도 시트로 함께 저장합니다.
    """
    output_path = RESULT_DIR / f"SMA_recon_result_{today_dt.strftime('%Y%m%d')}.xlsx"

    try:
        RESULT_DIR.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in report_sections.items():
                out = df if not df.empty else pd.DataFrame({'결과': ['대사 대상 거래 없음']})
                out.to_excel(writer, sheet_name=sheet_name, index=False)

            if REVIEW_LIST:
                review_df = pd.DataFrame(REVIEW_LIST).astype(str)
                review_df.to_excel(writer, sheet_name='검토필요(NaN)', index=False)

        print(f"[정보] 대사 결과 저장 완료: {output_path}")
    except Exception as e:
        print(f"[경고] 대사 결과 엑셀 저장 실패 ({e}). 저장 경로를 확인해 주세요: {output_path}")


# ==================================================================================
# Prelude(MS) 대사
# ==================================================================================

def load_pre_trade_excel(directory, prefix, date_str):
    """
    지정된 날짜의 pre-trade 엑셀 파일을 로드합니다.
    파일이 존재하지 않을 경우 가장 최근 수정 파일을 백업으로 활용합니다.
    """
    preferred = directory / f'{prefix} - {date_str}.xls'
    if preferred.exists():
        return pd.read_excel(preferred)

    candidates = sorted(directory.glob(f'{prefix} - *.xls'), key=lambda p: p.stat().st_mtime)
    if candidates:
        print(f"[WARN] {preferred.name} 파일이 존재하지 않아 최신 수정 파일인 {candidates[-1].name}을 로드합니다.")
        return pd.read_excel(candidates[-1])

    print(f"[WARN] {directory} 내에 '{prefix} - *.xls' 패턴의 파일이 존재하지 않습니다. 빈 데이터프레임을 반환합니다.")
    return pd.DataFrame()


def preprocess_ms(ms_df, file_label, required_cols):
    """Prelude(MS) pre-trade 엑셀을 대사 공통 스키마로 전처리합니다."""
    ms_rename = {
        'Buy/Sell': 'tradeTransactionType',
        'Stock Quantity': 'notionalQuantity',
        'Futures Contracts Quantity': 'notionalQuantity',
        'Net Notional in Swap Ccy': 'tradingNotionalNetProceeds',
        'Ric': 'rIC',
        'Swap Ccy': 'settleCurrency',
    }
    ms_df = check_and_filter_nan(ms_df, file_label, required_cols)
    return ms_df.rename(columns=ms_rename)


def prepare_prelude_frames(enf_raw_df, today_str):
    """Prelude(MS) 계좌의 ENF/브로커 데이터를 대사 공통 스키마로 준비합니다."""
    ms_df = load_pre_trade_excel(PRELUDE_DATA_DIR, 'Pre allocation Korea Stocks', today_str)
    ms_futures = load_pre_trade_excel(PRELUDE_DATA_DIR, 'Pre allocation Korea Futures', today_str)

    print("[작업 시작] Prelude 데이터 전처리 시작")

    # MS 데이터의 실제 Trade Date를 대사 대상 일자로 사용 (없으면 오늘 날짜)
    target_date_str = None
    if not ms_df.empty and 'Trade Date' in ms_df.columns:
        try:
            target_date_str = pd.to_datetime(ms_df['Trade Date']).dt.strftime('%Y-%m-%d').iloc[0]
        except Exception as e:
            print(f"[경고] MS Stocks의 Trade Date 파싱 실패: {e}")
    if not target_date_str:
        try:
            target_date_str = datetime.strptime(today_str, '%m%d%y').strftime('%Y-%m-%d')
        except ValueError:
            target_date_str = datetime.today().strftime('%Y-%m-%d')
    print(f"[정보] Prelude 대사 대상 일자 설정: {target_date_str}")

    enf_df = preprocess_enf(enf_raw_df, target_date_str, 'Prelude', account='Prelude')

    ms_combined = pd.concat([
        preprocess_ms(ms_df, 'Pre allocation Korea Stocks',
                      ['Ric', 'Buy/Sell', 'Stock Quantity', 'Net Notional in Swap Ccy']),
        preprocess_ms(ms_futures, 'Pre allocation Korea Futures',
                      ['Ric', 'Buy/Sell', 'Futures Contracts Quantity', 'Net Notional in Swap Ccy']),
    ], ignore_index=True)
    ms_combined = finalize_match_frame(ms_combined)

    print("[작업 완료] Prelude 데이터 전처리 완료")
    enf_df['계좌'] = 'Prelude'
    ms_combined['계좌'] = 'Prelude'
    return enf_df, ms_combined


# ==================================================================================
# GS PBS(SMA Swap/Cash) 대사
# ==================================================================================

def extract_trailing_date_token(stem):
    """
    파일명(확장자 제외)에 포함된 마지막 연속 숫자 구간의 끝 6자리(YYMMDD)를 추출합니다.
    예) 'GSETASIARECAP20260706' -> '260706'
        'KR - Execution Report from GS_260706 - Excel' -> '260706'
    """
    digit_runs = re.findall(r'\d+', stem)
    if not digit_runs or len(digit_runs[-1]) < 6:
        return None
    return digit_runs[-1][-6:]


def find_todays_file(directory, today_dt):
    """
    directory 내 엑셀 파일들 중 파일명에 오늘 날짜(YYMMDD)가 포함된 파일을 찾습니다.
    과거 파일로 대사하면 전부 가짜 누락으로 표시되므로, 오늘 날짜 파일이 없으면
    폴백 없이 None을 반환합니다 (= 당일 GS 거래 없음으로 처리).
    """
    target = today_dt.strftime('%y%m%d')
    for path in sorted(directory.glob('*.xls*')):
        if extract_trailing_date_token(path.stem) == target:
            return path

    print(f"[정보] {directory}에 오늘({target}) 날짜의 파일이 없습니다. 당일 거래 없음으로 처리합니다.")
    return None


def load_gs_file(directory, today_dt, label):
    """Swap/Cash 폴더에서 오늘 날짜(YYMMDD)의 GS 파일 경로를 찾습니다."""
    if not directory.exists():
        print(f"[WARN] {directory} 폴더가 존재하지 않습니다.")
        return None

    path = find_todays_file(directory, today_dt)
    if path is not None:
        print(f"[정보] {label} 파일 로드: {path.name}")
    return path


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
        for ric_col in row_str[row_str == 'RIC'].index:
            qty_col = ric_col + 5
            # 'RIC'에서 5칸 뒤에 'QUANTITY'가 있는 행만 실제 헤더 행으로 인정 (오탐 방지)
            if qty_col >= n_cols or str(raw.iat[i, qty_col]).strip() != 'QUANTITY':
                continue

            j = i + 1
            while j < n_rows:
                ric_val = raw.iat[j, ric_col]
                if pd.isna(ric_val) or str(ric_val).strip() == '':
                    break
                records.append({
                    'RIC': ric_val,
                    'SIDE': raw.iat[j, ric_col + 4],
                    'QUANTITY': raw.iat[j, qty_col],
                    'NET_PRINCIPAL': raw.iat[j, ric_col + 10],
                    'SETTLE_CCY': raw.iat[j, ric_col + 12],
                })
                j += 1

    return pd.DataFrame(records)


def preprocess_swap(swap_df):
    """GS 스왑 리캡 데이터를 대사 공통 스키마로 정규화합니다."""
    swap_df = check_and_filter_nan(swap_df, 'GS Swap Recap', ['RIC', 'B/S', 'QTY', 'Net Px'])
    if swap_df.empty:
        return finalize_match_frame(swap_df)

    qty = pd.to_numeric(swap_df['QTY'], errors='coerce')
    net_px = pd.to_numeric(swap_df['Net Px'], errors='coerce')
    out = pd.DataFrame({
        'rIC': swap_df['RIC'],
        'tradeTransactionType': swap_df['B/S'],
        # 스왑 거래는 Net Px(현지 통화 가격을 FX Rate로 나눈 USD 환산 가격) 기준 USD 결제
        'settleCurrency': 'USD',
        'notionalQuantity': qty,
        'tradingNotionalNetProceeds': qty * net_px,
    })
    return finalize_match_frame(out)


def preprocess_cash(cash_df):
    """GS Cash Execution Report 데이터를 대사 공통 스키마로 정규화합니다."""
    cash_df = check_and_filter_nan(cash_df, 'GS Cash Execution Report',
                                   ['RIC', 'SIDE', 'QUANTITY', 'NET_PRINCIPAL', 'SETTLE_CCY'])
    if cash_df.empty:
        return finalize_match_frame(cash_df)

    out = cash_df.rename(columns={
        'RIC': 'rIC',
        'SIDE': 'tradeTransactionType',
        'SETTLE_CCY': 'settleCurrency',
        'QUANTITY': 'notionalQuantity',
        'NET_PRINCIPAL': 'tradingNotionalNetProceeds',
    })
    return finalize_match_frame(out)


def prepare_qube_frames(enf_raw_df, today_dt):
    """Qube(GS PBS Swap + Cash) 계좌의 ENF/브로커 데이터를 대사 공통 스키마로 준비합니다."""
    swap_path = load_gs_file(SWAP_DIR, today_dt, '스왑')
    cash_path = load_gs_file(CASH_DIR, today_dt, 'Cash')

    swap_df = pd.read_excel(swap_path) if swap_path else pd.DataFrame()
    cash_df = parse_cash_execution_report(cash_path) if cash_path else pd.DataFrame()

    enf_df = preprocess_enf(
        enf_raw_df, today_dt.strftime('%Y-%m-%d'), 'GS PBS', account='Qube',
    )
    swap_norm = preprocess_swap(swap_df)
    cash_norm = preprocess_cash(cash_df)
    gs_df = pd.concat([swap_norm, cash_norm], ignore_index=True)

    print(f"[정보] Qube(GS PBS) 필터링 후 ENF 거래 건수: {len(enf_df)}")
    print(f"[정보] Swap 거래 건수: {len(swap_norm)}, Cash 거래 건수: {len(cash_norm)}")

    enf_df['계좌'] = 'Qube'
    gs_df['계좌'] = 'Qube'
    return enf_df, gs_df


# ==================================================================================
# 메인 실행 흐름
# ==================================================================================

def main():
    print("=" * 60)
    print("SMA 대사(Reconciliation) 프로그램 실행을 시작합니다.")
    print("=" * 60)

    # 1. EOD 리포트 실시간 페치 시도 (Prelude/GS PBS 공통으로 1회만 수행)
    try:
        fetch_eod_report(ENF_USERNAME, ENF_PASSWORD)
    except Exception as exc:
        print(f"[WARN] EOD 리포트 페치 실패 ({exc}). 기존 EOD_Trade.json 파일을 사용합니다.")

    if not EOD_TRADE_PATH.exists():
        print(f"[에러] {EOD_TRADE_PATH} 파일이 존재하지 않습니다. api.py를 먼저 실행해 주세요.")
        sys.exit(1)

    with open(EOD_TRADE_PATH, "r") as f:
        enf_raw_df = pd.json_normalize(json.load(f)['rows'])

    today_dt = datetime.today()

    # 어느 계좌(Prelude/Qube)에도 분류되지 않는 거래 파악 (예: NH 등 기타 브로커)
    # 제외하지 않고 계좌 '미분류'로 리포트에 포함시켜 거래가 누락되지 않게 합니다.
    assigned = pd.Series(False, index=enf_raw_df.index)
    for account in ACCOUNTS:
        assigned |= account_membership_mask(enf_raw_df, account)
    unassigned_raw = enf_raw_df[~assigned]
    if not unassigned_raw.empty:
        detail_cols = [c for c in ['fundName.value', 'counterPartyName.value'] if c in unassigned_raw.columns]
        print("[경고] Prelude/Qube 어느 계좌에도 분류되지 않는 거래가 있습니다. 계좌 '미분류'로 리포트에 포함합니다:")
        print(unassigned_raw[detail_cols].drop_duplicates().to_string(index=False))

    # 2. 계좌별 데이터 준비 (Prelude: MS recap / Qube: GS PBS Swap+Cash)
    print("\n" + "-" * 20 + " Prelude(MS) 데이터 준비 " + "-" * 20)
    prelude_enf, prelude_broker = prepare_prelude_frames(enf_raw_df, today_dt.strftime('%m%d%y'))

    print("\n" + "-" * 20 + " Qube(GS PBS) 데이터 준비 " + "-" * 20)
    qube_enf, qube_broker = prepare_qube_frames(enf_raw_df, today_dt)

    # 3. 통합 대사: ENF 전체 vs 브로커(Prelude + Qube) 전체
    # 미분류 거래(NH 등)도 ENF 측에 포함시켜 리포트에서 확인 가능하게 함
    unassigned_enf = preprocess_enf(unassigned_raw, today_dt.strftime('%Y-%m-%d'), '미분류')
    unassigned_enf['계좌'] = '미분류'

    enf_all = pd.concat([prelude_enf, qube_enf, unassigned_enf], ignore_index=True)
    broker_all = pd.concat([prelude_broker, qube_broker], ignore_index=True)
    report = reconcile_trades(enf_all, broker_all)

    # 4. 결과 리포트 출력
    print_reconciliation_report(report)

    # 5. 결과 엑셀 저장
    save_results_to_excel({'SMA_대사': report}, today_dt)

    # 6. 검토 필요 리스트 출력 (공통)
    if REVIEW_LIST:
        print("\n======== [검토 필요 리스트 (결측치 존재)] ========")
        review_df = pd.DataFrame(REVIEW_LIST)
        cols_to_show = ['출처파일', '검토사유'] + [c for c in review_df.columns if c not in ['출처파일', '검토사유']]
        review_df = review_df[cols_to_show]
        print(tabulate(review_df, headers='keys', tablefmt='pretty', showindex=False))
        print("====================================================")

    print("\n" + "=" * 60)
    print("SMA 대사 프로그램 실행이 완료되었습니다.")
    print("=" * 60)


if __name__ == "__main__":
    main()
