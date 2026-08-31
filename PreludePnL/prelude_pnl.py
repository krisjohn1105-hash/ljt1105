"""
Prelude(Morgan Stanley Blue Border) 리포트 기반 통합 손익/현금 리포트 생성기
==========================================================================

지정한 폴더(및 하위 폴더)에 있는 MS Prelude 추출 리포트(CSV)를 읽어서

  1) Swap / Cash Equity / FX / Cash 등 모든 자산군의 일일손익·누적손익
  2) 현재 현금잔고, 결제에 필요한 현금, 결제 후 현금
  3) 종목별 손익 상세, 거래내역, 검증(Recon)

을 하나의 엑셀 파일로 산출한다.
추가로 --organize 옵션으로 원본 리포트를 리포트 종류별 폴더로 분류·이동한다.

사용 예)
    python prelude_pnl.py --src "Z:/02.펀드/003.매매보고서 대사/Prelude_new"
    python prelude_pnl.py --src ... --organize --dry-run
    python prelude_pnl.py --src ... --organize --layout report-year

손익 계산 원리
--------------
모든 자산군에 동일한 시가평가(mark-to-market) 항등식을 적용한다.

    일일손익(자산군) = 평가액(t) - 평가액(t-1) + 해당 자산군에 귀속되는 현금흐름(t)

  * 평가액   : MAC001X(Global Positions Extract)의 'Market Value / Net Equity (Base)'
  * 현금흐름 : MAC002TDX(Normalized Trade Date Activity)의 'Net Amt Base' 중
               실제 현금원장(Position Type = PB / COLCASH)에 계상된 금액
               (매수 -, 매도 +, 스왑 리셋 수취 + ...)

  현금(Cash) 자산군은 잔여항으로 계산한다.
      일일손익(Cash) = ΔCash평가액 - Σ(타 자산군 현금흐름) - 외부 자금이동
      => 환평가손익 + 이자 + 배당 + 수수료가 여기에 집계된다.

  따라서  Σ 자산군 일일손익 = Δ총평가액 - 외부 자금이동  (완전 일치)

작성: Claude Code / 2026-08
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import re
import shutil
import sys
from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------

MONTHS = {m: i + 1 for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])}

# 파일명 규칙: "{CODE} - {Report Name} - {Account} - {DDMonYYYY}-{seq}.{ext}"
FILENAME_RE = re.compile(
    r"^(?P<label>.+?)\s*-\s*(?P<account>[A-Z0-9]{6,12})\s*-\s*"
    r"(?P<day>\d{2})(?P<mon>[A-Za-z]{3})(?P<year>\d{4})-(?P<seq>\d+)$")

# 자산군(MAC001X 'Asset Class') -> 리포트 버킷
ASSET_CLASS_TO_BUCKET = {
    "Equity Swaps": "Swap",
    "Cash Securities": "Cash Equity",
    "Fx Forwards": "FX",
    "Cash": "Cash",
}
BUCKET_ORDER = ["Swap", "Cash Equity", "FX", "Cash"]
BUCKET_KR = {
    "Swap": "스왑",
    "Cash Equity": "현물주식",
    "FX": "외환",
    "Cash": "현금/기타",
}

# 현금원장에 계상되는 Position Type (실제 현금이 움직이는 행)
CASH_LEDGER_POSITION_TYPES = {"PB", "COLCASH"}

# MAC002TDX Transaction Category Level 3 -> 버킷
L3_TO_BUCKET = {
    "Equity Swap Financing": "Swap",
    "Swap Activity": "Swap",
    "Swap Performance": "Swap",
    "Swap Collateral": "Swap",
    "Buy Long": "Cash Equity",
    "Sell Long": "Cash Equity",
    "Buy to Cover": "Cash Equity",
    "Sell Short": "Cash Equity",
    "Buy": "Cash Equity",
    "Sell": "Cash Equity",
    "FX Settlement": "FX",
    "Collateral": "Cash",        # PB <-> COLCASH 내부이체 (합계 0)
    "Margin Transfer": "Cash",   # PB <-> COLCASH 내부이체 (합계 0)
    "Dividends": "Cash",         # 손익 (현금 버킷에 귀속)
    "Interest": "Cash",
    "Custody Fee": "Cash",
    "Corporate Actions": "Cash Equity",
}
PRODUCT_TYPE_DESC_TO_BUCKET = {
    "EQUITY SWAP": "Swap",
    "Equity": "Cash Equity",
    "FX FORWARDS TRADES": "FX",
    "FX SPOT TRADES": "FX",
}

# 외부 자금이동(입출금)으로 볼 카테고리.
# 'Wires' = FUNDS PAID OR RECEIVED (펀드 외부로의 송금) -> 손익이 아니므로 제외한다.
DEFAULT_EXTERNAL_CATEGORIES: Tuple[str, ...] = ("Wires",)

# 직전 리포트일과의 간격이 이 일수를 넘으면 그 사이의 거래내역을 알 수 없으므로
# 해당 행의 손익을 산출하지 않는다(월말 스냅샷만 있는 구간 등).
DEFAULT_MAX_GAP_DAYS = 5

# 잔고 마커 행(손익계산에서 제외)
BALANCE_MARKER_CATEGORIES = {"STARTING CASH BALANCE", "ENDING CASH BALANCE"}

REPORT_CODES = {
    "positions": "MAC001X",          # Global Positions Extract
    "activity": "MAC002TDX",         # Normalized Trade Date Activity Extract
    "cash_forecast": "CASH005X",     # Next Five Days Activity Summary Extract
    "cash_detail": "CASH005DX",      # Next Five Days Activity Detail Extract
    "swap_mtm": "EQSWAP36X",         # Equity Swap MTM Summary Extract
    "swap_reset": "EQSWAP18SX",      # Equity Swap Unwind-Reset Detail Extract (일별)
    "swap_cashflow": "EQSWAP20MX",   # Equity Swap Cashflow Summary Extract
    "interest": "SW1003MX",          # Daily Interest Summary Extract By Currency
    "dividend": "MAC007X",           # Dividend Income
}


# ---------------------------------------------------------------------------
# 파일 스캔 / 파싱
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class SourceFile:
    path: str
    filename: str
    code: str          # 리포트 코드 (예: MAC001X). 없으면 ""
    label: str         # "MAC001X - Global Positions Extract"
    account: str
    date: Optional[dt.date]
    seq: int
    size: int

    @property
    def ext(self) -> str:
        return os.path.splitext(self.filename)[1].lower()


def long_path(p: str) -> str:
    r"""Windows 260자 경로 제한 우회(\\?\ 접두어). 그 외 OS에서는 그대로 반환."""
    if os.name != "nt":
        return p
    p = os.path.abspath(p)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def parse_filename(path: str) -> SourceFile:
    filename = os.path.basename(path)
    base = os.path.splitext(filename)[0]
    m = FILENAME_RE.match(base)
    if m:
        label = m.group("label").strip()
        try:
            date = dt.date(int(m.group("year")), MONTHS[m.group("mon").title()], int(m.group("day")))
        except (KeyError, ValueError):
            date = None
        account = m.group("account")
        seq = int(m.group("seq"))
    else:
        label, account, date, seq = base.strip(), "", None, 0

    # 리포트 코드: 라벨의 첫 토큰이 대문자+숫자로만 이루어져 있으면 코드로 인정
    first = re.split(r"\s*-\s*", label)[0].strip()
    code = first if re.fullmatch(r"[A-Z0-9]{4,12}", first) else ""

    return SourceFile(path=path, filename=filename, code=code, label=label,
                      account=account, date=date, seq=seq,
                      size=os.path.getsize(path))


def scan_files(root: str, exts: Sequence[str] = (".csv", ".html", ".htm", ".xlsx", ".xls", ".pdf", ".txt")) -> List[SourceFile]:
    """root 및 하위 폴더의 리포트 파일을 모두 수집한다(정리 후에도 동작하도록 재귀 스캔)."""
    found: List[SourceFile] = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith((".", "_output"))]
        for fn in filenames:
            if os.path.splitext(fn)[1].lower() in exts:
                found.append(parse_filename(os.path.join(dirpath, fn)))
    return found


_COUNT_RE = re.compile(r"^\s*count\s*=", re.IGNORECASE)


def read_report(path: str) -> pd.DataFrame:
    """Prelude CSV를 읽는다. 선두 preamble 행과 말미 'Count=' 푸터 행을 제거한다."""
    with open(long_path(path), "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
        rows = list(csv.reader(fh))
    if not rows:
        return pd.DataFrame()

    # 헤더 행 찾기: 비어있지 않은 셀이 3개 이상인 첫 행
    hdr_idx = 0
    for i, r in enumerate(rows[:5]):
        if sum(1 for c in r if c.strip()) >= 3:
            hdr_idx = i
            break
    header = [c.strip() for c in rows[hdr_idx]]
    ncol = len(header)

    data = []
    for r in rows[hdr_idx + 1:]:
        if not any(c.strip() for c in r):
            continue
        # 푸터: "REPORT - name", "Count = N"
        if any(_COUNT_RE.match(c) for c in r[:3] if c):
            continue
        r = (list(r) + [""] * ncol)[:ncol]
        data.append(r)

    df = pd.DataFrame(data, columns=header)
    # 중복 컬럼명 제거(뒤엣것 삭제)
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    return df


def to_num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).str.strip().replace({"": None, "nan": None}),
        errors="coerce").fillna(0.0)


def to_date(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce", format="mixed").dt.date


def index_by_code(files: Iterable[SourceFile]) -> Dict[str, List[SourceFile]]:
    idx: Dict[str, List[SourceFile]] = defaultdict(list)
    for f in files:
        if f.code:
            idx[f.code].append(f)
    return idx


def load_code(files_by_code: Dict[str, List[SourceFile]], code: str,
              verbose: bool = True) -> pd.DataFrame:
    """해당 리포트 코드의 모든 파일을 읽어 하나의 DataFrame으로 합친다(파일일자 컬럼 추가)."""
    frames = []
    for f in sorted(files_by_code.get(code, []), key=lambda x: (x.date or dt.date.min, x.seq)):
        if f.ext != ".csv" or f.date is None:
            continue
        try:
            df = read_report(f.path)
        except Exception as exc:  # pragma: no cover
            print(f"  ! {f.filename} 읽기 실패: {exc}", file=sys.stderr)
            continue
        if df.empty:
            continue
        df["_기준일"] = f.date
        df["_원본파일"] = f.filename
        frames.append(df)
    if not frames:
        if verbose:
            print(f"  - {code}: 파일 없음")
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True, sort=False)
    if verbose:
        print(f"  - {code}: {len(frames)}개 파일 / {len(out):,}행")
    return out


# ---------------------------------------------------------------------------
# 1) 포지션 (MAC001X)
# ---------------------------------------------------------------------------

def build_positions(raw: pd.DataFrame) -> pd.DataFrame:
    if raw.empty:
        return raw
    df = raw.copy()
    df["기준일"] = df["_기준일"]
    df["자산군_원본"] = df.get("Asset Class", "")
    df["버킷"] = df["자산군_원본"].map(ASSET_CLASS_TO_BUCKET).fillna(df["자산군_원본"])
    df["평가액_USD"] = to_num(df.get("Market Value / Net Equity (Base)", pd.Series(dtype=str)))
    df["총평가액_USD"] = to_num(df.get("Gross Market Value (Base)", pd.Series(dtype=str)))
    df["원가_USD"] = to_num(df.get("Notional Cost (Base)", pd.Series(dtype=str)))
    df["수량"] = to_num(df.get("Current Quantity", pd.Series(dtype=str)))
    df["결제수량"] = to_num(df.get("Settlement Date Quantity", pd.Series(dtype=str)))
    df["SD잔고_발행통화"] = to_num(df.get("S/D Balance (Issue)", pd.Series(dtype=str)))
    df["SD잔고_USD"] = to_num(df.get("S/D Balance (Base)", pd.Series(dtype=str)))
    df["발행통화"] = df.get("Issue Currency", "")
    df["종목명"] = df.get("Security Description", "")
    df["종목코드"] = df.get("Symbol", "")
    df["ISIN"] = df.get("ISIN", "")
    df["CUSIP"] = df.get("CUSIP", "")
    df["상품유형"] = df.get("Product Type", "")
    df["포지션유형"] = df.get("Position Type", "")
    df["서브계좌"] = df.get("Sub Account Number", "")
    df["FX_USD환산"] = to_num(df.get("Price (Base)", pd.Series(dtype=str)))
    return df


# ---------------------------------------------------------------------------
# 2) 거래/저널 (MAC002TDX)
# ---------------------------------------------------------------------------

def build_activity(raw: pd.DataFrame) -> pd.DataFrame:
    """월말에는 Daily/MTD 파일이 함께 존재하므로 Entry Date == 기준일 행만 사용하고 중복 제거."""
    if raw.empty:
        return raw
    df = raw.copy()
    df["기준일"] = df["_기준일"]
    df["입력일"] = to_date(df.get("Entry Date", pd.Series(dtype=str)))
    df["매매일"] = to_date(df.get("Trade Date", pd.Series(dtype=str)))
    df["결제일"] = to_date(df.get("Settle Date", pd.Series(dtype=str)))

    df["구분"] = df.get("Activity Category", "")
    df = df[~df["구분"].isin(BALANCE_MARKER_CATEGORIES)].copy()
    # 당일 입력분만 (MTD 파일의 과거분 제거)
    df = df[df["입력일"] == df["기준일"]].copy()

    dedup_cols = [c for c in df.columns if not c.startswith("_")]
    df = df.drop_duplicates(subset=dedup_cols).copy()

    df["포지션유형"] = df.get("Position Type", "")
    df["상품유형"] = df.get("Product Type Desc", "")
    df["대분류"] = df.get("Transaction Category Level 1", "")
    df["중분류"] = df.get("Transaction Category Level 2", "")
    df["소분류"] = df.get("Transaction Category Level 3", "")
    df["종목명"] = df.get("Security Description", "")
    df["종목코드"] = df.get("SYMBOL", "")
    df["ISIN"] = df.get("ISIN", "")
    df["CUSIP"] = df.get("CUSIP", "")
    df["매매구분"] = df.get("Buy Sell", "")
    df["수량"] = to_num(df.get("Quantity", pd.Series(dtype=str)))
    df["단가_USD"] = to_num(df.get("Price Base", pd.Series(dtype=str)))
    df["약정금액_USD"] = to_num(df.get("Principal Base", pd.Series(dtype=str)))
    df["수수료_USD"] = to_num(df.get("Commission Base", pd.Series(dtype=str)))
    df["세금_USD"] = to_num(df.get("Taxes Base", pd.Series(dtype=str)))
    df["이자_USD"] = to_num(df.get("Interest Base", pd.Series(dtype=str)))
    df["정산금액_USD"] = to_num(df.get("Net Amt Base", pd.Series(dtype=str)))
    df["정산금액_결제통화"] = to_num(df.get("Net Amt Settle", pd.Series(dtype=str)))
    df["결제통화"] = df.get("Settlement CCY", "")
    df["현금원장"] = df["포지션유형"].isin(CASH_LEDGER_POSITION_TYPES)
    df["버킷"] = df.apply(_bucket_of_activity, axis=1)
    return df


def _bucket_of_activity(row) -> str:
    l3 = str(row.get("소분류", "")).strip()
    if l3 in L3_TO_BUCKET:
        return L3_TO_BUCKET[l3]
    ptd = str(row.get("상품유형", "")).strip()
    if ptd in PRODUCT_TYPE_DESC_TO_BUCKET:
        return PRODUCT_TYPE_DESC_TO_BUCKET[ptd]
    pt = str(row.get("포지션유형", "")).strip()
    if pt in ("EQS", "SWAP"):
        return "Swap"
    if pt == "FX":
        return "FX"
    return "Cash"


# ---------------------------------------------------------------------------
# 3) 손익 계산
# ---------------------------------------------------------------------------

def compute_pnl(positions: pd.DataFrame, activity: pd.DataFrame,
                external_categories: Sequence[str] = (),
                max_gap_days: int = DEFAULT_MAX_GAP_DAYS) -> Dict[str, pd.DataFrame]:
    """자산군별 일일손익/누적손익을 계산한다."""
    if positions.empty:
        raise SystemExit("MAC001X(Global Positions Extract) 파일이 없어 손익을 계산할 수 없습니다.")

    dates = sorted(positions["기준일"].unique())

    # ---- 평가액 (일자 x 버킷)
    mv = (positions.groupby(["기준일", "버킷"], as_index=False)["평가액_USD"].sum()
          .pivot(index="기준일", columns="버킷", values="평가액_USD")
          .reindex(dates).fillna(0.0))
    for b in BUCKET_ORDER:
        if b not in mv.columns:
            mv[b] = 0.0
    mv = mv[[b for b in BUCKET_ORDER if b in mv.columns] +
            [c for c in mv.columns if c not in BUCKET_ORDER]]

    # ---- 현금흐름 (현금원장 행만, 일자 x 버킷)
    if activity.empty:
        flow = pd.DataFrame(0.0, index=mv.index, columns=mv.columns)
        ext = pd.Series(0.0, index=mv.index)
    else:
        cash_rows = activity[activity["현금원장"]].copy()
        ext_mask = cash_rows["소분류"].isin(external_categories) | cash_rows["중분류"].isin(external_categories)
        ext = (cash_rows[ext_mask].groupby("기준일")["정산금액_USD"].sum()
               .reindex(mv.index).fillna(0.0))
        cash_rows = cash_rows[~ext_mask]
        flow = (cash_rows.groupby(["기준일", "버킷"], as_index=False)["정산금액_USD"].sum()
                .pivot(index="기준일", columns="버킷", values="정산금액_USD")
                .reindex(mv.index).fillna(0.0))
        for c in mv.columns:
            if c not in flow.columns:
                flow[c] = 0.0
        flow = flow[mv.columns]

    # ---- 산출 가능 구간 판정
    idx = list(mv.index)
    prev_date = pd.Series([None] + idx[:-1], index=idx)
    gap_days = pd.Series(
        [float("nan")] + [(idx[i] - idx[i - 1]).days for i in range(1, len(idx))], index=idx)
    # 첫 기준일 + 직전 리포트일과 간격이 큰 행은 그 사이 거래내역을 알 수 없어 손익 산출 불가
    computable = pd.Series([False] + [g <= max_gap_days for g in gap_days.iloc[1:]], index=idx)

    d_mv = mv.diff()

    pnl = pd.DataFrame(0.0, index=mv.index, columns=mv.columns, dtype=float)
    other_buckets = [c for c in mv.columns if c != "Cash"]
    for b in other_buckets:
        pnl[b] = d_mv[b] + flow[b]
    if "Cash" in mv.columns:
        pnl["Cash"] = d_mv["Cash"] - flow[other_buckets].sum(axis=1) - ext
    pnl = pnl.where(computable, 0.0)

    note = pd.Series("", index=idx)
    note.iloc[0] = "기초일(직전 평가액 없음) - 손익 산출 제외"
    note[(~computable) & (note == "")] = "직전 리포트일과 간격이 커서 손익 산출 제외(월말 스냅샷 등)"

    total_mv = mv.sum(axis=1)
    daily = pd.DataFrame(index=mv.index)
    daily.index.name = "기준일"
    daily["직전 기준일"] = prev_date
    daily["경과일수"] = gap_days
    for b in mv.columns:
        daily[f"{BUCKET_KR.get(b, b)} 일일손익"] = pnl[b]
    daily["합계 일일손익"] = pnl.sum(axis=1)
    for b in mv.columns:
        daily[f"{BUCKET_KR.get(b, b)} 누적손익"] = pnl[b].cumsum()
    daily["합계 누적손익"] = pnl.sum(axis=1).cumsum()
    daily["총평가액(NAV)"] = total_mv
    daily["전일 총평가액"] = total_mv.shift()
    daily["외부 자금이동"] = ext.where(computable, 0.0)
    daily["일일수익률(%)"] = (daily["합계 일일손익"] /
                          daily["전일 총평가액"].replace(0, pd.NA) * 100).astype(float)
    daily.loc[~computable, "일일수익률(%)"] = float("nan")
    daily["비고"] = note

    # ---- 자산군별 상세(long)
    rows = []
    for i, d in enumerate(mv.index):
        for b in mv.columns:
            rows.append({
                "기준일": d,
                "자산군": BUCKET_KR.get(b, b),
                "자산군(원문)": b,
                "전일 평가액": mv[b].shift().iloc[i] if i else float("nan"),
                "당일 평가액": mv[b].iloc[i],
                "평가액 증감": d_mv[b].iloc[i],
                "현금흐름(매매/정산)": flow[b].iloc[i],
                "일일손익": pnl[b].iloc[i],
                "누적손익": pnl[b].cumsum().iloc[i],
                "비고": note.iloc[i],
            })
    detail = pd.DataFrame(rows)

    # ---- 월별 요약
    m = pd.DataFrame({"기준일": idx})
    m["연월"] = [f"{d:%Y-%m}" for d in idx]
    for b in mv.columns:
        m[BUCKET_KR.get(b, b)] = pnl[b].values
    m["합계"] = pnl.sum(axis=1).values
    monthly = m.drop(columns="기준일").groupby("연월", as_index=False).sum()
    nav_last = pd.DataFrame({"연월": m["연월"], "NAV": total_mv.values}).groupby(
        "연월", as_index=False).last()
    monthly = monthly.merge(nav_last, on="연월", how="left")
    monthly["누적손익"] = monthly["합계"].cumsum()

    # ---- 검증
    recon = pd.DataFrame(index=mv.index)
    recon.index.name = "기준일"
    recon["손익 산출 여부"] = computable
    recon["총평가액 증감"] = total_mv.diff().where(computable)
    recon["외부 자금이동"] = ext.where(computable, 0.0)
    recon["자산군 손익 합계"] = pnl.sum(axis=1)
    recon["차이(검증)"] = recon["총평가액 증감"] - recon["외부 자금이동"] - recon["자산군 손익 합계"]
    recon["일일수익률(%)"] = daily["일일수익률(%)"]
    recon["이상치(±3% 초과)"] = recon["일일수익률(%)"].abs() > 3
    recon["비고"] = note

    return {
        "daily": daily.reset_index(),
        "monthly": monthly,
        "detail": detail,
        "recon": recon.reset_index(),
        "mv": mv,
        "flow": flow,
        "pnl": pnl,
        "computable": computable,
    }


# ---------------------------------------------------------------------------
# 4) 종목별 손익
# ---------------------------------------------------------------------------

def _security_key(df: pd.DataFrame) -> pd.Series:
    key = df.get("ISIN", pd.Series("", index=df.index)).astype(str).str.strip()
    cusip = df.get("CUSIP", pd.Series("", index=df.index)).astype(str).str.strip()
    name = df.get("종목명", pd.Series("", index=df.index)).astype(str).str.strip()
    key = key.where(key != "", cusip)
    return key.where(key != "", name)


def _name_table(*frames: pd.DataFrame) -> pd.DataFrame:
    """종목키 -> 대표 종목명/코드/통화 (비어있지 않은 값 중 최빈값)."""
    parts = []
    for df in frames:
        if df is None or df.empty:
            continue
        cols = [c for c in ("종목키", "종목명", "종목코드", "발행통화") if c in df.columns]
        parts.append(df[cols])
    if not parts:
        return pd.DataFrame(columns=["종목키", "종목명", "종목코드", "발행통화"])
    allrows = pd.concat(parts, ignore_index=True, sort=False)
    for c in ("종목명", "종목코드", "발행통화"):
        if c not in allrows.columns:
            allrows[c] = ""
        allrows[c] = allrows[c].astype(str).str.strip()

    def best(s: pd.Series) -> str:
        s = s[(s != "") & (s.str.lower() != "nan")]
        # 스왑 파이낸싱 레그의 'FEDEF-1D' 같은 지수명은 종목코드로 부적절
        s = s[~s.str.upper().str.startswith(("FEDEF", "NONE"))]
        return s.mode().iloc[0] if len(s) else ""

    return (allrows.groupby("종목키", as_index=False)
            .agg(종목명=("종목명", best), 종목코드=("종목코드", best), 발행통화=("발행통화", best)))


def compute_security_pnl(positions: pd.DataFrame, activity: pd.DataFrame,
                         bucket: str, computable: Optional[pd.Series] = None) -> pd.DataFrame:
    """종목 단위 일일손익 = Δ평가액 + 현금흐름."""
    pos = positions[positions["버킷"] == bucket].copy()
    act = (activity[(activity["버킷"] == bucket) & activity["현금원장"]].copy()
           if not activity.empty else pd.DataFrame())
    if pos.empty and act.empty:
        return pd.DataFrame()
    if not pos.empty:
        pos["종목키"] = _security_key(pos)
    if not act.empty:
        act["종목키"] = _security_key(act)

    mv = (pos.groupby(["기준일", "종목키"], as_index=False)
          .agg(평가액=("평가액_USD", "sum"), 수량=("수량", "sum"))
          if not pos.empty else pd.DataFrame(columns=["기준일", "종목키", "평가액", "수량"]))
    fl = (act.groupby(["기준일", "종목키"], as_index=False)
          .agg(현금흐름=("정산금액_USD", "sum"), 매매수량=("수량", "sum"))
          if not act.empty else pd.DataFrame(columns=["기준일", "종목키", "현금흐름", "매매수량"]))

    dates = sorted(set(positions["기준일"].unique()))
    keys = sorted(set(mv["종목키"]) | set(fl["종목키"]))
    if not keys:
        return pd.DataFrame()
    grid = pd.MultiIndex.from_product([dates, keys], names=["기준일", "종목키"]).to_frame(index=False)

    out = (grid.merge(mv, on=["기준일", "종목키"], how="left")
               .merge(fl, on=["기준일", "종목키"], how="left"))
    for c in ("평가액", "현금흐름", "수량", "매매수량"):
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)
    out = out.merge(_name_table(pos, act), on="종목키", how="left")
    out = out.sort_values(["종목키", "기준일"])

    out["전일평가액"] = out.groupby("종목키")["평가액"].shift()
    out["일일손익"] = (out["평가액"] - out["전일평가액"].fillna(out["평가액"])) + out["현금흐름"]
    if computable is not None:
        bad = set(computable.index[~computable.astype(bool)])
        out.loc[out["기준일"].isin(bad), "일일손익"] = 0.0
    elif dates:
        out.loc[out["기준일"] == dates[0], "일일손익"] = 0.0
    out["누적손익"] = out.groupby("종목키")["일일손익"].cumsum()

    out = out[(out["평가액"] != 0) | (out["현금흐름"] != 0) | (out["일일손익"] != 0)]
    cols = ["기준일", "종목명", "종목코드", "종목키", "발행통화", "수량", "매매수량",
            "전일평가액", "평가액", "현금흐름", "일일손익", "누적손익"]
    return out[[c for c in cols if c in out.columns]].sort_values(["기준일", "종목명"])


def compute_swap_detail(swap_mtm: pd.DataFrame, swap_reset: pd.DataFrame,
                        computable: Optional[pd.Series] = None) -> pd.DataFrame:
    """EQSWAP36X(MTM) + EQSWAP18SX(리셋/언와인드)로 스왑 종목별 손익을 별도 산출(교차검증용)."""
    if swap_mtm.empty:
        return pd.DataFrame()
    df = swap_mtm.copy()
    df["기준일"] = df["_기준일"]
    df["종목명"] = df.get("Security Description", "")
    df["종목"] = df.get("Stock", "")
    df["스왑번호"] = df.get("Swap Number", "")
    df["레그"] = df.get("Leg Type", "")
    df["수량"] = to_num(df.get("Open Quantity", pd.Series(dtype=str)))
    df["MTM_USD"] = to_num(df.get("MTM Total Base", pd.Series(dtype=str)))
    df["평가손익(Performance)"] = to_num(df.get("Performance", pd.Series(dtype=str)))
    df["미실현금융비용"] = to_num(df.get("Pending Interest", pd.Series(dtype=str)))
    df["미실현배당"] = to_num(df.get("Pending Dividend", pd.Series(dtype=str)))
    df["명목금액"] = to_num(df.get("Mark Notional", pd.Series(dtype=str)))

    mtm = (df.groupby(["기준일", "스왑번호", "종목", "종목명"], as_index=False)
           .agg(수량=("수량", "sum"), MTM_USD=("MTM_USD", "sum"),
                명목금액=("명목금액", "sum"),
                미실현금융비용=("미실현금융비용", "sum"), 미실현배당=("미실현배당", "sum")))

    if not swap_reset.empty:
        rs = swap_reset.copy()
        rs["기준일"] = rs["_기준일"]
        rs["스왑번호"] = rs.get("Swap Number", "")
        rs["종목"] = rs.get("Stock", "")
        rs["실현손익"] = to_num(rs.get("Total", pd.Series(dtype=str)))
        rs["이벤트"] = rs.get("Event", "")
        real = (rs.groupby(["기준일", "스왑번호", "종목"], as_index=False)
                .agg(실현손익=("실현손익", "sum")))
    else:
        real = pd.DataFrame(columns=["기준일", "스왑번호", "종목", "실현손익"])

    out = mtm.merge(real, on=["기준일", "스왑번호", "종목"], how="outer")
    out[["MTM_USD", "실현손익", "수량", "명목금액"]] = out[["MTM_USD", "실현손익", "수량", "명목금액"]].fillna(0.0)
    out = out.sort_values(["스왑번호", "종목", "기준일"])
    out["전일MTM"] = out.groupby(["스왑번호", "종목"])["MTM_USD"].shift()
    dates = sorted(out["기준일"].dropna().unique())
    out["일일손익"] = (out["MTM_USD"] - out["전일MTM"].fillna(out["MTM_USD"])) + out["실현손익"]
    if computable is not None:
        bad = set(computable.index[~computable.astype(bool)])
        out.loc[out["기준일"].isin(bad), "일일손익"] = 0.0
    elif dates:
        out.loc[out["기준일"] == dates[0], "일일손익"] = 0.0
    out["누적손익"] = out.groupby(["스왑번호", "종목"])["일일손익"].cumsum()
    return out.sort_values(["기준일", "종목명"])


# ---------------------------------------------------------------------------
# 5) 현금잔고 / 결제
# ---------------------------------------------------------------------------

# CASH005X 컬럼 -> (구간명, 해당 결제일이 들어있는 컬럼)
# 'Prior Day' 컬럼이 리포트 기준일(D) 자체이고, 이후 컬럼이 D+1 .. D+4 이다.
CASH_FORECAST_COLS = [
    ("Prior Day Settlements", "D+0(기준일)", "Prior Day"),
    ("Current Day Settlements", "D+1", "Current Date"),
    ("Projected Settlement Day 1", "D+2", "Projected Settlement Date1"),
    ("Projected Settlement Day 2", "D+3", "Settlement Date2"),
    ("Projected Settlement Day 3", "D+4", "Settlement Date3"),
    ("Future Settlements", "이후전체", None),
]
FUTURE_BUCKETS = ["D+1", "D+2", "D+3", "D+4"]
STARTING_ROW = "Starting Balance"
ENDING_ROW = "Ending Balance"
TPB_ROW = "Total Projected Balance"


def build_cash_schedule(raw: pd.DataFrame) -> pd.DataFrame:
    """CASH005X를 long 형태(기준일/통화/카테고리/결제구간/결제일/금액)로 정리."""
    if raw.empty:
        return pd.DataFrame()
    df = raw.copy()
    df["기준일"] = df["_기준일"]
    df["통화"] = df.get("Settlement Currency", "")
    df["카테고리"] = df.get("Transaction Category", "").astype(str).str.strip()
    df["잔고유형"] = df.get("Balance Type", "")
    recs = []
    for src, bucket, datecol in CASH_FORECAST_COLS:
        if src not in df.columns:
            continue
        part = df[["기준일", "통화", "카테고리", "잔고유형"]].copy()
        part["결제구간"] = bucket
        part["결제일"] = to_date(df[datecol]) if datecol and datecol in df.columns else None
        part["금액"] = to_num(df[src])
        recs.append(part)
    if not recs:
        return pd.DataFrame()
    out = pd.concat(recs, ignore_index=True)
    order = {b: i for i, (_, b, _d) in enumerate(CASH_FORECAST_COLS)}
    out["_o"] = out["결제구간"].map(order)
    out["구분"] = "현금흐름"
    is_bal = (out["카테고리"].isin([STARTING_ROW, ENDING_ROW, TPB_ROW])
              | out["카테고리"].str.startswith("(")
              | out["카테고리"].str.contains("Net Projected Balance", na=False)
              | out["카테고리"].str.contains("Collateral Cash Balance", na=False))
    out.loc[is_bal, "구분"] = "잔고"
    out = out[out["금액"] != 0]
    cols = ["기준일", "통화", "구분", "카테고리", "잔고유형", "결제구간", "결제일", "금액"]
    return out.sort_values(["기준일", "통화", "_o", "구분", "카테고리"])[cols]


def build_cash_balance(positions: pd.DataFrame, cash_raw: pd.DataFrame) -> pd.DataFrame:
    """기준일 x 통화별: 현재 현금잔고 / 결제에 필요한 현금 / 결제 후 현금."""
    # (a) MAC001X 현금 포지션: 매매기준(현재) / 결제기준(결제완료) 잔고
    pos_cash = positions[positions["상품유형"].astype(str).str.upper() == "CASH"].copy()
    if not pos_cash.empty:
        base = (pos_cash.groupby(["기준일", "발행통화"], as_index=False)
                .agg(**{"현재잔고(매매기준)": ("수량", "sum"),
                        "현재잔고(결제기준)": ("SD잔고_발행통화", "sum"),
                        "현재잔고_USD": ("평가액_USD", "sum"),
                        "현재잔고(결제기준)_USD": ("SD잔고_USD", "sum")})
                .rename(columns={"발행통화": "통화"}))
    else:
        base = pd.DataFrame(columns=["기준일", "통화", "현재잔고(매매기준)", "현재잔고(결제기준)",
                                     "현재잔고_USD", "현재잔고(결제기준)_USD"])

    # (b) CASH005X: 향후 결제 스케줄
    if cash_raw is not None and not cash_raw.empty:
        cf = cash_raw.copy()
        cf["기준일"] = cf["_기준일"]
        cf["통화"] = cf.get("Settlement Currency", "")
        cf["카테고리"] = cf.get("Transaction Category", "").astype(str).str.strip()
        cf["잔고유형"] = cf.get("Balance Type", "").astype(str).str.strip()
        cf = cf[cf["잔고유형"].isin(["", "PB"])].copy()

        for src, bucket, _dc in CASH_FORECAST_COLS:
            cf[f"금액_{bucket}"] = to_num(cf[src]) if src in cf.columns else 0.0

        is_balance_row = (cf["카테고리"].isin([STARTING_ROW, ENDING_ROW, TPB_ROW])
                          | cf["카테고리"].str.startswith("(")
                          | cf["카테고리"].str.contains("Net Projected Balance", na=False))
        flows = cf[~is_balance_row].copy()
        fut = [f"금액_{b}" for b in FUTURE_BUCKETS]
        flows["_순액"] = flows[fut].sum(axis=1)
        flows["_유출"] = flows[fut].clip(upper=0).sum(axis=1)
        flows["_유입"] = flows[fut].clip(lower=0).sum(axis=1)
        flows["_이후"] = flows["금액_이후전체"]
        sched = (flows.groupby(["기준일", "통화"], as_index=False)
                 .agg(**{"결제예정 수취(D+1~D+4)": ("_유입", "sum"),
                         "결제필요현금(D+1~D+4)": ("_유출", "sum"),
                         "결제예정 순액(D+1~D+4)": ("_순액", "sum"),
                         "미도래 결제 순액(D+5 이후)": ("_이후", "sum")}))

        ends = cf[cf["카테고리"] == ENDING_ROW]
        if not ends.empty:
            endg = (ends.groupby(["기준일", "통화"], as_index=False)
                    .agg(**{"기준일 결제후잔고": ("금액_D+0(기준일)", "sum"),
                            "D+1 예상잔고": ("금액_D+1", "sum"),
                            "D+2 예상잔고": ("금액_D+2", "sum"),
                            "D+3 예상잔고": ("금액_D+3", "sum"),
                            "D+4 예상잔고": ("금액_D+4", "sum"),
                            "전체 결제후 잔고": ("금액_이후전체", "sum")}))
            sched = sched.merge(endg, on=["기준일", "통화"], how="outer")

        base = base.merge(sched, on=["기준일", "통화"], how="outer")

    # USD 환산율(발행통화 -> USD, 나누기 기준)
    if not pos_cash.empty:
        fx = (pos_cash.groupby(["기준일", "발행통화"], as_index=False)["FX_USD환산"].max()
              .rename(columns={"발행통화": "통화", "FX_USD환산": "USD환산율(나누기)"}))
        base = base.merge(fx, on=["기준일", "통화"], how="left")
        rate = base["USD환산율(나누기)"].replace(0, pd.NA)
        for col in ["결제필요현금(D+1~D+4)", "결제예정 순액(D+1~D+4)", "전체 결제후 잔고"]:
            if col in base.columns:
                base[f"{col}_USD"] = (base[col] / rate).astype(float)

    order = ["기준일", "통화",
             "현재잔고(매매기준)", "현재잔고(결제기준)", "현재잔고_USD", "현재잔고(결제기준)_USD",
             "결제예정 수취(D+1~D+4)", "결제필요현금(D+1~D+4)", "결제예정 순액(D+1~D+4)",
             "미도래 결제 순액(D+5 이후)",
             "기준일 결제후잔고", "D+1 예상잔고", "D+2 예상잔고", "D+3 예상잔고", "D+4 예상잔고",
             "전체 결제후 잔고", "USD환산율(나누기)",
             "결제필요현금(D+1~D+4)_USD", "결제예정 순액(D+1~D+4)_USD", "전체 결제후 잔고_USD"]
    base = base[[c for c in order if c in base.columns]]
    return base.sort_values(["기준일", "통화"]).reset_index(drop=True)


# ---------------------------------------------------------------------------
# 6) 엑셀 출력
# ---------------------------------------------------------------------------

NUM_FMT = "#,##0.00"
INT_FMT = "#,##0"
PCT_FMT = "0.00%"


def _autofit(ws, df: pd.DataFrame, wb, start_row: int = 0):
    money_cols = [c for c in df.columns if any(
        k in str(c) for k in ("손익", "평가액", "금액", "잔고", "현금", "USD", "NAV", "차이", "증감", "이동", "MTM", "명목"))]
    fmt_money = wb.add_format({"num_format": NUM_FMT})
    fmt_pct = wb.add_format({"num_format": "0.00"})
    for i, col in enumerate(df.columns):
        try:
            width = max(len(str(col)) + 2, int(df[col].astype(str).str.len().quantile(0.95)) + 2)
        except Exception:
            width = len(str(col)) + 2
        width = min(max(width, 10), 42)
        if col in money_cols:
            ws.set_column(i, i, width, fmt_money)
        elif "%" in str(col):
            ws.set_column(i, i, width, fmt_pct)
        else:
            ws.set_column(i, i, width)
    ws.freeze_panes(start_row + 1, 0)
    if len(df):
        ws.autofilter(start_row, 0, start_row + len(df), len(df.columns) - 1)


def _add_summary_chart(wb, ws_summary, daily: pd.DataFrame, sheet_name: str, anchor: str):
    """요약 시트에 누적손익/NAV 추이 차트를 추가한다."""
    if daily is None or len(daily) < 2:
        return
    cols = list(daily.columns)
    try:
        c_date = cols.index("기준일")
        c_cum = cols.index("합계 누적손익")
        c_nav = cols.index("총평가액(NAV)")
    except ValueError:
        return
    n = len(daily)
    chart = wb.add_chart({"type": "line"})
    chart.add_series({
        "name": "합계 누적손익 (USD)",
        "categories": [sheet_name, 1, c_date, n, c_date],
        "values": [sheet_name, 1, c_cum, n, c_cum],
        "line": {"color": "#1F3864", "width": 2.0},
    })
    chart2 = wb.add_chart({"type": "line"})
    chart2.add_series({
        "name": "총평가액 NAV (USD)",
        "categories": [sheet_name, 1, c_date, n, c_date],
        "values": [sheet_name, 1, c_nav, n, c_nav],
        "line": {"color": "#C00000", "width": 1.25, "dash_type": "dash"},
        "y2_axis": True,
    })
    chart.combine(chart2)
    chart.set_title({"name": "누적손익 및 NAV 추이"})
    chart.set_x_axis({"num_font": {"rotation": -45}})
    chart.set_y_axis({"name": "누적손익 (USD)", "num_format": "#,##0"})
    chart2.set_y2_axis({"name": "NAV (USD)", "num_format": "#,##0"})
    chart.set_size({"width": 900, "height": 380})
    chart.set_legend({"position": "bottom"})
    ws_summary.insert_chart(anchor, chart)


def write_excel(path: str, sheets: List[Tuple[str, pd.DataFrame]], meta: List[Tuple[str, str]]):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd",
                        date_format="yyyy-mm-dd") as xw:
        wb = xw.book
        hdr = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white",
                             "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        title = wb.add_format({"bold": True, "font_size": 14})
        label = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        value = wb.add_format({"border": 1})

        # 요약 시트
        ws = wb.add_worksheet("00_요약")
        xw.sheets["00_요약"] = ws
        ws.set_column(0, 0, 34)
        ws.set_column(1, 1, 60)
        ws.write(0, 0, "Prelude 통합 손익 · 현금 리포트", title)
        for r, (k, v) in enumerate(meta, start=2):
            ws.write(r, 0, k, label)
            ws.write(r, 1, v, value)

        for name, df in sheets:
            if df is None or len(df) == 0:
                ws2 = wb.add_worksheet(name[:31])
                xw.sheets[name[:31]] = ws2
                ws2.write(0, 0, "해당 데이터 없음")
                continue
            df = df.copy()
            for c in df.columns:
                if df[c].dtype == object:
                    df[c] = df[c].apply(lambda v: v if not isinstance(v, dt.date) or isinstance(v, dt.datetime) else v)
            df.to_excel(xw, sheet_name=name[:31], index=False, startrow=0)
            ws2 = xw.sheets[name[:31]]
            for i, col in enumerate(df.columns):
                ws2.write(0, i, str(col), hdr)
            _autofit(ws2, df, wb)
            if name.endswith("일일손익"):
                _add_summary_chart(wb, ws, df, name[:31], f"A{len(meta) + 5}")


# ---------------------------------------------------------------------------
# 7) 파일 정리(리포트별 폴더 이동)
# ---------------------------------------------------------------------------

def safe_folder_name(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", name).strip().rstrip(".")
    return name[:120] or "UNKNOWN"


def organize_files(files: Sequence[SourceFile], root: str, layout: str = "report",
                   mode: str = "move", dry_run: bool = False,
                   log_path: Optional[str] = None) -> pd.DataFrame:
    """
    리포트 종류별 폴더로 원본 파일을 분류한다.

    layout:
      report        -> <root>/<CODE - Report Name>/파일
      report-year   -> <root>/<CODE - Report Name>/<YYYY>/파일
      report-month  -> <root>/<CODE - Report Name>/<YYYY-MM>/파일
      date-report   -> <root>/<YYYY-MM-DD>/<CODE - Report Name>/파일
    """
    records = []
    op = shutil.move if mode == "move" else shutil.copy2
    for f in files:
        folder = safe_folder_name(f.label if f.label else "기타")
        if layout == "report":
            rel = folder
        elif layout == "report-year":
            rel = os.path.join(folder, f"{f.date.year}" if f.date else "날짜미상")
        elif layout == "report-month":
            rel = os.path.join(folder, f.date.strftime("%Y-%m") if f.date else "날짜미상")
        elif layout == "date-report":
            rel = os.path.join(f.date.isoformat() if f.date else "날짜미상", folder)
        else:
            raise ValueError(f"알 수 없는 layout: {layout}")

        dest_dir = os.path.join(root, rel)
        dest = os.path.join(dest_dir, f.filename)

        if os.path.abspath(dest) == os.path.abspath(f.path):
            status = "이미정리됨"
        elif os.path.exists(dest):
            status = "건너뜀(동일파일존재)"
        else:
            status = "예정" if dry_run else "완료"
            if not dry_run:
                try:
                    os.makedirs(long_path(dest_dir), exist_ok=True)
                    op(long_path(f.path), long_path(dest))
                except Exception as exc:  # pragma: no cover
                    status = f"실패: {exc}"

        records.append({"리포트": f.label, "코드": f.code, "기준일": f.date,
                        "파일명": f.filename, "이전경로": f.path,
                        "이동경로": dest, "처리": status})

    df = pd.DataFrame(records)
    if log_path and not dry_run and len(df):
        try:
            df.to_csv(log_path, index=False, encoding="utf-8-sig")
        except Exception:
            pass
    return df


# ---------------------------------------------------------------------------
# 메인
# ---------------------------------------------------------------------------

def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Prelude 리포트 → 통합 손익/현금 엑셀 생성 및 리포트별 폴더 정리",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--src", required=True, help="리포트 CSV가 있는 폴더 (하위 폴더 포함 스캔)")
    ap.add_argument("--out", default=None,
                    help="산출 엑셀 경로 (기본: <src>/_output/Prelude_PnL_<날짜범위>.xlsx)")
    ap.add_argument("--from-date", default=None, help="시작 기준일 YYYY-MM-DD")
    ap.add_argument("--to-date", default=None, help="종료 기준일 YYYY-MM-DD")
    ap.add_argument("--external-category", action="append", default=None,
                    help=f"외부 자금이동(입출금)으로 간주할 거래 카테고리. 여러 번 지정 가능 "
                         f"(기본: {', '.join(DEFAULT_EXTERNAL_CATEGORIES)})")
    ap.add_argument("--no-external", action="store_true",
                    help="외부 자금이동 분류를 사용하지 않음(모든 현금흐름을 손익에 포함)")
    ap.add_argument("--max-gap-days", type=int, default=DEFAULT_MAX_GAP_DAYS,
                    help=f"직전 리포트일과의 간격이 이 일수를 넘으면 손익 산출에서 제외 "
                         f"(기본: {DEFAULT_MAX_GAP_DAYS})")
    ap.add_argument("--organize", action="store_true", help="원본 파일을 리포트별 폴더로 정리")
    ap.add_argument("--layout", default="report",
                    choices=["report", "report-year", "report-month", "date-report"],
                    help="정리 폴더 구조 (기본: report)")
    ap.add_argument("--copy", action="store_true", help="이동 대신 복사")
    ap.add_argument("--dry-run", action="store_true", help="정리 시 실제로 옮기지 않고 계획만 출력")
    ap.add_argument("--no-excel", action="store_true", help="엑셀 생성 없이 정리만 수행")
    args = ap.parse_args(argv)

    if args.no_external:
        args.external_category = []
    elif args.external_category is None:
        args.external_category = list(DEFAULT_EXTERNAL_CATEGORIES)

    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    src = os.path.abspath(args.src)
    if not os.path.isdir(src):
        raise SystemExit(f"폴더를 찾을 수 없습니다: {src}")

    print(f"[1/6] 파일 스캔: {src}")
    files = scan_files(src)
    print(f"      총 {len(files):,}개 파일 / 리포트 종류 {len(set(f.label for f in files))}종")
    if not files:
        raise SystemExit("리포트 파일이 없습니다.")

    organize_df = pd.DataFrame()
    excel_path = None

    if not args.no_excel:
        by_code = index_by_code(files)

        print("[2/6] 리포트 로드")
        raw_pos = load_code(by_code, REPORT_CODES["positions"])
        raw_act = load_code(by_code, REPORT_CODES["activity"])
        raw_cash = load_code(by_code, REPORT_CODES["cash_forecast"])
        raw_cashd = load_code(by_code, REPORT_CODES["cash_detail"])
        raw_swap = load_code(by_code, REPORT_CODES["swap_mtm"])
        raw_reset = load_code(by_code, REPORT_CODES["swap_reset"])
        raw_int = load_code(by_code, REPORT_CODES["interest"])

        print("[3/6] 정규화")
        positions = build_positions(raw_pos)
        activity = build_activity(raw_act)

        d0 = dt.date.fromisoformat(args.from_date) if args.from_date else None
        d1 = dt.date.fromisoformat(args.to_date) if args.to_date else None

        def clip(df, col="기준일"):
            if df is None or df.empty or col not in df.columns:
                return df
            if d0:
                df = df[df[col] >= d0]
            if d1:
                df = df[df[col] <= d1]
            return df

        def clip_raw(df):
            if df is None or df.empty:
                return df
            if d0:
                df = df[df["_기준일"] >= d0]
            if d1:
                df = df[df["_기준일"] <= d1]
            return df

        positions = clip(positions)
        activity = clip(activity)
        raw_cash = clip_raw(raw_cash)
        raw_cashd = clip_raw(raw_cashd)
        raw_swap = clip_raw(raw_swap)
        raw_reset = clip_raw(raw_reset)
        raw_int = clip_raw(raw_int)

        print("[4/6] 손익 계산")
        res = compute_pnl(positions, activity,
                          external_categories=args.external_category,
                          max_gap_days=args.max_gap_days)
        computable = res["computable"]

        swap_sec = compute_security_pnl(positions, activity, "Swap", computable)
        cash_sec = compute_security_pnl(positions, activity, "Cash Equity", computable)
        fx_sec = compute_security_pnl(positions, activity, "FX", computable)
        swap_detail = compute_swap_detail(raw_swap, raw_reset, computable)

        print("[5/6] 현금/결제 정리")
        cash_bal = build_cash_balance(positions, raw_cash)
        cash_sched = build_cash_schedule(raw_cash)

        trades = pd.DataFrame()
        if not activity.empty:
            tcols = ["기준일", "입력일", "매매일", "결제일", "구분", "대분류", "중분류", "소분류",
                     "버킷", "포지션유형", "상품유형", "종목명", "종목코드", "ISIN", "매매구분",
                     "수량", "단가_USD", "약정금액_USD", "수수료_USD", "세금_USD", "이자_USD",
                     "정산금액_USD", "결제통화", "정산금액_결제통화", "현금원장", "_원본파일"]
            trades = activity[[c for c in tcols if c in activity.columns]].copy()
            trades["자산군"] = trades["버킷"].map(BUCKET_KR).fillna(trades["버킷"])
            trades = trades.sort_values(["기준일", "자산군", "종목명"])

        pos_snapshot = pd.DataFrame()
        if not positions.empty:
            last_d = positions["기준일"].max()
            pcols = ["기준일", "버킷", "자산군_원본", "서브계좌", "종목명", "종목코드", "ISIN",
                     "상품유형", "포지션유형", "발행통화", "수량", "결제수량",
                     "평가액_USD", "원가_USD", "SD잔고_발행통화", "SD잔고_USD"]
            pos_snapshot = positions[positions["기준일"] == last_d][
                [c for c in pcols if c in positions.columns]].copy()
            pos_snapshot["자산군"] = pos_snapshot["버킷"].map(BUCKET_KR).fillna(pos_snapshot["버킷"])

        inventory = pd.DataFrame([{
            "리포트": f.label, "코드": f.code, "계좌": f.account,
            "기준일": f.date, "파일명": f.filename,
            "크기(KB)": round(f.size / 1024, 1),
            "경로": os.path.relpath(f.path, src),
        } for f in files]).sort_values(["리포트", "기준일"])

        interest = pd.DataFrame()
        if not raw_int.empty:
            it = raw_int.copy()
            interest = pd.DataFrame({
                "기준일(파일)": it["_기준일"],
                "계좌": it.get("Account Number", ""),
                "통화": it.get("Currency", ""),
                "이자기산일": to_date(it.get("Value Date", pd.Series(dtype=str))),
                "차변잔액": to_num(it.get("Net Debit Balance", pd.Series(dtype=str))),
                "차변이율": to_num(it.get("Debit Rate", pd.Series(dtype=str))),
                "차변이자": to_num(it.get("Debit Interest", pd.Series(dtype=str))),
                "대변잔액": to_num(it.get("Net Credit Balance", pd.Series(dtype=str))),
                "대변이율": to_num(it.get("Credit Rate", pd.Series(dtype=str))),
                "대변이자": to_num(it.get("Credit Interest", pd.Series(dtype=str))),
            })
            # SW1003MX는 MTD 누적이라 파일 간 중복이 발생한다. 계좌/통화/기산일 기준 최신본만 남긴다.
            interest = (interest.sort_values("기준일(파일)")
                        .drop_duplicates(subset=["계좌", "통화", "이자기산일"], keep="last")
                        .sort_values(["통화", "계좌", "이자기산일"]))

        dates = sorted(positions["기준일"].unique())
        tag = f"{dates[0]:%Y%m%d}_{dates[-1]:%Y%m%d}" if dates else "all"
        excel_path = args.out or os.path.join(src, "_output", f"Prelude_PnL_{tag}.xlsx")

        daily = res["daily"]
        last = daily.iloc[-1] if len(daily) else None
        def first_val(col: str) -> str:
            if col in positions.columns and len(positions):
                s = positions[col].astype(str).replace("", pd.NA).dropna()
                if len(s):
                    return str(s.iloc[0])
            return ""

        acct = " / ".join(x for x in [first_val("Main Account Number"), first_val("Main Account Name")] if x)
        excluded = [str(d) for d in computable.index[~computable.astype(bool)]]
        cash_last = cash_bal[cash_bal["기준일"] == cash_bal["기준일"].max()] if len(cash_bal) else pd.DataFrame()
        cash_txt = " | ".join(
            f"{r['통화']} 현재 {r.get('현재잔고(매매기준)', 0):,.0f}"
            f" / 결제필요 {r.get('결제필요현금(D+1~D+4)', 0):,.0f}"
            f" / 결제후 {r.get('전체 결제후 잔고', 0):,.0f}"
            for _, r in cash_last.iterrows()) if len(cash_last) else "-"

        meta = [
            ("생성일시", dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("원본 폴더", src),
            ("계좌", acct),
            ("기준일 범위", f"{dates[0]} ~ {dates[-1]} ({len(dates)}개 기준일)" if dates else "-"),
            ("사용 리포트", ", ".join(sorted({f.code for f in files if f.code}))),
            ("기준통화", "USD (Client Base CCY)"),
            ("총평가액(NAV, 최종일)", f"{last['총평가액(NAV)']:,.2f}" if last is not None else "-"),
            ("누적손익(산출구간 합계)", f"{last['합계 누적손익']:,.2f}" if last is not None else "-"),
            ("최종일 일일손익", f"{last['합계 일일손익']:,.2f}" if last is not None else "-"),
            ("최종일 현금(통화별)", cash_txt),
            ("손익 산식", "일일손익 = Δ평가액 + 귀속 현금흐름 (현금 자산군은 잔여항)"),
            ("검증", "Σ자산군 손익 = Δ총평가액 − 외부 자금이동 (09_검증 시트 참조)"),
            ("외부 자금이동 분류", ", ".join(args.external_category) or "(없음)"),
            (f"손익 산출 제외일(간격>{args.max_gap_days}일)",
             ", ".join(excluded) if excluded else "없음"),
        ]

        sheets = [
            ("01_일일손익", daily),
            ("02_월별손익", res["monthly"]),
            ("03_자산군별상세", res["detail"]),
            ("04_Swap손익", swap_sec),
            ("05_현물주식손익", cash_sec),
            ("06_FX손익", fx_sec),
            ("07_현금잔고", cash_bal),
            ("08_결제스케줄", cash_sched),
            ("09_거래내역", trades),
            ("10_검증", res["recon"]),
            ("11_Swap상세(MTM)", swap_detail),
            ("12_포지션(최종일)", pos_snapshot),
            ("13_이자내역", interest),
            ("14_파일목록", inventory),
        ]

        print(f"[6/6] 엑셀 작성: {excel_path}")
        write_excel(excel_path, sheets, meta)
        print(f"      완료 → {excel_path}")

    if args.organize:
        print(f"\n[정리] 리포트별 폴더 분류 ({'복사' if args.copy else '이동'}"
              f"{', DRY-RUN' if args.dry_run else ''}, layout={args.layout})")
        skip = os.path.abspath(os.path.join(src, "_output"))
        targets = [f for f in files if not os.path.abspath(f.path).startswith(skip)]
        log = os.path.join(src, "_output", "_organize_log.csv")
        os.makedirs(os.path.dirname(log), exist_ok=True)
        organize_df = organize_files(targets, src, layout=args.layout,
                                     mode="copy" if args.copy else "move",
                                     dry_run=args.dry_run, log_path=log)
        summary = organize_df.groupby(["리포트", "처리"]).size().reset_index(name="건수")
        for _, r in summary.iterrows():
            print(f"      {r['처리']:<20} {r['건수']:>5}건  {r['리포트']}")
        print(f"      총 {len(organize_df):,}건" + ("" if args.dry_run else f" / 로그: {log}"))

    return 0


if __name__ == "__main__":
    sys.exit(main())
