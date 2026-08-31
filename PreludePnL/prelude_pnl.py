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

# 외부 자금이동(입출금)으로 볼 카테고리 - 기본값 없음. --external-category 로 추가.
DEFAULT_EXTERNAL_CATEGORIES: Tuple[str, ...] = ()

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
    with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
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
                external_categories: Sequence[str] = ()) -> Dict[str, pd.DataFrame]:
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

    d_mv = mv.diff()
    d_mv.iloc[0] = 0.0  # 첫 기준일은 직전 평가액이 없으므로 손익 산출 제외

    pnl = pd.DataFrame(index=mv.index, columns=mv.columns, dtype=float)
    other_buckets = [c for c in mv.columns if c != "Cash"]
    for b in other_buckets:
        pnl[b] = d_mv[b] + flow[b]
    if "Cash" in mv.columns:
        pnl["Cash"] = d_mv["Cash"] - flow[other_buckets].sum(axis=1) - ext
    pnl.iloc[0] = 0.0

    total_mv = mv.sum(axis=1)
    daily = pd.DataFrame(index=mv.index)
    daily.index.name = "기준일"
    for b in mv.columns:
        daily[f"{BUCKET_KR.get(b, b)} 일일손익"] = pnl[b]
    daily["합계 일일손익"] = pnl.sum(axis=1)
    for b in mv.columns:
        daily[f"{BUCKET_KR.get(b, b)} 누적손익"] = pnl[b].cumsum()
    daily["합계 누적손익"] = pnl.sum(axis=1).cumsum()
    daily["총평가액(NAV)"] = total_mv
    daily["전일 총평가액"] = total_mv.shift()
    daily["외부 자금이동"] = ext
    daily["일일수익률(%)"] = (daily["합계 일일손익"] / daily["전일 총평가액"].replace(0, pd.NA) * 100).astype(float)

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
            })
    detail = pd.DataFrame(rows)

    # ---- 검증
    recon = pd.DataFrame(index=mv.index)
    recon.index.name = "기준일"
    recon["총평가액 증감"] = total_mv.diff()
    recon["외부 자금이동"] = ext
    recon["자산군 손익 합계"] = pnl.sum(axis=1)
    recon["차이(검증)"] = recon["총평가액 증감"] - recon["외부 자금이동"] - recon["자산군 손익 합계"]
    recon["일일수익률(%)"] = daily["일일수익률(%)"]
    recon["이상치(±3% 초과)"] = recon["일일수익률(%)"].abs() > 3

    return {
        "daily": daily.reset_index(),
        "detail": detail,
        "recon": recon.reset_index(),
        "mv": mv,
        "flow": flow,
        "pnl": pnl,
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


def compute_security_pnl(positions: pd.DataFrame, activity: pd.DataFrame,
                         bucket: str) -> pd.DataFrame:
    """종목 단위 일일손익 = Δ평가액 + 현금흐름."""
    pos = positions[positions["버킷"] == bucket].copy()
    if pos.empty:
        return pd.DataFrame()
    pos["종목키"] = _security_key(pos)

    mv = (pos.groupby(["기준일", "종목키"], as_index=False)
          .agg(평가액=("평가액_USD", "sum"), 수량=("수량", "sum"),
               종목명=("종목명", "first"), 종목코드=("종목코드", "first"),
               발행통화=("발행통화", "first")))

    if not activity.empty:
        act = activity[(activity["버킷"] == bucket) & activity["현금원장"]].copy()
    else:
        act = pd.DataFrame()
    if not act.empty:
        act["종목키"] = _security_key(act)
        fl = (act.groupby(["기준일", "종목키"], as_index=False)
              .agg(현금흐름=("정산금액_USD", "sum"), 매매수량=("수량", "sum")))
    else:
        fl = pd.DataFrame(columns=["기준일", "종목키", "현금흐름", "매매수량"])

    dates = sorted(set(positions["기준일"].unique()))
    keys = sorted(set(mv["종목키"]) | set(fl["종목키"]))
    grid = pd.MultiIndex.from_product([dates, keys], names=["기준일", "종목키"]).to_frame(index=False)

    out = (grid.merge(mv, on=["기준일", "종목키"], how="left")
               .merge(fl, on=["기준일", "종목키"], how="left"))
    out[["평가액", "현금흐름", "수량", "매매수량"]] = out[["평가액", "현금흐름", "수량", "매매수량"]].fillna(0.0)
    out = out.sort_values(["종목키", "기준일"])
    names = (mv.dropna(subset=["종목명"]).groupby("종목키")
             .agg(종목명=("종목명", "last"), 종목코드=("종목코드", "last"), 발행통화=("발행통화", "last")))
    out = out.drop(columns=["종목명", "종목코드", "발행통화"]).merge(
        names, on="종목키", how="left")
    out = out.sort_values(["종목키", "기준일"])

    out["전일평가액"] = out.groupby("종목키")["평가액"].shift()
    first_date = dates[0] if dates else None
    out["일일손익"] = (out["평가액"] - out["전일평가액"].fillna(out["평가액"])) + out["현금흐름"]
    out.loc[out["기준일"] == first_date, "일일손익"] = 0.0
    out["누적손익"] = out.groupby("종목키")["일일손익"].cumsum()

    out = out[(out["평가액"] != 0) | (out["현금흐름"] != 0) | (out["일일손익"] != 0)]
    cols = ["기준일", "종목명", "종목코드", "종목키", "발행통화", "수량", "매매수량",
            "전일평가액", "평가액", "현금흐름", "일일손익", "누적손익"]
    return out[[c for c in cols if c in out.columns]].sort_values(["기준일", "종목명"])


def compute_swap_detail(swap_mtm: pd.DataFrame, swap_reset: pd.DataFrame) -> pd.DataFrame:
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
    if dates:
        out.loc[out["기준일"] == dates[0], "일일손익"] = 0.0
    out["누적손익"] = out.groupby(["스왑번호", "종목"])["일일손익"].cumsum()
    return out.sort_values(["기준일", "종목명"])


# ---------------------------------------------------------------------------
# 5) 현금잔고 / 결제
# ---------------------------------------------------------------------------

CASH_FORECAST_COLS = [
    ("Prior Day Settlements", "전일"),
    ("Current Day Settlements", "당일"),
    ("Projected Settlement Day 1", "D+1"),
    ("Projected Settlement Day 2", "D+2"),
    ("Projected Settlement Day 3", "D+3"),
    ("Future Settlements", "향후(전체)"),
]
STARTING_ROW = "Starting Balance"
ENDING_ROW = "Ending Balance"
TPB_ROW = "Total Projected Balance"


def build_cash_schedule(raw: pd.DataFrame) -> pd.DataFrame:
    """CASH005X를 long 형태(기준일/통화/카테고리/결제일버킷/금액)로 정리."""
    if raw.empty:
        return pd.DataFrame()
    df = raw.copy()
    df["기준일"] = df["_기준일"]
    df["통화"] = df.get("Settlement Currency", "")
    df["카테고리"] = df.get("Transaction Category", "").astype(str).str.strip()
    df["잔고유형"] = df.get("Balance Type", "")
    date_cols = {
        "전일": "Prior Day", "당일": "Current Date", "D+1": "Projected Settlement Date1",
        "D+2": "Settlement Date2", "D+3": "Settlement Date3", "향후(전체)": "Settlement Date3",
    }
    recs = []
    for src, bucket in CASH_FORECAST_COLS:
        if src not in df.columns:
            continue
        part = df[["기준일", "통화", "카테고리", "잔고유형"]].copy()
        part["결제구간"] = bucket
        part["결제일"] = to_date(df[date_cols[bucket]]) if date_cols[bucket] in df.columns else None
        part["금액"] = to_num(df[src])
        recs.append(part)
    if not recs:
        return pd.DataFrame()
    out = pd.concat(recs, ignore_index=True)
    order = {b: i for i, (_, b) in enumerate(CASH_FORECAST_COLS)}
    out["_o"] = out["결제구간"].map(order)
    return out.sort_values(["기준일", "통화", "_o", "카테고리"]).drop(columns="_o")


def build_cash_balance(positions: pd.DataFrame, cash_raw: pd.DataFrame) -> pd.DataFrame:
    """기준일 x 통화별 현재잔고 / 결제필요현금 / 결제후잔고."""
    rows = []

    # (a) MAC001X 현금 포지션: 매매기준 / 결제기준 잔고
    pos_cash = positions[positions["상품유형"].astype(str).str.upper() == "CASH"].copy()
    if not pos_cash.empty:
        g = (pos_cash.groupby(["기준일", "발행통화"], as_index=False)
             .agg(매매기준잔고=("수량", "sum"),
                  결제기준잔고=("SD잔고_발행통화", "sum"),
                  매매기준잔고_USD=("평가액_USD", "sum"),
                  결제기준잔고_USD=("SD잔고_USD", "sum")))
        rows.append(g.rename(columns={"발행통화": "통화"}))

    base = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(
        columns=["기준일", "통화", "매매기준잔고", "결제기준잔고", "매매기준잔고_USD", "결제기준잔고_USD"])

    # (b) CASH005X: 결제 스케줄
    if not cash_raw.empty:
        cf = cash_raw.copy()
        cf["기준일"] = cf["_기준일"]
        cf["통화"] = cf.get("Settlement Currency", "")
        cf["카테고리"] = cf.get("Transaction Category", "").astype(str).str.strip()
        cf["잔고유형"] = cf.get("Balance Type", "").astype(str).str.strip()
        cf = cf[cf["잔고유형"].isin(["", "PB"])]

        amt = {b: to_num(cf[src]) if src in cf.columns else 0.0
               for src, b in CASH_FORECAST_COLS}
        for b, v in amt.items():
            cf[f"금액_{b}"] = v

        flows = cf[~cf["카테고리"].isin([STARTING_ROW, ENDING_ROW, TPB_ROW])
                   & ~cf["카테고리"].str.startswith("(")
                   & ~cf["카테고리"].str.contains("Net Projected Balance", na=False)]
        fut_cols = ["금액_당일", "금액_D+1", "금액_D+2", "금액_D+3"]
        flows = flows.assign(_순액=flows[fut_cols].sum(axis=1),
                             _유출=flows[fut_cols].clip(upper=0).sum(axis=1),
                             _유입=flows[fut_cols].clip(lower=0).sum(axis=1),
                             _향후=flows["금액_향후(전체)"])
        sched = (flows.groupby(["기준일", "통화"], as_index=False)
                 .agg(결제예정_순액=("_순액", "sum"),
                      결제필요현금_유출=("_유출", "sum"),
                      결제예정_유입=("_유입", "sum"),
                      향후결제_순액=("_향후", "sum")))

        ends = cf[cf["카테고리"] == ENDING_ROW]
        if not ends.empty:
            endg = (ends.groupby(["기준일", "통화"], as_index=False)
                    .agg(당일말잔고=("금액_전일", "sum"),
                         D1말잔고=("금액_당일", "sum"),
                         D2말잔고=("금액_D+1", "sum"),
                         D3말잔고=("금액_D+2", "sum"),
                         결제후잔고=("금액_향후(전체)", "sum")))
            sched = sched.merge(endg, on=["기준일", "통화"], how="outer")

        base = base.merge(sched, on=["기준일", "통화"], how="outer")

    # USD 환산율(발행통화 -> USD)
    if not positions.empty:
        fx = (positions[positions["상품유형"].astype(str).str.upper() == "CASH"]
              .groupby(["기준일", "발행통화"], as_index=False)["FX_USD환산"].max()
              .rename(columns={"발행통화": "통화", "FX_USD환산": "USD환산율(나누기)"}))
        base = base.merge(fx, on=["기준일", "통화"], how="left")
        rate = base["USD환산율(나누기)"].replace(0, pd.NA)
        for col, new in [("결제필요현금_유출", "결제필요현금_유출_USD"),
                         ("결제예정_순액", "결제예정_순액_USD"),
                         ("결제후잔고", "결제후잔고_USD")]:
            if col in base.columns:
                base[new] = (base[col] / rate).astype(float)

    order = ["기준일", "통화",
             "매매기준잔고", "결제기준잔고", "매매기준잔고_USD", "결제기준잔고_USD",
             "결제예정_유입", "결제필요현금_유출", "결제예정_순액", "향후결제_순액",
             "당일말잔고", "D1말잔고", "D2말잔고", "D3말잔고", "결제후잔고",
             "USD환산율(나누기)", "결제필요현금_유출_USD", "결제예정_순액_USD", "결제후잔고_USD"]
    base = base[[c for c in order if c in base.columns]]
    return base.sort_values(["기준일", "통화"])


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
                os.makedirs(dest_dir, exist_ok=True)
                try:
                    op(f.path, dest)
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
    ap.add_argument("--external-category", action="append", default=list(DEFAULT_EXTERNAL_CATEGORIES),
                    help="외부 자금이동(입출금)으로 간주할 거래 카테고리. 여러 번 지정 가능")
    ap.add_argument("--organize", action="store_true", help="원본 파일을 리포트별 폴더로 정리")
    ap.add_argument("--layout", default="report",
                    choices=["report", "report-year", "report-month", "date-report"],
                    help="정리 폴더 구조 (기본: report)")
    ap.add_argument("--copy", action="store_true", help="이동 대신 복사")
    ap.add_argument("--dry-run", action="store_true", help="정리 시 실제로 옮기지 않고 계획만 출력")
    ap.add_argument("--no-excel", action="store_true", help="엑셀 생성 없이 정리만 수행")
    args = ap.parse_args(argv)

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
        res = compute_pnl(positions, activity, external_categories=args.external_category)

        swap_sec = compute_security_pnl(positions, activity, "Swap")
        cash_sec = compute_security_pnl(positions, activity, "Cash Equity")
        fx_sec = compute_security_pnl(positions, activity, "FX")
        swap_detail = compute_swap_detail(raw_swap, raw_reset)

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
            }).drop_duplicates().sort_values(["통화", "이자기산일"])

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
        meta = [
            ("생성일시", dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("원본 폴더", src),
            ("계좌", acct),
            ("기준일 범위", f"{dates[0]} ~ {dates[-1]} ({len(dates)}영업일)" if dates else "-"),
            ("사용 리포트", ", ".join(sorted({f.code for f in files if f.code}))),
            ("기준통화", "USD (Client Base CCY)"),
            ("총평가액(NAV, 최종일)", f"{last['총평가액(NAV)']:,.2f}" if last is not None else "-"),
            ("누적손익(기간 전체)", f"{last['합계 누적손익']:,.2f}" if last is not None else "-"),
            ("최종일 일일손익", f"{last['합계 일일손익']:,.2f}" if last is not None else "-"),
            ("손익 산식", "일일손익 = Δ평가액 + 귀속 현금흐름 (현금 자산군은 잔여항)"),
            ("검증", "Σ자산군 손익 = Δ총평가액 − 외부 자금이동 (09_검증 시트 참조)"),
        ]

        sheets = [
            ("01_일일손익", daily),
            ("02_자산군별상세", res["detail"]),
            ("03_Swap손익", swap_sec),
            ("04_현물주식손익", cash_sec),
            ("05_FX손익", fx_sec),
            ("06_현금잔고", cash_bal),
            ("07_결제스케줄", cash_sched),
            ("08_거래내역", trades),
            ("09_검증", res["recon"]),
            ("10_Swap상세(MTM)", swap_detail),
            ("11_포지션(최종일)", pos_snapshot),
            ("12_이자내역", interest),
            ("13_파일목록", inventory),
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
