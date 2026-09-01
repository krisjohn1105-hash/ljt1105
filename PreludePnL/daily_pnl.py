"""
Prelude 일일손익 누적 관리 파일 생성기
======================================

매일 실행하면 Prelude 리포트에서 그날의 손익을 계산해 **하나의 엑셀 파일에 누적**한다.

  * 2026-05-29 까지의 누적손익은 기존 EQSWAP.xlsx 의 Summary 시트에서 가져온다(시드).
  * 2026-06-01 부터는 Prelude_new 의 CSV 로 스왑 / 현물(Cash) / FX / IPO / 현금·이자
    5개 자산군의 일일손익·누적손익을 계산한다.

기존 EQSWAP.xlsx 는 스왑 전용이라 6월 Cash trade 개시 이후 일일손익이 부정확해졌다.
이 파일은 계좌 전체 평가액 변동을 기준으로 하므로 모든 자산군을 빠짐없이 담는다.

사용 예)
    python daily_pnl.py --src "Z:/02.펀드/003.매매보고서 대사/Prelude_new"
    python daily_pnl.py --src ... --rebuild          # 기존 누적분 무시하고 전량 재계산
    python daily_pnl.py --src ... --cutover 2026-06-01 --seed-date 2026-05-29

손익 계산 원리
--------------
    일일손익(자산군) = 평가액(t) - 평가액(t-1) + 해당 자산군 귀속 현금흐름(t)

  * 평가액   : MAC001X 'Market Value / Net Equity (Base)' (USD)
  * 현금흐름 : MAC002TDX 'Net Amt Base' 중 현금원장(Position Type = PB / COLCASH) 행
  * 현금·이자 자산군은 잔여항 -> 환평가손익 + 이자 + 배당 + 수수료가 모인다.
  => Σ 자산군 일일손익 = Δ 총평가액 - 외부 자금이동  (완전 일치, '검증' 시트 참조)

IPO 처리
--------
Prelude 상 IPO 배정주는 '대금 0원 Buy Long'(무상입고)으로 들어오고,
청약대금은 별도 'Wires' 로 빠져나간다. 따라서

    IPO 손익 = (배정주 평가액 변동 + 매도대금) - 청약대금

으로 계산된다. IPO 종목은 EQSWAP.xlsx 의 IPO 시트를 마스터로 삼고,
청약대금 Wire 는 그 시트의 Payment Amount(원화)와 금액을 대조해 찾아낸다.
시트에 없는 무상입고 건은 'IPO 후보' 로 보고만 하고 현물에 남긴다(--ipo-auto 로 포함 가능).

작성: Claude Code / 2026-08
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import sys
from typing import Dict, List, Optional, Sequence, Set, Tuple

import pandas as pd

from prelude_pnl import (  # 같은 폴더의 모듈 재사용
    ASSET_CLASS_TO_BUCKET, CASH_LEDGER_POSITION_TYPES,
    build_activity, build_cash_balance, build_positions,
    index_by_code, load_code, long_path, scan_files, to_num,
)

# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------

DEFAULT_CUTOVER = dt.date(2026, 6, 1)      # 이 날짜부터 Prelude 기반으로 계산
DEFAULT_PRINCIPAL = 10_000_000.0           # 기준가 산출용 원금 (EQSWAP.xlsx 와 동일)
DEFAULT_EQSWAP = "EQSWAP.xlsx"
DEFAULT_OUTPUT = "Prelude_Daily_PnL.xlsx"
MAX_GAP_DAYS = 5

BUCKETS = ["Swap", "Cash Equity", "FX", "IPO", "Cash"]
BUCKET_KR = {
    "Swap": "스왑",
    "Cash Equity": "현물",
    "FX": "FX",
    "IPO": "IPO",
    "Cash": "현금·이자·기타",
}

DAILY_SHEET = "01_일일손익"
IPO_WIRE_CATEGORIES = ("Wires",)


# ---------------------------------------------------------------------------
# EQSWAP.xlsx 읽기 (시드 누적손익 + IPO 마스터)
# ---------------------------------------------------------------------------

def read_eqswap_summary(path: str) -> pd.DataFrame:
    """EQSWAP.xlsx Summary 시트를 (기준일, 누적손익) 형태로 읽는다."""
    import openpyxl
    wb = openpyxl.load_workbook(long_path(path), read_only=True, data_only=True)
    try:
        ws = wb["Summary"]
        recs = []
        for row in ws.iter_rows(values_only=True):
            d, cum = row[0], row[1]
            if isinstance(d, dt.datetime):
                d = d.date()
            if isinstance(d, dt.date) and isinstance(cum, (int, float)):
                recs.append({
                    "기준일": d,
                    "누적손익": float(cum),
                    "미실현_평가": _f(row[2]), "미실현_대기": _f(row[3]), "미실현_이자": _f(row[4]),
                    "실현_누적": _f(row[5]), "실현_당일": _f(row[6]),
                    "금융비용": _f(row[7]), "IPO": _f(row[8]), "FX": _f(row[9]),
                })
    finally:
        wb.close()
    return pd.DataFrame(recs).sort_values("기준일").reset_index(drop=True)


def read_eqswap_ipo(path: str) -> pd.DataFrame:
    """EQSWAP.xlsx IPO 시트(청약 마스터)를 읽는다."""
    import openpyxl
    wb = openpyxl.load_workbook(long_path(path), read_only=True, data_only=True)
    try:
        if "IPO" not in wb.sheetnames:
            return pd.DataFrame()
        rows = list(wb["IPO"].iter_rows(values_only=True))
    finally:
        wb.close()
    if not rows:
        return pd.DataFrame()
    hdr = [str(c).strip() if c is not None else "" for c in rows[0]]
    df = pd.DataFrame(rows[1:], columns=hdr)
    df = df[df["Stock Description"].notna()].copy()
    for c in ("Listing Date", "Trade Date", "Settle Date"):
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
    for c in ("Allocation Shares", "IPO Price", "Payment Amount",
              "Settled Net Amount", "Settled Net Amount ($)", "PnL", "PnL ($)"):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    df["ticker"] = df.get("ticker", "").astype(str).str.strip()
    return df.reset_index(drop=True)


def _f(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ---------------------------------------------------------------------------
# IPO 식별
# ---------------------------------------------------------------------------

def _norm_name(s) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def _norm_ticker(s) -> str:
    return str(s).strip().upper().split(".")[0]


def detect_free_deliveries(activity: pd.DataFrame) -> pd.DataFrame:
    """대금 0원 매수(무상입고) = IPO 배정 후보를 찾는다."""
    if activity.empty:
        return pd.DataFrame()
    a = activity[
        activity["포지션유형"].eq("PB")
        & activity["소분류"].isin(["Buy Long", "Buy"])
        & (activity["수량"] > 0)
        & (activity["정산금액_USD"].abs() < 0.005)
        & (activity["정산금액_결제통화"].abs() < 1.0)
    ]
    if a.empty:
        return pd.DataFrame()
    return (a.groupby(["종목명", "종목코드"], as_index=False)
            .agg(입고일=("기준일", "min"), 수량=("수량", "sum")))


def build_ipo_universe(ipo_master: pd.DataFrame, activity: pd.DataFrame,
                       extra_tickers: Sequence[str], use_auto: bool
                       ) -> Tuple[Set[str], Set[str], pd.DataFrame]:
    """
    IPO 종목 식별자 집합을 만든다.
    반환: (티커 집합, 정규화 종목명 집합, 자동탐지 후보 표)
    """
    tickers: Set[str] = {_norm_ticker(t) for t in extra_tickers if str(t).strip()}
    names: Set[str] = set()

    if ipo_master is not None and not ipo_master.empty:
        for _, r in ipo_master.iterrows():
            if str(r.get("ticker", "")).strip():
                tickers.add(_norm_ticker(r["ticker"]))
            if str(r.get("Stock Description", "")).strip():
                names.add(_norm_name(r["Stock Description"]))

    candidates = detect_free_deliveries(activity)
    if not candidates.empty:
        known = tickers | names
        candidates["등록여부"] = candidates.apply(
            lambda r: "등록됨" if (_norm_ticker(r["종목코드"]) in tickers
                                or _norm_name(r["종목명"]) in names
                                or any(n and n in _norm_name(r["종목명"]) for n in names)) else "미등록",
            axis=1)
        if use_auto:
            for _, r in candidates[candidates["등록여부"] == "미등록"].iterrows():
                tickers.add(_norm_ticker(r["종목코드"]))
                names.add(_norm_name(r["종목명"]))
    return tickers, names, candidates


def is_ipo_security(name, code, tickers: Set[str], names: Set[str]) -> bool:
    if _norm_ticker(code) in tickers:
        return True
    n = _norm_name(name)
    if not n:
        return False
    if n in names:
        return True
    # 'Piece Peace Studio' (시트) vs 'PIECE PEACE STUDIO CO LTD' (Prelude)
    return any(m and (m in n or n in m) for m in names)


def find_ipo_wires(activity: pd.DataFrame, ipo_master: pd.DataFrame,
                   tolerance: float = 1.0) -> pd.DataFrame:
    """IPO 청약대금 Wire 를 찾아낸다(원화 금액을 IPO 시트 Payment Amount 와 대조)."""
    if activity.empty:
        return pd.DataFrame()
    wires = activity[activity["소분류"].isin(IPO_WIRE_CATEGORIES) & activity["현금원장"]].copy()
    if wires.empty:
        return pd.DataFrame()
    pays = []
    if ipo_master is not None and not ipo_master.empty and "Payment Amount" in ipo_master.columns:
        for _, r in ipo_master.iterrows():
            amt = r.get("Payment Amount")
            if pd.notna(amt) and float(amt) != 0:
                pays.append((float(amt), r.get("Stock Description", ""), r.get("Trade Date")))

    def match(row):
        for amt, name, td in pays:
            if abs(abs(row["정산금액_결제통화"]) - amt) <= tolerance:
                return pd.Series({"IPO종목": name, "청약대금": amt, "IPO매매일": td})
        return pd.Series({"IPO종목": None, "청약대금": None, "IPO매매일": None})

    wires = pd.concat([wires.reset_index(drop=True),
                       wires.apply(match, axis=1).reset_index(drop=True)], axis=1)
    return wires


# ---------------------------------------------------------------------------
# 자산군 재분류 (IPO 분리)
# ---------------------------------------------------------------------------

def apply_ipo_buckets(positions: pd.DataFrame, activity: pd.DataFrame,
                      tickers: Set[str], names: Set[str],
                      ipo_wire_index: Set[int]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    pos = positions.copy()
    act = activity.copy()

    if not pos.empty:
        mask = (pos["버킷"] == "Cash Equity") & pos.apply(
            lambda r: is_ipo_security(r["종목명"], r["종목코드"], tickers, names), axis=1)
        pos.loc[mask, "버킷"] = "IPO"

    if not act.empty:
        mask = act["버킷"].isin(["Cash Equity"]) & act.apply(
            lambda r: is_ipo_security(r["종목명"], r["종목코드"], tickers, names), axis=1)
        act.loc[mask, "버킷"] = "IPO"
        if ipo_wire_index:
            w = act.index.isin(ipo_wire_index)
            act.loc[w, "버킷"] = "IPO"
            act.loc[w, "종목명"] = "IPO 청약대금"
            act.loc[w, "종목코드"] = ""
    return pos, act


# ---------------------------------------------------------------------------
# 일일손익 계산
# ---------------------------------------------------------------------------

def compute_daily(positions: pd.DataFrame, activity: pd.DataFrame,
                  external_index: Set[int], cutover: dt.date,
                  max_gap_days: int = MAX_GAP_DAYS) -> Dict[str, pd.DataFrame]:
    dates = sorted(positions["기준일"].unique())
    if not dates:
        raise SystemExit("MAC001X 포지션 데이터가 없습니다.")

    mv = (positions.groupby(["기준일", "버킷"], as_index=False)["평가액_USD"].sum()
          .pivot(index="기준일", columns="버킷", values="평가액_USD")
          .reindex(dates).fillna(0.0))
    for b in BUCKETS:
        if b not in mv.columns:
            mv[b] = 0.0
    mv = mv[BUCKETS + [c for c in mv.columns if c not in BUCKETS]]

    cash_rows = activity[activity["현금원장"]] if not activity.empty else pd.DataFrame()
    if len(cash_rows):
        ext_mask = cash_rows.index.isin(external_index)
        ext = (cash_rows[ext_mask].groupby("기준일")["정산금액_USD"].sum()
               .reindex(mv.index).fillna(0.0))
        flow = (cash_rows[~ext_mask].groupby(["기준일", "버킷"], as_index=False)["정산금액_USD"].sum()
                .pivot(index="기준일", columns="버킷", values="정산금액_USD")
                .reindex(mv.index).fillna(0.0))
    else:
        ext = pd.Series(0.0, index=mv.index)
        flow = pd.DataFrame(0.0, index=mv.index, columns=mv.columns)
    for c in mv.columns:
        if c not in flow.columns:
            flow[c] = 0.0
    flow = flow[mv.columns]

    idx = list(mv.index)
    prev_date = pd.Series([None] + idx[:-1], index=idx)
    gap = pd.Series([float("nan")] + [(idx[i] - idx[i - 1]).days for i in range(1, len(idx))], index=idx)
    computable = pd.Series([d >= cutover and (i > 0 and gap.iloc[i] <= max_gap_days)
                            for i, d in enumerate(idx)], index=idx)

    first = next((d for d in idx if computable[d]), None)
    d_mv = mv.diff()
    pnl = pd.DataFrame(0.0, index=mv.index, columns=mv.columns)
    others = [c for c in mv.columns if c != "Cash"]
    for b in others:
        pnl[b] = d_mv[b] + flow[b]
    pnl["Cash"] = d_mv["Cash"] - flow[others].sum(axis=1) - ext
    pnl = pnl.where(computable, 0.0)

    note = pd.Series("", index=idx)
    note[[d for d in idx if d < cutover]] = "컷오버 이전(EQSWAP 시드 구간)"
    note[[d for d in idx if d >= cutover and not computable[d]]] = \
        "직전 리포트일과 간격이 커서 산출 제외"

    total_mv = mv.sum(axis=1)
    daily = pd.DataFrame(index=mv.index)
    daily.index.name = "기준일"
    daily["직전 기준일"] = prev_date
    daily["경과일수"] = gap
    for b in BUCKETS:
        daily[BUCKET_KR[b]] = pnl[b]
    daily["일일손익 합계"] = pnl[BUCKETS].sum(axis=1)
    daily["MS계좌 총평가액"] = total_mv
    daily["외부 자금이동"] = ext.where(computable, 0.0)
    daily["산출대상"] = computable
    daily["비고"] = note

    detail = []
    for i, d in enumerate(idx):
        for b in BUCKETS:
            detail.append({
                "기준일": d, "자산군": BUCKET_KR[b],
                "전일 평가액": mv[b].shift().iloc[i] if i else float("nan"),
                "당일 평가액": mv[b].iloc[i],
                "평가액 증감": d_mv[b].iloc[i],
                "현금흐름": flow[b].iloc[i],
                "일일손익": pnl[b].iloc[i],
                "산출대상": bool(computable.iloc[i]),
            })

    recon = pd.DataFrame(index=mv.index)
    recon.index.name = "기준일"
    recon["산출대상"] = computable
    recon["총평가액 증감"] = total_mv.diff().where(computable)
    recon["외부 자금이동"] = ext.where(computable, 0.0)
    recon["자산군 손익 합계"] = pnl[BUCKETS].sum(axis=1)
    recon["차이(검증)"] = (recon["총평가액 증감"] - recon["외부 자금이동"]
                       - recon["자산군 손익 합계"])
    recon["비고"] = note

    return {"daily": daily.reset_index(), "detail": pd.DataFrame(detail),
            "recon": recon.reset_index(), "mv": mv, "flow": flow, "pnl": pnl,
            "computable": computable}


def compute_security_pnl(positions: pd.DataFrame, activity: pd.DataFrame,
                         computable: pd.Series) -> pd.DataFrame:
    """자산군·종목 단위 일일손익."""
    pos = positions.copy()
    act = activity[activity["현금원장"]].copy() if not activity.empty else pd.DataFrame()
    key_cols = ["버킷", "종목명"]
    mv = (pos.groupby(["기준일"] + key_cols, as_index=False)
          .agg(평가액=("평가액_USD", "sum"), 수량=("수량", "sum"),
               종목코드=("종목코드", "last"))
          if not pos.empty else pd.DataFrame())
    fl = (act.groupby(["기준일"] + key_cols, as_index=False)
          .agg(현금흐름=("정산금액_USD", "sum"))
          if len(act) else pd.DataFrame(columns=["기준일"] + key_cols + ["현금흐름"]))

    dates = sorted(positions["기준일"].unique())
    keys = pd.concat([mv[key_cols] if len(mv) else pd.DataFrame(columns=key_cols),
                      fl[key_cols] if len(fl) else pd.DataFrame(columns=key_cols)]
                     ).drop_duplicates()
    if keys.empty:
        return pd.DataFrame()
    grid = keys.merge(pd.DataFrame({"기준일": dates}), how="cross")
    out = grid.merge(mv, on=["기준일"] + key_cols, how="left") \
              .merge(fl, on=["기준일"] + key_cols, how="left")
    for c in ("평가액", "현금흐름", "수량"):
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)
    out = out.sort_values(key_cols + ["기준일"])
    out["전일평가액"] = out.groupby(key_cols)["평가액"].shift()
    out["일일손익"] = (out["평가액"] - out["전일평가액"].fillna(out["평가액"])) + out["현금흐름"]
    bad = set(computable.index[~computable.astype(bool)])
    out.loc[out["기준일"].isin(bad), "일일손익"] = 0.0
    out["누적손익"] = out.groupby(key_cols)["일일손익"].cumsum()
    out["자산군"] = out["버킷"].map(BUCKET_KR).fillna(out["버킷"])
    # 종목이 붙지 않는 현금원장 저널(스왑 리셋/파이낸싱 정산, 담보이체 등) 구분
    out["구분"] = "종목"
    journal = (out["평가액"] == 0) & (out["수량"] == 0) & (out["현금흐름"] != 0)
    out.loc[journal, "구분"] = "현금흐름(종목없음)"
    out = out[(out["평가액"] != 0) | (out["현금흐름"] != 0) | (out["일일손익"] != 0)]
    cols = ["기준일", "자산군", "구분", "종목명", "종목코드", "수량",
            "전일평가액", "평가액", "현금흐름", "일일손익", "누적손익"]
    return out[cols].sort_values(["기준일", "자산군", "구분", "종목명"])


# ---------------------------------------------------------------------------
# 누적 병합 (기존 파일 + 신규 계산분)
# ---------------------------------------------------------------------------

def merge_history(new_rows: pd.DataFrame, out_path: str, rebuild: bool) -> pd.DataFrame:
    """기존 산출 파일의 일일손익을 읽어 신규 계산분과 병합한다(같은 날짜는 신규가 우선)."""
    if rebuild or not os.path.exists(long_path(out_path)):
        return new_rows.copy()
    try:
        old = pd.read_excel(out_path, sheet_name=DAILY_SHEET)
    except Exception as exc:
        print(f"  ! 기존 파일 읽기 실패({exc}) - 신규 계산분만 사용합니다.", file=sys.stderr)
        return new_rows.copy()
    if old.empty or "기준일" not in old.columns:
        return new_rows.copy()
    old["기준일"] = pd.to_datetime(old["기준일"], errors="coerce").dt.date
    old = old[old["기준일"].notna()]

    # 새로 계산했지만 산출 불가(직전 기준일 없음 등)인 날은 기존 값을 덮어쓰지 않는다.
    # 과거 원본을 다른 곳으로 옮겨도 이미 쌓아둔 손익이 0 으로 지워지지 않게 하는 안전장치.
    if "산출대상" in new_rows.columns:
        usable = new_rows[new_rows["산출대상"].astype(bool)]
        dropped = new_rows[~new_rows["산출대상"].astype(bool)]
        overwritten = set(usable["기준일"])
        readd = dropped[~dropped["기준일"].isin(set(old["기준일"]))]
        new_rows = pd.concat([usable, readd], ignore_index=True, sort=False)
    else:
        overwritten = set(new_rows["기준일"])

    keep = old[~old["기준일"].isin(overwritten | set(new_rows["기준일"]))]
    cols = [c for c in new_rows.columns if c in keep.columns]
    merged = pd.concat([keep[cols], new_rows], ignore_index=True, sort=False)
    merged = merged.drop_duplicates(subset="기준일", keep="last")
    return merged.sort_values("기준일").reset_index(drop=True)


def add_cumulative(daily: pd.DataFrame, seed_cum: float, seed_date: Optional[dt.date],
                   principal: float) -> pd.DataFrame:
    d = daily.sort_values("기준일").reset_index(drop=True).copy()
    # 인수인계 기준점(EQSWAP 시드)을 첫 행으로 넣어 누적 추이가 이어지게 한다
    if seed_date is not None and seed_date not in set(d["기준일"]):
        seed_row = {c: 0.0 for c in d.columns if d[c].dtype.kind in "if"}
        seed_row.update({"기준일": seed_date, "직전 기준일": pd.NaT, "경과일수": float("nan"),
                         "산출대상": False, "비고": "EQSWAP.xlsx 인수 기준점(시드)"})
        d = pd.concat([pd.DataFrame([seed_row]), d], ignore_index=True, sort=False)
        d = d.sort_values("기준일").reset_index(drop=True)
    d["누적손익"] = seed_cum + d["일일손익 합계"].fillna(0).cumsum()
    for b in BUCKETS:
        kr = BUCKET_KR[b]
        if kr in d.columns:
            d[f"{kr} 누적"] = d[kr].fillna(0).cumsum()
    d["기준가(달러)"] = (principal + d["누적손익"]) / principal
    d["AUM(원금+누적손익)"] = principal + d["누적손익"]
    d["일일수익률(%)"] = (d["일일손익 합계"] /
                     (principal + d["누적손익"].shift().fillna(seed_cum))* 100).astype(float)
    if seed_date is not None:
        d.loc[d["기준일"] <= seed_date, "일일수익률(%)"] = float("nan")
    order = (["기준일", "직전 기준일", "경과일수"]
             + [BUCKET_KR[b] for b in BUCKETS]
             + ["일일손익 합계", "누적손익", "기준가(달러)", "AUM(원금+누적손익)", "일일수익률(%)"]
             + [f"{BUCKET_KR[b]} 누적" for b in BUCKETS]
             + ["MS계좌 총평가액", "외부 자금이동", "산출대상", "비고"])
    return d[[c for c in order if c in d.columns]]


# ---------------------------------------------------------------------------
# 엑셀 출력
# ---------------------------------------------------------------------------

def write_workbook(path: str, sheets: List[Tuple[str, pd.DataFrame]],
                   meta: List[Tuple[str, str]]):
    os.makedirs(long_path(os.path.dirname(os.path.abspath(path))), exist_ok=True)
    with pd.ExcelWriter(path, engine="xlsxwriter",
                        datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as xw:
        wb = xw.book
        hdr = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white",
                             "border": 1, "align": "center", "valign": "vcenter",
                             "text_wrap": True})
        title = wb.add_format({"bold": True, "font_size": 14})
        label = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        value = wb.add_format({"border": 1})
        money = wb.add_format({"num_format": "#,##0.00"})
        px = wb.add_format({"num_format": "0.00000000"})

        ws = wb.add_worksheet("00_요약")
        xw.sheets["00_요약"] = ws
        ws.set_column(0, 0, 34); ws.set_column(1, 1, 66)
        ws.write(0, 0, "Prelude 일일손익 누적 관리", title)
        for r, (k, v) in enumerate(meta, start=2):
            ws.write(r, 0, k, label); ws.write(r, 1, v, value)

        for name, df in sheets:
            sn = name[:31]
            if df is None or len(df) == 0:
                w = wb.add_worksheet(sn); xw.sheets[sn] = w
                w.write(0, 0, "해당 데이터 없음")
                continue
            df.to_excel(xw, sheet_name=sn, index=False)
            w = xw.sheets[sn]
            for i, c in enumerate(df.columns):
                w.write(0, i, str(c), hdr)
                try:
                    width = max(len(str(c)) + 2,
                                int(df[c].astype(str).str.len().quantile(0.95)) + 2)
                except Exception:
                    width = len(str(c)) + 2
                width = min(max(width, 11), 40)
                if "기준가" in str(c):
                    w.set_column(i, i, width, px)
                elif any(k in str(c) for k in ("손익", "평가액", "금액", "잔고", "현금",
                                               "AUM", "누적", "증감", "이동", "대금")):
                    w.set_column(i, i, width, money)
                else:
                    w.set_column(i, i, width)
            w.freeze_panes(1, 1)
            w.autofilter(0, 0, len(df), len(df.columns) - 1)

            if sn == DAILY_SHEET and len(df) > 2:
                cols = list(df.columns)
                if "누적손익" in cols and "기준일" in cols:
                    ci, cc = cols.index("기준일"), cols.index("누적손익")
                    ch = wb.add_chart({"type": "line"})
                    ch.add_series({"name": "누적손익 (USD)",
                                   "categories": [sn, 1, ci, len(df), ci],
                                   "values": [sn, 1, cc, len(df), cc],
                                   "line": {"color": "#1F3864", "width": 2.0}})
                    ch.set_title({"name": "누적손익 추이 (EQSWAP 시드 + Prelude 계산)"})
                    ch.set_y_axis({"num_format": "#,##0"})
                    ch.set_size({"width": 900, "height": 380})
                    ch.set_legend({"position": "bottom"})
                    ws.insert_chart(f"A{len(meta) + 5}", ch)


# ---------------------------------------------------------------------------
# 메인
# ---------------------------------------------------------------------------

def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Prelude 리포트 기반 일일손익/누적손익 누적 관리 엑셀 생성",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--src", required=True, help="Prelude 리포트 폴더 (하위 폴더 포함)")
    ap.add_argument("--out", default=None,
                    help=f"산출 엑셀 (기본: <src>/_output/{DEFAULT_OUTPUT})")
    ap.add_argument("--eqswap", default=None,
                    help=f"기존 스왑 관리 파일 (기본: <src>/{DEFAULT_EQSWAP})")
    ap.add_argument("--cutover", default=DEFAULT_CUTOVER.isoformat(),
                    help=f"Prelude 기반 계산 시작일 (기본 {DEFAULT_CUTOVER})")
    ap.add_argument("--seed-date", default=None,
                    help="EQSWAP 누적손익을 가져올 기준일 (기본: 컷오버 직전 영업일)")
    ap.add_argument("--seed-value", type=float, default=None,
                    help="시드 누적손익을 직접 지정 (EQSWAP 파일 대신)")
    ap.add_argument("--principal", type=float, default=DEFAULT_PRINCIPAL,
                    help=f"기준가 산출 원금 (기본 {DEFAULT_PRINCIPAL:,.0f})")
    ap.add_argument("--ipo-ticker", action="append", default=[],
                    help="IPO 종목 티커 추가 지정 (여러 번 사용 가능)")
    ap.add_argument("--ipo-auto", action="store_true",
                    help="IPO 시트에 없는 무상입고 종목도 IPO 로 자동 편입")
    ap.add_argument("--rebuild", action="store_true",
                    help="기존 산출 파일을 무시하고 전체 재계산")
    ap.add_argument("--max-gap-days", type=int, default=MAX_GAP_DAYS,
                    help=f"직전 리포트일과 간격이 이 일수를 넘으면 산출 제외 (기본 {MAX_GAP_DAYS})")
    args = ap.parse_args(argv)

    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    src = os.path.abspath(args.src)
    if not os.path.isdir(src):
        raise SystemExit(f"폴더를 찾을 수 없습니다: {src}")
    cutover = dt.date.fromisoformat(args.cutover)
    out_path = args.out or os.path.join(src, "_output", DEFAULT_OUTPUT)
    eqswap_path = args.eqswap or os.path.join(src, DEFAULT_EQSWAP)

    print(f"[1/7] 리포트 스캔: {src}")
    files = scan_files(src)
    by_code = index_by_code(files)
    print(f"      {len(files):,}개 파일")

    print("[2/7] 리포트 로드")
    positions = build_positions(load_code(by_code, "MAC001X"))
    activity = build_activity(load_code(by_code, "MAC002TDX"))
    raw_cash = load_code(by_code, "CASH005X")
    if positions.empty:
        raise SystemExit("MAC001X 파일이 없어 계산할 수 없습니다.")

    print("[3/7] EQSWAP.xlsx 시드 읽기")
    seed_cum, seed_date, ipo_master = 0.0, None, pd.DataFrame()
    if args.seed_value is not None:
        seed_cum = args.seed_value
        seed_date = dt.date.fromisoformat(args.seed_date) if args.seed_date else None
        print(f"      시드 직접 지정: {seed_cum:,.2f}")
    elif os.path.exists(long_path(eqswap_path)):
        summary = read_eqswap_summary(eqswap_path)
        ipo_master = read_eqswap_ipo(eqswap_path)
        if args.seed_date:
            want = dt.date.fromisoformat(args.seed_date)
            cand = summary[summary["기준일"] == want]
        else:
            cand = summary[summary["기준일"] < cutover]
        if cand.empty:
            raise SystemExit(f"EQSWAP.xlsx Summary 에서 시드 행을 찾지 못했습니다 ({args.seed_date or f'< {cutover}'})")
        row = cand.iloc[-1]
        seed_cum, seed_date = float(row["누적손익"]), row["기준일"]
        print(f"      시드: {seed_date} 누적손익 {seed_cum:,.2f} USD "
              f"(IPO 마스터 {len(ipo_master)}건)")
    else:
        print(f"      ! EQSWAP.xlsx 없음({eqswap_path}) - 시드 0 으로 진행")

    print("[4/7] IPO 식별")
    tickers, names, candidates = build_ipo_universe(
        ipo_master, activity, args.ipo_ticker, args.ipo_auto)
    wires = find_ipo_wires(activity, ipo_master)
    ipo_wire_idx: Set[int] = set()
    external_idx: Set[int] = set()
    pre_cutover_ipo_pay = []
    if not wires.empty:
        for _, r in wires.iterrows():
            orig = activity.index[(activity["기준일"] == r["기준일"])
                                  & (activity["정산금액_결제통화"] == r["정산금액_결제통화"])
                                  & (activity["소분류"] == r["소분류"])]
            if r["기준일"] < cutover:
                if r["IPO종목"]:
                    pre_cutover_ipo_pay.append((r["기준일"], r["IPO종목"], r["정산금액_USD"]))
                continue          # 컷오버 이전 건은 시드 구간이므로 손익에 영향 없음
            if r["IPO종목"]:
                ipo_wire_idx.update(orig)
            else:
                external_idx.update(orig)
        print(f"      IPO 청약 Wire {len(ipo_wire_idx)}건 / 외부 자금이동 {len(external_idx)}건 (컷오버 이후)")
        for d, nm, amt in pre_cutover_ipo_pay:
            print(f"      · 컷오버 이전 청약대금 {d} {nm} {abs(amt):,.2f} USD (참고, 손익 미반영)")
    unreg = candidates[(candidates.get("등록여부", "") == "미등록")
                       & (candidates["입고일"] >= cutover)] if len(candidates) else pd.DataFrame()
    if len(unreg):
        print(f"      ! IPO 시트 미등록 무상입고 {len(unreg)}건 "
              f"({', '.join(unreg['종목명'].head(3))}) → 현물에 남김"
              f"{' (--ipo-auto 로 편입 가능)' if not args.ipo_auto else ''}")

    positions, activity = apply_ipo_buckets(positions, activity, tickers, names, ipo_wire_idx)

    print("[5/7] 일일손익 계산")
    res = compute_daily(positions, activity, external_idx, cutover,
                        max_gap_days=args.max_gap_days)
    new_daily = res["daily"][res["daily"]["기준일"] >= cutover].copy()
    print(f"      계산 구간: {new_daily['기준일'].min()} ~ {new_daily['기준일'].max()} "
          f"({int(new_daily['산출대상'].sum())}영업일)")

    print("[6/7] 기존 누적분 병합")
    merged = merge_history(new_daily.drop(columns=["산출대상"], errors="ignore")
                           .assign(산출대상=new_daily["산출대상"].values),
                           out_path, args.rebuild)
    daily = add_cumulative(merged, seed_cum, seed_date, args.principal)
    last = daily.iloc[-1]

    sec = compute_security_pnl(positions, activity, res["computable"])
    sec = sec[sec["기준일"] >= cutover]
    detail = res["detail"][res["detail"]["기준일"] >= cutover]
    recon = res["recon"][res["recon"]["기준일"] >= cutover]
    cash_bal = build_cash_balance(positions, raw_cash)
    cash_bal = cash_bal[cash_bal["기준일"] >= cutover] if len(cash_bal) else cash_bal

    ipo_rows = pd.DataFrame()
    if len(sec):
        ipo_rows = sec[sec["자산군"] == "IPO"].copy()
    ipo_sheet = pd.DataFrame()
    if len(ipo_master):
        ipo_sheet = ipo_master[[c for c in ["Trade Date", "Settle Date", "Stock Description",
                                            "ticker", "Allocation Shares", "IPO Price",
                                            "Payment Amount", "Settled Net Amount ($)",
                                            "PnL ($)"] if c in ipo_master.columns]].copy()

    trades = pd.DataFrame()
    if not activity.empty:
        tcols = ["기준일", "매매일", "결제일", "구분", "중분류", "소분류", "버킷", "포지션유형",
                 "상품유형", "종목명", "종목코드", "매매구분", "수량", "단가_USD",
                 "약정금액_USD", "정산금액_USD", "결제통화", "정산금액_결제통화", "현금원장"]
        trades = activity[[c for c in tcols if c in activity.columns]].copy()
        trades["자산군"] = trades["버킷"].map(BUCKET_KR).fillna(trades["버킷"])
        trades = trades[trades["기준일"] >= cutover].sort_values(["기준일", "자산군", "종목명"])

    monthly = daily[daily["기준일"] >= cutover].copy()
    monthly["연월"] = pd.to_datetime(monthly["기준일"]).dt.strftime("%Y-%m")
    magg = (monthly.groupby("연월", as_index=False)[[BUCKET_KR[b] for b in BUCKETS]
                                                    + ["일일손익 합계"]].sum())
    mlast = monthly.groupby("연월", as_index=False).last()[["연월", "누적손익", "기준가(달러)",
                                                          "MS계좌 총평가액"]]
    monthly = magg.merge(mlast, on="연월", how="left")

    meta = [
        ("생성일시", dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("원본 폴더", src),
        ("산출 파일", out_path),
        ("시드 (EQSWAP.xlsx)", f"{seed_date} 누적손익 {seed_cum:,.2f} USD" if seed_date
         else f"{seed_cum:,.2f} USD"),
        ("Prelude 계산 시작일", str(cutover)),
        ("최종 기준일", str(last["기준일"])),
        ("누적손익 (최종)", f"{last['누적손익']:,.2f} USD"),
        ("기준가(달러)", f"{last['기준가(달러)']:.8f}"),
        ("AUM (원금+누적손익)", f"{last['AUM(원금+누적손익)']:,.2f} USD"),
        ("컷오버 이후 손익", f"{last['누적손익'] - seed_cum:,.2f} USD"),
        ("자산군별 누적 (컷오버 이후)",
         " | ".join(f"{BUCKET_KR[b]} {last.get(f'{BUCKET_KR[b]} 누적', 0):,.0f}" for b in BUCKETS)),
        ("IPO 종목", ", ".join(sorted(ipo_rows["종목명"].unique())) if len(ipo_rows) else "없음"),
        ("IPO 시트 미등록 무상입고", ", ".join(unreg["종목명"]) if len(unreg) else "없음"),
        ("컷오버 이전 지급 청약대금(참고)",
         " | ".join(f"{d} {nm} {abs(a):,.2f}" for d, nm, a in pre_cutover_ipo_pay)
         if pre_cutover_ipo_pay else "없음"),
        ("손익 산식", "일일손익 = Δ평가액 + 귀속 현금흐름 (현금·이자는 잔여항)"),
        ("검증", "Σ자산군 = Δ총평가액 − 외부 자금이동 (04_검증 시트)"),
    ]

    sheets = [
        (DAILY_SHEET, daily),
        ("02_월별손익", monthly),
        ("03_자산군별상세", detail),
        ("04_검증", recon),
        ("05_종목별손익", sec),
        ("06_IPO상세", ipo_rows),
        ("07_IPO마스터(EQSWAP)", ipo_sheet),
        ("08_IPO후보(무상입고)", candidates),
        ("09_현금잔고", cash_bal),
        ("10_거래내역", trades),
    ]

    print(f"[7/7] 엑셀 작성: {out_path}")
    write_workbook(out_path, sheets, meta)
    print(f"      완료 · 최종 {last['기준일']} 누적손익 {last['누적손익']:,.2f} USD "
          f"/ 기준가 {last['기준가(달러)']:.6f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
