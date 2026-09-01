# -*- coding: utf-8 -*-
"""권리배정내역 엑셀(.xls/.xlsx)을 읽어 지시서용 레코드로 정규화한다."""
from __future__ import annotations  # Python 3.9 호환 (X | None 표기)
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# 엑셀 헤더가 2줄(상위 병합 + 하위 세부)로 되어 있어 직접 파싱한다.
HEADER_KEYS = ("순번", "기준일", "권리구분")
DATA_EXTENSIONS = (".xls", ".xlsx", ".xlsm")


@dataclass
class Record:
    """엑셀 한 줄(대차 체결 건 하나)."""
    seq: int
    base_date: date          # 기준일 (배당기준일)
    pay_date: date           # 지급일 (수기운용지시 반영일자)
    right_type: str          # 권리구분  예) 현금배당
    manage_type: str         # 운용구분  예) 차입 / 대여
    security_type: str       # 증권구분  예) 주식
    fund_code: str
    fund_name: str
    ticker: str              # 기준종목명
    qty: int                 # 대차수량
    amount: int              # 대금 (현금배당 지급액)
    broker: str              # 중개기관명
    trustee: str             # 수탁기관명
    alloc_name: str = ""     # 권리배정내역 > 종목명 (분할·배당으로 배정받는 종목)
    alloc_qty: int = 0       # 권리배정내역 > 배정수량
    odd_amount: int = 0      # 권리배정내역 > 단주대금

    @property
    def per_share(self) -> float:
        return self.amount / self.qty if self.qty else 0.0


def find_input_files(target: Path) -> list[Path]:
    """디렉터리면 안의 엑셀 전부, 파일이면 그 파일 하나를 돌려준다."""
    target = Path(target)
    if target.is_file():
        return [target]
    if not target.is_dir():
        raise FileNotFoundError(f"입력 경로를 찾을 수 없습니다: {target}")
    files = sorted(
        p for p in target.iterdir()
        if p.suffix.lower() in DATA_EXTENSIONS and not p.name.startswith("~$")
    )
    if not files:
        raise FileNotFoundError(f"엑셀 파일이 없습니다: {target}")
    return files


def _read_raw(path: Path) -> pd.DataFrame:
    engine = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"
    return pd.read_excel(path, header=None, dtype=object, engine=engine)


def _build_columns(raw: pd.DataFrame) -> tuple[int, dict[str, int]]:
    """헤더 2줄을 합쳐 {컬럼명: 열번호} 와 데이터 시작 행을 만든다."""
    header_row = None
    for r in range(min(10, len(raw))):
        values = [str(v).strip() for v in raw.iloc[r].tolist()]
        if all(any(k == v for v in values) for k in HEADER_KEYS):
            header_row = r
            break
    if header_row is None:
        raise ValueError("권리배정내역 헤더(순번/기준일/권리구분)를 찾지 못했습니다.")

    top = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[header_row]]
    sub_row = header_row + 1
    sub = ([str(v).strip() if pd.notna(v) else "" for v in raw.iloc[sub_row]]
           if sub_row < len(raw) else [""] * len(top))
    # 하위 헤더가 실제 세부명인지(= 데이터 행이 아닌지) 확인
    if sum(1 for v in sub if v and v != "nan") < 2:
        sub = [""] * len(top)
        data_start = sub_row
    else:
        data_start = sub_row + 1

    columns: dict[str, int] = {}
    for idx, (t, s) in enumerate(zip(top, sub)):
        name = s if s and s != "nan" else t
        if name and name != "nan" and name not in columns:
            columns[name] = idx
    return data_start, columns


def _to_date(value) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = re.sub(r"[^0-9]", "", str(value))
    if len(text) != 8:
        return None
    return datetime.strptime(text, "%Y%m%d").date()


def _to_int(value) -> int:
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return 0
    if isinstance(value, str):
        value = re.sub(r"[^0-9.\-]", "", value) or "0"
    return int(round(float(value)))


def _text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_records(path: Path) -> list[Record]:
    raw = _read_raw(Path(path))
    data_start, columns = _build_columns(raw)

    def col(row, name, default=None):
        idx = columns.get(name)
        return row[idx] if idx is not None and idx < len(row) else default

    records: list[Record] = []
    for r in range(data_start, len(raw)):
        row = raw.iloc[r].tolist()
        seq_raw = col(row, "순번")
        if seq_raw is None or (isinstance(seq_raw, float) and pd.isna(seq_raw)):
            continue
        try:
            seq = int(float(seq_raw))
        except (TypeError, ValueError):
            continue  # 소계/그룹 헤더 행

        pay_date = _to_date(col(row, "지급일")) or _to_date(col(row, "기준일"))
        if pay_date is None:
            continue

        records.append(Record(
            seq=seq,
            base_date=_to_date(col(row, "기준일")) or pay_date,
            pay_date=pay_date,
            right_type=_text(col(row, "권리구분")),
            manage_type=_text(col(row, "운용구분")),
            security_type=_text(col(row, "증권구분")),
            fund_code=_text(col(row, "펀드코드")),
            fund_name=_text(col(row, "펀드명")),
            ticker=_text(col(row, "기준종목명")) or _text(col(row, "종목명")),
            qty=_to_int(col(row, "대차수량")),
            amount=_to_int(col(row, "대금")),
            broker=_text(col(row, "중개기관명")),
            trustee=_text(col(row, "수탁기관명")),
            alloc_name=_text(col(row, "종목명")),
            alloc_qty=_to_int(col(row, "배정수량")),
            odd_amount=_to_int(col(row, "단주대금")),
        ))
    return records


def read_all(target: Path) -> list[Record]:
    records: list[Record] = []
    for path in find_input_files(target):
        records.extend(read_records(path))
    return records
