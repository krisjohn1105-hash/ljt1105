# -*- coding: utf-8 -*-
"""읽어들인 레코드를 '지시서 1건' 단위로 묶고 문서에 채울 값을 만든다."""
from __future__ import annotations  # Python 3.9 호환 (X | None 표기)
import re
import sys
from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

try:
    from . import config
    from .reader import Record
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src import config
    from src.reader import Record

TRUSTEE_PATTERN = re.compile(r"^\s*(?P<pbs>[^(]+?)\s*\(\s*(?P<mgr>[^-()]+?)\s*-\s*(?P<bank>[^()]+?)\s*\)\s*$")


@dataclass
class Line:
    """지시서 내역 표의 한 줄."""
    fund_name: str
    fund_code: str
    apply_date: date
    qty: int
    amount: int
    trade_name: str
    ticker: str
    per_share: float

    @property
    def fund_label(self) -> str:
        return f"{self.fund_name} ({self.fund_code})" if self.fund_code else self.fund_name

    @property
    def detail(self) -> str:
        return f"(종목명: {self.ticker}, 주당 {_won(self.per_share)}원)"


@dataclass
class Instruction:
    """지시서 1건 = 문서 1개."""
    pay_date: date
    trustee: str
    trade_name: str
    broker: str
    lines: list[Line] = field(default_factory=list)
    seq_no: int = 1

    # ------------------------------------------------------------ 파생값
    @property
    def bank(self) -> str:
        return _split_trustee(self.trustee)[0]

    @property
    def pbs(self) -> str:
        return _split_trustee(self.trustee)[1]

    @property
    def bank_short(self) -> str:
        return re.sub(r"(은행|증권|銀行)$", "", self.bank) or self.bank

    @property
    def total_qty(self) -> int:
        return sum(l.qty for l in self.lines)

    @property
    def total_amount(self) -> int:
        return sum(l.amount for l in self.lines)

    @property
    def doc_no(self) -> str:
        return f"{self.pay_date:%Y%m%d}-{self.seq_no:02d}"

    @property
    def receive_account(self) -> str:
        account = config.RECEIVE_ACCOUNTS.get(self.broker)
        if account:
            return account
        base = re.sub(r"\(.*\)$", "", self.broker).strip()
        account = config.RECEIVE_ACCOUNTS.get(base)
        if account:
            return account
        print(f"  [경고] 수취계좌 미등록 중개기관: {self.broker!r} "
              f"-> config.RECEIVE_ACCOUNTS 에 추가하세요.")
        return f"[[수취계좌 미등록: {self.broker}]]"

    @property
    def filename(self) -> str:
        return config.FILENAME_FORMAT.format(
            date=f"{self.pay_date:%Y%m%d}",
            company=config.COMPANY_SHORT,
            bank_short=self.bank_short,
            trade_key=self.trade_name.replace(" ", ""),
        )

    def context(self) -> dict:
        """템플릿 {{KEY}} 치환용 값."""
        return {
            "CONTACT_NAME": config.CONTACT_NAME,
            "CONTACT_PHONE": config.CONTACT_PHONE,
            "DOC_NO_PREFIX": config.DOC_NO_PREFIX,
            "DOC_NO": self.doc_no,
            "DOC_DATE": f"{self.pay_date.year}. {self.pay_date.month}. {self.pay_date.day}.",
            "RECIPIENT": config.RECIPIENT_FORMAT.format(bank=self.bank),
            "CC": config.CC_FORMAT.format(pbs=self.pbs),
            "SUBJECT": (f"{config.COMPANY_SHORT} {self.pay_date.year}년 "
                        f"{self.pay_date.month}월 {self.pay_date.day}일자 수기운용지시의 건 "),
            "TRADE_TYPE": self.trade_name,
            "ORDER_DATE_KR": f"{self.pay_date.month}월 {self.pay_date.day}일자",
            "RECEIVE_ACCOUNT": self.receive_account,
            "TOTAL_AMOUNT": f"{self.total_amount:,}원",
            "COMPANY_NAME": config.COMPANY_NAME,
            "CEO_NAME": config.CEO_NAME,
        }


def _won(value: float) -> str:
    """주당 금액처럼 소수가 나올 수 있는 값을 보기 좋게."""
    return f"{value:,.0f}" if abs(value - round(value)) < 1e-9 else f"{value:,.2f}"


def _split_trustee(trustee: str) -> tuple[str, str]:
    """수탁기관명 -> (수탁은행, PBS명)."""
    if trustee in config.TRUSTEE_OVERRIDES:
        return config.TRUSTEE_OVERRIDES[trustee]
    m = TRUSTEE_PATTERN.match(trustee or "")
    if m:
        return m.group("bank").strip(), m.group("pbs").strip()
    return (trustee or "").strip(), (trustee or "").strip()


def trade_name_of(record: Record) -> str:
    key = (record.manage_type, record.right_type, record.security_type)
    if key in config.TRADE_TYPE_NAMES:
        return config.TRADE_TYPE_NAMES[key]
    print(f"  [경고] 거래유형명 미등록: {key} -> config.TRADE_TYPE_NAMES 에 추가하세요.")
    return f"{record.manage_type}{record.security_type} {record.right_type}".strip()


def build_instructions(records: list[Record], start_seq: int = 1) -> list[Instruction]:
    """(지급일, 수탁기관, 거래유형) 별로 문서를 나누고, 펀드·종목별로 줄을 합산한다."""
    grouped: "OrderedDict[tuple, Instruction]" = OrderedDict()
    line_bucket: dict[tuple, dict[tuple, Line]] = {}

    for rec in records:
        trade_name = trade_name_of(rec)
        doc_key = (rec.pay_date, rec.trustee, trade_name)
        if doc_key not in grouped:
            grouped[doc_key] = Instruction(
                pay_date=rec.pay_date, trustee=rec.trustee,
                trade_name=trade_name, broker=rec.broker,
            )
            line_bucket[doc_key] = OrderedDict()

        line_key = (rec.fund_code, rec.fund_name, rec.ticker, round(rec.per_share, 6))
        bucket = line_bucket[doc_key]
        if line_key in bucket:
            line = bucket[line_key]
            line.qty += rec.qty
            line.amount += rec.amount
        else:
            bucket[line_key] = Line(
                fund_name=rec.fund_name, fund_code=rec.fund_code,
                apply_date=rec.pay_date, qty=rec.qty, amount=rec.amount,
                trade_name=trade_name, ticker=rec.ticker, per_share=rec.per_share,
            )

    instructions = []
    per_date_seq: dict[date, int] = {}
    for doc_key, instruction in grouped.items():
        instruction.lines = list(line_bucket[doc_key].values())
        # 문서번호 일련번호는 지급일 기준으로 01 부터 매긴다.
        seq = per_date_seq.get(instruction.pay_date, start_seq)
        instruction.seq_no = seq
        per_date_seq[instruction.pay_date] = seq + 1
        instructions.append(instruction)
    return instructions
