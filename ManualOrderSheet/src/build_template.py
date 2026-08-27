# -*- coding: utf-8 -*-
"""사용자가 직접 작성한 지시서(.docx)를 자리표시자 템플릿으로 변환한다.

원본 서식(글꼴/표 너비/머리글/도장 라인)을 그대로 두고 값만 {{KEY}} 로 바꾼다.
양식 자체가 바뀌면 새 원본으로 이 스크립트를 다시 돌리면 된다.

    python -m src.build_template [원본.docx]
"""
from __future__ import annotations  # Python 3.9 호환 (X | None 표기)
import sys
from pathlib import Path

from docx import Document
from docx.table import _Cell

try:
    from . import config
    from .docx_util import set_paragraph_text, set_cell_text
except ImportError:  # 스크립트로 직접 실행할 때
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src import config
    from src.docx_util import set_paragraph_text, set_cell_text

DEFAULT_SOURCE = (config.TEMPLATE_DIR /
                  "20260825_두나미스자산운용 수기운용지시서_신한수탁_차입주식배당지급.docx")


def tc(table, row_idx, tc_idx):
    return _Cell(table.rows[row_idx]._tr.tc_lst[tc_idx], table)


def build(source: Path, dest: Path) -> Path:
    doc = Document(str(source))
    head, items = doc.tables[0], doc.tables[1]

    # --- 머리글: 담당자 ------------------------------------------------
    for paragraph in doc.sections[0].header.paragraphs:
        if "담당:" in paragraph.text:
            head_text = paragraph.text
            head_text = head_text.replace(config.CONTACT_NAME, "{{CONTACT_NAME}}")
            head_text = head_text.replace(config.CONTACT_PHONE, "{{CONTACT_PHONE}}")
            set_paragraph_text(paragraph, head_text)

    # --- 상단 표: 문서번호 / 일자 / 수신 / 참조 / 제목 -----------------
    set_cell_text(tc(head, 1, 1), " {{DOC_NO_PREFIX}} 제{{DOC_NO}}")
    set_cell_text(tc(head, 1, 2), "{{DOC_DATE}}")
    set_cell_text(tc(head, 2, 1), " {{RECIPIENT}}")
    set_cell_text(tc(head, 3, 1), " {{CC}}")
    set_cell_text(tc(head, 4, 1), " {{SUBJECT}}")

    # --- 본문 문단 ----------------------------------------------------
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text.startswith("2."):
            set_paragraph_text(
                paragraph,
                "2. 당사 금일 발생한 {{TRADE_TYPE}} 관련하여 "
                "{{ORDER_DATE_KR}} 수기운용지시를 반영 요청 드립니다.",
            )
        elif text.startswith("수취계좌:"):
            set_paragraph_text(paragraph, "수취계좌: {{RECEIVE_ACCOUNT}}")
        elif text.startswith("주식회사"):
            set_paragraph_text(paragraph, "{{COMPANY_NAME}}")
        elif text.startswith("대표이사"):
            set_paragraph_text(paragraph, "대표이사 {{CEO_NAME}} (인)")

    # --- 내역 표: 1행이 반복 행 템플릿, 마지막 행이 합계 ---------------
    for idx, key in enumerate(
        ["NO", "FUND", "APPLY_DATE", "QTY", "AMOUNT", "TRADE_DESC"]
    ):
        set_cell_text(tc(items, 1, idx), "{{%s}}" % key)
    # TRADE_DESC 는 두 줄(거래유형명 / 상세)로 쓰므로 분리해 둔다.
    set_cell_text(tc(items, 1, 5), "{{TRADE_NAME}}\n{{TRADE_DETAIL}}")
    set_cell_text(tc(items, 2, 2), "{{TOTAL_AMOUNT}}")

    dest.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dest))
    return dest


def main():
    source = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SOURCE
    if not source.exists():
        raise SystemExit(f"원본 지시서를 찾을 수 없습니다: {source}")
    out = build(source, config.TEMPLATE_PATH)
    print(f"템플릿 생성 완료: {out}")


if __name__ == "__main__":
    main()
