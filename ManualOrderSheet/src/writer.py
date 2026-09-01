# -*- coding: utf-8 -*-
"""템플릿(.docx)에 값을 채워 수기운용지시서 파일을 만든다."""
from __future__ import annotations  # Python 3.9 호환 (X | None 표기)
import sys
from pathlib import Path

from docx import Document
from docx.table import _Cell

try:
    from . import config
    from .docx_util import (replace_everywhere, set_cell_text, clone_row,
                            delete_row, trim_spacers_after)
    from .instruction import Instruction
    from . import seal as seal_mod
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src import config
    from src.docx_util import (replace_everywhere, set_cell_text, clone_row,
                              delete_row, trim_spacers_after)
    from src.instruction import Instruction
    from src import seal as seal_mod

ITEM_TABLE_INDEX = 1   # 0: 문서번호/수신/참조/제목,  1: 내역 표
ITEM_ROW_INDEX = 1     # 0: 머리행, 1: 반복 행 템플릿, 2: 합계

# 문서 종류별 내역 표의 열 구성 (build_template.SOURCES 와 같아야 한다)
COLUMNS = {
    "dividend": ["NO", "FUND", "APPLY_DATE", "QTY", "AMOUNT", "TRADE_DESC"],
    "odd_lot": ["NO", "FUND", "APPLY_DATE", "AMOUNT", "TRADE_DESC"],
}

# 윗줄과 값이 같으면 비워두는 칸
DITTO_KEYS = ("FUND", "APPLY_DATE", "TRADE_DESC")


def _cells(table, row_idx):
    return [_Cell(tc, table) for tc in table.rows[row_idx]._tr.tc_lst]


def resolve_output_dir(instruction: Instruction, output_dir: Path,
                       subdir_by_date: bool | None = None) -> Path:
    """출력 폴더. 기본은 <출력폴더>/<지급일 YYYYMMDD>/ 로 한 단계 내려간다."""
    if subdir_by_date is None:
        subdir_by_date = config.OUTPUT_SUBDIR_BY_DATE
    output_dir = Path(output_dir)
    if subdir_by_date:
        output_dir = output_dir / instruction.pay_date.strftime(config.OUTPUT_SUBDIR_FORMAT)
    return output_dir


def render(instruction: Instruction, output_dir: Path,
           template_path: Path | None = None, overwrite: bool = True,
           subdir_by_date: bool | None = None,
           insert_seal: bool | None = None,
           trim_spacers: int | None = None) -> Path:
    template_path = Path(template_path or instruction.template_path)
    if not template_path.exists():
        raise FileNotFoundError(
            f"템플릿이 없습니다: {template_path}\n"
            f"  python -m src.build_template 으로 먼저 생성하세요."
        )

    doc = Document(str(template_path))
    table = doc.tables[ITEM_TABLE_INDEX]

    # 내역 행: 템플릿 행을 필요한 만큼 복제한 뒤 값을 채운다.
    extra_rows = max(0, len(instruction.lines) - 1)
    for offset in range(extra_rows):
        clone_row(table, ITEM_ROW_INDEX + offset)

    # 표 아래 빈 줄을 걷어내 서명부가 다음 장으로 밀리지 않게 한다.
    # trim_spacers 를 주면 그 개수만큼만, 안 주면 늘어난 줄 수만큼 걷어낸다.
    trim_spacers_after(table, extra_rows if trim_spacers is None else trim_spacers)

    previous = None
    for idx, line in enumerate(instruction.lines):
        cells = _cells(table, ITEM_ROW_INDEX + idx)
        values = {
            "NO": str(idx + 1),
            "FUND": line.fund_label,
            "APPLY_DATE": f"{line.apply_date:%Y-%m-%d}",
            "QTY": f"{line.qty:,}주",
            "AMOUNT": f"{line.amount:,}원",
            "TRADE_DESC": f"{line.trade_name}\n{line.detail}",
        }
        # 손으로 쓰던 방식: 해당 펀드 / 반영일자 / 거래유형명은 윗줄과 같으면 비워둔다.
        shown = dict(values)
        if previous is not None:
            for key in DITTO_KEYS:
                if values[key] == previous[key]:
                    shown[key] = ""
        previous = values
        for col_idx, key in enumerate(COLUMNS[instruction.kind]):
            set_cell_text(cells[col_idx], shown[key])

    if not instruction.lines:                      # 내역이 없으면 템플릿 행 제거
        delete_row(table, ITEM_ROW_INDEX)

    replace_everywhere(doc, instruction.context())

    # 사용인감
    if insert_seal is None:
        insert_seal = config.INSERT_SEAL
    if insert_seal:
        if seal_mod.count(doc) == 0:
            print("   [경고] 템플릿에 사용인감 그림이 없습니다. 템플릿을 확인하세요.")
        elif not Path(config.SEAL_IMAGE).exists():
            print(f"   [참고] 인감 원본을 못 찾아 템플릿의 이미지를 사용합니다: {config.SEAL_IMAGE}")
        else:
            seal_mod.refresh(doc)
    else:
        seal_mod.remove(doc)

    output_dir = resolve_output_dir(instruction, output_dir, subdir_by_date)
    output_dir.mkdir(parents=True, exist_ok=True)
    out_path = output_dir / instruction.filename
    if not overwrite:
        stem, suffix, n = out_path.stem, out_path.suffix, 2
        while out_path.exists():
            out_path = output_dir / f"{stem}({n}){suffix}"
            n += 1
    doc.save(str(out_path))
    return out_path
