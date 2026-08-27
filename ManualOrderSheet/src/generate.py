# -*- coding: utf-8 -*-
"""수기운용지시서 생성 CLI.

    python -m src.generate 20260825
    python -m src.generate data/input/20260825
    python -m src.generate data/input/20260825/권리배정내역_20260825093039.xls -o data/output
"""
from __future__ import annotations  # Python 3.9 호환 (X | None 표기)
import argparse
import sys
import tempfile
from pathlib import Path

import pymupdf

try:
    from . import config
    from .reader import read_all, find_input_files
    from .instruction import build_instructions
    from .writer import render
    from .share import copy_outputs
    from .pdf import WordConverter, make_pdfs
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src import config
    from src.reader import read_all, find_input_files
    from src.instruction import build_instructions
    from src.writer import render
    from src.share import copy_outputs
    from src.pdf import WordConverter, make_pdfs


def latest_input_dir() -> Path:
    """data/input 아래에서 가장 최근 날짜 폴더를 고른다. (편집기에서 인자 없이 실행할 때)"""
    candidates = sorted(
        (p for p in config.INPUT_DIR.iterdir() if p.is_dir() and p.name.isdigit()),
        key=lambda p: p.name,
    )
    if not candidates:
        raise SystemExit(f"{config.INPUT_DIR} 아래에 날짜 폴더가 없습니다.")
    return candidates[-1]


def resolve_input(value: str | None) -> Path:
    """'20260825' 처럼 폴더명만 줘도 data/input 아래에서 찾아준다. 생략하면 최신 폴더."""
    if not value:
        target = latest_input_dir()
        print(f"(입력 미지정 → 최신 폴더 사용: {target.name})")
        return target
    path = Path(value)
    if path.exists():
        return path
    candidate = config.INPUT_DIR / value
    if candidate.exists():
        return candidate
    raise SystemExit(f"입력 경로를 찾을 수 없습니다: {value}")


MAX_TRIM = 4          # 표 아래에서 걷어낼 수 있는 빈 줄의 상한


def default_trim(instruction) -> int:
    """Word 없이 쓰는 어림값. 늘어난 줄 수보다 하나 적게 걷어낸다."""
    return max(0, len(instruction.lines) - 2)


def fit_spacers(conv, instruction, args) -> int:
    """빈 줄을 몇 개 걷어내면 한 장에 들어가는지 실제로 재본다.

    임시 폴더에 문서를 만들어 Word 로 페이지 수를 확인한다. 어림값에서 출발해
    한 장이 될 때까지 하나씩 늘리므로 보통 한 번만 변환한다.
    """
    start = default_trim(instruction)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        for trim in range(start, MAX_TRIM + 1):
            probe = render(instruction, tmp_dir, Path(args.template),
                           subdir_by_date=False, insert_seal=not args.no_seal,
                           trim_spacers=trim)
            pdf = tmp_dir / "probe.pdf"
            if conv.to_pdf(probe, pdf) is None:
                return start
            with pymupdf.open(str(pdf)) as doc:
                pages = len(doc)
            pdf.unlink(missing_ok=True)
            if pages <= 1:
                return trim
    print(f"   [경고] 한 장에 담기지 않습니다 (내역 {len(instruction.lines)}줄). "
          f"템플릿의 표 글자 크기나 여백을 조정하세요.")
    return MAX_TRIM


def main(argv=None):
    parser = argparse.ArgumentParser(description="권리배정내역 엑셀 -> 수기운용지시서(.docx)")
    parser.add_argument("input", nargs="?", default=None,
                        help="엑셀 파일 또는 폴더 (data/input 하위 폴더명만 써도 됨). "
                             "생략하면 data/input 의 최신 날짜 폴더")
    parser.add_argument("-o", "--output", default=str(config.OUTPUT_DIR), help="출력 폴더")
    parser.add_argument("-t", "--template", default=str(config.TEMPLATE_PATH), help="템플릿 .docx")
    parser.add_argument("--seq", type=int, default=1, help="문서번호 시작 일련번호 (기본 1)")
    parser.add_argument("--no-overwrite", action="store_true", help="같은 이름이 있으면 (2) 를 붙여 저장")
    parser.add_argument("--flat", action="store_true",
                        help="날짜 하위 폴더를 만들지 않고 출력 폴더에 바로 저장")
    parser.add_argument("--no-pdf", action="store_true", help="PDF 를 만들지 않음")
    parser.add_argument("--keep-text-pdf", action="store_true",
                        help="이미지 PDF 외에 텍스트 PDF(.pdf)도 남김 (인감 추출 가능)")
    parser.add_argument("--pdf-dpi", type=int, default=None,
                        help=f"이미지 PDF 해상도 (기본 {config.PDF_DPI})")
    parser.add_argument("--no-seal", action="store_true",
                        help="사용인감 도장을 넣지 않고 '(인)' 만 남김 (초안용)")
    parser.add_argument("--no-shared", action="store_true",
                        help="공유 드라이브 사본 저장을 하지 않음")
    parser.add_argument("--shared-dir", default=None,
                        help=f"공유 드라이브 경로 (기본 {config.SHARED_OUTPUT_DIR})")
    parser.add_argument("--overwrite-shared", action="store_true",
                        help="공유 폴더에 같은 이름 파일이 있어도 덮어씀 (기본은 건너뜀)")
    parser.add_argument("--dry-run", action="store_true",
                        help="공유 드라이브에 실제로 쓰지 않고 무엇이 복사될지만 출력")
    args = parser.parse_args(argv)

    target = resolve_input(args.input)
    print(f"입력: {target}")
    source_files = find_input_files(target)
    for f in source_files:
        print(f"  - {f.name}")

    records = read_all(target)
    if not records:
        raise SystemExit("읽어들인 권리배정 내역이 없습니다.")
    print(f"내역 {len(records)}건")

    instructions = build_instructions(records, start_seq=args.seq)
    outputs = []
    rendered = []
    shared_seen = set()
    pdfs = {}

    make_pdf = config.MAKE_PDF and not args.no_pdf
    with WordConverter() as conv:
        if make_pdf and not conv.available:
            print(f"\n[경고] PDF 변환을 건너뜁니다 - {conv.error}")

        for ins in instructions:
            print(f"\n[지시서] 제{ins.doc_no} | {ins.trade_name} | 수신 {ins.bank} / 참조 {ins.pbs}")
            for i, line in enumerate(ins.lines, 1):
                print(f"   {i}. {line.fund_label} {line.qty:,}주 {line.amount:,}원 {line.detail}")
            print(f"   합계 {ins.total_qty:,}주 / {ins.total_amount:,}원")

            trim = fit_spacers(conv, ins, args) if conv.available else default_trim(ins)
            path = render(ins, Path(args.output), Path(args.template),
                          overwrite=not args.no_overwrite,
                          subdir_by_date=not args.flat,
                          insert_seal=not args.no_seal,
                          trim_spacers=trim)
            outputs.append(path)
            rendered.append((ins, path))
            print(f"   -> {path}")

            if make_pdf and conv.available:
                made = make_pdfs(conv, path, dpi=args.pdf_dpi,
                                 image_only=config.PDF_IMAGE_ONLY,
                                 keep_text_pdf=args.keep_text_pdf or None)
                pdfs[path] = made
                for m in made:
                    print(f"   -> {m.name}")
                    outputs.append(m)

    # --- 공유 드라이브 사본 --------------------------------------------
    if not args.no_shared and config.SHARED_OUTPUT_DIR:
        print("\n[공유 드라이브 사본]")
        for ins, path in rendered:
            copy_outputs(ins, [path] + pdfs.get(path, []), source_files,
                         root=args.shared_dir,
                         overwrite=args.overwrite_shared or None,
                         dry_run=args.dry_run,
                         seen=shared_seen)

    print(f"\n총 {len(outputs)}개 파일 생성 완료.")
    return outputs


if __name__ == "__main__":
    main()
