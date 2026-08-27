# -*- coding: utf-8 -*-
"""생성된 지시서를 공유 드라이브(Z:)에도 사본 저장한다.

공유 폴더는 실제 업무 공문이 쌓이는 곳이므로 기본 정책은 다음과 같다.
  - 같은 이름 파일이 이미 있으면 덮어쓰지 않고 건너뛴다 (서명본 보호)
  - 드라이브가 연결돼 있지 않으면 경고만 남기고 로컬 저장은 그대로 성공 처리
"""
from __future__ import annotations  # Python 3.9 호환 (X | None 표기)
import shutil
import sys
from pathlib import Path

try:
    from . import config
    from .instruction import Instruction
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src import config
    from src.instruction import Instruction


def subdir_name(instruction: Instruction) -> str:
    return config.SHARED_SUBDIR_FORMAT.format(
        date=f"{instruction.pay_date:%Y%m%d}",
        trade_key=instruction.trade_name.replace(" ", ""),
    )


def resolve_shared_dir(instruction: Instruction, root: Path) -> Path:
    """공유 폴더 안의 대상 폴더.

    기존 폴더명이 'YYYYMMDD_차입주식 배당지급' 처럼 띄어쓰기가 다를 수 있어,
    같은 날짜·같은 거래유형의 기존 폴더가 있으면 새로 만들지 않고 그대로 쓴다.
    """
    wanted = subdir_name(instruction)
    exact = root / wanted
    if exact.is_dir():
        return exact

    date_prefix = f"{instruction.pay_date:%Y%m%d}_"
    trade_key = instruction.trade_name.replace(" ", "")
    try:
        for existing in sorted(p for p in root.iterdir() if p.is_dir()):
            if not existing.name.startswith(date_prefix):
                continue
            if existing.name.replace(" ", "").endswith(trade_key):
                return existing
    except OSError:
        pass
    return exact


def copy_outputs(instruction: Instruction, local_files: list[Path],
                 source_files: list[Path] | None = None,
                 root: Path | None = None,
                 overwrite: bool | None = None,
                 dry_run: bool = False,
                 seen: set | None = None) -> list[Path]:
    """local_files 를 공유 폴더에 복사하고 복사된 경로 목록을 돌려준다.

    seen: 한 번의 실행에서 이미 처리한 (폴더, 파일명) 집합. 여러 지시서가 같은
    공유 폴더를 쓸 때 원본 엑셀이 중복 보고되지 않게 한다.
    """
    root = Path(root) if root else config.SHARED_OUTPUT_DIR
    if root is None:
        return []
    if overwrite is None:
        overwrite = config.SHARED_OVERWRITE

    if not root.exists():
        print(f"   [경고] 공유 드라이브에 접근할 수 없어 사본을 건너뜁니다: {root}")
        return []

    target_dir = resolve_shared_dir(instruction, root)
    to_copy = list(local_files)
    if config.SHARED_COPY_INPUT and source_files:
        to_copy += [p for p in source_files if p.exists()]

    label = "[예정] " if dry_run else ""
    if not target_dir.exists():
        print(f"   {label}공유 폴더 생성: {target_dir}")
    else:
        print(f"   {label}공유 폴더: {target_dir}")

    copied: list[Path] = []
    for src in to_copy:
        dest = target_dir / src.name
        if seen is not None:
            if dest in seen:
                continue
            seen.add(dest)
        if dest.exists() and not overwrite:
            print(f"   {label}건너뜀(이미 있음): {dest.name}")
            continue
        verb = "덮어씀" if dest.exists() else "복사"
        print(f"   {label}{verb}: {dest.name}")
        if not dry_run:
            target_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dest)
        copied.append(dest)
    return copied
