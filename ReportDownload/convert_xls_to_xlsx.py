"""
Qube-RT 매매보고서 대사 폴더의 구형 엑셀(.xls) 파일을 .xlsx 포맷으로 일괄 변환합니다.

대상 .xls 파일들은 확장자만 .xls가 아니라 실제 구형 BIFF8 바이너리(D0 CF 11 E0)이므로
단순 이름 변경(rename)으로는 파일이 손상됩니다. 따라서 Excel COM API의 SaveAs를 사용해
실제 포맷 변환(xlOpenXMLWorkbook)을 수행합니다.

사용 예:
    # 변환 대상만 확인 (파일 변경 없음)
    python convert_xls_to_xlsx.py --dry-run

    # 202607 폴더의 모든 날짜별 하위 폴더 변환 (원본 .xls 보존)
    python convert_xls_to_xlsx.py

    # 다른 월 폴더 변환
    python convert_xls_to_xlsx.py --root "Z:/02.펀드/003.매매보고서 대사/Qube-RT/202608"

    # 변환 성공 후 원본 .xls 삭제 (되돌릴 수 없으므로 주의)
    python convert_xls_to_xlsx.py --delete-original
"""

import argparse
import os
import sys
import zipfile
from pathlib import Path

import pythoncom
import win32com.client

# Z 드라이브 네트워크 경로 상수 정의 (사용자 규칙 7 준수)
Z_DRIVE = Path("Z:/")

# 기본 변환 대상 루트 (Qube-RT 202607)
DEFAULT_ROOT = Z_DRIVE / "02.펀드" / "003.매매보고서 대사" / "Qube-RT" / "202607"

# Excel SaveAs FileFormat 상수: 51 = xlOpenXMLWorkbook (.xlsx)
XL_OPEN_XML_WORKBOOK = 51

# 변환 대상 확장자 (소문자 기준)
DEFAULT_SOURCE_EXTS = (".xls",)


def is_valid_xlsx(path: Path) -> bool:
    """
    .xlsx 파일이 정상적인 OOXML(zip) 구조인지 검증합니다.
    원본 삭제 여부를 판단하기 전 안전장치로 사용합니다.
    """
    try:
        if not path.is_file() or path.stat().st_size == 0:
            return False
        with zipfile.ZipFile(path) as zf:
            return "[Content_Types].xml" in zf.namelist()
    except Exception:
        return False


def collect_targets(root: Path, source_exts, overwrite: bool):
    """
    루트 아래의 모든 하위 폴더(날짜별 폴더)를 순회하여 변환 대상 파일 목록을 만듭니다.

    반환값: (변환할 파일 리스트, 이미 변환되어 건너뛴 파일 리스트)
    """
    targets = []
    skipped = []

    for src in sorted(root.rglob("*")):
        if not src.is_file():
            continue
        if src.suffix.lower() not in source_exts:
            continue
        # 엑셀 임시 잠금 파일(~$xxx.xls) 제외
        if src.name.startswith("~$"):
            continue

        dst = src.with_suffix(".xlsx")
        if dst.exists() and not overwrite:
            skipped.append(src)
            continue

        targets.append((src, dst))

    return targets, skipped


def start_excel():
    """
    백그라운드 Excel 인스턴스를 생성하고 변환 중 팝업이 뜨지 않도록 설정합니다.
    """
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False          # 포맷 변환/덮어쓰기 경고 팝업 방지
    excel.AskToUpdateLinks = False       # 외부 링크 업데이트 여부 질문 방지
    excel.AlertBeforeOverwriting = False
    excel.EnableEvents = False           # 통합문서 이벤트 매크로 실행 방지
    return excel


def quit_excel(excel):
    """
    Excel 프로세스를 확실하게 종료합니다.
    """
    if excel is None:
        return
    try:
        excel.Quit()
    except Exception as close_err:
        print(f"[경고] 엑셀 프로세스 종료 시 에러 발생: {close_err}")


def convert_one(excel, src: Path, dst: Path) -> None:
    """
    단일 .xls 파일을 .xlsx로 변환합니다. 실패 시 예외를 그대로 전파합니다.
    """
    wb = None
    try:
        # 읽기 전용으로 열어 원본 변경 및 파일 잠금을 방지
        wb = excel.Workbooks.Open(
            str(src.resolve()),
            UpdateLinks=0,
            ReadOnly=True,
        )

        # 덮어쓰기 모드에서 기존 산출물이 남아 있으면 SaveAs가 실패할 수 있으므로 미리 제거
        if dst.exists():
            dst.unlink()

        wb.SaveAs(Filename=str(dst.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass

    if not is_valid_xlsx(dst):
        raise RuntimeError("변환된 .xlsx 파일이 정상적인 OOXML 구조가 아닙니다.")


def convert_all(root: Path, source_exts, overwrite: bool, delete_original: bool,
                dry_run: bool, restart_every: int) -> int:
    """
    루트 아래 모든 대상 파일을 변환합니다. 반환값은 실패 건수입니다.
    """
    if not root.is_dir():
        print(f"[오류] 대상 디렉토리가 존재하지 않습니다: {root}")
        return 1

    print(f"[작업 시작] 변환 대상 스캔: {root}")
    targets, skipped = collect_targets(root, source_exts, overwrite)

    if skipped:
        print(f"[정보] 이미 .xlsx가 존재하여 건너뛴 파일: {len(skipped)}개 "
              f"(--overwrite 옵션으로 재변환 가능)")

    if not targets:
        print("[작업 완료] 변환할 파일이 없습니다.")
        return 0

    print(f"[정보] 변환 대상: 총 {len(targets)}개 파일")

    if dry_run:
        current_dir = None
        for src, dst in targets:
            if src.parent != current_dir:
                current_dir = src.parent
                print(f"\n[{current_dir.relative_to(root)}]")
            print(f"  {src.name} -> {dst.name}")
        print(f"\n[모의 실행] 실제 변환은 수행하지 않았습니다. 대상 {len(targets)}개.")
        return 0

    success = 0
    failed = []
    deleted = 0
    excel = None
    current_dir = None

    try:
        # COM 라이브러리 초기화
        pythoncom.CoInitialize()
        excel = start_excel()

        for index, (src, dst) in enumerate(targets, start=1):
            # 날짜별 폴더가 바뀔 때마다 구분해서 로그 출력
            if src.parent != current_dir:
                current_dir = src.parent
                print(f"\n[작업 시작] 폴더 변환: {current_dir.relative_to(root)}")

            # 장시간 실행 시 Excel 메모리 누수를 방지하기 위해 주기적으로 재시작
            if restart_every > 0 and index > 1 and (index - 1) % restart_every == 0:
                quit_excel(excel)
                excel = start_excel()
                print(f"[정보] {index - 1}개 처리 후 엑셀 인스턴스를 재시작했습니다.")

            try:
                convert_one(excel, src, dst)
                success += 1
                print(f"  [{index}/{len(targets)}] 변환 성공: {src.name} -> {dst.name}")

                if delete_original:
                    # 산출물 검증(is_valid_xlsx)을 통과한 경우에만 원본을 삭제
                    try:
                        os.chmod(src, 0o666)  # 읽기 전용 속성으로 인한 삭제 실패 방지
                        src.unlink()
                        deleted += 1
                    except Exception as del_err:
                        print(f"  [경고] 원본 삭제 실패: {src.name} ({del_err})")

            except Exception as conv_err:
                failed.append((src, conv_err))
                print(f"  [오류] 변환 실패: {src.name} ({conv_err})")
                # 손상된 산출물이 남지 않도록 정리
                if dst.exists() and not is_valid_xlsx(dst):
                    try:
                        dst.unlink()
                    except Exception:
                        pass
                # 실패로 Excel 상태가 불안정해질 수 있으므로 인스턴스를 재시작
                quit_excel(excel)
                excel = start_excel()

    finally:
        # Excel 자원 해제 및 COM 객체 완전 해제 (try-finally 패턴)
        quit_excel(excel)
        excel = None
        pythoncom.CoUninitialize()

    print("\n" + "=" * 60)
    print(f"[작업 완료] 변환 성공 {success}개 / 실패 {len(failed)}개 / 건너뜀 {len(skipped)}개")
    if delete_original:
        print(f"[작업 완료] 원본 .xls 삭제: {deleted}개")
    if failed:
        print("[오류] 실패 목록:")
        for src, err in failed:
            print(f"  - {src} : {err}")
    print("=" * 60)

    return len(failed)


def parse_args():
    parser = argparse.ArgumentParser(
        description="날짜별 폴더의 .xls 파일을 .xlsx 포맷으로 일괄 변환합니다."
    )
    parser.add_argument(
        "--root",
        type=Path,
        default=DEFAULT_ROOT,
        help=f"변환 대상 루트 디렉토리 (기본값: {DEFAULT_ROOT})",
    )
    parser.add_argument(
        "--ext",
        action="append",
        dest="exts",
        metavar=".xls",
        help="변환 대상 확장자. 여러 번 지정 가능 (기본값: .xls)",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="같은 이름의 .xlsx가 이미 있어도 다시 변환합니다.",
    )
    parser.add_argument(
        "--delete-original",
        action="store_true",
        help="변환 성공 후 원본 .xls를 삭제합니다. (되돌릴 수 없으므로 주의)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="변환 대상 목록만 출력하고 실제 변환은 하지 않습니다.",
    )
    parser.add_argument(
        "--restart-every",
        type=int,
        default=100,
        help="지정한 파일 수마다 엑셀 인스턴스를 재시작합니다. 0이면 재시작하지 않습니다. (기본값: 100)",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    exts = tuple(
        e.lower() if e.startswith(".") else f".{e.lower()}"
        for e in (args.exts or DEFAULT_SOURCE_EXTS)
    )

    print("[시스템 시작] 엑셀 포맷 일괄 변환(.xls -> .xlsx) 작업을 시작합니다.")
    print(f"[정보] 대상 확장자: {', '.join(exts)} / 원본 삭제: {args.delete_original}")

    failed_count = convert_all(
        root=args.root,
        source_exts=exts,
        overwrite=args.overwrite,
        delete_original=args.delete_original,
        dry_run=args.dry_run,
        restart_every=args.restart_every,
    )

    print("[시스템 종료] 엑셀 포맷 변환 작업이 종료되었습니다.")
    return 1 if failed_count else 0


if __name__ == "__main__":
    sys.exit(main())
