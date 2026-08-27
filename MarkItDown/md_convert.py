#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
문서 및 이미지 -> Markdown 일괄 변환기 (MarkItDown + Tesseract OCR 기반)

사용법 (Windows, 사내 PC):
    pip install "markitdown[all]" pywin32 pytesseract Pillow pandas pdf2image
    python md_convert.py "Z:\\AX\\Operations\\001.수요예측"

주요 옵션:
    --out  <경로>         결과 저장 위치 (기본: 원본폴더명 + "_md" 를 바탕화면에 생성)
    --inplace             원본 폴더 안에 _md 폴더를 만들어 저장
    --skip-doc            구형 .doc 변환(Word 자동화)을 건너뜀
    --tesseract-cmd <경로> Tesseract 실행 파일(tesseract.exe) 직접 지정
    --lang <언어코드>      OCR 언어 설정 (기본값: kor+eng)
    --no-ocr              OCR 기능 비활성화

동작 원리:
  - 폴더 구조를 그대로 유지한 채 각 문서/이미지를 .md 로 변환
  - 이미지 파일(.png, .jpg 등): Tesseract OCR을 통해 텍스트를 추출하여 .md 로 변환
  - 텍스트 미추출 스캔 PDF: pdf2image + Tesseract OCR을 통해 폴백 변환 시도
  - 구형 .doc: 설치된 MS Word COM 자동화로 .docx 변환 후 파싱
  - 결과 요약(_conversion_log.md, _conversion_log.csv): 성공/실패/검토 필요 문서를 pandas로 집계 및 정리
"""

import argparse
import os
import shutil
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

# Windows 콘솔 출력 인코딩 안전화
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# 지원하는 파일 확장자 목록
DOC_EXTS = {
    ".pdf", ".docx", ".doc", ".xlsx", ".xlsm", ".xls", ".pptx", ".ppt",
    ".msg", ".html", ".htm", ".csv", ".tsv", ".json", ".xml", ".txt",
    ".epub", ".zip",
}
IMAGE_EXTS = {
    ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp",
}
SUPPORTED = DOC_EXTS | IMAGE_EXTS

# 이 글자 수 미만이면 본문 부족으로 간주 -> 스캔 PDF의 경우 OCR 시도 또는 '검토 필요' 처리
EMPTY_THRESHOLD = 120

# Tesseract Windows 기본 탐색 경로 목록
TESSERACT_DEFAULT_PATHS = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Tesseract-OCR" / "tesseract.exe",
]


def init_tesseract(custom_cmd: Optional[str] = None) -> bool:
    """
    Tesseract OCR 환경을 초기화하고 실행 가능 여부를 확인합니다.

    Args:
        custom_cmd: 사용자가 지정한 tesseract.exe 경로 (선택 사항)

    Returns:
        bool: Tesseract 사용 가능 여부
    """
    print("[초기화] Tesseract OCR 환경 확인 중...")
    try:
        import pytesseract

        # 1. 사용자 지정 경로 확인
        if custom_cmd:
            cmd_path = Path(custom_cmd)
            if cmd_path.is_file():
                pytesseract.pytesseract.tesseract_cmd = str(cmd_path.resolve())
                print(f"[초기화] 사용자 지정 Tesseract 경로 적용: {cmd_path}")
                return True
            else:
                print(f"[경고] 지정된 Tesseract 경로를 찾을 수 없습니다: {custom_cmd}")

        # 2. 기본 설치 경로 탐색
        for p in TESSERACT_DEFAULT_PATHS:
            if p.is_file():
                pytesseract.pytesseract.tesseract_cmd = str(p.resolve())
                print(f"[초기화] Tesseract 기본 경로 감지: {p}")
                return True

        # 3. 시스템 PATH 확인 (버전 호출 테스트)
        pytesseract.get_tesseract_version()
        print("[초기화] 시스템 PATH의 Tesseract 감지 성공")
        return True

    except Exception as e:
        print(f"[안내] Tesseract OCR을 사용할 수 없습니다 (OCR 기능 비활성화): {e}")
        return False


def ocr_image(img_path: Path, lang: str = "kor+eng") -> str:
    """
    단일 이미지 파일에서 Tesseract OCR을 이용하여 텍스트를 추출합니다.

    Args:
        img_path: 대상 이미지 파일 경로
        lang: Tesseract 언어 설정 (기본값: 'kor+eng')

    Returns:
        str: OCR로 추출된 텍스트
    """
    try:
        from PIL import Image
        import pytesseract

        with Image.open(img_path) as img:
            # RGB 변환 (RGBA 또는 Palette 모드 호환)
            if img.mode not in ("RGB", "L"):
                img = img.convert("RGB")
            extracted_text = pytesseract.image_to_string(img, lang=lang)
            return extracted_text.strip()
    except Exception as e:
        print(f"[예외 발생] 이미지 OCR 실패 ({img_path.name}): {e}")
        return ""


def ocr_pdf(pdf_path: Path, lang: str = "kor+eng", tmpdir: Optional[Path] = None) -> str:
    """
    스캔된 PDF 문서의 각 페이지를 이미지로 렌더링한 후 Tesseract OCR로 텍스트를 추출합니다.

    Args:
        pdf_path: 대상 PDF 파일 경로
        lang: Tesseract 언어 설정
        tmpdir: 임시 파일 저장 디렉토리

    Returns:
        str: 추출된 전체 페이지 텍스트 (Markdown 페이지 구분 포함)
    """
    try:
        from pdf2image import convert_from_path
        import pytesseract

        pages = convert_from_path(str(pdf_path.resolve()), dpi=200, output_folder=str(tmpdir) if tmpdir else None)
        extracted_pages = []
        for idx, page_img in enumerate(pages, 1):
            text = pytesseract.image_to_string(page_img, lang=lang).strip()
            if text:
                extracted_pages.append(f"### [페이지 {idx}]\n\n{text}")
        return "\n\n---\n\n".join(extracted_pages)
    except Exception as e:
        print(f"[예외 발생] PDF OCR 실패 ({pdf_path.name}): {e}")
        return ""


def convert_legacy_doc(src: Path, tmpdir: Path) -> Path:
    """
    구형 .doc 문서를 MS Word COM 자동화를 통해 .docx 로 변환합니다.

    Args:
        src: 원본 .doc 파일 경로
        tmpdir: 변환된 .docx를 저장할 임시 디렉토리

    Returns:
        Path: 생성된 .docx 파일 경로
    """
    try:
        import win32com.client as win32
    except ImportError:
        raise RuntimeError("pywin32 미설치 (pip install pywin32) - .doc 변환 불가")

    out = tmpdir / (src.stem + ".docx")
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(src.resolve()), ReadOnly=True, ConfirmConversions=False)
        # 12 = wdFormatXMLDocument (.docx)
        doc.SaveAs2(str(out.resolve()), FileFormat=12)
        doc.Close(SaveChanges=0)
    finally:
        try:
            word.Quit()
        except Exception:
            pass

    if not out.exists():
        raise RuntimeError("Word 변환 결과 파일(.docx)이 생성되지 않음")
    return out


def convert_single_file(
    src: Path,
    root: Path,
    out_root: Path,
    tmpdir: Path,
    md_engine,
    ocr_available: bool,
    ocr_lang: str,
    skip_doc: bool,
) -> Dict[str, object]:
    """
    개별 파일(문서 또는 이미지)을 Markdown(.md) 파일로 변환합니다.

    Args:
        src: 원본 파일 경로
        root: 원본 루트 폴더 경로
        out_root: 결과 루트 폴더 경로
        tmpdir: 임시 폴더 경로
        md_engine: MarkItDown 인스턴스 (또는 None)
        ocr_available: OCR 사용 가능 여부
        ocr_lang: OCR 언어 코드
        skip_doc: 구형 doc 건너뛰기 여부

    Returns:
        Dict: 변환 결과 정보 (파일, 상태, 추출문자수, 비고, 변환방식)
    """
    rel = src.relative_to(root)
    dst = out_root / rel.parent / (src.stem + ".md")
    dst.parent.mkdir(parents=True, exist_ok=True)

    status = "OK"
    note = ""
    chars = 0
    method = "MarkItDown"
    suffix = src.suffix.lower()

    try:
        text = ""

        # 1. 이미지 파일인 경우 -> Tesseract OCR 직접 수행
        if suffix in IMAGE_EXTS:
            if ocr_available:
                method = f"Tesseract OCR ({ocr_lang})"
                text = ocr_image(src, lang=ocr_lang)
                chars = len(text.strip())
                if chars < EMPTY_THRESHOLD:
                    status = "검토 필요"
                    note = f"이미지 텍스트가 매우 적거나 없음({chars}자) - 해상도 또는 원본 확인 필요"
                else:
                    note = f"이미지 OCR 완료 ({chars}자 추출)"
            else:
                status = "검토 필요"
                note = "Tesseract OCR 미설치/비활성화로 이미지 텍스트 미추출"
                text = f"*(이미지 파일: OCR 비활성화로 텍스트가 추출되지 않았습니다: `{src.name}`)*"

        # 2. 구형 .doc 파일인 경우 -> Word COM 자동화로 docx 경유 변환
        elif suffix == ".doc":
            if skip_doc:
                raise RuntimeError("건너뜀(--skip-doc): 구형 .doc")
            target_docx = convert_legacy_doc(src, tmpdir)
            note = "Word로 .docx 경유 변환"
            if md_engine:
                text = md_engine.convert(str(target_docx)).text_content or ""
            chars = len(text.strip())

        # 3. 일반 문서 파일 (PDF, docx, xlsx, pptx, txt 등) -> MarkItDown 변환
        else:
            if md_engine:
                try:
                    text = md_engine.convert(str(src)).text_content or ""
                except Exception as ex_md:
                    # MarkItDown 실패 시 일반 텍스트 파일이면 직접 읽기 시도
                    if suffix in {".txt", ".csv", ".tsv", ".json", ".xml", ".html", ".htm"}:
                        try:
                            text = src.read_text(encoding="utf-8", errors="replace")
                            method = "Direct Text Read"
                        except Exception:
                            text = ""
                            note = f"MarkItDown 실패({ex_md})"
                    else:
                        text = ""
                        note = f"MarkItDown 실패({ex_md})"
            else:
                if suffix in {".txt", ".csv", ".tsv", ".json", ".xml", ".html", ".htm"}:
                    text = src.read_text(encoding="utf-8", errors="replace")
                    method = "Direct Text Read"

            chars = len(text.strip())

            # PDF 문서인데 추출 글자 수가 너무 적은 경우 (스캔/날인 문서) -> OCR 우회 시도
            if suffix == ".pdf" and chars < EMPTY_THRESHOLD and ocr_available:
                ocr_text = ocr_pdf(src, lang=ocr_lang, tmpdir=tmpdir)
                ocr_chars = len(ocr_text.strip())
                if ocr_chars > chars:
                    text = ocr_text
                    chars = ocr_chars
                    method = f"MarkItDown -> PDF OCR 폴백 ({ocr_lang})"
                    note = f"스캔 PDF 감지되어 OCR 적용 ({chars}자 추출)"

            # 본문 추출 결과 상태 판정 (규칙 6 준수: 결측/부족 시 '검토 필요' 분류)
            if chars < EMPTY_THRESHOLD:
                status = "검토 필요"
                extra = f"추출 텍스트가 매우 적음({chars}자) - 스캔/서식 확인 권장"
                note = (note + " / " if note else "") + extra

        # Markdown 파일 헤더 생성 및 파일 쓰기 (규칙 8 준수: 원본 보존, 별도 파일 저장)
        header = (
            f"---\n"
            f'source: "{rel.as_posix()}"\n'
            f"source_bytes: {src.stat().st_size}\n"
            f"converted_by: {method}\n"
            f"extracted_characters: {chars}\n"
            f"---\n\n"
        )
        dst.write_text(header + text, encoding="utf-8")

    except Exception as e:
        status = "실패"
        note = f"{type(e).__name__}: {e}"
        chars = 0

    return {
        "파일": rel.as_posix(),
        "상태": status,
        "추출문자수": chars,
        "변환방식": method,
        "비고": note,
    }


def save_conversion_reports(rows: List[Dict[str, object]], root: Path, out_root: Path) -> None:
    """
    변환 결과를 pandas DataFrame을 활용하여 집계하고 요약 Markdown 및 CSV 리포트를 생성합니다.
    (규칙 4: pandas 벡터화 집계, 규칙 5/6: 검토 필요 및 실패 내역 명시)

    Args:
        rows: 파일별 변환 결과 리스트
        root: 원본 디렉토리 경로
        out_root: 출력 디렉토리 경로
    """
    print("\n[리포트 생성] 변환 결과 요약 보고서 작성 중...")

    if not rows:
        print("[경고] 변환 대상 파일이 없습니다.")
        return

    df = pd.DataFrame(rows)

    # 상태별 집계 (pandas 벡터화 연산)
    total_count = len(df)
    status_counts = df["상태"].value_counts().to_dict()
    ok_count = status_counts.get("OK", 0)
    review_count = status_counts.get("검토 필요", 0)
    fail_count = status_counts.get("실패", 0)

    # Markdown 보고서 작성
    log_md = out_root / "_conversion_log.md"
    with log_md.open("w", encoding="utf-8") as f:
        f.write("# 변환 결과 요약 보고서\n\n")
        f.write(f"- **원본 폴더**: `{root.resolve()}`\n")
        f.write(f"- **결과 폴더**: `{out_root.resolve()}`\n")
        f.write(
            f"- **총 파일 수**: {total_count}건 (정상: **{ok_count}**건 / "
            f"검토 필요: **{review_count}**건 / 실패: **{fail_count}**건)\n\n"
        )

        f.write("## 상세 변환 내역\n\n")
        f.write("| 파일 | 상태 | 추출 문자수 | 변환 방식 | 비고 |\n")
        f.write("|---|:---:|---:|---|---|\n")
        for _, r in df.iterrows():
            f.write(
                f"| {r['파일']} | {r['상태']} | {r['추출문자수']:,} | {r['변환방식']} | {r['비고']} |\n"
            )

        if review_count > 0:
            f.write("\n## ⚠️ 검토 필요 파일 목록\n\n")
            review_df = df[df["상태"] == "검토 필요"]
            for _, r in review_df.iterrows():
                f.write(f"- `{r['파일']}` ({r['추출문자수']}자): {r['비고']}\n")

        if fail_count > 0:
            f.write("\n## ❌ 변환 실패 파일 목록\n\n")
            fail_df = df[df["상태"] == "실패"]
            for _, r in fail_df.iterrows():
                f.write(f"- `{r['파일']}`: {r['비고']}\n")

    # CSV 보고서 작성 (Excel 호환 utf-8-sig)
    log_csv = out_root / "_conversion_log.csv"
    df.to_csv(log_csv, index=False, encoding="utf-8-sig")

    print(f"[리포트 완료] 요약 MD: {log_md}")
    print(f"[리포트 완료] 요약 CSV: {log_csv}")
    print(f"최종 집계: 총 {total_count}건 -> 정상 {ok_count} / 검토 필요 {review_count} / 실패 {fail_count}")


def main():
    """메인 실행 함수: CLI 인자 파싱, 변환 파이프라인 제어 및 진행 상황 출력"""
    ap = argparse.ArgumentParser(description="문서 및 이미지 파일의 Markdown 일괄 변환기 (Tesseract OCR 통합)")
    ap.add_argument("root", help="변환할 원본 폴더 경로")
    ap.add_argument("--out", default=None, help="결과 저장 폴더 경로")
    ap.add_argument("--inplace", action="store_true", help="원본 폴더 내부에 _md 폴더 생성하여 저장")
    ap.add_argument("--skip-doc", action="store_true", help="구형 .doc 변환(Word 자동화) 건너뛰기")
    ap.add_argument("--tesseract-cmd", default=None, help="Tesseract 실행 파일(tesseract.exe) 경로 직접 지정")
    ap.add_argument("--lang", default="kor+eng", help="OCR 대상 언어 (기본값: kor+eng)")
    ap.add_argument("--no-ocr", action="store_true", help="OCR 기능 비활성화")
    args = ap.parse_args()

    print("==================================================")
    print("      문서 및 이미지 -> Markdown 일괄 변환기       ")
    print("==================================================")

    root = Path(args.root)
    if not root.is_dir():
        sys.exit(f"[오류] 폴더를 찾을 수 없습니다: {root}")

    # 출력 폴더 경로 결정 (규칙 7, 8 준수)
    if args.out:
        out_root = Path(args.out)
    elif args.inplace:
        out_root = root / "_md"
    else:
        desktop = Path.home() / "Desktop"
        out_root = (desktop if desktop.is_dir() else Path.cwd()) / (root.name + "_md")
    out_root.mkdir(parents=True, exist_ok=True)

    # 1. Tesseract OCR 환경 초기화
    ocr_available = False
    if not args.no_ocr:
        ocr_available = init_tesseract(args.tesseract_cmd)
    else:
        print("[안내] --no-ocr 옵션에 의해 OCR 기능이 비활성화되었습니다.")

    # 2. MarkItDown 엔진 로드 (규칙 11 준수: try-except)
    md_engine = None
    try:
        from markitdown import MarkItDown
        md_engine = MarkItDown(enable_plugins=True)
        print("[초기화] MarkItDown 엔진 로드 완료")
    except ImportError:
        print("[경고] markitdown 라이브러리가 설치되지 않았습니다. 일반 문서 변환이 제한될 수 있습니다.")

    # 3. 대상 파일 목록 탐색
    files = [
        p for p in sorted(root.rglob("*"))
        if p.is_file()
        and p.suffix.lower() in SUPPORTED
        and "_md" not in p.parts
        and not p.name.startswith("~$")
    ]

    print(f"\n[작업 시작] 변환 대상 파일: 총 {len(files)}건")
    print(f"[작업 경로] 원본: {root.resolve()}")
    print(f"[작업 경로] 출력: {out_root.resolve()}\n")

    rows = []
    tmpdir = Path(tempfile.mkdtemp(prefix="md_convert_"))

    try:
        for i, src in enumerate(files, 1):
            result = convert_single_file(
                src=src,
                root=root,
                out_root=out_root,
                tmpdir=tmpdir,
                md_engine=md_engine,
                ocr_available=ocr_available,
                ocr_lang=args.lang,
                skip_doc=args.skip_doc,
            )
            rows.append(result)

            status_mark = {
                "OK": "  OK  ",
                "검토 필요": " CHK? ",
                "실패": " FAIL ",
            }.get(result["상태"], "  ??  ")

            note_str = f" ({result['비고']})" if result["비고"] else ""
            print(f"[{i:>3}/{len(files)}] [{status_mark}] {result['파일']} [{result['변환방식']}]{note_str}")

    finally:
        # 임시 디렉토리 정리
        shutil.rmtree(tmpdir, ignore_errors=True)

    # 4. 요약 보고서 생성
    save_conversion_reports(rows, root, out_root)
    print("\n[작업 완료] 모든 변환 및 리포트 작성이 완료되었습니다.")


if __name__ == "__main__":
    main()