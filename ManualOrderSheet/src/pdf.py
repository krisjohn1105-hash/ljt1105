# -*- coding: utf-8 -*-
"""docx -> PDF 변환. 인감·텍스트를 복사할 수 없도록 전체를 이미지로 굽는다.

두 단계로 처리한다.
  1) Word 로 docx -> PDF (텍스트 PDF, 임시)
  2) 각 페이지를 이미지로 렌더링해 이미지만 담은 PDF 로 다시 저장

결과물은 텍스트 레이어가 없어 글자 선택·복사가 되지 않고, 인감도 페이지 전체
이미지의 일부라 따로 빼낼 수 없다. 기존 공문의 '_E.pdf' 와 같은 형태다.
"""
from __future__ import annotations  # Python 3.9 호환 (X | None 표기)
import sys
import tempfile
from pathlib import Path

try:
    from . import config
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src import config

WD_EXPORT_FORMAT_PDF = 17


class WordConverter:
    """Word 를 한 번만 띄워 여러 문서를 변환한다.

        with WordConverter() as conv:
            conv.to_pdf(docx_path, pdf_path)

    Word 가 없거나 COM 을 못 쓰면 available=False 가 되고, to_pdf 는 None 을 준다.
    """

    def __init__(self):
        self._app = None
        self.error = None

    @property
    def available(self) -> bool:
        return self._app is not None

    def __enter__(self):
        try:
            import win32com.client as win32
        except ImportError:
            self.error = "pywin32 가 설치되지 않았습니다 (pip install pywin32)"
            return self
        try:
            # DispatchEx: 사용자가 열어둔 Word 창을 건드리지 않도록 별도 인스턴스
            app = win32.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            self._app = app
        except Exception as exc:  # Word 미설치 등
            self.error = f"Word 를 실행할 수 없습니다: {exc}"
        return self

    def __exit__(self, *exc_info):
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None
        return False

    def to_pdf(self, docx_path, pdf_path):
        if not self.available:
            return None
        docx_path, pdf_path = Path(docx_path), Path(pdf_path)
        pdf_path.parent.mkdir(parents=True, exist_ok=True)
        doc = None
        try:
            doc = self._app.Documents.Open(
                str(docx_path.resolve()), ConfirmConversions=False,
                ReadOnly=True, AddToRecentFiles=False, Visible=False,
            )
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path.resolve()),
                ExportFormat=WD_EXPORT_FORMAT_PDF,
            )
        except Exception as exc:
            print(f"   [경고] PDF 변환 실패: {docx_path.name} -> {exc}")
            return None
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
        return pdf_path if pdf_path.exists() else None


def rasterize(src_pdf, dest_pdf, dpi=None):
    """PDF 각 페이지를 이미지로 굽어 텍스트 레이어가 없는 PDF 로 저장한다.

    페이지 수를 돌려준다.
    """
    import pymupdf

    dpi = dpi or config.PDF_DPI
    src = pymupdf.open(str(src_pdf))
    out = pymupdf.open()
    try:
        for page in src:
            pixmap = page.get_pixmap(dpi=dpi)
            new_page = out.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(new_page.rect, pixmap=pixmap)
        # Word 가 남긴 작성자·경로 등 메타데이터는 남기지 않는다.
        out.set_metadata({})
        dest_pdf = Path(dest_pdf)
        dest_pdf.parent.mkdir(parents=True, exist_ok=True)
        out.save(str(dest_pdf), deflate=True, garbage=4)
        pages = len(out)
    finally:
        out.close()
        src.close()
    return pages


def make_pdfs(converter, docx_path, dpi=None, image_only=True, keep_text_pdf=None):
    """docx 하나에서 PDF 를 만들고 생성된 경로 목록을 돌려준다."""
    docx_path = Path(docx_path)
    if keep_text_pdf is None:
        keep_text_pdf = config.KEEP_TEXT_PDF

    text_pdf = docx_path.with_suffix(".pdf")
    made = []

    if not image_only:
        result = converter.to_pdf(docx_path, text_pdf)
        return [result] if result else []

    # 이미지 PDF 를 만들려면 중간 텍스트 PDF 가 필요하다.
    if keep_text_pdf:
        staging = text_pdf
    else:
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp.close()
        staging = Path(tmp.name)

    result = converter.to_pdf(docx_path, staging)
    if result is None:
        if not keep_text_pdf:
            staging.unlink(missing_ok=True)
        return []
    if keep_text_pdf:
        made.append(text_pdf)

    image_pdf = docx_path.with_name(docx_path.stem + config.PDF_IMAGE_SUFFIX + ".pdf")
    try:
        pages = rasterize(staging, image_pdf, dpi)
        made.append(image_pdf)
        if pages > 1:
            # 내역이 많으면 두 장이 된다(기존 공문도 마찬가지). 오류는 아니지만
            # 서명부가 어디에 찍혔는지 눈으로 확인하는 게 좋다.
            print(f"   [참고] {pages}페이지 문서입니다 - 서명부 위치를 확인하세요.")
    except Exception as exc:
        print(f"   [경고] 이미지화 실패: {docx_path.name} -> {exc}")
    finally:
        if not keep_text_pdf:
            staging.unlink(missing_ok=True)
    return made
