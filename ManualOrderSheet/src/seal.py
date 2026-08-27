# -*- coding: utf-8 -*-
"""사용인감(도장) 이미지 처리.

도장은 서명부 문단('대표이사 ... (인)' 바로 위 회사명 문단)에 떠 있는(anchor)
그림으로 템플릿에 들어 있다. 위치·크기는 원본 지시서의 값을 그대로 쓴다.
여기서는 두 가지만 한다.
  - refresh(): 템플릿 안의 도장 이미지를 공유 드라이브의 원본 PNG 로 갱신
  - remove():  도장을 빼고 (인) 만 남긴 초안용 문서 생성
"""
from __future__ import annotations  # Python 3.9 호환 (X | None 표기)
import sys
from pathlib import Path

try:
    from . import config
except ImportError:
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from src import config

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
DRAWING_TAGS = ("drawing", "pict", "object", "AlternateContent")


def _drawings(doc):
    """본문 문단 안의 그림 요소 목록. (문단, 그림요소) 튜플."""
    found = []
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for child in list(run._element):
                if child.tag.split("}")[-1] in DRAWING_TAGS:
                    found.append((paragraph, child))
    return found


def _embed_ids(element) -> list[str]:
    """그림 요소가 참조하는 이미지 관계 ID 목록."""
    ids = []
    for node in element.iter():
        rid = node.get(R_NS + "embed") or node.get(R_NS + "link")
        if rid and rid not in ids:
            ids.append(rid)
    return ids


def count(doc) -> int:
    return len(_drawings(doc))


def remove(doc) -> int:
    """도장 그림을 모두 제거한다. 제거한 개수를 돌려준다."""
    removed = 0
    for _, element in _drawings(doc):
        element.getparent().remove(element)
        removed += 1
    return removed


def refresh(doc, image_path: Path | None = None) -> bool:
    """템플릿에 박혀 있는 도장 이미지를 원본 PNG 내용으로 갈아끼운다.

    원본 파일이 없으면(공유 드라이브 미연결 등) 템플릿에 들어 있는 이미지를
    그대로 쓰고 False 를 돌려준다.
    """
    image_path = Path(image_path or config.SEAL_IMAGE)
    if not image_path.exists():
        return False
    data = image_path.read_bytes()

    updated = False
    for _, element in _drawings(doc):
        for rid in _embed_ids(element):
            try:
                part = doc.part.related_parts[rid]
            except KeyError:
                continue
            if part.blob != data:
                part._blob = data
                updated = True
    return updated
