# -*- coding: utf-8 -*-
"""수기운용지시서 생성에 필요한 고정값 모음.

실무에서 바뀌는 값(담당자, 계좌, 거래유형 문구 등)은 전부 이 파일에서만 수정한다.
"""
from __future__ import annotations  # Python 3.9 호환 (X | None 표기)
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent

TEMPLATE_DIR = BASE_DIR / "templates"
INPUT_DIR = BASE_DIR / "data" / "input"
OUTPUT_DIR = BASE_DIR / "data" / "output"

TEMPLATE_PATH = TEMPLATE_DIR / "수기운용지시서_template.docx"

# ---------------------------------------------------------------- 발신 정보
COMPANY_NAME = "주식회사 두나미스자산운용"
COMPANY_SHORT = "두나미스자산운용"
CEO_NAME = "김 대 욱"
CONTACT_NAME = "이종택"
CONTACT_PHONE = "02-501-3982"

# 문서번호 접두 (운용지원 제YYYYMMDD-NN)
DOC_NO_PREFIX = "운용지원"

# ------------------------------------------------------- 수취계좌 (중개기관별)
# 권리대금을 실제로 보내는 계좌. 새로운 중개기관이 나오면 여기에 추가한다.
RECEIVE_ACCOUNTS = {
    "한국예탁결제원(증권대차)": "신한은행 305-04-003414 (계좌주: 한국예탁결제원)",
    "한국예탁결제원": "신한은행 305-04-003414 (계좌주: 한국예탁결제원)",
}

# ------------------------------------------------------------- 거래유형명 매핑
# key: (운용구분, 권리구분, 증권구분)  ->  지시서 '거래유형명' 칸에 찍히는 문구
# 실제로 작성해 본 유형만 등록한다. 등록 안 된 유형이 들어오면 콘솔에 경고가 뜬다.
TRADE_TYPE_NAMES = {
    ("차입", "현금배당", "주식"): "차입주식 배당지급",
}

# ------------------------------------------------------------- 수탁기관 매핑
# 엑셀의 '수탁기관명'이 "PBS명(운용사-수탁은행)" 형태면 자동 파싱한다.
# 형식이 다른 기관은 아래에 {수탁기관명: (수탁은행, PBS명)} 으로 직접 등록한다.
TRUSTEE_OVERRIDES = {
    # "예시수탁기관명": ("국민은행", "미래에셋증권"),
}

# 수신/참조 문구 서식
RECIPIENT_FORMAT = "{bank} 수탁부 귀하"
CC_FORMAT = "{pbs} PBS"

# 출력 파일명 서식
FILENAME_FORMAT = "{date}_{company} 수기운용지시서_{bank_short}수탁_{trade_key}.docx"

# True 면 출력 폴더 안에 지급일(YYYYMMDD) 하위 폴더를 만들어 저장한다. (입력 폴더와 동일한 구조)
OUTPUT_SUBDIR_BY_DATE = True
OUTPUT_SUBDIR_FORMAT = "%Y%m%d"

# ------------------------------------------------------ 공유 드라이브 사본 저장
# data/output 에 저장한 뒤 이 경로에도 같이 복사한다. None 이면 복사하지 않는다.
SHARED_OUTPUT_DIR = Path(r"Z:\02.펀드\006.수탁사(수기운용지시공문)")

# 공유 드라이브 하위 폴더명. 기존 정리 방식 "YYYYMMDD_차입주식배당지급" 을 따른다.
SHARED_SUBDIR_FORMAT = "{date}_{trade_key}"

# 원본 권리배정내역 엑셀도 같은 폴더에 복사한다 (기존 폴더들의 관례).
SHARED_COPY_INPUT = True

# ---------------------------------------------------------------- PDF 변환
# docx 와 함께 PDF 도 만든다. Word(pywin32) + pymupdf 필요.
MAKE_PDF = True

# PDF 를 페이지 전체 이미지로 굽는다 (텍스트·인감 복사 불가). 기존 '_E.pdf' 와 동일.
PDF_IMAGE_ONLY = True
PDF_IMAGE_SUFFIX = "_E"

# 이미지 해상도. 기존 공문이 288dpi(2382x3368) 이므로 같은 값을 쓴다.
PDF_DPI = 288

# 텍스트 PDF(접미사 없는 .pdf)도 함께 남긴다. 이 파일은 인감을 빼낼 수 있으니 주의.
KEEP_TEXT_PDF = False

# ------------------------------------------------------------------ 사용인감
# 템플릿에 이미 도장이 들어 있고, 실행할 때 아래 원본 PNG 내용으로 갱신한다.
# 위치·크기는 원본 지시서의 값을 그대로 쓴다 (변경은 워드에서 템플릿을 직접 수정).
SEAL_IMAGE = SHARED_OUTPUT_DIR / "사용인감도장1.png"

# False 면 도장을 빼고 '(인)' 만 남긴 문서를 만든다. CLI --no-seal 과 동일.
INSERT_SEAL = True

# 공유 폴더에 같은 이름 파일이 이미 있으면 덮어쓰지 않고 건너뛴다.
# 서명·수정된 실제 공문을 날리지 않기 위한 안전장치. CLI --overwrite-shared 로만 해제.
SHARED_OVERWRITE = False
