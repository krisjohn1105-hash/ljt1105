import pandas as pd

def normalize(value):
    """
    데이터 정규화를 위한 커스텀 함수.
    사용자 전역 규칙(3)에 따라 숫자형 데이터는 소수점 3째 자리에서 반올림하여 2째자리까지 표기합니다.
    """
    # 결측치나 문자열 예외 처리
    if pd.isna(value) or value == "COLUMN NOT FOUND":
        return value
        
    try:
        # 숫자로 변환 가능한 경우 소수점 둘째 자리까지 반올림
        return round(float(value), 2)
    except (ValueError, TypeError):
        # 문자열일 경우 양쪽 공백 제거
        if isinstance(value, str):
            return value.strip()
        return value
