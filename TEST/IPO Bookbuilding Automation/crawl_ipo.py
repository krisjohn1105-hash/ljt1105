import requests
from bs4 import BeautifulSoup
import pandas as pd
import urllib3
import re

# SSL 경고 무시
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def get_soup(url):
    """지정된 URL에서 BeautifulSoup 객체를 가져옵니다."""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    # verify=False 를 통해 SSL 인증서 에러를 우회합니다.
    try:
        response = requests.get(url, headers=headers, verify=False)
        # 응답 바이트를 바로 euc-kr로 디코드하되, 에러가 발생한 문자는 '?' 등으로 치환(replace)합니다.
        decoded_text = response.content.decode('euc-kr', 'replace')
        return BeautifulSoup(decoded_text, 'html.parser')
    except Exception as e:
        print(f"[오류] {url} 요청 중 에러 발생: {e}")
        return None

def extract_detail_info(detail_url):
    """상세 페이지에서 각 핵심 정보를 추출합니다."""
    soup = get_soup(detail_url)
    if soup is None:
        return {}

    info = {}
    
    # 텍스트 정제 함수
    def clean_text(text):
        if not text: return ""
        return text.replace("\xa0", " ").replace("\t", "").replace("\r", "").replace("\n", "").strip()

    def get_value_by_label(label_keywords):
        for t in soup.find_all(['th', 'td']):
            t_text = clean_text(t.get_text())
            t_text_nospace = t_text.replace(" ", "")
            
            for k in label_keywords:
                k_nospace = k.replace(" ", "")
                # 키워드가 완전히 일치하거나, 포함되더라도 '청구' 등의 수식어가 없는 짧은 헤더인지 확인
                if k_nospace == t_text_nospace or (k_nospace in t_text_nospace and '청구' not in t_text_nospace and len(t_text_nospace) < 10):
                    sib = t.find_next_sibling(['td', 'th'])
                    if sib:
                        return clean_text(sib.get_text())
        return ""

    info['종목명'] = get_value_by_label(['종목명', '기업명'])
    info['종목코드'] = get_value_by_label(['종목코드'])
    info['시장구분'] = get_value_by_label(['시장구분'])
    info['희망공모가액'] = get_value_by_label(['희망공모가'])
    info['수요예측일'] = get_value_by_label(['수요예측일'])
    info['공모청약일'] = get_value_by_label(['공모청약일'])
    info['기관투자자 최대배정수량'] = get_value_by_label(['기관투자자등'])
    
    # 주간사 추출 (보통 '주간사' 뒤에 텍스트가 있으나 종종 본문에 섞임)
    주간사_val = get_value_by_label(['주간사'])
    if "주식수:" in 주간사_val:
        주간사_val = 주간사_val.split("주식수:")[0].strip()
    info['주간사'] = 주간사_val

    return info

def run_ipo_crawler(existing_items=None):
    if existing_items is None:
        existing_items = []
        
    base_url = 'https://www.38.co.kr'
    list_url = f'{base_url}/html/fund/index.htm?o=r'
    
    print("수요예측 일정 페이지 접속 중...")
    soup = get_soup(list_url)
    if not soup:
        print("[에러] 수요예측 목록 페이지를 불러올 수 없습니다.")
        return []
    
    # 수요예측일정 테이블 탐색
    tables = soup.find_all('table', summary='수요예측일정')
    
    IPO_list = []
    
    if tables:
        rows = tables[0].find_all('tr')
        # 첫 줄은 헤더이므로 생략 (또는 th, td 구조에 맞게 순회)
        for row in rows:
            cols = row.find_all('td')
            if not cols:
                continue
            
            # 종목명 찾기 (a 태그 안에 있음)
            a_tag = cols[0].find('a')
            if not a_tag:
                continue
                
            item_name = a_tag.text.strip().replace('\xa0', ' ')
            # 링크가 "./" 와 같은 상대경로로 넘어오므로 절대경로로 올바르게 변환합니다.
            href = a_tag['href']
            if href.startswith('./'):
                href = href[2:]
            detail_link = f"{base_url}/html/fund/{href}"
            
            # 엑셀시트에 없는 종목만 찾아서
            if item_name not in existing_items:
                print(f"[{item_name}] 상세 페이지 크롤링 진행 중...")
                detail_data = extract_detail_info(detail_link)
                # 추출되지 않은 데이터가 있을 수 있으므로 이름 병합
                if '종목명' not in detail_data:
                    detail_data['종목명'] = item_name
                    
                IPO_list.append(detail_data)
            else:
                print(f"[{item_name}] 이미 엑셀(existing_items)에 존재하므로 스킵합니다.")
                
    else:
        print("수요예측일정 테이블을 찾지 못했습니다.")
        
    print(f"\n최종 {len(IPO_list)}건의 신규 크롤링 완료.")
    
    # 결과를 Pandas DataFrame으로 변환 (기존 호출 하위호환 유지용도)
    if IPO_list:
        df = pd.DataFrame(IPO_list)
        columns_order = [
            '종목명', '종목코드', '시장구분', '희망공모가액', 
            '수요예측일', '공모청약일', '기관투자자 최대배정수량', '주간사'
        ]
        final_cols = [c for c in columns_order if c in df.columns]
        df = df[final_cols]
        result_file = 'ipo_crawling_result.csv'
        df.to_csv(result_file, index=False, encoding='utf-8-sig')
        print(f"[{result_file}] 파일로 성공적으로 저장되었습니다.")
        
    return IPO_list

if __name__ == '__main__':
    import traceback
    try:
        ipo_df = run_ipo_crawler()
        if ipo_df is not None:
            print("크롤링 성공!")
    except Exception as e:
        error_msg = f"실행 중 치명적 에러 발생:\n{traceback.format_exc()}"
        print(error_msg)
        with open('error_log.txt', 'w', encoding='utf-8') as f:
            f.write(error_msg)

