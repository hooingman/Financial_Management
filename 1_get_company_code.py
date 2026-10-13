import os
import io
import zipfile
import xml.etree.ElementTree as ET
import requests
from dotenv import load_dotenv
from company_config import get_company_name

# 1. .env 파일에서 API KEY 로드
load_dotenv()
API_KEY = os.getenv("DART_API_KEY")

if not API_KEY:
    raise ValueError(".env 파일에서 'DART_API_KEY'를 찾지 못했습니다.")

# DART 고정 보고서 코드 (회사 구분 없음)
REPRT_CODES = {
    "1분기보고서": "11013",
    "반기보고서": "11012",
    "3분기보고서": "11014",
    "사업보고서(연간)": "11011"
}

def get_company_info(company_name: str):
    print(f"1. OpenDART 전체 회사 고유번호 파일 다운로드 중... (잠시만 기다려주세요)")
    url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={API_KEY}"
    
    try:
        response = requests.get(url)
        if response.status_code != 200:
            print("❌ API Key가 잘못되었거나 DART 서버 응답이 실패했습니다.")
            return

        print("2. 압축 해제 및 회사 목록 검색 중...")
        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            xml_data = z.read('CORPCODE.xml')
        
        root = ET.fromstring(xml_data)
        
        for list_node in root.findall('list'):
            current_name = list_node.findtext('corp_name')
            
            # 입력한 회사명과 정확히 일치하는 경우
            if current_name == company_name:
                corp_code = list_node.findtext('corp_code')
                stock_code = list_node.findtext('stock_code').strip()
                
                if not stock_code:
                    stock_code = "비상장사"

                print("\n" + "="*40)
                print(f"  [{company_name}] 검색 결과")
                print("="*40)
                print(f"• 회사 고유번호 (corp_code) : {corp_code}")
                print(f"• 주식 종목코드 (stock_code): {stock_code}")
                print("-"*40)
                print("[재무제표 요청 시 사용할 reprt_code 고정값]")
                for name, code in REPRT_CODES.items():
                    print(f"  - {name}: {code}")
                print("="*40 + "\n")
                return

        print(f"❌ '{company_name}' 회사를 DART 목록에서 찾지 못했습니다.")

    except Exception as e:
        print(f"오류 발생: {e}")

# 파이썬 파일 실행 시 직접 수행되는 블록
if __name__ == "__main__":
    get_company_info(get_company_name())
