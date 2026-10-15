import os
import io
import time
import zipfile
import xml.etree.ElementTree as ET
import requests
import pandas as pd
from dotenv import load_dotenv
from company_config import get_company_name

# 1. .env 파일에서 DART API KEY 로드
load_dotenv()
API_KEY = os.getenv("DART_API_KEY")

if not API_KEY:
    raise ValueError(".env 파일에서 'DART_API_KEY'를 찾지 못했습니다.")

# 보고서 코드 정의 (기본값: 사업보고서/연간)
REPRT_CODES = {
    "1분기보고서": "11013",
    "반기보고서": "11012",
    "3분기보고서": "11014",
    "사업보고서(연간)": "11011"
}


def get_corp_info(company_name: str):
    """
    OpenDART 전체 회사 목록(CORPCODE.xml)에서 입력된 회사명과 정확히 일치하는
    corp_code(고유번호)와 stock_code(종목코드)를 검색하여 반환합니다.
    STOCK_CODE 설정으로 동명 법인을 구분하며, 후보가 여러 개면 선택하지 않습니다.
    """
    print(f"1. OpenDART 전체 회사 목록 다운로드 중...")
    url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={API_KEY}"
    
    try:
        response = requests.get(url, timeout=30)
        if response.status_code != 200:
            print("❌ DART 서버 응답 실패 또는 API Key 오류입니다.")
            return None, None

        print("2. 압축 해제 및 회사 고유번호 검색 중...")
        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            xml_data = z.read('CORPCODE.xml')
        
        root = ET.fromstring(xml_data)
        
        stock_filter = os.getenv("STOCK_CODE", "").strip()
        candidates = [
            ((node.findtext("corp_code") or "").strip(),
             (node.findtext("stock_code") or "").strip())
            for node in root.findall("list")
            if (node.findtext("corp_name") or "").strip() == company_name
        ]
        if stock_filter:
            candidates = [item for item in candidates if item[1] == stock_filter]
        if not candidates:
            print(f"❌ '{company_name}' 회사를 DART 목록에서 찾지 못했습니다. COMPANY_NAME과 STOCK_CODE를 확인하세요.")
            return None, None
        if len(candidates) != 1:
            print(f"❌ '{company_name}' 이름의 회사가 {len(candidates)}개입니다.")
            for corp_code, stock_code in candidates:
                print(f"• corp_code: {corp_code}, stock_code: {stock_code or '비상장사'}")
            print(".env에 원하는 회사의 STOCK_CODE를 설정하세요. 예: STOCK_CODE=028260")
            return None, None

        corp_code, stock_code = candidates[0]
        stock_code = stock_code or "비상장사"
        print("\n" + "=" * 50)
        print(f"  [{company_name}] 기업 정보 검색 성공")
        print("=" * 50)
        print(f"• 회사 고유번호 (corp_code) : {corp_code}")
        print(f"• 주식 종목코드 (stock_code): {stock_code}")
        print("=" * 50 + "\n")
        return corp_code, stock_code

    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        return None, None


def collect_financial_statements(company_name: str, start_year: int, end_year: int, fs_div: str = "CFS"):
    """
    회사 이름과 기간(시작연도~종료연도)을 받아 연도별 재무제표를 수집하여 각 연도별 시트로 저장합니다.
    - fs_div: 'CFS' (연결재무제표) 또는 'OFS' (별도/개별재무제표)
    """
    # 1. 회사 고유번호(corp_code) 가져오기
    corp_code, stock_code = get_corp_info(company_name)
    if not corp_code:
        print("❌ 프로세스를 종료합니다.")
        return

    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json"
    yearly_dfs = {}

    print(f"🚀 [{company_name}] ({start_year}년 ~ {end_year}년) 재무제표 데이터 수집 시작...")

    # 2. 지정된 기간 동안 연도별 반복 수집
    for year in range(start_year, end_year + 1):
        year_str = str(year)
        print(f"⏳ {year_str}년 재무제표 요청 중...", end=" ")
        
        params = {
            "crtfc_key": API_KEY,
            "corp_code": corp_code,
            "bsns_year": year_str,
            "reprt_code": REPRT_CODES["사업보고서(연간)"],  # 11011
            "fs_div": fs_div
        }

        try:
            response = requests.get(url, params=params, timeout=15)
            if response.status_code == 200:
                data = response.json()
                status = data.get("status")
                
                if status == "000":
                    raw_data = data.get("list", [])
                    df = pd.DataFrame(raw_data)
                    
                    # 필요한 컬럼 정제
                    cols = ['sj_nm', 'account_id', 'account_nm', 'thstrm_amount', 'frmtrm_amount']
                    existing_cols = [c for c in cols if c in df.columns]
                    df_filtered = df[existing_cols].copy()
                    
                    df_filtered.rename(columns={
                        'sj_nm': '재무제표구분',
                        'account_id': '계정ID',
                        'account_nm': '계정과목명',
                        'thstrm_amount': f'당기금액({year})',
                        'frmtrm_amount': f'전기금액({year-1})'
                    }, inplace=True)
                    
                    yearly_dfs[year_str] = df_filtered
                    print(f"✅ 성공 ({len(df_filtered)}개 계정 수집)")
                else:
                    print(f"❌ 실패 (DART 메시지: {data.get('message')})")
            else:
                print(f"❌ HTTP 통신 오류 ({response.status_code})")
                
        except Exception as e:
            print(f"❌ 에러 발생: {e}")

        # DART 서버 요청 제한 방지 대기
        time.sleep(0.5)

    if not yearly_dfs:
        print("\n❌ 수집된 재무제표 데이터가 없습니다.")
        return

    # 3. 각 연도별 원본 시트로 엑셀 저장
    file_name = f"{company_name}_재무제표_{start_year}_{end_year}.xlsx"
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, file_name)

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        for year, df in yearly_dfs.items():
            df.to_excel(writer, sheet_name=f"{year}년", index=False)

    print(f"\n🎉 [완료] [{company_name}] 재무제표 수집 완료!")
    print(f"📁 엑셀 파일 저장 경로: {file_path}\n")


# ==================================================
# 💡 실행부: 회사이름은 .env, 기간은 아래에서 설정하세요.
# ==================================================
if __name__ == "__main__":
    COMPANY_NAME = get_company_name()  # .env의 COMPANY_NAME
    START_YEAR = 2016          # 시작 연도
    END_YEAR = 2025            # 종료 연도
    FS_TYPE = "CFS"            # CFS: 연결재무제표, OFS: 별도/개별재무제표

    collect_financial_statements(
        company_name=COMPANY_NAME,
        start_year=START_YEAR,
        end_year=END_YEAR,
        fs_div=FS_TYPE
    )
