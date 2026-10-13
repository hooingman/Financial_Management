import os
import io
import time
import zipfile
import xml.etree.ElementTree as ET
import requests
import pandas as pd
from dotenv import load_dotenv

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


def load_all_corp_codes():
    """
    OpenDART 전체 회사 목록(CORPCODE.xml)을 다운로드하여 
    {회사명: (corp_code, stock_code)} 형태의 딕셔너리로 반환합니다.
    """
    print("⏳ [초기화] OpenDART 전체 기업 고유번호 목록 다운로드 중... (최초 1회)")
    url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={API_KEY}"
    
    try:
        response = requests.get(url, timeout=30)
        if response.status_code != 200:
            print("❌ DART 서버 응답 실패 또는 API Key 오류입니다.")
            return {}

        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            xml_data = z.read('CORPCODE.xml')
        
        root = ET.fromstring(xml_data)
        corp_map = {}
        
        for list_node in root.findall('list'):
            name = list_node.findtext('corp_name').strip()
            corp_code = list_node.findtext('corp_code').strip()
            stock_code = list_node.findtext('stock_code').strip()
            if not stock_code:
                stock_code = "비상장사"
            
            # 동일 회사명이 존재할 경우 주식 종목코드가 있는(상장사) 고유번호를 우선 등록
            if name not in corp_map or stock_code != "비상장사":
                corp_map[name] = (corp_code, stock_code)
                
        print(f"✅ DART 전체 {len(corp_map):,}개 기업 고유번호 로드 완료!\n")
        return corp_map

    except Exception as e:
        print(f"❌ 전체 기업 목록 로드 중 오류 발생: {e}")
        return {}


def collect_multi_companies_financials(companies: list, start_year: int, end_year: int, fs_div: str = "CFS"):
    """
    여러 회사와 기간(시작연도~종료연도)을 받아 원천 재무제표 데이터를 일괄 수집하고 엑셀로 저장합니다.
    - companies: ['동진쎄미켐', '삼성전자', 'SK하이닉스'] 등 대상 회사명 리스트
    - fs_div: 'CFS' (연결재무제표) 또는 'OFS' (별도/개별재무제표)
    """
    # 1. DART 전체 기업 매핑 로드
    corp_map = load_all_corp_codes()
    if not corp_map:
        print("❌ 기업 목록을 불러오지 못해 프로세스를 종료합니다.")
        return

    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json"
    combined_list = []
    
    print("==================================================")
    print(f"🚀 총 {len(companies)}개 기업 재무제표 수집 시작 ({start_year}년 ~ {end_year}년)")
    print("==================================================\n")

    # 2. 기업별 루프
    for corp_name in companies:
        if corp_name not in corp_map:
            print(f"⚠️ [{corp_name}] : DART 등록 목록에서 찾을 수 없어 스킵합니다.")
            continue
            
        corp_code, stock_code = corp_map[corp_name]
        print(f"🏢 [{corp_name}] (고유번호: {corp_code} | 종목코드: {stock_code}) 수집 중...")

        # 3. 연도별 루프
        for year in range(start_year, end_year + 1):
            year_str = str(year)
            print(f"  └ ⏳ {year_str}년 데이터 요청...", end=" ")

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
                        items = data.get("list", [])
                        for item in items:
                            # 💡 컬럼 이름을 고정값('당기금액', '전기금액')으로 변경
                            combined_list.append({
                                "회사명": corp_name,
                                "종목코드": stock_code,
                                "사업연도": year_str,
                                "재무제표구분": item.get("sj_nm"),
                                "계정ID": item.get("account_id"),
                                "계정과목명": item.get("account_nm"),
                                "당기금액": item.get("thstrm_amount"),
                                "전기금액": item.get("frmtrm_amount")
                            })
                        print(f"✅ 성공 ({len(items)}개 계정)")
                    else:
                        print(f"❌ 실패 ({data.get('message')})")
                else:
                    print(f"❌ 통신 에러 ({response.status_code})")

            except Exception as e:
                print(f"❌ 오류: {e}")

            # DART API 트래픽 제한 방지 (0.4초 대기)
            time.sleep(0.4)
            
        print("-" * 50)

    if not combined_list:
        print("\n❌ 수집된 재무제표 데이터가 없습니다.")
        return

    # 4. 원천 데이터 저장
    df_all = pd.DataFrame(combined_list)

    file_name = f"다중기업_재무제표_원천데이터_{start_year}_{end_year}.xlsx"
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, file_name)

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        # 시트 1: 수집된 전체 원천 데이터
        df_all.to_excel(writer, sheet_name="전체_원천데이터", index=False)
        
        # 시트 2~N: 기업별 개별 시트 분리 저장
        for corp_name in companies:
            df_corp = df_all[df_all["회사명"] == corp_name]
            if not df_corp.empty:
                df_corp.to_excel(writer, sheet_name=corp_name[:30], index=False)

    print(f"\n🎉 [완료] 모든 기업의 재무제표 원천 데이터 수집이 완료되었습니다!")
    print(f"📁 저장된 엑셀 파일: {file_path}\n")


# ==================================================
# 💡 실행부
# ==================================================
if __name__ == "__main__":
    TARGET_COMPANIES = ["현대건설", "삼성전자", "SK하이닉스"]
    START_YEAR = 2021
    END_YEAR = 2025
    FS_TYPE = "CFS"

    collect_multi_companies_financials(
        companies=TARGET_COMPANIES,
        start_year=START_YEAR,
        end_year=END_YEAR,
        fs_div=FS_TYPE
    )
