"""단일 기업 스크립트가 공유하는 .env 설정."""

import os
from pathlib import Path

from dotenv import load_dotenv


PROJECT_DIR = Path(__file__).resolve().parent
load_dotenv(PROJECT_DIR / ".env")


def get_company_name():
    company_name = os.getenv("COMPANY_NAME", "").strip()
    if not company_name:
        raise ValueError("프로젝트 .env 파일에 COMPANY_NAME을 설정하세요. 예: COMPANY_NAME=현대건설")
    if any(character in company_name for character in '/\\:*?"<>|') or any(
        ord(character) < 32 for character in company_name
    ):
        raise ValueError("COMPANY_NAME에 파일명으로 사용할 수 없는 문자가 있습니다.")
    return company_name
