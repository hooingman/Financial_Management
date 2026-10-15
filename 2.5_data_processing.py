"""연도별 DART 수집 Excel을 단일 기업 분석용 3단계 리포트로 정제합니다."""

from pathlib import Path
from company_config import get_company_name
import pandas as pd
from openpyxl.styles import Font, PatternFill


# 주요 계정은 전체 기업 기준으로 검증합니다(지배기업 귀속 금액 제외).
METRIC_IDS = {
    "매출액": ["ifrs_Revenue", "ifrs-full_Revenue"],
    "영업이익": ["dart_OperatingIncomeLoss"],
    "당기순이익": ["ifrs_ProfitLoss", "ifrs-full_ProfitLoss"],
    "자산총계": ["ifrs_Assets", "ifrs-full_Assets"],
    "자본총계": ["ifrs_Equity", "ifrs-full_Equity"],
    "부채총계": ["ifrs_Liabilities", "ifrs-full_Liabilities"],
    "유동자산": ["ifrs_CurrentAssets", "ifrs-full_CurrentAssets"],
    "유동부채": ["ifrs_CurrentLiabilities", "ifrs-full_CurrentLiabilities"],
}
BALANCE_METRICS = {"자산총계", "자본총계", "부채총계", "유동자산", "유동부채"}


def integrate_sheets(input_path, company_name, start_year, end_year):
    """원본 값을 문자열로 보존하고 원본 시트 및 행 번호를 기록합니다."""
    if start_year > end_year:
        raise ValueError("시작 연도는 종료 연도보다 클 수 없습니다.")
    frames = []
    with pd.ExcelFile(input_path) as workbook:
        for year in range(start_year, end_year + 1):
            sheet = f"{year}년"
            if sheet not in workbook.sheet_names:
                raise ValueError(f"입력 파일에 {sheet} 시트가 없습니다.")
            df = pd.read_excel(workbook, sheet_name=sheet, dtype=str)
            required = {
                "재무제표구분", "계정ID", "계정과목명",
                f"당기금액({year})", f"전기금액({year - 1})",
            }
            missing = required - set(df.columns)
            if missing:
                raise ValueError(f"{sheet} 필수 열 누락: {sorted(missing)}")
            df = df.rename(columns={
                "계정ID": "account_id",
                f"당기금액({year})": "당기금액",
                f"전기금액({year - 1})": "전기금액",
            })
            df.insert(0, "연도", sheet)
            df.insert(1, "회사명", company_name)
            df["원본시트"] = sheet
            df["원본행"] = range(2, len(df) + 2)
            frames.append(df)
    return pd.concat(frames, ignore_index=True)


def clean_amounts(raw):
    """금액의 쉼표와 괄호 음수를 처리하고 결측값은 그대로 둡니다."""
    clean = raw.copy()
    for column in ["account_id", "계정과목명", "재무제표구분"]:
        clean[column] = clean[column].astype("string").str.strip()
    for column in ["당기금액", "전기금액"]:
        values = clean[column].astype("string").str.strip()
        values = values.str.replace(",", "", regex=False).str.replace("−", "-", regex=False)
        values = values.str.replace(r"^\((.+)\)$", r"-\1", regex=True)
        values = values.mask(values.isin(["", "-", "—", "–"]))
        numeric = pd.to_numeric(values, errors="coerce")
        invalid = values.notna() & numeric.isna()
        if invalid.any():
            examples = clean.loc[invalid, ["원본시트", "원본행", column]]
            raise ValueError(f"{column} 숫자 변환 불가:\n{examples.head(10).to_string(index=False)}")
        clean[column] = numeric
    return clean


def prepare_analysis_data(clean):
    """지분 구분이 없는 자본변동표를 제외하고 전체 기준 계정을 우선합니다."""
    final = clean.loc[clean["재무제표구분"] != "자본변동표"].copy()
    priority_ids = set(METRIC_IDS["자본총계"] + METRIC_IDS["당기순이익"])
    final["_우선순위"] = final["account_id"].isin(priority_ids).map({True: 0, False: 1})
    final["_표순서"] = final["재무제표구분"].map({
        "재무상태표": 0, "손익계산서": 1, "포괄손익계산서": 2, "현금흐름표": 3,
    }).fillna(4)
    # 기존 분석 코드가 여러 후보 중 첫 행을 선택하므로 전체 자본/순이익을 먼저 배치.
    return (
        final.sort_values(["연도", "_표순서", "_우선순위", "원본행"], kind="stable")
        .drop(columns=["_우선순위", "_표순서"])
        .reset_index(drop=True)
    )


def validate_data(raw, clean, final, start_year, end_year):
    """주요 8개 계정의 유일성과 자산=부채+자본을 검사합니다."""
    audit = []
    for year in range(start_year, end_year + 1):
        label = f"{year}년"
        group = final.loc[final["연도"] == label]
        row = {
            "연도": label,
            "원본행수": int((raw["연도"] == label).sum()),
            "분석행수": len(group),
            "제외_자본변동표행수": int(
                ((clean["연도"] == label) & (clean["재무제표구분"] == "자본변동표")).sum()
            ),
            "당기금액_결측행수": int(group["당기금액"].isna().sum()),
        }
        for metric, ids in METRIC_IDS.items():
            sections = ["재무상태표"] if metric in BALANCE_METRICS else ["손익계산서", "포괄손익계산서"]
            matches = group.loc[
                group["account_id"].isin(ids)
                & group["재무제표구분"].isin(sections)
                & group["당기금액"].notna()
            ]
            if metric == "당기순이익":
                # 두 표에 같은 순이익이 반복되면 손익계산서를 우선 사용합니다.
                # 같은 표 안의 중복과 표 사이 금액 불일치는 허용하지 않습니다.
                if matches.groupby("재무제표구분").size().gt(1).any():
                    raise ValueError(f"{label} {metric}: 같은 재무제표에 유효 계정이 여러 개입니다.")
                if matches["당기금액"].nunique() > 1:
                    raise ValueError(
                        f"{label} {metric}: 손익계산서와 포괄손익계산서의 당기금액이 다릅니다."
                    )
                preferred = matches.loc[matches["재무제표구분"] == "손익계산서"]
                if not preferred.empty:
                    matches = preferred
            if len(matches) != 1:
                raise ValueError(
                    f"{label} {metric}: 유효 계정이 {len(matches)}개입니다. "
                    "원본 계정과 METRIC_IDS를 확인하세요. 임의 선택하거나 합산하지 않습니다."
                )
            row[metric] = matches.iloc[0]["당기금액"]
        row["자산_부채_자본_차이"] = row["자산총계"] - row["부채총계"] - row["자본총계"]
        if row["자산_부채_자본_차이"] != 0:
            raise ValueError(f"{label} 자산=부채+자본 검증 실패: {row['자산_부채_자본_차이']}")
        audit.append(row)
    return pd.DataFrame(audit)


def process_financial_data(company_name, start_year, end_year, file_dir, output_path=None):
    """검증 성공 후 새 리포트를 저장합니다. 동일 파일명은 덮어씁니다."""
    file_dir = Path(file_dir)
    input_path = file_dir / f"{company_name}_재무제표_{start_year}_{end_year}.xlsx"
    output_path = Path(output_path) if output_path is not None else (
        file_dir / f"{company_name}_재무제표_3단계_완전정제_리포트.xlsx"
    )
    if input_path.resolve() == output_path.resolve():
        raise ValueError("입력 파일과 출력 파일은 달라야 합니다.")
    raw = integrate_sheets(input_path, company_name, start_year, end_year)
    clean = clean_amounts(raw)
    final = prepare_analysis_data(clean)
    audit = validate_data(raw, clean, final, start_year, end_year)
    notes = pd.DataFrame([
        ("원본 파일", input_path.name),
        ("대상 기간", f"{start_year}~{end_year}년"),
        ("1단계", "연도별 원본 통합. 금액은 원본 문자열 보존. 원본시트/원본행으로 추적 가능."),
        ("2단계", "열 이름 통일, 문자열 앞뒤 공백 제거, 당기/전기 금액 숫자 변환. 결측치는 0으로 채우지 않음."),
        ("3단계", "자본변동표는 지분 구분 정보가 없어 분석용에서 제외. 전체 자료는 1/2단계에 보존. 다른 재무제표 행은 유지."),
        ("계정 선택", "재무상태표의 전체 자본 및 손익계산서의 전체 순이익을 해당 표 안에서 앞에 배치. 지배기업 귀속 후보보다 먼저 선택하도록 함."),
        ("검증", "각 연도 8개 주요 계정의 유일성과 자산=부채+자본 확인. 당기순이익은 손익계산서 우선, 없으면 포괄손익계산서 사용. 두 표의 당기금액 불일치 및 같은 표 안 중복은 오류 처리."),
        ("금액 단위", "원본 금액 단위 유지. 분석 코드는 원 단위를 가정하므로 원본 공시 단위 확인 필요."),
        ("공시 범위", "원본에는 연결/별도 및 통화 정보가 없으므로 이 파일만으로 독립 검증할 수 없음."),
        ("전기금액", "원본 전기금액 보존. 다음 연도 비교 수치로 과거 당기금액을 덮어쓰지 않음."),
        ("중복 처리", "자본변동표 외 행은 임의 합산/삭제하지 않음."),
    ], columns=["항목", "설명"])
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, data in [
            ("정제_설명", notes), ("1_원본_통합본", raw), ("2_금액_정제본", clean),
            ("3_최종_정제통합본", final), ("검증_연도별", audit),
        ]:
            data.to_excel(writer, sheet_name=name, index=False)
            sheet = writer.sheets[name]
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1B365D")
            for column in sheet.columns:
                header = column[0].value
                width = 70 if header == "설명" else 48 if header == "account_id" else 28 if header == "계정과목명" else 23
                sheet.column_dimensions[column[0].column_letter].width = width
                if header in ["당기금액", "전기금액"] or (name == "검증_연도별" and header in METRIC_IDS):
                    for cell in column[1:]:
                        cell.number_format = "#,##0;[Red]-#,##0"
    print(f"생성 완료: {output_path.resolve()}")
    print(f"원본 {len(raw):,}행 → 분석용 {len(final):,}행. {len(audit)}개 연도 검증 완료.")
    return output_path


if __name__ == "__main__":
    # 회사는 .env에서 설정합니다. 기간 변경 시 수집/분석 코드도 맞춰주세요.
    # 재실행 시 같은 이름의 정제 리포트를 덮어씁니다(원본 수집 파일은 유지).
    COMPANY_NAME = get_company_name()
    START_YEAR = 2016
    END_YEAR = 2025
    FILE_DIR = Path(__file__).resolve().parent

    process_financial_data(COMPANY_NAME, START_YEAR, END_YEAR, FILE_DIR)
