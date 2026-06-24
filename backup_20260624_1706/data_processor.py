# data_processor.py
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import pandas as pd
import numpy as np
import os

SAMPLE_DIR = os.path.join(os.path.dirname(__file__), '샘플')

# 캐시 변수
_GRADE_CACHE = None
_PERF_CACHE = None

def clear_cache():
    """캐시 데이터 초기화"""
    global _GRADE_CACHE, _PERF_CACHE
    _GRADE_CACHE = None
    _PERF_CACHE = None

def load_safety_grade_map():
    """안전등급.xlsx에서 사업자번호 → 안전등급 매핑 딕셔너리 반환"""
    global _GRADE_CACHE
    if _GRADE_CACHE is not None:
        return _GRADE_CACHE

    grade_path = os.path.join(SAMPLE_DIR, '안전등급.xlsx')
    if not os.path.exists(grade_path):
        print(f"오류: {grade_path} 파일이 존재하지 않습니다.")
        return {}

    df = pd.read_excel(grade_path)

    # 컬럼 인덱스로 접근 (한글 컬럼 안정성 확보)
    # [11] = 사업자등록번호, [18] = ESG평가.1 (안전등급: SH1, SH2, SA1 등)
    if len(df.columns) <= 18:
        print("오류: 안전등급.xlsx 컬럼 수가 부족합니다.")
        return {}

    col_biz_no = df.columns[11]   # 사업자등록번호
    col_grade  = df.columns[18]   # 안전등급 (ESG평가.1 컬럼)

    # 첫 두 행은 헤더 중복이므로 제거
    df = df.dropna(subset=[col_biz_no])
    df = df[df[col_biz_no].astype(str).str.match(r'\d{3}-\d{2}-\d{5}')]

    grade_map = {}
    for _, row in df.iterrows():
        biz_no = str(row[col_biz_no]).strip().replace('-', '')
        grade  = str(row[col_grade]).strip() if pd.notna(row[col_grade]) else ''
        if biz_no and grade and grade != 'nan':
            grade_map[biz_no] = grade

    _GRADE_CACHE = grade_map
    return grade_map


def load_performance_score_map():
    """수행도평가.xlsx에서 업체명 → 수행도 안전관리 점수 매핑 (과거 모든 계약 평균)"""
    global _PERF_CACHE
    if _PERF_CACHE is not None:
        return _PERF_CACHE

    perf_path = os.path.join(SAMPLE_DIR, '수행도평가.xlsx')
    if not os.path.exists(perf_path):
        print(f"오류: {perf_path} 파일이 존재하지 않습니다.")
        return {}

    df = pd.read_excel(perf_path)

    # [4]=업체명, [12]=평가종류, [19]=공사.5(안전관리 점수), [20]=공사.6(안전관리 배점)
    if len(df.columns) <= 20:
        print("오류: 수행도평가.xlsx 컬럼 수가 부족합니다.")
        return {}

    col_company      = df.columns[4]
    col_eval_type    = df.columns[12]
    col_safety_score = df.columns[19]
    col_safety_max   = df.columns[20]

    # 실제 데이터 행만 필터링 (번호 컬럼이 숫자인 행)
    df = df[pd.to_numeric(df.iloc[:, 0], errors='coerce').notna()]

    company_scores = {}
    for _, row in df.iterrows():
        # 시공평가에 해당하는 안전관리 점수만 사용 (착공평가나 준공평가는 제외)
        eval_type = str(row[col_eval_type]).strip() if pd.notna(row[col_eval_type]) else ''
        if eval_type != '시공평가':
            continue

        company = str(row[col_company]).strip()
        score   = row[col_safety_score]
        max_pts = row[col_safety_max]
        if company and pd.notna(score) and pd.notna(max_pts):
            score_val = float(score)
            max_val   = float(max_pts)
            if max_val > 0:
                normalized_score = (score_val / max_val) * 40.0
                clean_name = company.replace('(주)', '').replace('주식회사', '').replace(' ', '').strip()
                if clean_name not in company_scores:
                    company_scores[clean_name] = []
                company_scores[clean_name].append(normalized_score)

    perf_map = {}
    for clean_name, scores in company_scores.items():
        if scores:
            perf_map[clean_name] = sum(scores) / len(scores)

    _PERF_CACHE = perf_map
    return perf_map


def build_contracts():
    """
    입찰실적.xlsx를 공고번호별로 그룹화하고,
    안전등급 및 수행도 점수를 JOIN하여 INITIAL_CONTRACTS 형식의 리스트 반환.
    """
    bid_path = os.path.join(SAMPLE_DIR, '입찰실적.xlsx')
    if not os.path.exists(bid_path):
        print(f"오류: {bid_path} 파일이 존재하지 않습니다.")
        return []

    df_bid = pd.read_excel(bid_path)

    # 컬럼 인덱스 안정성 체크
    if len(df_bid.columns) <= 39:
        print("오류: 입찰실적.xlsx 컬럼 수가 부족합니다.")
        return []

    # 컬럼 인덱스로 접근
    col_id      = df_bid.columns[0]   # 공고번호
    col_type    = df_bid.columns[17]  # 공종
    col_budget  = df_bid.columns[21]  # 발주예산
    col_biz_no  = df_bid.columns[22]  # 사업자번호
    col_company = df_bid.columns[23]  # 입찰업체명
    col_bid_amt = df_bid.columns[32]  # 입찰금액(최종)
    col_rank    = df_bid.columns[39]  # 순위

    grade_map = load_safety_grade_map()
    perf_map  = load_performance_score_map()

    contracts = []
    grouped   = df_bid.groupby(col_id)

    for contract_id, group in grouped:
        # 순위 기준 정렬
        group = group.sort_values(col_rank)

        # 공고번호별 대표 값
        type_val   = str(group[col_type].iloc[0]).strip()
        budget_val = int(group[col_budget].iloc[0]) if pd.notna(group[col_budget].iloc[0]) else 0

        company_names   = []
        bids            = []
        safety_grades   = []
        safety_scores   = []

        for _, row in group.iterrows():
            # 업체명
            company = str(row[col_company]).strip() if pd.notna(row[col_company]) else ''
            # 입찰금액
            bid_amt = int(row[col_bid_amt]) if pd.notna(row[col_bid_amt]) else 0
            # 사업자번호로 안전등급 조회
            biz_no  = str(row[col_biz_no]).strip().replace('-', '') if pd.notna(row[col_biz_no]) else ''
            grade   = grade_map.get(biz_no, '')
            # 업체명 기준 수행도 안전관리 점수 조회
            clean_name = company.replace('(주)', '').replace('주식회사', '').replace(' ', '').strip()
            score   = perf_map.get(clean_name, None)

            company_names.append(company)
            bids.append(bid_amt)
            safety_grades.append(grade)
            safety_scores.append(round(score, 2) if score is not None else None)

        contracts.append({
            'id':               str(contract_id),
            'type':             type_val,
            'budget':           budget_val,
            'companyNames':     company_names,
            'bids':             bids,
            'safetyGrades':     safety_grades,
            'safetyEvalScores': safety_scores,
        })

    return contracts
