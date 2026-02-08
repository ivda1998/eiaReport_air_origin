"""
논문 검색 결과 병합, 중복 제거, 랭킹, 포맷팅

역할:
- 다중 소스(Semantic Scholar, CORE) 결과 병합
- DOI 및 제목 기반 중복 제거
- 인용 수 + 최신성 기반 랭킹
- 프롬프트 삽입용 텍스트 포맷팅
- 섹션별 맞춤 검색 쿼리 실행
"""

import re
import time

from .semantic_scholar_client import search_semantic_scholar
from .core_client import search_core


# ==========================================
# 섹션별 검색 쿼리 매핑
# ==========================================
# 환경영향평가서 대기질 부문의 각 섹션에 맞는 학술 검색어

SECTION_QUERIES = {
    "2": [
        "ambient air quality monitoring Korea PM2.5 PM10",
        "대기질 현황측정 환경기준 모니터링",
    ],
    "3": [
        "AERMOD atmospheric dispersion modeling validation",
        "AERMOD 대기확산 모델링 환경영향평가",
        "gaussian plume model regulatory application EPA",
    ],
    "4": [
        "AERMOD prediction results environmental impact assessment",
        "대기질 영향 예측 정온시설 sensitive receptor concentration",
        "maximum ground-level concentration dispersion model",
    ],
    "5": [
        "air quality standards compliance assessment Korea",
        "대기환경기준 적합성 평가 환경영향평가",
    ],
    "6": [
        "fugitive dust mitigation construction site PM",
        "비산먼지 저감 방안 공사장 대기질",
        "air pollution control measures construction",
    ],
}

# 일반 검색 쿼리 (config에서 지정하지 않은 경우 사용)
DEFAULT_GENERAL_QUERIES = [
    "AERMOD atmospheric dispersion modeling environmental impact",
    "환경영향평가 대기질 모델링 예측",
]


# ==========================================
# 중복 제거 및 랭킹
# ==========================================

def _normalize_title(title):
    """제목을 정규화하여 비교 가능한 형태로 변환"""
    if not title:
        return set()
    # 소문자 변환, 특수문자 제거, 단어 집합 반환
    normalized = re.sub(r'[^\w\s]', '', title.lower())
    return set(normalized.split())


def _title_similarity(title_a, title_b):
    """두 제목의 Jaccard 유사도 계산

    Returns:
        float: 0.0 ~ 1.0
    """
    words_a = _normalize_title(title_a)
    words_b = _normalize_title(title_b)

    if not words_a or not words_b:
        return 0.0

    intersection = words_a & words_b
    union = words_a | words_b

    return len(intersection) / len(union) if union else 0.0


def _calculate_score(paper):
    """논문 랭킹 점수 계산

    score = citation_count * 0.6 + recency_score * 0.4
    recency_score = max(0, (year - 2010)) * 5
    """
    citation_count = paper.get('citation_count', 0) or 0
    year = paper.get('year') or 2020

    recency_score = max(0, (year - 2010)) * 5
    score = citation_count * 0.6 + recency_score * 0.4

    return score


def merge_and_rank_papers(papers_a, papers_b, max_results=10):
    """두 소스의 논문 결과를 병합, 중복 제거, 랭킹

    중복 판단:
    1. DOI가 같으면 중복 (정확 매칭)
    2. 제목 Jaccard 유사도 > 0.7이면 중복

    중복 시 인용 수가 높은 쪽을 유지.

    Args:
        papers_a: 첫 번째 소스 결과
        papers_b: 두 번째 소스 결과
        max_results: 최대 반환 수
    Returns:
        list[dict]: 중복 제거 및 정렬된 PaperDict 목록
    """
    all_papers = list(papers_a) + list(papers_b)

    if not all_papers:
        return []

    # DOI 기반 중복 제거
    seen_dois = {}
    unique_papers = []

    for paper in all_papers:
        doi = paper.get('doi')
        if doi:
            doi_lower = doi.lower().strip()
            if doi_lower in seen_dois:
                # 이미 본 DOI → 인용 수 더 높은 쪽 유지
                existing_idx = seen_dois[doi_lower]
                if paper.get('citation_count', 0) > unique_papers[existing_idx].get('citation_count', 0):
                    unique_papers[existing_idx] = paper
                continue
            else:
                seen_dois[doi_lower] = len(unique_papers)

        unique_papers.append(paper)

    # 제목 유사도 기반 중복 제거
    final_papers = []
    for paper in unique_papers:
        is_duplicate = False
        for existing in final_papers:
            if _title_similarity(paper.get('title', ''), existing.get('title', '')) > 0.7:
                # 중복 → 인용 수 높은 쪽 유지
                if paper.get('citation_count', 0) > existing.get('citation_count', 0):
                    final_papers[final_papers.index(existing)] = paper
                is_duplicate = True
                break
        if not is_duplicate:
            final_papers.append(paper)

    # 점수 기반 정렬 (높은 순)
    final_papers.sort(key=_calculate_score, reverse=True)

    return final_papers[:max_results]


# ==========================================
# 프롬프트 포맷팅
# ==========================================

def format_papers_for_prompt(papers, max_total_chars=8000, max_per_paper=500):
    """논문 목록을 프롬프트 삽입용 텍스트로 변환

    형식:
    [논문: Title (FirstAuthor, Year)]
    Abstract text...

    Args:
        papers: PaperDict 목록
        max_total_chars: 전체 최대 문자 수
        max_per_paper: 논문당 최대 문자 수 (초록)
    Returns:
        str: 포맷된 텍스트
    """
    if not papers:
        return ""

    parts = []
    total_chars = 0

    for paper in papers:
        title = paper.get('title', 'Unknown')
        authors = paper.get('authors', 'Unknown')
        year = paper.get('year', 'n.d.')
        abstract = paper.get('abstract', '')
        citations = paper.get('citation_count', 0)
        doi = paper.get('doi', '')
        venue = paper.get('venue', '')

        # 첫 번째 저자만 추출
        first_author = authors.split(';')[0].strip() if authors else 'Unknown'

        # 헤더
        header = f"[논문: {title} ({first_author}, {year})]"

        # 메타정보
        meta_parts = []
        if venue:
            meta_parts.append(f"학술지: {venue}")
        if citations:
            meta_parts.append(f"인용: {citations}회")
        if doi:
            meta_parts.append(f"DOI: {doi}")
        meta = ' | '.join(meta_parts)

        # 초록 제한
        if abstract and len(abstract) > max_per_paper:
            abstract = abstract[:max_per_paper] + '...'

        # 조립
        entry = f"{header}\n{meta}\n{abstract}" if abstract else f"{header}\n{meta}"

        # 전체 문자 수 제한 체크
        if total_chars + len(entry) + 2 > max_total_chars:
            break

        parts.append(entry)
        total_chars += len(entry) + 2  # +2 for \n\n

    return '\n\n'.join(parts)


# ==========================================
# 섹션별 검색 오케스트레이션
# ==========================================

def search_papers_for_section(section_id, config):
    """특정 섹션에 맞는 논문 검색 및 포맷

    Args:
        section_id: 섹션 번호 ("1", "2", ..., "8")
        config: ProjectConfig 인스턴스
    Returns:
        str: 해당 섹션용 포맷된 논문 텍스트
    """
    search_config = config.config.get('search', {})
    queries = SECTION_QUERIES.get(section_id, [])

    if not queries:
        return ""

    max_per_query = search_config.get('max_results_per_query', 5)
    max_total = search_config.get('max_results_total', 10)
    max_total_chars = search_config.get('total_max_chars', 8000)
    max_per_paper = search_config.get('max_chars_per_paper', 500)

    all_ss_papers = []
    all_core_papers = []

    for query in queries:
        # Semantic Scholar 검색
        ss_papers = search_semantic_scholar(query, config, max_results=max_per_query)
        all_ss_papers.extend(ss_papers)

        # CORE 검색
        core_papers = search_core(query, config, max_results=max_per_query)
        all_core_papers.extend(core_papers)

        # API rate limit 방지
        time.sleep(1)

    # 병합 및 랭킹
    merged = merge_and_rank_papers(all_ss_papers, all_core_papers, max_results=max_total)

    # 포맷팅
    return format_papers_for_prompt(merged, max_total_chars=max_total_chars,
                                     max_per_paper=max_per_paper)


def search_all_papers(config):
    """모든 섹션별 + 일반 논문 검색을 한번에 실행

    Args:
        config: ProjectConfig 인스턴스
    Returns:
        dict: {section_id: formatted_paper_text}
            예: {"general": "...", "3": "...", "4": "...", ...}
    """
    search_config = config.config.get('search', {})

    if not search_config.get('enabled', False):
        return {}

    max_per_query = search_config.get('max_results_per_query', 5)
    max_total = search_config.get('max_results_total', 10)
    max_total_chars = search_config.get('total_max_chars', 8000)
    max_per_paper = search_config.get('max_chars_per_paper', 500)

    results = {}

    # 1. 일반 검색 (config의 queries 또는 기본값)
    general_queries = search_config.get('queries', DEFAULT_GENERAL_QUERIES)

    print(f"\n🔍 학술 논문 검색 시작...")

    all_general_ss = []
    all_general_core = []

    for query in general_queries:
        ss_papers = search_semantic_scholar(query, config, max_results=max_per_query)
        all_general_ss.extend(ss_papers)

        core_papers = search_core(query, config, max_results=max_per_query)
        all_general_core.extend(core_papers)

        time.sleep(1)

    general_merged = merge_and_rank_papers(all_general_ss, all_general_core, max_results=max_total)
    results['general'] = format_papers_for_prompt(
        general_merged, max_total_chars=max_total_chars, max_per_paper=max_per_paper
    )

    if general_merged:
        print(f"  ✓ 일반 검색: {len(general_merged)}건")

    # 2. 섹션별 검색 (쿼리가 있는 섹션만)
    # 이미 검색된 쿼리를 캐싱하여 중복 호출 방지
    query_cache = {}

    for section_id, queries in SECTION_QUERIES.items():
        section_ss = []
        section_core = []

        for query in queries:
            # 캐시 확인
            if query in query_cache:
                cached_ss, cached_core = query_cache[query]
                section_ss.extend(cached_ss)
                section_core.extend(cached_core)
                continue

            ss_papers = search_semantic_scholar(query, config, max_results=max_per_query)
            core_papers = search_core(query, config, max_results=max_per_query)

            query_cache[query] = (ss_papers, core_papers)
            section_ss.extend(ss_papers)
            section_core.extend(core_papers)

            time.sleep(1)

        section_merged = merge_and_rank_papers(section_ss, section_core, max_results=max_total)
        section_text = format_papers_for_prompt(
            section_merged, max_total_chars=max_total_chars // 2,  # 섹션별은 절반
            max_per_paper=max_per_paper
        )

        if section_text:
            results[section_id] = section_text
            print(f"  ✓ 섹션 {section_id} 검색: {len(section_merged)}건")

    total_sections = sum(1 for k, v in results.items() if v and k != 'general')
    print(f"✓ 학술 논문 검색 완료 (일반 + {total_sections}개 섹션)")

    return results
