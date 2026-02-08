"""
SearchAgent: 학술 논문 검색 에이전트

역할:
- Semantic Scholar API를 통한 논문 검색
- CORE API를 통한 오픈액세스 논문 검색 (API 키 필요)
- 검색 결과 병합, 중복 제거, 랭킹
- 섹션별 맞춤 검색 쿼리 실행
"""

from .semantic_scholar_client import search_semantic_scholar
from .core_client import search_core
from .paper_merger import (
    merge_and_rank_papers,
    format_papers_for_prompt,
    search_papers_for_section,
    search_all_papers,
)
