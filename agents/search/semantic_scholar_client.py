"""
Semantic Scholar API 클라이언트

역할:
- Semantic Scholar Graph API를 통한 학술 논문 검색
- 무료 사용 가능 (API 키 선택사항, 있으면 rate limit 향상)
- 한글 + 영어 검색 지원

API 문서: https://api.semanticscholar.org/api-docs/
"""

import json
import time
import urllib.request
import urllib.parse
import urllib.error


# Semantic Scholar API 설정
SS_API_BASE = "https://api.semanticscholar.org/graph/v1"
SS_SEARCH_FIELDS = "title,abstract,authors,year,citationCount,venue,externalIds,isOpenAccess"

# 재시도 설정
MAX_RETRIES = 3
RETRY_DELAYS = [5, 10, 20]  # 초


def search_semantic_scholar(query, config, year_range=None, min_citations=0,
                            fields_of_study=None, max_results=10):
    """Semantic Scholar API 검색

    Args:
        query: 검색어 (한글+영어 가능)
        config: ProjectConfig 인스턴스
        year_range: 연도 범위 (예: "2015-2026")
        min_citations: 최소 인용 수
        fields_of_study: 학문 분야 필터 리스트
        max_results: 최대 결과 수
    Returns:
        list[dict]: 논문 메타데이터 목록
            각 dict: {title, authors, year, abstract, citation_count,
                      doi, venue, source, is_open_access}
    """
    search_config = config.config.get('search', {})
    ss_config = search_config.get('semantic_scholar', {})

    if not ss_config.get('enabled', True):
        return []

    # 검색 파라미터 구성
    params = {
        'query': query,
        'fields': SS_SEARCH_FIELDS,
        'limit': min(max_results, 100),  # API 최대 100
    }

    # 연도 필터
    if year_range:
        params['year'] = year_range
    elif search_config.get('year_range'):
        params['year'] = search_config['year_range']

    # 최소 인용 수
    effective_min_citations = min_citations or search_config.get('min_citations', 0)
    if effective_min_citations > 0:
        params['minCitationCount'] = str(effective_min_citations)

    # 학문 분야 필터
    if fields_of_study:
        params['fieldsOfStudy'] = ','.join(fields_of_study)

    # URL 구성
    url = f"{SS_API_BASE}/paper/search?{urllib.parse.urlencode(params)}"

    # 헤더 구성
    headers = {
        'Accept': 'application/json',
        'User-Agent': 'EIA-Report-Generator/2.2',
    }

    # API 키 (선택사항)
    api_key = ss_config.get('api_key', '')
    if api_key:
        headers['x-api-key'] = api_key

    # API 호출 (재시도 로직)
    for attempt in range(MAX_RETRIES):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=30) as response:
                data = json.loads(response.read().decode('utf-8'))

            papers = _parse_response(data, search_config)

            if papers:
                print(f"  ✓ Semantic Scholar: '{query[:30]}...' → {len(papers)}건")
            return papers

        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt < MAX_RETRIES - 1:
                delay = RETRY_DELAYS[attempt]
                print(f"  ⏳ Semantic Scholar rate limit. {delay}초 대기... ({attempt+1}/{MAX_RETRIES})")
                time.sleep(delay)
            elif e.code == 429:
                print(f"  ⚠️ Semantic Scholar rate limit 초과. 검색 건너뜀: {query[:30]}")
                return []
            else:
                print(f"  ⚠️ Semantic Scholar HTTP 오류 ({e.code}): {query[:30]}")
                return []

        except urllib.error.URLError as e:
            print(f"  ⚠️ Semantic Scholar 연결 실패: {e.reason}")
            return []

        except json.JSONDecodeError:
            print(f"  ⚠️ Semantic Scholar 응답 파싱 실패")
            return []

        except Exception as e:
            print(f"  ⚠️ Semantic Scholar 검색 오류: {e}")
            return []

    return []


def _parse_response(data, search_config):
    """API 응답을 표준 PaperDict 형식으로 변환

    Args:
        data: API JSON 응답
        search_config: config의 search 섹션
    Returns:
        list[dict]: PaperDict 목록
    """
    max_chars_per_paper = search_config.get('max_chars_per_paper', 500)
    papers = []

    raw_papers = data.get('data', [])
    if not raw_papers:
        return []

    for raw in raw_papers:
        try:
            # 저자 포맷팅: "Kim, S.; Lee, J." 형식
            authors_list = raw.get('authors', [])
            if authors_list:
                author_names = []
                for a in authors_list[:5]:  # 최대 5명
                    name = a.get('name', '')
                    if name:
                        author_names.append(name)
                authors_str = '; '.join(author_names)
                if len(authors_list) > 5:
                    authors_str += ' et al.'
            else:
                authors_str = 'Unknown'

            # DOI 추출
            external_ids = raw.get('externalIds', {}) or {}
            doi = external_ids.get('DOI', None)

            # 초록 처리
            abstract = raw.get('abstract', '') or ''
            if len(abstract) > max_chars_per_paper:
                abstract = abstract[:max_chars_per_paper] + '...'

            paper = {
                'title': raw.get('title', ''),
                'authors': authors_str,
                'year': raw.get('year'),
                'abstract': abstract,
                'citation_count': raw.get('citationCount', 0) or 0,
                'doi': doi,
                'venue': raw.get('venue', '') or '',
                'source': 'semantic_scholar',
                'is_open_access': raw.get('isOpenAccess', False) or False,
            }

            # 제목이 있는 경우만 추가
            if paper['title']:
                papers.append(paper)

        except Exception:
            continue

    return papers
