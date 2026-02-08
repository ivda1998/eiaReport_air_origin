"""
CORE API 클라이언트

역할:
- CORE API를 통한 오픈액세스 학술 논문 검색
- API 키 필수 (https://core.ac.uk 에서 발급)
- 키 없으면 자동으로 비활성화

API 문서: https://api.core.ac.uk/docs/v3
"""

import json
import time
import urllib.request
import urllib.parse
import urllib.error


# CORE API 설정
CORE_API_BASE = "https://api.core.ac.uk/v3"

# 재시도 설정
MAX_RETRIES = 3
RETRY_DELAYS = [5, 10, 20]


def search_core(query, config, year_range=None, max_results=10):
    """CORE API 검색

    Args:
        query: 검색어
        config: ProjectConfig 인스턴스
        year_range: 연도 범위 (예: "2015-2026")
        max_results: 최대 결과 수
    Returns:
        list[dict]: 논문 메타데이터 목록 (semantic_scholar와 동일 구조)
    """
    search_config = config.config.get('search', {})
    core_config = search_config.get('core', {})

    if not core_config.get('enabled', True):
        return []

    # API 키 확인
    api_key = core_config.get('api_key', '')
    if not api_key:
        # 키 없으면 조용히 건너뜀 (최초 1회만 안내)
        if not hasattr(search_core, '_warned'):
            print("  ℹ️ CORE API 키 미설정. CORE 검색 비활성화 (Semantic Scholar만 사용)")
            search_core._warned = True
        return []

    # 쿼리 구성
    search_query = query
    if year_range:
        parts = year_range.split('-')
        if len(parts) == 2:
            search_query = f"({query}) AND yearPublished>={parts[0]} AND yearPublished<={parts[1]}"
    elif search_config.get('year_range'):
        yr = search_config['year_range']
        parts = yr.split('-')
        if len(parts) == 2:
            search_query = f"({query}) AND yearPublished>={parts[0]} AND yearPublished<={parts[1]}"

    # URL 구성
    params = {
        'q': search_query,
        'limit': min(max_results, 100),
    }
    url = f"{CORE_API_BASE}/search/works?{urllib.parse.urlencode(params)}"

    # 헤더
    headers = {
        'Accept': 'application/json',
        'Authorization': f'Bearer {api_key}',
        'User-Agent': 'EIA-Report-Generator/2.2',
    }

    # API 호출 (재시도 로직)
    for attempt in range(MAX_RETRIES):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=30) as response:
                data = json.loads(response.read().decode('utf-8'))

            papers = _parse_response(data, search_config)

            if papers:
                print(f"  ✓ CORE: '{query[:30]}...' → {len(papers)}건")
            return papers

        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt < MAX_RETRIES - 1:
                delay = RETRY_DELAYS[attempt]
                print(f"  ⏳ CORE rate limit. {delay}초 대기... ({attempt+1}/{MAX_RETRIES})")
                time.sleep(delay)
            elif e.code == 401 or e.code == 403:
                print(f"  ⚠️ CORE API 인증 실패. API 키를 확인하세요.")
                return []
            elif e.code == 429:
                print(f"  ⚠️ CORE rate limit 초과. 검색 건너뜀: {query[:30]}")
                return []
            else:
                print(f"  ⚠️ CORE HTTP 오류 ({e.code}): {query[:30]}")
                return []

        except urllib.error.URLError as e:
            print(f"  ⚠️ CORE 연결 실패: {e.reason}")
            return []

        except json.JSONDecodeError:
            print(f"  ⚠️ CORE 응답 파싱 실패")
            return []

        except Exception as e:
            print(f"  ⚠️ CORE 검색 오류: {e}")
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

    raw_papers = data.get('results', [])
    if not raw_papers:
        return []

    for raw in raw_papers:
        try:
            # 저자 포맷팅
            authors_list = raw.get('authors', [])
            if authors_list:
                author_names = []
                for a in authors_list[:5]:
                    if isinstance(a, dict):
                        name = a.get('name', '')
                    elif isinstance(a, str):
                        name = a
                    else:
                        name = str(a)
                    if name:
                        author_names.append(name)
                authors_str = '; '.join(author_names)
                if len(authors_list) > 5:
                    authors_str += ' et al.'
            else:
                authors_str = 'Unknown'

            # DOI 추출
            doi = raw.get('doi', None)

            # 초록 처리
            abstract = raw.get('abstract', '') or raw.get('description', '') or ''
            if len(abstract) > max_chars_per_paper:
                abstract = abstract[:max_chars_per_paper] + '...'

            # 연도 추출
            year = raw.get('yearPublished', None)
            if year is None:
                date_str = raw.get('publishedDate', '') or ''
                if date_str and len(date_str) >= 4:
                    try:
                        year = int(date_str[:4])
                    except ValueError:
                        year = None

            paper = {
                'title': raw.get('title', ''),
                'authors': authors_str,
                'year': year,
                'abstract': abstract,
                'citation_count': raw.get('citationCount', 0) or 0,
                'doi': doi,
                'venue': raw.get('publisher', '') or raw.get('journals', [{}])[0].get('title', '') if raw.get('journals') else '',
                'source': 'core',
                'is_open_access': True,  # CORE는 기본적으로 오픈액세스
            }

            if paper['title']:
                papers.append(paper)

        except Exception:
            continue

    return papers
