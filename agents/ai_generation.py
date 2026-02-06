"""
AIGenerationAgent: AI 기반 보고서 생성

역할:
- Google Gemini AI 클라이언트 초기화
- 단일 AI 호출 (재시도 로직)
- 컨텍스트 데이터 수집
- 템플릿 기반 섹션별 보고서 생성
- 기존 단일 프롬프트 방식 보고서 생성 (fallback)
- 분석 총괄 (analyze_output)
"""

import os
import time

from .data_loader import (
    load_sample_reports,
    load_facility_data,
    load_measurement_data,
    load_reference_papers,
)
from .isopleth_agent import generate_all_isopleths


# 모듈 레벨 AI 클라이언트 (init_ai_client로 초기화)
_client = None


def init_ai_client(config):
    """AI 클라이언트 초기화

    Args:
        config: ProjectConfig 인스턴스
    Returns:
        Client: Google GenAI 클라이언트
    """
    global _client
    from google.genai import Client
    _client = Client(api_key=config.config['ai']['api_key'])
    return _client


def call_ai(prompt, config):
    """단일 AI 호출 (재시도 로직 포함)

    Args:
        prompt: AI 프롬프트 텍스트
        config: ProjectConfig 인스턴스
    Returns:
        str or None: AI 응답 텍스트
    """
    global _client
    if _client is None:
        init_ai_client(config)

    ai_config = config.config['ai']
    max_retries = ai_config['max_retries']
    retry_delay = ai_config['retry_delay']

    for attempt in range(max_retries):
        try:
            response = _client.models.generate_content(
                model=ai_config['model'],
                contents=prompt,
                config={
                    "max_output_tokens": ai_config['max_output_tokens'],
                    "temperature": ai_config['temperature']
                }
            )

            if response and hasattr(response, 'text') and response.text:
                return response.text.strip()
            else:
                if attempt < max_retries - 1:
                    print(f"   ⚠️ AI 응답 비어있음. 재시도 중... ({attempt+1}/{max_retries})")
                    time.sleep(10)
                    continue
                else:
                    raise Exception("AI 응답이 생성되지 않았습니다.")

        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg and attempt < max_retries - 1:
                print(f"   ⏳ API 할당량 초과. {retry_delay}초 대기 중... ({attempt+1}/{max_retries})")
                time.sleep(retry_delay)
            else:
                raise Exception(f"AI 호출 실패: {error_msg}")
    return None


def prepare_context_data(config):
    """보고서 생성에 필요한 모든 컨텍스트 데이터 수집

    Args:
        config: ProjectConfig 인스턴스
    Returns:
        dict: 컨텍스트 데이터
    """
    report_config = config.config['report']
    max_chars = report_config['aermod_result_max_chars']

    with open(config.OUTPUT_FILE, 'r', encoding='utf-8', errors='ignore') as f:
        aermod_data = f.read()[-max_chars:]

    print(f"\n📚 참조 자료 로딩 중...")
    sample_reports = load_sample_reports(config)
    facility_data = load_facility_data(config)
    measurement_data = load_measurement_data(config)
    reference_papers = load_reference_papers(config)

    sample_text = "\n\n".join([
        f"[샘플 보고서: {item['filename']}]\n{item['content']}"
        for item in sample_reports
    ]) if sample_reports else "샘플 보고서 없음"

    return {
        'aermod_data': aermod_data,
        'sample_text': sample_text,
        'facility_data': facility_data if facility_data else '',
        'measurement_data': measurement_data if measurement_data else '',
        'reference_papers': reference_papers if reference_papers else '',
        'project_name': config.config['project']['name'],
    }


def analyze_with_templates(context_data, isopleth_images, config, templates):
    """템플릿 기반 섹션별 보고서 생성

    Args:
        context_data: prepare_context_data()의 반환값
        isopleth_images: 등농도곡선 이미지 경로 목록
        config: ProjectConfig 인스턴스
        templates: TemplateManager 인스턴스
    Returns:
        tuple: (보고서 텍스트, 이미지 경로 목록)
    """
    sections = templates.get_section_list()
    if not sections:
        print("⚠️ 템플릿 섹션 목록이 비어있습니다. 기존 방식으로 전환합니다.")
        return analyze_legacy(context_data, isopleth_images, config)

    # 부록 제외 (본문 섹션만)
    main_sections = [s for s in sections if s.get('id') not in ['부록']]
    total = len(main_sections)

    full_report_parts = []
    previous_summary = ""

    print(f"\n🤖 템플릿 기반 보고서 작성 시작 (총 {total}개 섹션)")

    for idx, section in enumerate(main_sections):
        section_id = section['id']
        section_title = section.get('title', '')

        print(f"\n  📝 [{idx+1}/{total}] 섹션 {section_id}: {section_title} 작성 중...")

        prompt = templates.get_section_prompt(
            section_id, context_data, previous_summary
        )

        try:
            section_text = call_ai(prompt, config)
            if section_text:
                full_report_parts.append(section_text)
                # 이전 섹션 요약 업데이트 (결론용)
                if len(section_text) > 500:
                    previous_summary += f"\n[{section_id}. {section_title}] {section_text[:300]}..."
                else:
                    previous_summary += f"\n[{section_id}. {section_title}] {section_text}"
                print(f"  ✓ 섹션 {section_id} 완료 ({len(section_text):,}자)")
            else:
                print(f"  ⚠️ 섹션 {section_id} 생성 실패 (건너뜀)")
        except Exception as e:
            print(f"  ❌ 섹션 {section_id} 오류: {str(e)}")
            # 에러 발생해도 나머지 섹션 계속 진행

        # API 레이트 리밋 방지
        if idx < total - 1:
            time.sleep(3)

    if not full_report_parts:
        raise Exception("모든 섹션 생성에 실패했습니다.")

    full_report = "\n\n---\n\n".join(full_report_parts)
    print(f"\n✓ 전체 보고서 생성 완료 ({len(full_report):,}자, {len(full_report_parts)}개 섹션)")

    return full_report, isopleth_images


def analyze_legacy(context_data, isopleth_images, config):
    """기존 단일 프롬프트 방식 보고서 생성 (fallback)

    Args:
        context_data: prepare_context_data()의 반환값
        isopleth_images: 등농도곡선 이미지 경로 목록
        config: ProjectConfig 인스턴스
    Returns:
        tuple: (보고서 텍스트, 이미지 경로 목록)
    """
    sample_text = context_data['sample_text']
    facility_data = context_data['facility_data']
    measurement_data = context_data['measurement_data']
    reference_papers = context_data['reference_papers']
    aermod_data = context_data['aermod_data']

    prompt = f"""
당신은 환경영향평가 대기질 분야 전문가입니다.
제공된 AERMOD 모델링 결과와 참조자료를 바탕으로 환경영향평가서의 대기질 부분을 작성하세요.

[작성 지침]
1. 샘플 보고서의 양식과 구성을 참조하여 동일한 형식으로 작성
2. 정온시설 위치 정보를 활용하여 영향 예측 대상을 명시
3. 현황측정자료를 참조하여 현황과 예측결과를 비교·분석
4. 오염물질별 최대 착지농도와 발생 위치(좌표)를 명확히 제시
5. 환경기준 대비 평가 결과 및 적부 판정
6. 문체는 "~함", "~임"을 사용하고 전문용어를 정확히 사용
7. 보고서로 바로 활용 가능하도록 체계적으로 작성
8. **중요**: 예측 결과 해석 시 다음 사항 반드시 포함:
   - 관련 선행연구나 논문의 연구결과를 인용하여 예측결과의 타당성 검증
   - 국내외 학술논문, 정부 보고서, 유사 환경영향평가서 등을 적극 인용
   - 인용 형식: 저자(연도) 또는 기관명(연도) 형식 사용

[참조 1: 샘플 보고서 양식]
{sample_text[:15000]}

[참조 2: 정온시설 위치 정보]
{facility_data[:10000] if facility_data else "정온시설 정보 없음"}

[참조 3: 현황측정자료]
{measurement_data[:5000] if measurement_data else "현황측정자료 없음"}

[참조 4: 외부 전문자료 및 선행연구]
{reference_papers[:5000] if reference_papers else "외부 전문자료가 없습니다. 대신 일반적인 대기질 관련 학술 연구와 정부 보고서를 인용하여 작성하세요."}

[AERMOD 모델링 결과]
{aermod_data}

위 정보를 종합하여 환경영향평가서 대기질 부분을 작성하세요.
**반드시 선행연구와 전문자료를 인용하여 예측결과의 신뢰성을 높이세요.**
"""

    print("\n🤖 AI 보고서 작성 중... (최대 1-2분 소요)")

    result_text = call_ai(prompt, config)
    if result_text:
        print(f"✓ AI 응답 생성 완료 ({len(result_text):,}자)")
        return result_text, isopleth_images
    else:
        raise Exception("AI 보고서 생성 실패")


def analyze_output(config, templates=None, use_templates=None):
    """AERMOD 결과를 분석하여 환경영향평가서 작성

    Args:
        config: ProjectConfig 인스턴스
        templates: TemplateManager 인스턴스 (없으면 기존 방식)
        use_templates: True=템플릿 사용, False=기존방식, None=자동판단
    Returns:
        tuple: (보고서 텍스트, 이미지 경로 목록)
    """
    if not os.path.exists(config.OUTPUT_FILE):
        return "❌ AERMOD 출력 파일을 찾을 수 없습니다. 먼저 모델링을 실행하세요.", []

    print(f"\n📊 AERMOD 결과 분석 중...")

    isopleth_images = generate_all_isopleths(config)
    context_data = prepare_context_data(config)

    # 템플릿 사용 여부 결정
    if use_templates is None:
        use_templates = templates is not None

    if use_templates and templates is not None:
        print("📋 템플릿 기반 섹션별 생성 모드")
        return analyze_with_templates(context_data, isopleth_images, config, templates)
    else:
        print("📄 기존 단일 프롬프트 생성 모드")
        return analyze_legacy(context_data, isopleth_images, config)
