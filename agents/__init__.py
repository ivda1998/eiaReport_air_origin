"""
환경영향평가 대기질 보고서 자동화 시스템 - 에이전트 모듈

7개의 에이전트로 구성:
- config_agent: 프로젝트 설정 및 템플릿 관리
- data_loader: 참조 자료 로딩 (샘플 보고서, 정온시설, 현황측정자료)
- aermod_agent: AERMOD 모델링 실행
- isopleth_agent: 등농도곡선 생성 (DXF 베이스맵, 정온시설 표시)
- ai_generation: AI 기반 보고서 생성 (섹션별 템플릿 / 기존 단일 프롬프트)
- report_format: 보고서 저장 및 DOCX 스타일링
- main: CLI 오케스트레이터 (메뉴 루프)
"""

from .config_agent import ProjectConfig, TemplateManager
from .data_loader import (
    load_sample_reports,
    load_facility_data,
    load_measurement_data,
    load_reference_papers,
)
from .aermod_agent import run_aermod
from .isopleth_agent import (
    parse_aermod_concentrations,
    load_base_dxf,
    load_facility_locations,
    create_isopleth,
    generate_all_isopleths,
)
from .ai_generation import (
    init_ai_client,
    call_ai,
    prepare_context_data,
    analyze_with_templates,
    analyze_legacy,
    analyze_output,
)
from .report_format import save_report

__version__ = "2.2 (에이전트 아키텍처)"
