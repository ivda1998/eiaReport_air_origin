"""
환경영향평가 대기질 보고서 자동 작성 시스템 v2.2

에이전트 아키텍처 기반 실행 엔트리포인트.
실제 로직은 agents/ 패키지에 분리되어 있습니다.

구조:
  agents/
  ├── __init__.py          # 패키지 초기화 및 공개 API
  ├── config_agent.py      # ProjectConfig + TemplateManager
  ├── data_loader.py       # 샘플 보고서, 정온시설, 현황측정자료 로딩
  ├── aermod_agent.py      # AERMOD 모델링 실행
  ├── isopleth_agent.py    # 등농도곡선 생성 (DXF 베이스맵, 정온시설)
  ├── ai_generation.py     # AI 보고서 생성 (템플릿/레거시)
  ├── report_format.py     # DOCX 스타일링 + 파일 저장
  └── main.py              # CLI 오케스트레이터 (메뉴 루프)

사용법:
  python t2_v2.py           # 메인 메뉴 실행
  python -m agents.main     # 동일하게 메인 메뉴 실행
"""

from agents.main import init_system, menu_loop


if __name__ == "__main__":
    config, templates = init_system()
    menu_loop(config, templates)
