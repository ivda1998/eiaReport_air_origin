"""
CLI 오케스트레이터: 메인 메뉴 루프 및 에이전트 통합

역할:
- 전역 설정 초기화 (ProjectConfig, TemplateManager, AI Client)
- 사용자 메뉴 인터페이스
- 각 에이전트 호출 및 흐름 제어
- 새 프로젝트 초기화
"""

import os
import time
import traceback

from .config_agent import ProjectConfig, TemplateManager
from .aermod_agent import run_aermod
from .isopleth_agent import generate_all_isopleths
from .ai_generation import init_ai_client, analyze_output
from .report_format import save_report
from .search import search_all_papers, format_papers_for_prompt

__version__ = "2.3 (학술 검색 에이전트 추가)"


def init_system(config_path="config.yaml"):
    """시스템 초기화: 설정, 템플릿, AI 클라이언트

    Args:
        config_path: config.yaml 경로
    Returns:
        tuple: (config, templates)
    """
    # 1. 프로젝트 설정 로딩
    try:
        config = ProjectConfig(config_path)
    except Exception as e:
        print(f"❌ 설정 초기화 실패: {e}")
        print("   config.yaml 파일을 확인하고 다시 시도하세요.")
        exit(1)

    # 2. 템플릿 로딩
    templates = None
    try:
        templates_dir = os.path.join(config.BASE_DIR, "templates")
        templates = TemplateManager(templates_dir)
    except Exception as e:
        print(f"⚠️ 템플릿 로딩 실패 (기존 방식으로 동작): {e}")

    # 3. AI 클라이언트 초기화
    try:
        init_ai_client(config)
    except Exception as e:
        print(f"⚠️ AI 클라이언트 초기화 실패: {e}")

    return config, templates


def init_new_project():
    """새로운 프로젝트 초기화"""
    print("\n" + "="*70)
    print("   새 프로젝트 초기화")
    print("="*70)

    project_name = input("\n프로젝트 이름을 입력하세요: ").strip()
    if not project_name:
        print("❌ 프로젝트 이름이 필요합니다.")
        return

    project_dir = input("프로젝트 폴더 경로 (Enter=현재 폴더): ").strip()
    if not project_dir:
        project_dir = os.path.join(os.getcwd(), project_name.replace(" ", "_"))

    if os.path.exists(project_dir):
        confirm = input(f"⚠️ 폴더가 이미 존재합니다: {project_dir}\n계속하시겠습니까? (y/n): ").strip().lower()
        if confirm != 'y':
            print("프로젝트 초기화를 취소합니다.")
            return
    else:
        os.makedirs(project_dir)

    print(f"\n📁 프로젝트 폴더 생성 중: {project_dir}")

    subdirs = ['input', 'output', 'sample', 'data', 'engine', 'data/references']
    for subdir in subdirs:
        subdir_path = os.path.join(project_dir, subdir)
        if not os.path.exists(subdir_path):
            os.makedirs(subdir_path)
            print(f"  ✓ {subdir} 폴더 생성")

    # config.yaml 생성
    config_template = f"""# 환경영향평가 대기질 보고서 자동화 시스템 - 프로젝트 설정 파일

# 프로젝트 기본 정보
project:
  name: "{project_name}"
  description: "환경영향평가 대기질 모델링"
  date: "{time.strftime('%Y-%m-%d')}"

# 경로 설정
paths:
  base_dir: "{project_dir.replace(chr(92), chr(92)+chr(92))}"
  input_dir: "input"
  output_dir: "output"
  sample_dir: "sample"
  data_dir: "data"
  engine_dir: "engine"

# AERMOD 설정
aermod:
  exe_name: "aermod.exe"
  input_file: "project.inp"
  output_file: "project.out"
  timeout: 300

# 데이터 파일 설정
data_files:
  base_map: "base.dxf"
  facilities: "정온시설.xlsx"
  measurements: "현황측정자료.xlsx"

  facilities_structure:
    header_rows: 2
    name_column: 3
    x_column: 7
    y_column: 8

# 오염물질 설정
pollutants:
  - name: "PM2.5"
    display_name: "PM₂.₅"
    unit: "μg/m³"
    standard: 25.0

  - name: "PM10"
    display_name: "PM₁₀"
    unit: "μg/m³"
    standard: 50.0

  - name: "NO2"
    display_name: "NO₂"
    unit: "μg/m³"
    standard: 60.0

  - name: "SO2"
    display_name: "SO₂"
    unit: "μg/m³"
    standard: 50.0

# 등농도곡선 기본 설정
isopleth:
  default_level_count: 10
  default_level_min: null
  default_level_max: null
  default_level_interval: null

  grid_resolution: 200
  interpolation_method: "linear"

  visualization:
    figure_width: 14
    figure_height: 12
    dpi: 300
    contour_color: "blue"
    contour_linewidth: 1.2
    facility_color: "red"
    facility_size: 8
    label_fontsize: 6
    base_map_color: "darkgray"
    base_map_linewidth: 0.3
    base_map_alpha: 0.5

# AI 보고서 생성 설정
ai:
  api_key: "YOUR_API_KEY_HERE"
  model: "gemini-2.5-flash"
  max_output_tokens: 16000
  temperature: 0.3
  max_retries: 3
  retry_delay: 35

# 보고서 설정
report:
  formats:
    - "txt"
    - "docx"

  sample_max_pages: 50
  sample_max_chars: 30000
  reference_max_pages: 10
  reference_max_chars: 5000
  aermod_result_max_chars: 20000

  docx:
    title: "환경영향평가서 - 대기질 부문"
    font_name: "맑은 고딕"
    font_size: 10
    line_spacing: 1.5
    image_width: 6.0

# 학술 논문 검색 설정
search:
  enabled: true
  semantic_scholar:
    enabled: true
    api_key: ""
  core:
    enabled: true
    api_key: ""
  max_results_per_query: 5
  max_results_total: 10
  max_chars_per_paper: 500
  total_max_chars: 8000
  year_range: "2015-2026"
  min_citations: 3
  queries:
    - "AERMOD atmospheric dispersion modeling environmental impact"
    - "환경영향평가 대기질 모델링 예측"

# 한글 폰트 설정
font:
  family: "Malgun Gothic"
  unicode_minus: false
"""

    config_path = os.path.join(project_dir, "config.yaml")
    with open(config_path, 'w', encoding='utf-8') as f:
        f.write(config_template)
    print(f"  ✓ config.yaml 생성")

    print(f"\n✅ 프로젝트 초기화 완료!")
    print(f"\n📌 다음 단계:")
    print(f"   1. {config_path} 파일을 열어 API 키 등을 설정하세요")
    print(f"   2. 필요한 파일들(inp, dxf, xlsx 등)을 해당 폴더에 추가하세요")
    print(f"   3. AERMOD 실행 파일을 engine 폴더에 복사하세요")
    print(f"   4. python -m agents.main 또는 python t2_v2.py로 실행하세요")


def menu_loop(config, templates):
    """메인 메뉴 루프

    Args:
        config: ProjectConfig 인스턴스
        templates: TemplateManager 인스턴스
    """
    print("\n" + "="*70)
    print(f"   환경영향평가 대기질 보고서 자동 작성 시스템 {__version__}")
    print("="*70)
    print(f"   프로젝트: {config.config['project']['name']}")
    print("="*70)

    while True:
        print("\n[명령어]")
        print("  1. 모델링    - AERMOD 시뮬레이션 실행")
        print("  2. 분석      - 결과 분석 및 보고서 작성 (TXT, 기존방식)")
        print("  3. 전체      - 모델링 + 분석 한번에 실행 (TXT)")
        print("  4. DOCX      - 보고서를 DOCX 형식으로 작성 (기존방식)")
        print("  5. 등농도선  - 등농도곡선만 생성")
        print("  6. 종료      - 프로그램 종료")
        print("  7. 템플릿    - 템플릿 기반 고품질 보고서 (DOCX)")
        print("  8. 템플릿TXT - 템플릿 기반 보고서 (TXT)")
        print("  9. 논문검색  - 학술 논문 검색 테스트")

        command = input("\n명령을 입력하세요: ").strip()

        if command == "종료" or command == "6":
            print("\n👋 프로그램을 종료합니다.")
            break

        elif command == "모델링" or command == "1":
            result = run_aermod(config)
            print(result)

        elif command == "분석" or command == "2":
            _cmd_analyze_txt(config, templates, use_templates=False)

        elif command == "전체" or command == "3":
            result = run_aermod(config)
            print(result)
            if "SUCCESS" in result or "WARNING" in result:
                _cmd_analyze_txt(config, templates, use_templates=False)
            else:
                print("\n❌ 모델링 실패로 분석을 진행할 수 없습니다.")

        elif command == "DOCX" or command == "docx" or command == "4":
            _cmd_analyze_docx(config, templates, use_templates=False)

        elif command == "등농도선" or command == "5":
            _cmd_isopleth(config)

        elif command == "템플릿" or command == "7":
            if templates is None:
                print("\n❌ 템플릿이 로딩되지 않았습니다. templates 폴더를 확인하세요.")
                continue
            _cmd_template_docx(config, templates)

        elif command == "템플릿TXT" or command == "8":
            if templates is None:
                print("\n❌ 템플릿이 로딩되지 않았습니다. templates 폴더를 확인하세요.")
                continue
            _cmd_analyze_txt(config, templates, use_templates=True)

        elif command == "논문검색" or command == "9":
            _cmd_search_papers(config)

        else:
            print("⚠️ 올바른 명령을 입력하세요.")


def _cmd_analyze_txt(config, templates, use_templates=False):
    """분석 + TXT 저장 명령"""
    try:
        result_data = analyze_output(config, templates, use_templates=use_templates)
        if result_data:
            report, isopleth_images = result_data
            print("\n" + "="*70)
            label = "템플릿 기반" if use_templates else "AI 작성"
            print(f"[{label} 환경영향평가서 - 대기질]")
            print("="*70)
            if len(report) > 1000:
                print(report[:1000])
                print(f"\n... (총 {len(report):,}자, 이하 생략) ...\n")
            else:
                print(report)

            result = save_report(report, config, templates if use_templates else None,
                                format='txt', isopleth_images=isopleth_images)
            if result:
                print(f"✓ 보고서 작성 완료")
        else:
            print("❌ 보고서 내용이 생성되지 않았습니다.")
    except Exception as e:
        print(f"\n❌ 오류 발생: {str(e)}")
        traceback.print_exc()


def _cmd_analyze_docx(config, templates, use_templates=False):
    """분석 + DOCX 저장 명령"""
    try:
        if os.path.exists(config.OUTPUT_FILE):
            result_data = analyze_output(config, templates, use_templates=use_templates)
            if result_data:
                report, isopleth_images = result_data
                print("\n" + "="*70)
                print("[AI 작성 환경영향평가서 - 대기질]")
                print("="*70)
                print(f"보고서 길이: {len(report):,}자")
                print(report[:500] + "..." if len(report) > 500 else report)

                result = save_report(report, config, templates if use_templates else None,
                                    format='docx', isopleth_images=isopleth_images)
                if result:
                    print(f"✓ DOCX 파일 생성 완료")
            else:
                print("❌ 보고서 내용이 생성되지 않았습니다.")
        else:
            print("\n❌ AERMOD 출력 파일이 없습니다. 먼저 '모델링' 또는 '전체'를 실행하세요.")
    except Exception as e:
        print(f"\n❌ DOCX 생성 오류: {str(e)}")
        traceback.print_exc()


def _cmd_template_docx(config, templates):
    """템플릿 기반 DOCX + TXT 저장 명령"""
    try:
        if os.path.exists(config.OUTPUT_FILE):
            result_data = analyze_output(config, templates, use_templates=True)
            if result_data:
                report, isopleth_images = result_data
                print("\n" + "="*70)
                print("[템플릿 기반 환경영향평가서 - 대기질]")
                print("="*70)
                print(f"보고서 길이: {len(report):,}자")
                if len(report) > 500:
                    print(report[:500] + "...")
                else:
                    print(report)

                # TXT 저장
                save_report(report, config, templates, format='txt', isopleth_images=isopleth_images)
                # DOCX 저장
                result = save_report(report, config, templates, format='docx', isopleth_images=isopleth_images)
                if result:
                    print(f"✓ 템플릿 기반 DOCX 파일 생성 완료")
            else:
                print("❌ 보고서 내용이 생성되지 않았습니다.")
        else:
            print("\n❌ AERMOD 출력 파일이 없습니다. 먼저 '모델링' 또는 '전체'를 실행하세요.")
    except Exception as e:
        print(f"\n❌ 템플릿 보고서 생성 오류: {str(e)}")
        traceback.print_exc()


def _cmd_isopleth(config):
    """등농도곡선 생성 명령"""
    try:
        if os.path.exists(config.OUTPUT_FILE):
            print("\n🎨 등농도곡선 생성을 시작합니다...")
            print("\n[옵션 설정]")
            print("  Enter를 누르면 자동으로 설정됩니다.")

            level_min_str = input("  최소 농도 레벨 (μg/m³, 기본값: 자동): ").strip()
            level_min = float(level_min_str) if level_min_str else None

            level_max_str = input("  최대 농도 레벨 (μg/m³, 기본값: 자동): ").strip()
            level_max = float(level_max_str) if level_max_str else None

            print("\n  등농도선 설정 방법:")
            print("    1. 간격으로 설정 (예: 5 μg/m³ 간격)")
            print("    2. 개수로 설정 (예: 10개)")
            interval_choice = input("  선택 (1 또는 2, 기본값: 2): ").strip()

            level_interval = None
            level_count = 10

            if interval_choice == "1":
                level_interval_str = input("  등농도선 간격 (μg/m³): ").strip()
                level_interval = float(level_interval_str) if level_interval_str else None
            else:
                level_count_str = input("  등농도선 개수 (기본값: 10): ").strip()
                level_count = int(level_count_str) if level_count_str else 10

            images = generate_all_isopleths(config,
                                            level_min=level_min, level_max=level_max,
                                            level_count=level_count, level_interval=level_interval)
            if images:
                print(f"\n✓ 총 {len(images)}개의 등농도곡선이 생성되었습니다:")
                for img in images:
                    print(f"  - {img}")
            else:
                print("\n⚠️ 등농도곡선 생성에 실패했습니다.")
        else:
            print("\n❌ AERMOD 출력 파일이 없습니다. 먼저 '모델링' 또는 '전체'를 실행하세요.")
    except ValueError as ve:
        print(f"\n❌ 입력 오류: 숫자를 입력해주세요.")
    except Exception as e:
        print(f"\n❌ 등농도곡선 생성 오류: {str(e)}")
        traceback.print_exc()


def _cmd_search_papers(config):
    """학술 논문 검색 테스트 명령"""
    try:
        search_config = config.config.get('search', {})
        if not search_config.get('enabled', False):
            print("\n⚠️ 학술 논문 검색이 비활성화되어 있습니다.")
            print("   config.yaml에서 search.enabled: true로 설정하세요.")
            return

        print("\n🔍 학술 논문 검색을 시작합니다...")
        results = search_all_papers(config)

        if not results:
            print("\n⚠️ 검색 결과가 없습니다.")
            return

        print("\n" + "="*70)
        print("   학술 논문 검색 결과")
        print("="*70)

        for key, text in results.items():
            if not text:
                continue
            label = f"일반 검색" if key == 'general' else f"섹션 {key}"
            print(f"\n--- [{label}] ---")
            # 각 섹션별로 최대 1000자만 미리보기
            if len(text) > 1000:
                print(text[:1000])
                print(f"  ... ({len(text):,}자 중 1000자만 표시)")
            else:
                print(text)

        total_chars = sum(len(v) for v in results.values() if v)
        print(f"\n✓ 검색 완료: {len(results)}개 카테고리, 총 {total_chars:,}자")

    except Exception as e:
        print(f"\n❌ 논문 검색 오류: {str(e)}")
        traceback.print_exc()


# ==========================================
# 직접 실행 지원
# ==========================================
if __name__ == "__main__":
    config, templates = init_system()
    menu_loop(config, templates)
