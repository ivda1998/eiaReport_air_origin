import os
import time
import subprocess
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import yaml
from google.genai import Client

try:
    from PyPDF2 import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Mm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    Document = None


# ==========================================
# 설정 로딩
# ==========================================
class ProjectConfig:
    """프로젝트 설정 관리 클래스"""

    def __init__(self, config_path="config.yaml"):
        self.config_path = config_path
        self.config = self.load_config()
        self.setup_paths()
        self.setup_matplotlib()

    def load_config(self):
        """config.yaml 파일 로딩"""
        if not os.path.exists(self.config_path):
            raise FileNotFoundError(
                f"설정 파일을 찾을 수 없습니다: {self.config_path}\n"
                f"config.yaml 파일을 프로젝트 루트에 생성해주세요."
            )

        with open(self.config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)

        print(f"✓ 설정 파일 로딩 완료: {self.config_path}")
        print(f"  프로젝트명: {config['project']['name']}")
        return config

    def setup_paths(self):
        """경로 설정"""
        paths = self.config['paths']
        base_dir = paths['base_dir']

        self.BASE_DIR = base_dir
        self.INPUT_DIR = os.path.join(base_dir, paths['input_dir'])
        self.OUTPUT_DIR = os.path.join(base_dir, paths['output_dir'])
        self.SAMPLE_DIR = os.path.join(base_dir, paths['sample_dir'])
        self.DATA_DIR = os.path.join(base_dir, paths['data_dir'])
        self.ENGINE_DIR = os.path.join(base_dir, paths['engine_dir'])

        aermod = self.config['aermod']
        self.AERMOD_EXE = os.path.join(self.ENGINE_DIR, aermod['exe_name'])
        self.INPUT_FILE = os.path.join(self.INPUT_DIR, aermod['input_file'])
        self.OUTPUT_FILE = os.path.join(self.INPUT_DIR, aermod['output_file'])

    def setup_matplotlib(self):
        """matplotlib 한글 폰트 설정"""
        font_config = self.config['font']
        matplotlib.rcParams['font.family'] = font_config['family']
        matplotlib.rcParams['axes.unicode_minus'] = font_config['unicode_minus']

    def get_pollutant_config(self, pollutant_name):
        """오염물질 설정 조회"""
        for p in self.config['pollutants']:
            if p['name'] == pollutant_name:
                return p
        return None


# ==========================================
# 템플릿 관리 클래스
# ==========================================
class TemplateManager:
    """YAML 템플릿 로딩 및 프롬프트 구성 관리"""

    def __init__(self, templates_dir):
        self.templates_dir = templates_dir
        self.toc = None
        self.detailed = None
        self.rules = None
        self.style_guide = None
        self.style = None
        self.fonts = None
        self.report_structure = None
        self.load_all()

    def load_all(self):
        """모든 YAML 템플릿 로딩"""
        if not os.path.exists(self.templates_dir):
            print(f"⚠️ 템플릿 폴더를 찾을 수 없습니다: {self.templates_dir}")
            return

        file_map = {
            'toc': 'toc_structure.yaml',
            'detailed': 'report_template_detailed.yaml',
            'rules': 'section_rules.yaml',
            'style_guide': 'style_guide.yaml',
            'style': 'style.yaml',
            'fonts': 'fonts.yaml',
            'report_structure': 'report_structure.yaml',
        }

        loaded_count = 0
        for attr, filename in file_map.items():
            data = self._load_yaml(filename)
            setattr(self, attr, data)
            if data:
                loaded_count += 1

        print(f"✓ 템플릿 로딩 완료: {loaded_count}/{len(file_map)}개")

    def _load_yaml(self, filename):
        """단일 YAML 파일 로딩"""
        filepath = os.path.join(self.templates_dir, filename)
        if not os.path.exists(filepath):
            return None
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            print(f"⚠️ 템플릿 로딩 실패 ({filename}): {e}")
            return None

    def get_section_list(self):
        """toc_structure.yaml에서 섹션 목록 반환"""
        if not self.toc or 'sections' not in self.toc:
            return []
        return self.toc['sections']

    def get_writing_rules_summary(self):
        """style_guide.yaml에서 작성 규칙 요약 반환"""
        if not self.style_guide:
            return ""

        lines = []
        # 작성 원칙
        principles = self.style_guide.get('writing_principles', {})
        rules = principles.get('rules', [])
        if rules:
            lines.append("[작성 규칙]")
            for r in rules:
                lines.append(f"- {r}")

        # 금지 사항
        prohibited = self.style_guide.get('prohibited', [])
        if prohibited:
            lines.append("\n[금지 사항]")
            for p in prohibited:
                lines.append(f"- {p}")

        # 수치 표기
        num_fmt = self.style_guide.get('number_formatting', {})
        if num_fmt:
            lines.append(f"\n[수치 표기] 소수점 {num_fmt.get('decimals', 2)}자리, 천단위 쉼표")

        # 확실성 표현
        certainty = self.style_guide.get('certainty_levels', {})
        if certainty:
            lines.append("\n[확실성 표현]")
            for level, phrases in certainty.items():
                lines.append(f"  {level}: {', '.join(phrases[:2])}")

        return '\n'.join(lines)

    def get_standard_phrases(self, section_id):
        """해당 섹션에 맞는 표준 문구 반환"""
        if not self.detailed:
            return ""

        phrases = self.detailed.get('standard_phrases', {})
        section_templates = self.detailed.get('section_templates', {})

        lines = []

        # 섹션별 매핑
        section_num = section_id.split('.')[0] if '.' in section_id else section_id

        if section_num == "1":
            # 사업 개요 - opening 문구
            opening = phrases.get('opening', [])
            if opening:
                lines.append("[표준 서두 문구]")
                for p in opening:
                    lines.append(f"- {p}")

        elif section_num == "2":
            # 대기질 현황 - methodology.survey
            survey = phrases.get('methodology', {}).get('survey', [])
            if survey:
                lines.append("[조사 방법론 표준 문구]")
                for p in survey:
                    lines.append(f"- {p}")
            # 섹션 템플릿에서 추가 문구
            for key in ["2.0", "2.1", "2.2", "2.3", "2.4", "2.5"]:
                tmpl = section_templates.get(key, {})
                sp = tmpl.get('standard_phrases', [])
                if sp:
                    lines.append(f"\n[{key} 표준 문구]")
                    for p in sp:
                        lines.append(f"- {p}")

        elif section_num == "3":
            # 영향 예측 방법 - methodology.modeling
            modeling = phrases.get('methodology', {}).get('modeling', [])
            if modeling:
                lines.append("[모델링 방법론 표준 문구]")
                for p in modeling:
                    lines.append(f"- {p}")
            # 섹션별 표준 문구
            for key in ["3.0", "3.1", "3.2", "3.3", "3.4", "3.5", "3.6"]:
                tmpl = section_templates.get(key, {})
                sp = tmpl.get('standard_phrases', [])
                if sp:
                    lines.append(f"\n[{key} 표준 문구]")
                    for p in sp:
                        lines.append(f"- {p}")
                # subsections의 표준 문구도
                for sub_key, sub_val in tmpl.get('subsections', {}).items():
                    if isinstance(sub_val, dict):
                        sub_sp = sub_val.get('standard_phrases', [])
                        if sub_sp:
                            lines.append(f"\n[{sub_key} 표준 문구]")
                            for p in sub_sp:
                                lines.append(f"- {p}")

        elif section_num == "4":
            # 영향 예측 결과 - results
            conc = phrases.get('results', {}).get('concentration', [])
            comp = phrases.get('results', {}).get('compliance', [])
            if conc:
                lines.append("[예측 결과 표준 문구]")
                for p in conc:
                    lines.append(f"- {p}")
            if comp:
                lines.append("\n[적합성 평가 표준 문구]")
                for p in comp:
                    lines.append(f"- {p}")

        elif section_num == "5":
            # 환경기준 적합성 - validation
            val = phrases.get('validation', [])
            if val:
                lines.append("[타당성 검증 표준 문구]")
                for p in val:
                    lines.append(f"- {p}")

        elif section_num == "6":
            # 저감방안 - mitigation
            const = phrases.get('mitigation', {}).get('construction', [])
            oper = phrases.get('mitigation', {}).get('operation', [])
            if const:
                lines.append("[공사시 저감방안]")
                for p in const:
                    lines.append(f"- {p}")
            if oper:
                lines.append("\n[운영시 저감방안]")
                for p in oper:
                    lines.append(f"- {p}")

        elif section_num == "7":
            # 결론
            concl = phrases.get('conclusion', [])
            if concl:
                lines.append("[결론 표준 문구]")
                for p in concl:
                    lines.append(f"- {p}")

        return '\n'.join(lines)

    def get_section_rules(self, section_id):
        """section_rules.yaml에서 해당 섹션 구조 요건 반환"""
        if not self.rules:
            return ""

        sections = self.rules.get('sections', {})
        section_num = section_id.split('.')[0] if '.' in section_id else section_id
        section_data = sections.get(section_num, {})

        if not section_data:
            return ""

        lines = []
        lines.append(f"[섹션 구조 요건: {section_data.get('title', '')}]")
        lines.append(f"목적: {section_data.get('purpose', '')}")

        subsections = section_data.get('subsections', {})
        if isinstance(subsections, dict):
            for sub_id, sub_data in subsections.items():
                if isinstance(sub_data, dict):
                    lines.append(f"\n  {sub_id} {sub_data.get('title', '')}")
                    length = sub_data.get('length', '')
                    if length:
                        lines.append(f"    분량: {length}")

                    req = sub_data.get('required_elements', [])
                    if req:
                        lines.append(f"    필수 요소:")
                        for r in req:
                            if isinstance(r, str):
                                lines.append(f"      - {r}")
                            elif isinstance(r, dict):
                                lines.append(f"      - [{r.get('type', '')}] {r.get('name', r.get('title', ''))}")

                    guide = sub_data.get('writing_guide', [])
                    if guide:
                        lines.append(f"    작성 가이드:")
                        for g in guide:
                            lines.append(f"      - {g}")

        # 작성 주의사항
        notes = self.rules.get('writing_notes', {})
        critical = notes.get('critical_sections', [])
        for c in critical:
            if c.get('section', '').startswith(section_num):
                lines.append(f"\n⚠️ {c['section']}: {c['note']}")

        return '\n'.join(lines)

    def get_toc_structure_text(self, section_id):
        """toc_structure.yaml에서 해당 섹션의 목차 구조 텍스트 반환"""
        if not self.toc:
            return ""

        sections = self.toc.get('sections', [])
        for sec in sections:
            if sec.get('id') == section_id:
                lines = [f"{sec['id']}. {sec.get('title', '')}"]
                for sub in sec.get('subsections', []):
                    lines.append(f"  {sub['id']} {sub.get('title', '')}")
                    for content_item in sub.get('content', []):
                        lines.append(f"    - {content_item}")
                return '\n'.join(lines)
        return ""

    def get_section_prompt(self, section_id, context_data, previous_sections_summary=""):
        """섹션별 AI 프롬프트 조립"""
        writing_rules = self.get_writing_rules_summary()
        toc_text = self.get_toc_structure_text(section_id)
        section_rules = self.get_section_rules(section_id)
        standard_phrases = self.get_standard_phrases(section_id)

        # 섹션별 데이터 매핑
        section_num = section_id.split('.')[0] if '.' in section_id else section_id
        data_section = ""

        if section_num in ["1"]:
            data_section = f"""[사업 정보]
프로젝트명: {context_data.get('project_name', '미정')}
"""

        elif section_num in ["2"]:
            measurement = context_data.get('measurement_data', '현황측정자료 없음')
            data_section = f"""[현황측정자료]
{measurement[:5000] if measurement else '현황측정자료 없음'}
"""

        elif section_num in ["3"]:
            aermod = context_data.get('aermod_data', '')
            facility = context_data.get('facility_data', '')
            data_section = f"""[AERMOD 입력 조건 (상단부)]
{aermod[:5000] if aermod else 'AERMOD 데이터 없음'}

[정온시설 데이터]
{facility[:5000] if facility else '정온시설 정보 없음'}
"""

        elif section_num in ["4", "5"]:
            aermod = context_data.get('aermod_data', '')
            facility = context_data.get('facility_data', '')
            measurement = context_data.get('measurement_data', '')
            data_section = f"""[AERMOD 모델링 결과]
{aermod[:15000] if aermod else 'AERMOD 데이터 없음'}

[정온시설 데이터]
{facility[:8000] if facility else '정온시설 정보 없음'}

[현황측정자료]
{measurement[:3000] if measurement else '현황측정자료 없음'}
"""

        elif section_num in ["6"]:
            data_section = ""  # 표준 문구만 사용

        elif section_num in ["7"]:
            data_section = f"""[이전 섹션 요약]
{previous_sections_summary[:5000] if previous_sections_summary else '이전 섹션 정보 없음'}
"""

        elif section_num in ["8"]:
            data_section = ""  # 참고문헌 형식만

        # 샘플 보고서 참조 (섹션 1~5에만)
        sample_ref = ""
        if section_num in ["1", "2", "3", "4", "5"]:
            sample_text = context_data.get('sample_text', '')
            if sample_text:
                sample_ref = f"\n[참조: 샘플 보고서 양식]\n{sample_text[:5000]}\n"

        prompt = f"""당신은 환경영향평가 대기질 분야 전문가입니다.
아래 지침에 따라 환경영향평가서의 "{toc_text.split(chr(10))[0] if toc_text else section_id}" 섹션을 작성하세요.

{writing_rules}

[섹션 구조]
{toc_text}

{section_rules}

{f'[표준 문구 (적극 활용하세요)]' + chr(10) + standard_phrases if standard_phrases else ''}

{data_section}
{sample_ref}

[출력 형식]
- 마크다운 형식으로 작성 (# 제목, ## 소제목, | 표 |)
- 표는 마크다운 표 형식 사용
- 문체: ~함, ~임 서술형
- 수치는 소수점 2자리, 단위 뒤 공백
- 선행연구 인용 시 저자(연도) 형식 사용

위 정보를 바탕으로 해당 섹션을 작성하세요.
"""
        return prompt

    def get_docx_styles(self):
        """style.yaml + fonts.yaml 병합하여 DOCX 스타일 반환"""
        merged = {}
        if self.style:
            merged.update(self.style)
        if self.fonts:
            # fonts.yaml의 주요 항목 병합
            merged['heading_styles'] = self.fonts.get('heading_styles', {})
            merged['body_styles'] = self.fonts.get('body_styles', {})
            merged['table_styles'] = self.fonts.get('table_styles', {})
            merged['figure_styles'] = self.fonts.get('figure_styles', {})
            merged['cover_fonts'] = self.fonts.get('cover', {})
            merged['font_colors'] = self.fonts.get('colors', {})
        return merged


# 전역 설정 객체
try:
    CONFIG = ProjectConfig()
    client = Client(api_key=CONFIG.config['ai']['api_key'])
except Exception as e:
    print(f"❌ 설정 초기화 실패: {e}")
    print("   config.yaml 파일을 확인하고 다시 시도하세요.")
    exit(1)

# 템플릿 관리 객체
try:
    TEMPLATES_DIR = os.path.join(CONFIG.BASE_DIR, "templates")
    TEMPLATES = TemplateManager(TEMPLATES_DIR)
except Exception as e:
    print(f"⚠️ 템플릿 로딩 실패 (기존 방식으로 동작): {e}")
    TEMPLATES = None


# ==========================================
# 1. 참조 자료 로딩 함수
# ==========================================
def load_sample_reports():
    """sample 폴더의 PDF 보고서를 텍스트로 추출"""
    samples = []

    if not os.path.exists(CONFIG.SAMPLE_DIR):
        print(f"⚠️ 샘플 폴더를 찾을 수 없습니다: {CONFIG.SAMPLE_DIR}")
        return samples

    report_config = CONFIG.config['report']
    max_pages = report_config['sample_max_pages']
    max_chars = report_config['sample_max_chars']

    print(f"\n📂 샘플 보고서 로딩 중...")
    for filename in os.listdir(CONFIG.SAMPLE_DIR):
        if filename.endswith('.pdf'):
            filepath = os.path.join(CONFIG.SAMPLE_DIR, filename)
            try:
                if PdfReader is None:
                    print(f"  ⚠️ PyPDF2가 설치되지 않아 {filename}을 읽을 수 없습니다.")
                    continue

                reader = PdfReader(filepath)
                text = ""
                page_count = min(max_pages, len(reader.pages))
                for page_num in range(page_count):
                    text += reader.pages[page_num].extract_text()

                samples.append({
                    'filename': filename,
                    'content': text[:max_chars]
                })
                print(f"  ✓ {filename} 로딩 완료 ({len(reader.pages)} 페이지)")
            except Exception as e:
                print(f"  ✗ {filename} 로딩 실패: {e}")

    return samples


def load_facility_data():
    """data 폴더의 정온시설 위치 정보 로딩"""
    data_files = CONFIG.config['data_files']
    facility_file = os.path.join(CONFIG.DATA_DIR, data_files['facilities'])

    if not os.path.exists(facility_file):
        print(f"⚠️ 정온시설 파일을 찾을 수 없습니다: {facility_file}")
        return None

    try:
        df = pd.read_excel(facility_file)
        header_rows = data_files['facilities_structure']['header_rows']

        if len(df) > header_rows:
            df = df.iloc[header_rows:]

        actual_count = len(df)
        print(f"✓ 정온시설 데이터 로딩 완료 ({actual_count}개 시설)")

        pd.set_option('display.max_colwidth', None)
        pd.set_option('display.max_rows', None)
        result = df.to_string(index=False, max_colwidth=100)
        pd.reset_option('display.max_colwidth')
        pd.reset_option('display.max_rows')

        return result
    except Exception as e:
        print(f"✗ 정온시설 데이터 로딩 실패: {e}")
        return None


def load_measurement_data():
    """data 폴더의 현황측정자료 로딩"""
    data_files = CONFIG.config['data_files']
    measurement_file = os.path.join(CONFIG.DATA_DIR, data_files['measurements'])

    if not os.path.exists(measurement_file):
        print(f"⚠️ 현황측정자료를 찾을 수 없습니다: {measurement_file}")
        return None

    try:
        df = pd.read_excel(measurement_file)
        print(f"✓ 현황측정자료 로딩 완료 ({len(df)}개 데이터)")
        return df.to_string(index=False)
    except Exception as e:
        print(f"✗ 현황측정자료 로딩 실패: {e}")
        return None


def load_reference_papers():
    """외부 전문자료 폴더에서 논문/보고서 로딩"""
    references_dir = os.path.join(CONFIG.DATA_DIR, "references")
    references = []

    if not os.path.exists(references_dir):
        print(f"ℹ️ references 폴더가 없습니다. 외부 전문자료 없이 진행합니다.")
        return None

    report_config = CONFIG.config['report']
    max_pages = report_config['reference_max_pages']
    max_chars = report_config['reference_max_chars']

    print(f"\n📚 외부 전문자료 로딩 중...")
    for filename in os.listdir(references_dir):
        if filename.endswith(('.txt', '.pdf', '.docx')):
            filepath = os.path.join(references_dir, filename)
            try:
                if filename.endswith('.txt'):
                    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                        references.append({
                            'filename': filename,
                            'content': content[:max_chars]
                        })
                        print(f"  ✓ {filename} 로딩 완료")
                elif filename.endswith('.pdf') and PdfReader:
                    reader = PdfReader(filepath)
                    text = ""
                    page_count = min(max_pages, len(reader.pages))
                    for page_num in range(page_count):
                        text += reader.pages[page_num].extract_text()
                    references.append({
                        'filename': filename,
                        'content': text[:max_chars]
                    })
                    print(f"  ✓ {filename} 로딩 완료")
            except Exception as e:
                print(f"  ✗ {filename} 로딩 실패: {e}")

    if references:
        print(f"✓ 총 {len(references)}개 외부 전문자료 로딩 완료")
        return "\n\n".join([
            f"[참고문헌: {item['filename']}]\n{item['content']}"
            for item in references
        ])
    return None


# ==========================================
# 2. 등농도곡선 생성 함수
# ==========================================
def parse_aermod_concentrations(output_file):
    """AERMOD 출력 파일에서 농도 격자 데이터 추출"""
    try:
        input_dir = os.path.dirname(output_file)
        dat_files = [f for f in os.listdir(input_dir) if f.endswith('.dat')]

        if dat_files:
            dat_file = os.path.join(input_dir, dat_files[0])
            print(f"✓ 격자 데이터 파일 발견: {dat_files[0]}")

            concentrations = []
            x_coords = []
            y_coords = []

            with open(dat_file, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()

            data_started = False
            for line in lines:
                if '____________' in line:
                    data_started = True
                    continue

                if data_started and line.strip() and not line.startswith('*'):
                    parts = line.strip().split()
                    if len(parts) >= 3:
                        try:
                            x = float(parts[0])
                            y = float(parts[1])
                            conc = float(parts[2])

                            if conc > 0:
                                x_coords.append(x)
                                y_coords.append(y)
                                concentrations.append(conc)
                        except ValueError:
                            continue

            if concentrations:
                print(f"✓ .dat 파일에서 {len(concentrations)}개 격자 데이터 파싱 완료")
                print(f"   X 범위: {min(x_coords):.1f} ~ {max(x_coords):.1f}")
                print(f"   Y 범위: {min(y_coords):.1f} ~ {max(y_coords):.1f}")
                print(f"   농도 범위: {min(concentrations):.3f} ~ {max(concentrations):.3f}")

                return {
                    'x': np.array(x_coords),
                    'y': np.array(y_coords),
                    'conc': np.array(concentrations)
                }

        print("⚠️ .dat 파일을 찾을 수 없습니다. .out 파일에서 요약 데이터를 사용합니다.")

        with open(output_file, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()

        concentrations = []
        x_coords = []
        y_coords = []

        import re
        pattern = r'VALUE IS\s+([\d.]+)\s+AT\s+\(\s*([\d.]+),\s*([\d.]+),'
        matches = re.findall(pattern, content)

        if matches:
            print(f"✓ {len(matches)}개의 농도 데이터 발견")
            for conc, x, y in matches:
                concentrations.append(float(conc))
                x_coords.append(float(x))
                y_coords.append(float(y))

        if not concentrations:
            return None

        print(f"✓ 최종 파싱 완료: {len(concentrations)}개 데이터")
        print(f"   X 범위: {min(x_coords):.1f} ~ {max(x_coords):.1f}")
        print(f"   Y 범위: {min(y_coords):.1f} ~ {max(y_coords):.1f}")
        print(f"   농도 범위: {min(concentrations):.3f} ~ {max(concentrations):.3f}")

        return {
            'x': np.array(x_coords),
            'y': np.array(y_coords),
            'conc': np.array(concentrations)
        }

    except Exception as e:
        print(f"❌ AERMOD 농도 데이터 파싱 오류: {e}")
        import traceback
        traceback.print_exc()
        return None


def load_base_dxf():
    """DXF 베이스 맵 로드"""
    try:
        data_files = CONFIG.config['data_files']
        dxf_file = os.path.join(CONFIG.DATA_DIR, data_files['base_map'])

        if not os.path.exists(dxf_file):
            print(f"⚠️ DXF 파일을 찾을 수 없습니다: {dxf_file}")
            return None

        try:
            import ezdxf
        except ImportError:
            print("⚠️ ezdxf 라이브러리가 없습니다. pip install ezdxf를 실행하세요.")
            return None

        doc = ezdxf.readfile(dxf_file)
        modelspace = doc.modelspace()

        print(f"✓ DXF 파일 로드 완료: {len(list(modelspace))}개 객체")
        return doc, modelspace

    except Exception as e:
        print(f"⚠️ DXF 파일 로드 실패: {e}")
        return None


def load_facility_locations():
    """정온시설 위치 데이터 로드"""
    try:
        data_files = CONFIG.config['data_files']
        facility_file = os.path.join(CONFIG.DATA_DIR, data_files['facilities'])

        if not os.path.exists(facility_file):
            print(f"⚠️ 정온시설 파일을 찾을 수 없습니다.")
            return None

        df = pd.read_excel(facility_file)

        structure = data_files['facilities_structure']
        header_rows = structure['header_rows']
        name_col = structure['name_column']
        x_col = structure['x_column']
        y_col = structure['y_column']

        facilities = []
        for idx, row in df.iterrows():
            if idx < header_rows:
                continue
            try:
                x_val = row.iloc[x_col]
                y_val = row.iloc[y_col]
                name = row.iloc[name_col] if pd.notna(row.iloc[name_col]) else f"시설{idx-header_rows+1}"

                if pd.notna(x_val) and pd.notna(y_val):
                    x = float(x_val)
                    y = float(y_val)
                    facilities.append({'name': name, 'x': x, 'y': y})
            except:
                continue

        if facilities:
            print(f"✓ 정온시설 {len(facilities)}개 위치 로드 완료")
            return facilities
        return None

    except Exception as e:
        print(f"⚠️ 정온시설 데이터 로드 실패: {e}")
        return None


def create_isopleth(data, output_image_path, pollutant_name="오염물질", unit="μg/m³",
                   level_min=None, level_max=None, level_count=10, level_interval=None):
    """등농도곡선 생성"""
    try:
        x = data['x']
        y = data['y']
        conc = data['conc']

        if len(x) < 4:
            print(f"⚠️ 데이터 포인트가 너무 적습니다 ({len(x)}개). 최소 4개 필요.")
            return None

        dxf_data = load_base_dxf()
        facilities = load_facility_locations()

        # 설정에서 격자 해상도 가져오기
        isopleth_config = CONFIG.config['isopleth']
        grid_res = isopleth_config['grid_resolution']

        xi = np.linspace(x.min(), x.max(), grid_res)
        yi = np.linspace(y.min(), y.max(), grid_res)
        Xi, Yi = np.meshgrid(xi, yi)

        # 보간 방법 설정
        interp_method = isopleth_config['interpolation_method']
        from scipy.interpolate import griddata
        try:
            Zi = griddata((x, y), conc, (Xi, Yi), method=interp_method, fill_value=0)
        except Exception as e:
            print(f"⚠️ {interp_method} 보간 실패, nearest 방법 시도 중...")
            Zi = griddata((x, y), conc, (Xi, Yi), method='nearest', fill_value=0)

        # 시각화 설정
        vis_config = isopleth_config['visualization']
        fig, ax = plt.subplots(
            figsize=(vis_config['figure_width'], vis_config['figure_height']),
            facecolor='white'
        )
        ax.set_facecolor('white')

        # DXF 베이스맵 그리기
        if dxf_data:
            doc, modelspace = dxf_data
            entity_count = 0
            for entity in modelspace:
                try:
                    if entity.dxftype() == 'LINE':
                        start = entity.dxf.start
                        end = entity.dxf.end
                        ax.plot([start.x, end.x], [start.y, end.y],
                               color=vis_config['base_map_color'],
                               linewidth=vis_config['base_map_linewidth'],
                               alpha=vis_config['base_map_alpha'], zorder=1)
                        entity_count += 1

                    elif entity.dxftype() == 'LWPOLYLINE':
                        points = list(entity.get_points())
                        if len(points) > 1:
                            xs = [p[0] for p in points]
                            ys = [p[1] for p in points]
                            if entity.closed:
                                xs.append(xs[0])
                                ys.append(ys[0])
                            ax.plot(xs, ys, color=vis_config['base_map_color'],
                                   linewidth=vis_config['base_map_linewidth'],
                                   alpha=vis_config['base_map_alpha'], zorder=1)
                            entity_count += 1

                    elif entity.dxftype() == 'POLYLINE':
                        points = list(entity.points())
                        if len(points) > 1:
                            xs = [p[0] for p in points]
                            ys = [p[1] for p in points]
                            ax.plot(xs, ys, color=vis_config['base_map_color'],
                                   linewidth=vis_config['base_map_linewidth'],
                                   alpha=vis_config['base_map_alpha'], zorder=1)
                            entity_count += 1

                    elif entity.dxftype() == 'CIRCLE':
                        center = entity.dxf.center
                        radius = entity.dxf.radius
                        circle = plt.Circle((center.x, center.y), radius,
                                          color=vis_config['base_map_color'], fill=False,
                                          linewidth=vis_config['base_map_linewidth'],
                                          alpha=vis_config['base_map_alpha'], zorder=1)
                        ax.add_patch(circle)
                        entity_count += 1

                except Exception as e:
                    continue

            print(f"  ✓ DXF 객체 {entity_count}개 렌더링 완료")

        # 등농도선 레벨 설정
        if level_min is None:
            level_min = max(conc.min(), 0.001)
        if level_max is None:
            level_max = conc.max()

        if level_interval is not None:
            levels = np.arange(level_min, level_max + level_interval, level_interval)
            level_count = len(levels)
            print(f"  ✓ 간격 {level_interval} {unit}로 {level_count}개 레벨 생성")
        else:
            levels = np.linspace(level_min, level_max, level_count)

        # 등농도선 그리기
        contour = ax.contour(Xi, Yi, Zi, levels=levels,
                            colors=vis_config['contour_color'],
                            linewidths=vis_config['contour_linewidth'],
                            alpha=0.8, zorder=2)

        ax.clabel(contour, inline=True, fontsize=8, fmt=f'%.2f {unit}')

        # 최대 농도 지점 표시
        max_idx = np.argmax(conc)
        ax.plot(x[max_idx], y[max_idx], 'r*', markersize=18,
                label=f'최대농도: {conc[max_idx]:.3f} {unit}', zorder=4)

        # 정온시설 표시
        if facilities:
            for idx, facility in enumerate(facilities, start=1):
                ax.plot(facility['x'], facility['y'], 'o',
                       color=vis_config['facility_color'],
                       markersize=vis_config['facility_size'],
                       markeredgecolor='black',
                       markeredgewidth=0.5, alpha=0.8, zorder=3)
                ax.text(facility['x'], facility['y'], str(idx),
                       fontsize=vis_config['label_fontsize'],
                       fontweight='bold', color='white',
                       ha='center', va='center', zorder=4)

            ax.plot([], [], 'o', color=vis_config['facility_color'],
                   markersize=vis_config['facility_size'],
                   markeredgecolor='black', markeredgewidth=0.5,
                   label=f'정온시설 ({len(facilities)}개)')

        ax.set_xlabel('X 좌표 (m)', fontsize=12, fontweight='bold')
        ax.set_ylabel('Y 좌표 (m)', fontsize=12, fontweight='bold')

        if level_interval is not None:
            title_str = f'{pollutant_name} 등농도곡선\n({level_min:.3f} ~ {level_max:.3f} {unit}, 간격 {level_interval} {unit})'
        else:
            title_str = f'{pollutant_name} 등농도곡선\n({level_min:.3f} ~ {level_max:.3f} {unit}, {level_count}단계)'

        ax.set_title(title_str, fontsize=15, fontweight='bold', pad=20)
        ax.legend(loc='upper right', fontsize=10, framealpha=0.9)
        ax.grid(True, alpha=0.2, linestyle='--', color='gray')
        ax.set_aspect('equal')

        x_margin = (x.max() - x.min()) * 0.05
        y_margin = (y.max() - y.min()) * 0.05
        ax.set_xlim(x.min() - x_margin, x.max() + x_margin)
        ax.set_ylim(y.min() - y_margin, y.max() + y_margin)

        plt.tight_layout()
        plt.savefig(output_image_path, dpi=vis_config['dpi'],
                   bbox_inches='tight', facecolor='white')
        plt.close()

        print(f"✓ 등농도곡선 생성 완료: {output_image_path}")
        print(f"   레벨 범위: {level_min:.3f} ~ {level_max:.3f} {unit}")
        if level_interval is not None:
            print(f"   간격: {level_interval} {unit}, 총 {level_count}개")
        else:
            print(f"   레벨 개수: {level_count}개")
        return output_image_path

    except ImportError as ie:
        print(f"❌ 필요한 라이브러리가 없습니다: {ie}")
        print("   pip install scipy ezdxf를 실행하세요.")
        return None
    except Exception as e:
        print(f"❌ 등농도곡선 생성 오류: {e}")
        import traceback
        traceback.print_exc()
        return None


def generate_all_isopleths(level_min=None, level_max=None, level_count=10, level_interval=None):
    """AERMOD 결과로부터 모든 오염물질의 등농도곡선 생성"""
    if not os.path.exists(CONFIG.OUTPUT_FILE):
        print("❌ AERMOD 출력 파일을 찾을 수 없습니다.")
        return []

    print("\n🎨 등농도곡선 생성 중...")

    pollutant_name = "PM2.5"
    try:
        with open(CONFIG.OUTPUT_FILE, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read(5000)
            import re
            match = re.search(r'POLLUTID\s+(\S+)', content)
            if match:
                pollutant_name = match.group(1)
                print(f"✓ 오염물질: {pollutant_name}")
    except:
        pass

    data = parse_aermod_concentrations(CONFIG.OUTPUT_FILE)
    if data is None:
        return []

    if not os.path.exists(CONFIG.OUTPUT_DIR):
        os.makedirs(CONFIG.OUTPUT_DIR)

    image_paths = []

    # 설정에서 오염물질 정보 가져오기
    pollutant_config = CONFIG.get_pollutant_config(pollutant_name)
    if pollutant_config:
        display_name = pollutant_config['display_name']
        unit = pollutant_config['unit']
    else:
        display_name = pollutant_name
        unit = "μg/m³"

    # 레벨 범위 자동 설정
    if level_min is None and level_max is None:
        conc = data['conc']
        if pollutant_name == "PM2.5":
            level_max = max(50.0, conc.max())
            level_min = 0.1
        else:
            level_max = conc.max()
            level_min = max(0.1, conc.min())

        print(f"✓ 자동 레벨 범위: {level_min:.3f} ~ {level_max:.3f} {unit}")

    image_path = os.path.join(CONFIG.OUTPUT_DIR, f"등농도곡선_{pollutant_name}.png")
    result = create_isopleth(data, image_path, pollutant_name=display_name, unit=unit,
                            level_min=level_min, level_max=level_max,
                            level_count=level_count, level_interval=level_interval)
    if result:
        image_paths.append(result)

    return image_paths


# ==========================================
# 3. AERMOD 실행 함수
# ==========================================
def run_aermod():
    """AERMOD 모델링 실행"""
    if not os.path.exists(CONFIG.AERMOD_EXE):
        return f"❌ AERMOD 실행 파일을 찾을 수 없습니다: {CONFIG.AERMOD_EXE}"

    if not os.path.exists(CONFIG.INPUT_FILE):
        return f"❌ 입력 파일을 찾을 수 없습니다: {CONFIG.INPUT_FILE}"

    print(f"\n🚀 AERMOD 시뮬레이션 시작...")
    print(f"   입력 파일: {CONFIG.INPUT_FILE}")

    try:
        original_dir = os.getcwd()
        os.chdir(CONFIG.INPUT_DIR)

        aermod_config = CONFIG.config['aermod']
        timeout = aermod_config['timeout']

        result = subprocess.run(
            [CONFIG.AERMOD_EXE, aermod_config['input_file']],
            capture_output=True,
            text=True,
            timeout=timeout
        )

        os.chdir(original_dir)

        if result.returncode == 0:
            print("✅ AERMOD 모델링 완료")
            return "SUCCESS"
        else:
            print(f"⚠️ AERMOD 실행 중 경고 발생 (코드: {result.returncode})")
            return f"WARNING: {result.stderr[:500]}"

    except subprocess.TimeoutExpired:
        os.chdir(original_dir)
        return f"❌ AERMOD 실행 시간 초과 ({timeout}초)"
    except Exception as e:
        os.chdir(original_dir)
        return f"❌ AERMOD 실행 오류: {str(e)}"


# ==========================================
# 4. 결과 분석 및 보고서 작성
# ==========================================
def _call_ai(prompt):
    """단일 AI 호출 (재시도 로직 포함)"""
    ai_config = CONFIG.config['ai']
    max_retries = ai_config['max_retries']
    retry_delay = ai_config['retry_delay']

    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
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


def _prepare_context_data():
    """보고서 생성에 필요한 모든 컨텍스트 데이터 수집"""
    report_config = CONFIG.config['report']
    max_chars = report_config['aermod_result_max_chars']

    with open(CONFIG.OUTPUT_FILE, 'r', encoding='utf-8', errors='ignore') as f:
        aermod_data = f.read()[-max_chars:]

    print(f"\n📚 참조 자료 로딩 중...")
    sample_reports = load_sample_reports()
    facility_data = load_facility_data()
    measurement_data = load_measurement_data()
    reference_papers = load_reference_papers()

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
        'project_name': CONFIG.config['project']['name'],
    }


def _analyze_with_templates(context_data, isopleth_images):
    """템플릿 기반 섹션별 보고서 생성"""
    sections = TEMPLATES.get_section_list()
    if not sections:
        print("⚠️ 템플릿 섹션 목록이 비어있습니다. 기존 방식으로 전환합니다.")
        return _analyze_legacy(context_data, isopleth_images)

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

        prompt = TEMPLATES.get_section_prompt(
            section_id, context_data, previous_summary
        )

        try:
            section_text = _call_ai(prompt)
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


def _analyze_legacy(context_data, isopleth_images):
    """기존 단일 프롬프트 방식 보고서 생성 (fallback)"""
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

    result_text = _call_ai(prompt)
    if result_text:
        print(f"✓ AI 응답 생성 완료 ({len(result_text):,}자)")
        return result_text, isopleth_images
    else:
        raise Exception("AI 보고서 생성 실패")


def analyze_output(use_templates=None):
    """AERMOD 결과를 분석하여 환경영향평가서 작성

    Args:
        use_templates: True=템플릿 사용, False=기존방식, None=자동판단
    """
    if not os.path.exists(CONFIG.OUTPUT_FILE):
        return "❌ AERMOD 출력 파일을 찾을 수 없습니다. 먼저 모델링을 실행하세요.", []

    print(f"\n📊 AERMOD 결과 분석 중...")

    isopleth_images = generate_all_isopleths()
    context_data = _prepare_context_data()

    # 템플릿 사용 여부 결정
    if use_templates is None:
        use_templates = TEMPLATES is not None

    if use_templates and TEMPLATES is not None:
        print("📋 템플릿 기반 섹션별 생성 모드")
        return _analyze_with_templates(context_data, isopleth_images)
    else:
        print("📄 기존 단일 프롬프트 생성 모드")
        return _analyze_legacy(context_data, isopleth_images)


# ==========================================
# DOCX 스타일링 헬퍼 함수
# ==========================================
def _parse_pt(value):
    """'16pt' 같은 문자열에서 숫자 추출"""
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        return int(value.replace('pt', '').replace('mm', '').strip())
    return 10

def _parse_color(hex_color):
    """'#2E5090' -> RGBColor"""
    if not hex_color or not isinstance(hex_color, str):
        return RGBColor(0x21, 0x21, 0x21)
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

def _setup_page(doc, styles_config):
    """페이지 설정 (여백)"""
    page_config = styles_config.get('document', {}).get('page', {})
    if not page_config:
        page_config = styles_config.get('page', {})
    margins = page_config.get('margins', {})

    section = doc.sections[0]
    if margins:
        section.top_margin = Mm(int(str(margins.get('top', '25')).replace('mm', '')))
        section.bottom_margin = Mm(int(str(margins.get('bottom', '25')).replace('mm', '')))
        section.left_margin = Mm(int(str(margins.get('left', '30')).replace('mm', '')))
        section.right_margin = Mm(int(str(margins.get('right', '30')).replace('mm', '')))

def _setup_heading_styles(doc, styles_config):
    """H1~H3 제목 스타일 설정"""
    heading_config = styles_config.get('heading_styles', {})
    if not heading_config:
        heading_config = styles_config.get('headings', {})
    if not heading_config:
        return

    for level, key in [(1, 'h1'), (2, 'h2'), (3, 'h3')]:
        h_conf = heading_config.get(key, {})
        if not h_conf:
            continue
        try:
            style = doc.styles[f'Heading {level}']
            font = style.font
            font.name = h_conf.get('font_family', h_conf.get('font_name', '맑은 고딕'))
            font.size = Pt(_parse_pt(h_conf.get('font_size', 16 - (level-1)*2)))
            font.bold = h_conf.get('font_weight') == 'bold'
            font.color.rgb = _parse_color(h_conf.get('color', '#212121'))

            pf = style.paragraph_format
            pf.space_before = Pt(_parse_pt(h_conf.get('spacing_before', 12)))
            pf.space_after = Pt(_parse_pt(h_conf.get('spacing_after', 6)))
        except Exception:
            pass

def _setup_body_style(doc, styles_config):
    """본문 스타일 설정"""
    body_config = styles_config.get('body_styles', {}).get('normal', {})
    if not body_config:
        body_config = styles_config.get('body', {})
    if not body_config:
        return

    try:
        style = doc.styles['Normal']
        font = style.font
        font.name = body_config.get('font_family', '맑은 고딕')
        font.size = Pt(_parse_pt(body_config.get('font_size', 10)))
        font.color.rgb = _parse_color(body_config.get('color', '#212121'))

        pf = style.paragraph_format
        spacing = body_config.get('line_spacing', 1.6)
        if isinstance(spacing, (int, float)):
            pf.line_spacing = spacing
        pf.space_after = Pt(_parse_pt(body_config.get('spacing_after', 6)))
    except Exception:
        pass

def _add_cover_page(doc, styles_config, project_config):
    """표지 페이지 생성"""
    cover = styles_config.get('cover', styles_config.get('cover_fonts', {}))
    if not cover:
        return

    # 빈 줄 추가로 표지 상단 여백
    for _ in range(6):
        doc.add_paragraph()

    # 제목
    title_conf = cover.get('title', {})
    title_text = project_config.get('name', '환경영향평가서')
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(title_text)
    title_run.font.size = Pt(_parse_pt(title_conf.get('font_size', 24)))
    title_run.font.bold = title_conf.get('font_weight') == 'bold'
    title_run.font.color.rgb = _parse_color(title_conf.get('color', '#2E5090'))
    title_run.font.name = title_conf.get('font_family', '맑은 고딕')

    # 부제목
    subtitle_conf = cover.get('subtitle', {})
    subtitle_para = doc.add_paragraph()
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle_para.add_run("환경영향평가서 - 대기질 부문")
    subtitle_run.font.size = Pt(_parse_pt(subtitle_conf.get('font_size', 18)))
    subtitle_run.font.color.rgb = _parse_color(subtitle_conf.get('color', '#4A7BA7'))
    subtitle_run.font.name = subtitle_conf.get('font_family', '맑은 고딕')

    # 빈 줄
    for _ in range(4):
        doc.add_paragraph()

    # 프로젝트 정보
    info_conf = cover.get('info', {})
    info_items = [
        f"사업명: {project_config.get('name', '')}",
        f"작성일: {project_config.get('date', time.strftime('%Y-%m-%d'))}",
        f"작성 도구: AERMOD + AI 자동 작성 시스템",
    ]
    for item in info_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(item)
        r.font.size = Pt(_parse_pt(info_conf.get('font_size', 12)))
        r.font.name = info_conf.get('font_family', '맑은 고딕')
        r.font.color.rgb = _parse_color(info_conf.get('color', '#212121'))

    # 페이지 나누기
    doc.add_page_break()

def _add_header_footer(doc, styles_config, project_name):
    """머리글/바닥글 설정"""
    try:
        header_conf = styles_config.get('header', {})
        footer_conf = styles_config.get('footer', {})

        section = doc.sections[0]

        # 머리글
        if header_conf.get('enabled', False):
            header = section.header
            header.is_linked_to_previous = False
            p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            p.text = ""
            # 왼쪽: 프로젝트명
            run_left = p.add_run(f"{project_name}")
            run_left.font.size = Pt(_parse_pt(header_conf.get('font_size', 9)))
            run_left.font.color.rgb = _parse_color(header_conf.get('color', '#757575'))
            run_left.font.name = header_conf.get('font_family', '맑은 고딕')
            # 구분자
            run_sep = p.add_run("    |    ")
            run_sep.font.size = Pt(9)
            run_sep.font.color.rgb = _parse_color('#BDBDBD')
            # 오른쪽: 문서명
            run_right = p.add_run("환경영향평가서")
            run_right.font.size = Pt(_parse_pt(header_conf.get('font_size', 9)))
            run_right.font.color.rgb = _parse_color(header_conf.get('color', '#757575'))
            run_right.font.name = header_conf.get('font_family', '맑은 고딕')

        # 바닥글
        if footer_conf.get('enabled', False):
            footer = section.footer
            footer.is_linked_to_previous = False
            p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            p.text = ""
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(f"{project_name} | {time.strftime('%Y년 %m월 %d일')}")
            run.font.size = Pt(_parse_pt(footer_conf.get('font_size', 9)))
            run.font.color.rgb = _parse_color(footer_conf.get('color', '#757575'))
            run.font.name = footer_conf.get('font_family', '맑은 고딕')
    except Exception as e:
        print(f"  ⚠️ 머리글/바닥글 설정 실패: {e}")

def _style_table(table, styles_config):
    """표에 스타일 적용 (헤더 색상, 테두리 등)"""
    table_conf = styles_config.get('table_styles', {}).get('default', {})
    if not table_conf:
        table_conf = styles_config.get('tables', {}).get('default', {})
    if not table_conf:
        return

    header_conf = table_conf.get('header', {})
    bg_color = header_conf.get('background_color', '#2E5090').lstrip('#')
    text_color = header_conf.get('text_color', '#FFFFFF')

    # 첫 행(헤더)에 배경색 적용
    if table.rows:
        for cell in table.rows[0].cells:
            try:
                # 셀 배경색
                tc_pr = cell._tc.get_or_add_tcPr()
                shading = OxmlElement('w:shd')
                shading.set(qn('w:fill'), bg_color)
                shading.set(qn('w:val'), 'clear')
                tc_pr.append(shading)

                # 텍스트 색상 및 볼드
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in paragraph.runs:
                        run.font.bold = True
                        run.font.color.rgb = _parse_color(text_color)
                        run.font.size = Pt(_parse_pt(table_conf.get('font_size', 9)))
                        run.font.name = table_conf.get('font_family', '맑은 고딕')
            except Exception:
                pass

    # 데이터 행 스타일
    body_conf = table_conf.get('body', {})
    alt_color = body_conf.get('alternate_color', '#F5F5F5').lstrip('#')
    use_alt = body_conf.get('alternate_rows', False)

    for row_idx, row in enumerate(table.rows[1:], 1):
        for cell in row.cells:
            try:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(_parse_pt(table_conf.get('font_size', 9)))
                        run.font.name = table_conf.get('font_family', '맑은 고딕')

                # 교대 행 색상
                if use_alt and row_idx % 2 == 0:
                    tc_pr = cell._tc.get_or_add_tcPr()
                    shading = OxmlElement('w:shd')
                    shading.set(qn('w:fill'), alt_color)
                    shading.set(qn('w:val'), 'clear')
                    tc_pr.append(shading)
            except Exception:
                pass

def _apply_docx_styles(doc, styles_config):
    """DOCX 문서에 전체 스타일 적용"""
    _setup_page(doc, styles_config)
    _setup_heading_styles(doc, styles_config)
    _setup_body_style(doc, styles_config)


def save_report(content, format='txt', isopleth_images=None):
    """보고서를 파일로 저장"""
    if not os.path.exists(CONFIG.OUTPUT_DIR):
        os.makedirs(CONFIG.OUTPUT_DIR)

    if not content or not content.strip():
        print("⚠️ 저장할 내용이 비어있습니다.")
        return None

    if format == 'txt':
        report_file = os.path.join(CONFIG.OUTPUT_DIR, "대기질_환경영향평가서.txt")
        try:
            with open(report_file, "w", encoding="utf-8") as f:
                f.write(content)
                if isopleth_images:
                    f.write("\n\n[등농도곡선 이미지]\n")
                    for img_path in isopleth_images:
                        f.write(f"- {img_path}\n")
            print(f"\n💾 보고서 저장 완료: {report_file}")
            print(f"   파일 크기: {len(content):,}자")
            return report_file
        except Exception as e:
            print(f"❌ TXT 파일 저장 실패: {e}")
            return None

    elif format == 'docx':
        if Document is None:
            print("\n⚠️ python-docx가 설치되지 않았습니다. TXT 파일로 저장합니다.")
            return save_report(content, 'txt', isopleth_images)

        timestamp = time.strftime("%Y%m%d_%H%M%S")
        report_file = os.path.join(CONFIG.OUTPUT_DIR, f"대기질_환경영향평가서_{timestamp}.docx")

        base_report_file = os.path.join(CONFIG.OUTPUT_DIR, "대기질_환경영향평가서.docx")
        if os.path.exists(base_report_file):
            print(f"\n⚠️ 기존 파일이 있습니다. 새 파일명으로 저장합니다: {os.path.basename(report_file)}")

        try:
            doc = Document()

            docx_config = CONFIG.config['report']['docx']

            # 템플릿 스타일 적용
            if TEMPLATES is not None:
                docx_styles = TEMPLATES.get_docx_styles()
                _apply_docx_styles(doc, docx_styles)
                _add_cover_page(doc, docx_styles, CONFIG.config['project'])
                _add_header_footer(doc, docx_styles, CONFIG.config['project']['name'])
            else:
                title = doc.add_heading(docx_config['title'], level=1)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER

            lines = content.split('\n')
            i = 0

            while i < len(lines):
                line = lines[i].strip()

                if not line:
                    i += 1
                    continue

                if line.startswith('|'):
                    table_lines = []
                    while i < len(lines) and lines[i].strip().startswith('|'):
                        table_lines.append(lines[i].strip())
                        i += 1

                    if len(table_lines) > 1:
                        try:
                            header_cells = [cell.strip() for cell in table_lines[0].split('|')[1:-1]]
                            num_cols = len(header_cells)

                            data_lines = [table_lines[0]]
                            for tl in table_lines[1:]:
                                if not tl.replace('|', '').replace('-', '').replace(' ', '').replace(':', ''):
                                    continue
                                data_lines.append(tl)

                            if len(data_lines) > 0:
                                table = doc.add_table(rows=len(data_lines), cols=num_cols)
                                table.style = 'Table Grid'

                                for row_idx, data_line in enumerate(data_lines):
                                    cells = [cell.strip() for cell in data_line.split('|')[1:-1]]
                                    for col_idx, cell_text in enumerate(cells):
                                        if col_idx < num_cols:
                                            cell = table.rows[row_idx].cells[col_idx]
                                            cell.text = cell_text

                                # 템플릿 스타일 적용
                                if TEMPLATES is not None:
                                    _style_table(table, TEMPLATES.get_docx_styles())
                        except Exception as e:
                            p = doc.add_paragraph('\n'.join(table_lines))
                            p.paragraph_format.line_spacing = 1.0

                    continue

                if line.startswith('#'):
                    level = min(line.count('#', 0, 4), 3)
                    text = line.lstrip('#').strip()
                    if text:
                        doc.add_heading(text, level=level)
                    i += 1
                    continue

                if line[0:3].replace('.', '').replace(')', '').replace('(', '').replace('가', '').replace('나', '').strip():
                    first_chars = line.split()[0] if line.split() else ""
                    if any(c in first_chars for c in ['.', ')', '(']):
                        doc.add_heading(line, level=2)
                        i += 1
                        continue

                para_lines = [line]
                i += 1
                while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith('#') and not lines[i].strip().startswith('|'):
                    para_lines.append(lines[i].strip())
                    i += 1

                para_text = ' '.join(para_lines)
                if para_text:
                    p = doc.add_paragraph(para_text)
                    p.paragraph_format.line_spacing = docx_config['line_spacing']

                    for run in p.runs:
                        run.font.name = docx_config['font_name']
                        run.font.size = Pt(docx_config['font_size'])

            if isopleth_images:
                doc.add_page_break()
                doc.add_heading('등농도곡선', level=1)
                for img_path in isopleth_images:
                    if os.path.exists(img_path):
                        try:
                            img_filename = os.path.basename(img_path)
                            doc.add_paragraph(f"[그림] {img_filename.replace('.png', '')}")
                            doc.add_picture(img_path, width=Inches(docx_config['image_width']))
                            last_paragraph = doc.paragraphs[-1]
                            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            doc.add_paragraph()
                        except Exception as e:
                            print(f"⚠️ 이미지 삽입 실패 ({img_filename}): {e}")

            max_retries = 3
            for attempt in range(max_retries):
                try:
                    doc.save(report_file)
                    print(f"\n💾 보고서 저장 완료: {report_file}")
                    print(f"   파일명: {os.path.basename(report_file)}")
                    print("   ℹ️ DOCX 파일은 한글(HWP)에서 열어서 HWP 형식으로 저장할 수 있습니다.")
                    return report_file
                except PermissionError as pe:
                    if attempt < max_retries - 1:
                        print(f"\n⚠️ 파일 저장 실패 (시도 {attempt+1}/{max_retries})")
                        print("   파일이 다른 프로그램에서 열려있을 수 있습니다.")
                        timestamp = time.strftime("%Y%m%d_%H%M%S") + f"_{attempt+1}"
                        report_file = os.path.join(CONFIG.OUTPUT_DIR, f"대기질_환경영향평가서_{timestamp}.docx")
                        print(f"   새 파일명으로 재시도: {os.path.basename(report_file)}")
                        time.sleep(1)
                    else:
                        raise pe

        except PermissionError as pe:
            print(f"\n❌ DOCX 파일 저장 실패 - 권한 오류")
            print("   해결 방법:")
            print("   1. 기존 DOCX 파일이 Word, 한글 등에서 열려있다면 닫아주세요")
            print("   2. output 폴더의 DOCX 파일을 삭제하거나 이름을 변경하세요")
            print("   3. 다시 시도해주세요")
            return None

        except Exception as e:
            print(f"\n⚠️ DOCX 생성 실패: {e}")
            import traceback
            traceback.print_exc()
            print("   TXT 파일로 대체 저장합니다.")
            return save_report(content, 'txt', isopleth_images)

    elif format == 'hwp':
        print("\n📝 HWP 형식으로 저장하기 위해 먼저 DOCX를 생성합니다.")
        result = save_report(content, 'docx', isopleth_images)
        print("\n   📌 HWP 파일이 필요하면:")
        print("      1. 생성된 DOCX 파일을 한글(HWP)에서 열기")
        print("      2. '파일 > 다른 이름으로 저장 > HWP 형식' 선택")
        return result


# ==========================================
# 5. 새 프로젝트 초기화 함수
# ==========================================
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

    # 하위 폴더 생성
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

# 한글 폰트 설정
font:
  family: "Malgun Gothic"
  unicode_minus: false
"""

    config_path = os.path.join(project_dir, "config.yaml")
    with open(config_path, 'w', encoding='utf-8') as f:
        f.write(config_template)
    print(f"  ✓ config.yaml 생성")

    # README.md 생성
    readme_content = f"""# {project_name}

환경영향평가 대기질 보고서 자동화 시스템

## 프로젝트 구조

```
{project_name}/
├── config.yaml              # 프로젝트 설정 파일
├── t2_v2.py                # 메인 스크립트
├── input/                  # AERMOD 입력 파일 폴더
│   ├── project.inp        # AERMOD 입력 파일
│   └── project.out        # AERMOD 출력 파일 (자동 생성)
├── output/                # 결과 파일 폴더
│   ├── 대기질_환경영향평가서.txt
│   ├── 대기질_환경영향평가서.docx
│   └── 등농도곡선_*.png
├── sample/                # 샘플 보고서 폴더 (PDF)
├── data/                  # 프로젝트 데이터 폴더
│   ├── base.dxf          # 베이스맵 DXF 파일
│   ├── 정온시설.xlsx      # 정온시설 위치 정보
│   ├── 현황측정자료.xlsx   # 현황 측정 자료
│   └── references/       # 참조 논문/보고서 폴더
└── engine/               # AERMOD 엔진 폴더
    └── aermod.exe       # AERMOD 실행 파일
```

## 시작하기

### 1. 필요한 파일 준비

1. **AERMOD 입력 파일**: `input/project.inp`
2. **AERMOD 실행 파일**: `engine/aermod.exe`
3. **베이스맵**: `data/base.dxf`
4. **정온시설 정보**: `data/정온시설.xlsx`
5. **현황측정자료**: `data/현황측정자료.xlsx`
6. **샘플 보고서**: `sample/*.pdf` (선택사항)
7. **참조 문헌**: `data/references/*.pdf` (선택사항)

### 2. 설정 파일 수정

`config.yaml` 파일을 열어 다음 항목을 수정하세요:

- **ai.api_key**: Google Gemini API 키 입력
- **pollutants**: 분석할 오염물질 추가/수정
- **data_files**: 데이터 파일명이 다른 경우 수정
- **isopleth**: 등농도곡선 시각화 설정 조정

### 3. 스크립트 실행

```bash
python t2_v2.py
```

### 4. 명령어

1. **모델링**: AERMOD 시뮬레이션 실행
2. **분석**: 결과 분석 및 보고서 작성 (TXT)
3. **전체**: 모델링 + 분석 한번에 실행 (TXT)
4. **DOCX**: 보고서를 DOCX 형식으로 작성
5. **등농도선**: 등농도곡선만 생성
6. **종료**: 프로그램 종료

## 사용 팁

- 샘플 보고서(PDF)를 sample 폴더에 추가하면 AI가 양식을 학습합니다
- references 폴더에 관련 논문/보고서를 추가하면 더 전문적인 보고서가 생성됩니다
- config.yaml에서 등농도곡선 색상, 크기 등을 자유롭게 조정할 수 있습니다
- 정온시설 엑셀 파일 구조가 다른 경우 config.yaml의 facilities_structure를 수정하세요

## 문의

프로젝트 생성일: {time.strftime('%Y-%m-%d')}
"""

    readme_path = os.path.join(project_dir, "README.md")
    with open(readme_path, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print(f"  ✓ README.md 생성")

    print(f"\n✅ 프로젝트 초기화 완료!")
    print(f"\n📌 다음 단계:")
    print(f"   1. {config_path} 파일을 열어 API 키 등을 설정하세요")
    print(f"   2. 필요한 파일들(inp, dxf, xlsx 등)을 해당 폴더에 추가하세요")
    print(f"   3. AERMOD 실행 파일을 engine 폴더에 복사하세요")
    print(f"   4. python t2_v2.py로 스크립트를 실행하세요")


# ==========================================
# 6. 메인 실행
# ==========================================
if __name__ == "__main__":
    print("\n" + "="*70)
    print("   환경영향평가 대기질 보고서 자동 작성 시스템 v2.1 (템플릿 지원)")
    print("="*70)
    print(f"   프로젝트: {CONFIG.config['project']['name']}")
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

        command = input("\n명령을 입력하세요: ").strip()

        if command == "종료" or command == "6":
            print("\n👋 프로그램을 종료합니다.")
            break

        elif command == "모델링" or command == "1":
            result = run_aermod()
            print(result)

        elif command == "분석" or command == "2":
            try:
                result_data = analyze_output()
                if result_data:
                    report, isopleth_images = result_data
                    print("\n" + "="*70)
                    print("[AI 작성 환경영향평가서 - 대기질]")
                    print("="*70)
                    if len(report) > 1000:
                        print(report[:1000])
                        print(f"\n... (총 {len(report):,}자, 이하 생략) ...\n")
                    else:
                        print(report)

                    result = save_report(report, format='txt', isopleth_images=isopleth_images)
                    if result:
                        print(f"✓ 보고서 작성 완료")
                else:
                    print("❌ 보고서 내용이 생성되지 않았습니다.")
            except Exception as e:
                print(f"\n❌ 오류 발생: {str(e)}")
                import traceback
                traceback.print_exc()

        elif command == "전체" or command == "3":
            result = run_aermod()
            print(result)

            if "SUCCESS" in result or "WARNING" in result:
                try:
                    result_data = analyze_output()
                    if result_data:
                        report, isopleth_images = result_data
                        print("\n" + "="*70)
                        print("[AI 작성 환경영향평가서 - 대기질]")
                        print("="*70)
                        if len(report) > 1000:
                            print(report[:1000])
                            print(f"\n... (총 {len(report):,}자, 이하 생략) ...\n")
                        else:
                            print(report)

                        result = save_report(report, format='txt', isopleth_images=isopleth_images)
                        if result:
                            print(f"✓ 보고서 작성 완료")
                    else:
                        print("❌ 보고서 내용이 생성되지 않았습니다.")
                except Exception as e:
                    print(f"\n❌ 보고서 작성 오류: {str(e)}")
                    import traceback
                    traceback.print_exc()
            else:
                print("\n❌ 모델링 실패로 분석을 진행할 수 없습니다.")

        elif command == "DOCX" or command == "docx" or command == "4":
            try:
                if os.path.exists(CONFIG.OUTPUT_FILE):
                    result_data = analyze_output()
                    if result_data:
                        report, isopleth_images = result_data
                        print("\n" + "="*70)
                        print("[AI 작성 환경영향평가서 - 대기질]")
                        print("="*70)
                        print(f"보고서 길이: {len(report):,}자")
                        print(report[:500] + "..." if len(report) > 500 else report)

                        result = save_report(report, format='docx', isopleth_images=isopleth_images)
                        if result:
                            print(f"✓ DOCX 파일 생성 완료")
                    else:
                        print("❌ 보고서 내용이 생성되지 않았습니다.")
                else:
                    print("\n❌ AERMOD 출력 파일이 없습니다. 먼저 '모델링' 또는 '전체'를 실행하세요.")
            except Exception as e:
                print(f"\n❌ DOCX 생성 오류: {str(e)}")
                import traceback
                traceback.print_exc()

        elif command == "등농도선" or command == "5":
            try:
                if os.path.exists(CONFIG.OUTPUT_FILE):
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

                    images = generate_all_isopleths(level_min=level_min, level_max=level_max,
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
                import traceback
                traceback.print_exc()

        elif command == "템플릿" or command == "7":
            try:
                if TEMPLATES is None:
                    print("\n❌ 템플릿이 로딩되지 않았습니다. templates 폴더를 확인하세요.")
                    continue
                if os.path.exists(CONFIG.OUTPUT_FILE):
                    result_data = analyze_output(use_templates=True)
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
                        save_report(report, format='txt', isopleth_images=isopleth_images)
                        # DOCX 저장
                        result = save_report(report, format='docx', isopleth_images=isopleth_images)
                        if result:
                            print(f"✓ 템플릿 기반 DOCX 파일 생성 완료")
                    else:
                        print("❌ 보고서 내용이 생성되지 않았습니다.")
                else:
                    print("\n❌ AERMOD 출력 파일이 없습니다. 먼저 '모델링' 또는 '전체'를 실행하세요.")
            except Exception as e:
                print(f"\n❌ 템플릿 보고서 생성 오류: {str(e)}")
                import traceback
                traceback.print_exc()

        elif command == "템플릿TXT" or command == "8":
            try:
                if TEMPLATES is None:
                    print("\n❌ 템플릿이 로딩되지 않았습니다. templates 폴더를 확인하세요.")
                    continue
                if os.path.exists(CONFIG.OUTPUT_FILE):
                    result_data = analyze_output(use_templates=True)
                    if result_data:
                        report, isopleth_images = result_data
                        print("\n" + "="*70)
                        print("[템플릿 기반 환경영향평가서 - 대기질]")
                        print("="*70)
                        if len(report) > 1000:
                            print(report[:1000])
                            print(f"\n... (총 {len(report):,}자, 이하 생략) ...\n")
                        else:
                            print(report)

                        result = save_report(report, format='txt', isopleth_images=isopleth_images)
                        if result:
                            print(f"✓ 템플릿 기반 TXT 보고서 작성 완료")
                    else:
                        print("❌ 보고서 내용이 생성되지 않았습니다.")
                else:
                    print("\n❌ AERMOD 출력 파일이 없습니다. 먼저 '모델링' 또는 '전체'를 실행하세요.")
            except Exception as e:
                print(f"\n❌ 템플릿 보고서 생성 오류: {str(e)}")
                import traceback
                traceback.print_exc()

        else:
            print("⚠️ 올바른 명령을 입력하세요.")
