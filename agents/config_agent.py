"""
ConfigAgent: 프로젝트 설정 관리 및 YAML 템플릿 관리

역할:
- config.yaml 로딩 및 경로 설정
- matplotlib 한글 폰트 설정
- YAML 템플릿 로딩 및 프롬프트 구성
"""

import os
import yaml
import matplotlib


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
            opening = phrases.get('opening', [])
            if opening:
                lines.append("[표준 서두 문구]")
                for p in opening:
                    lines.append(f"- {p}")

        elif section_num == "2":
            survey = phrases.get('methodology', {}).get('survey', [])
            if survey:
                lines.append("[조사 방법론 표준 문구]")
                for p in survey:
                    lines.append(f"- {p}")
            for key in ["2.0", "2.1", "2.2", "2.3", "2.4", "2.5"]:
                tmpl = section_templates.get(key, {})
                sp = tmpl.get('standard_phrases', [])
                if sp:
                    lines.append(f"\n[{key} 표준 문구]")
                    for p in sp:
                        lines.append(f"- {p}")

        elif section_num == "3":
            modeling = phrases.get('methodology', {}).get('modeling', [])
            if modeling:
                lines.append("[모델링 방법론 표준 문구]")
                for p in modeling:
                    lines.append(f"- {p}")
            for key in ["3.0", "3.1", "3.2", "3.3", "3.4", "3.5", "3.6"]:
                tmpl = section_templates.get(key, {})
                sp = tmpl.get('standard_phrases', [])
                if sp:
                    lines.append(f"\n[{key} 표준 문구]")
                    for p in sp:
                        lines.append(f"- {p}")
                for sub_key, sub_val in tmpl.get('subsections', {}).items():
                    if isinstance(sub_val, dict):
                        sub_sp = sub_val.get('standard_phrases', [])
                        if sub_sp:
                            lines.append(f"\n[{sub_key} 표준 문구]")
                            for p in sub_sp:
                                lines.append(f"- {p}")

        elif section_num == "4":
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
            val = phrases.get('validation', [])
            if val:
                lines.append("[타당성 검증 표준 문구]")
                for p in val:
                    lines.append(f"- {p}")

        elif section_num == "6":
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
        special_rules = ""

        if section_num in ["1"]:
            data_section = f"""[사업 정보]
프로젝트명: {context_data.get('project_name', '[사업자가 직접 작성]')}

⚠️ 중요 규칙:
- 제공된 데이터에 없는 사업 정보(사업자명, 사업위치, 사업목적, 사업규모 등)는 절대 임의로 작성하지 마세요.
- 실제 데이터가 없는 항목은 반드시 "[사업자가 직접 작성]" 또는 "[OO 정보 입력 필요]"로 placeholder를 남기세요.
- 예: 사업위치 → "[사업위치 입력 필요]", 사업규모 → "[사업규모 입력 필요]"
"""

        elif section_num in ["2"]:
            measurement = context_data.get('measurement_data', '')
            if measurement:
                data_section = f"""[현황측정자료]
{measurement[:5000]}

⚠️ 중요 규칙:
- 위 측정자료에 포함된 데이터만 사용하세요.
- 측정자료에 없는 항목(기상현황 등)은 "[측정자료 확인 필요]"로 placeholder를 남기세요.
"""
            else:
                data_section = """[현황측정자료]
현황측정자료가 제공되지 않았습니다.

⚠️ 중요 규칙:
- 현황측정자료가 없으므로 임의의 수치를 생성하지 마세요.
- 모든 수치 항목에 "[현황측정자료 입력 필요]"로 placeholder를 남기세요.
"""

        elif section_num in ["3"]:
            aermod = context_data.get('aermod_data', '')
            facility = context_data.get('facility_data', '')
            data_section = f"""[AERMOD 입력 조건]
{aermod[:5000] if aermod else '[AERMOD 입력조건 확인 필요]'}

[정온시설 데이터]
{facility[:5000] if facility else '[정온시설 정보 입력 필요]'}

⚠️ 중요 규칙:
- 제공된 AERMOD 입력조건과 정온시설 데이터만 사용하세요.
- 수용점 위치도는 별도의 DXF 파일로 생성되므로 "[수용점 배치도 - 별도 첨부]"로 표기하세요.
- 데이터에 없는 배출원 정보, 기상관측소 정보 등은 placeholder로 남기세요.
"""

        elif section_num in ["4"]:
            aermod = context_data.get('aermod_data', '')
            facility = context_data.get('facility_data', '')
            measurement = context_data.get('measurement_data', '')
            data_section = f"""[AERMOD 모델링 결과]
{aermod[:15000] if aermod else '[AERMOD 결과 데이터 없음]'}

[정온시설 데이터]
{facility[:8000] if facility else '[정온시설 정보 없음]'}

[현황측정자료]
{measurement[:3000] if measurement else '[현황측정자료 없음]'}
"""
            special_rules = """
⚠️ 매우 중요한 규칙 - 반드시 준수할 것:

1. **최대착지농도 ≠ 정온시설별 예측농도**: 이 두 개념은 완전히 다릅니다.
   - "최대착지농도"는 AERMOD 결과에서 격자(Grid) 수용점 전체에서 가장 높은 농도입니다.
   - "정온시설별 예측농도"는 각 정온시설 위치(이산 수용점)에서의 예측농도입니다.
   - 최대착지농도를 정온시설에 적용하면 안 됩니다.

2. **AERMOD 결과에서 구분하여 서술하세요**:
   - 4.1절: 오염물질별 최대착지농도 (격자 수용점 기준) → 환경기준 대비 평가
   - 4.2절: 정온시설별 예측농도 (이산 수용점 기준) → 시설별 환경기준 대비 평가
   - 4.3절: 현황농도 + 정온시설별 예측농도 합산 → 누적 영향 평가

3. **데이터에 없는 수치는 절대 임의 생성하지 마세요.**
   - AERMOD 결과에서 읽을 수 없는 수치는 "[AERMOD 결과 확인 필요]"로 표기
   - 등농도 분포도는 별도 생성되므로 "[등농도 분포도 - 별도 첨부]"로 표기
"""

        elif section_num in ["5"]:
            aermod = context_data.get('aermod_data', '')
            measurement = context_data.get('measurement_data', '')
            data_section = f"""[AERMOD 모델링 결과 요약]
{aermod[:8000] if aermod else '[AERMOD 결과 데이터 없음]'}

[현황측정자료]
{measurement[:3000] if measurement else '[현황측정자료 없음]'}

⚠️ 중요 규칙:
- 환경기준 적합성은 AERMOD 결과에 나타난 수치만 사용하세요.
- 데이터에 없는 수치는 "[확인 필요]"로 placeholder를 남기세요.
- 최대착지농도와 정온시설별 예측농도를 혼동하지 마세요.
"""

        elif section_num in ["6"]:
            data_section = """⚠️ 중요 규칙:
- 저감방안은 표준 문구를 활용하되, 구체적인 수치(3m, 3회/일 등)를 포함하세요.
- 사업 유형에 특화된 내용이 필요하나 사업 정보가 부족하면 "[사업유형에 맞는 저감방안 보완 필요]"로 표기하세요.
"""

        elif section_num in ["7"]:
            data_section = f"""[이전 섹션 요약]
{previous_sections_summary[:5000] if previous_sections_summary else '[이전 섹션 정보 없음]'}

⚠️ 중요 규칙:
- 결론은 앞서 작성된 내용의 요약이어야 합니다.
- 새로운 수치나 분석을 추가하지 마세요.
- 이전 섹션에서 placeholder로 남긴 항목은 결론에서도 동일하게 placeholder로 유지하세요.
"""

        elif section_num in ["8"]:
            data_section = """⚠️ 중요 규칙:
- 실제 인용한 문헌만 기재하세요.
- 임의의 참고문헌을 만들어내지 마세요.
- 본문에서 인용하지 않은 문헌은 포함하지 마세요.
- 형식은 APA 스타일을 따르세요.
"""

        # 샘플 보고서 참조 (섹션 2~5에만 - 섹션1은 임의 데이터 방지를 위해 제외)
        sample_ref = ""
        if section_num in ["2", "3", "4", "5"]:
            sample_text = context_data.get('sample_text', '')
            if sample_text:
                sample_ref = f"\n[참조: 샘플 보고서 양식 (구조와 문체만 참조하고, 수치는 반드시 실제 데이터 사용)]\n{sample_text[:5000]}\n"

        prompt = f"""당신은 환경영향평가 대기질 분야 전문가입니다.
아래 지침에 따라 환경영향평가서의 "{toc_text.split(chr(10))[0] if toc_text else section_id}" 섹션을 작성하세요.

{writing_rules}

[핵심 원칙]
- 제공된 실제 데이터만 사용하세요. 임의의 내용이나 수치를 절대 생성하지 마세요.
- 데이터가 없는 항목은 "[OO 입력 필요]" 형태의 placeholder로 남기세요.
- 그림, 위치도 등 별도 첨부물은 "[OO - 별도 첨부]"로 표기하세요.

[섹션 구조]
{toc_text}

{section_rules}

{f'[표준 문구 (적극 활용하세요)]' + chr(10) + standard_phrases if standard_phrases else ''}

{data_section}
{special_rules}
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
            merged['heading_styles'] = self.fonts.get('heading_styles', {})
            merged['body_styles'] = self.fonts.get('body_styles', {})
            merged['table_styles'] = self.fonts.get('table_styles', {})
            merged['figure_styles'] = self.fonts.get('figure_styles', {})
            merged['cover_fonts'] = self.fonts.get('cover', {})
            merged['font_colors'] = self.fonts.get('colors', {})
        return merged
