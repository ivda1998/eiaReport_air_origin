"""
환경영향평가 대기질 보고서 자동화 시스템
새 프로젝트 초기화 스크립트

사용법:
    python init_project.py

이 스크립트는 새로운 환경영향평가 프로젝트를 위한 폴더 구조와 설정 파일을 생성합니다.
"""

import os
import time


def init_new_project():
    """새로운 프로젝트 초기화"""
    print("\n" + "="*70)
    print("   환경영향평가 대기질 보고서 자동화 시스템")
    print("   새 프로젝트 초기화")
    print("="*70)

    # 프로젝트 정보 입력
    project_name = input("\n프로젝트 이름을 입력하세요: ").strip()
    if not project_name:
        print("❌ 프로젝트 이름이 필요합니다.")
        return

    project_description = input("프로젝트 설명 (Enter=기본값): ").strip()
    if not project_description:
        project_description = "환경영향평가 대기질 모델링"

    project_dir = input("프로젝트 폴더 경로 (Enter=현재 폴더): ").strip()
    if not project_dir:
        project_dir = os.path.join(os.getcwd(), project_name.replace(" ", "_"))

    # 폴더 확인
    if os.path.exists(project_dir):
        confirm = input(f"\n⚠️ 폴더가 이미 존재합니다: {project_dir}\n계속하시겠습니까? (y/n): ").strip().lower()
        if confirm != 'y':
            print("프로젝트 초기화를 취소합니다.")
            return
    else:
        os.makedirs(project_dir)

    print(f"\n📁 프로젝트 폴더 생성 중: {project_dir}")

    # 하위 폴더 생성
    subdirs = [
        'input',
        'output',
        'sample',
        'data',
        'data/references',
        'engine'
    ]

    for subdir in subdirs:
        subdir_path = os.path.join(project_dir, subdir)
        if not os.path.exists(subdir_path):
            os.makedirs(subdir_path)
            print(f"  ✓ {subdir} 폴더 생성")

    # API 키 입력
    print("\n🔑 Google Gemini API 키를 입력하세요")
    print("   (API 키가 없다면 Enter를 누르고 나중에 config.yaml에서 설정하세요)")
    api_key = input("API 키: ").strip()
    if not api_key:
        api_key = "YOUR_API_KEY_HERE"

    # config.yaml 생성
    config_template = f"""# 환경영향평가 대기질 보고서 자동화 시스템 - 프로젝트 설정 파일
# 이 파일을 수정하여 프로젝트에 맞게 설정을 조정할 수 있습니다.

# 프로젝트 기본 정보
project:
  name: "{project_name}"
  description: "{project_description}"
  date: "{time.strftime('%Y-%m-%d')}"

# 경로 설정 (상대 경로 또는 절대 경로)
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
  timeout: 300  # 초 단위 (5분)

# 데이터 파일 설정
data_files:
  base_map: "base.dxf"  # DXF 베이스맵 파일명
  facilities: "정온시설.xlsx"  # 정온시설 위치 파일
  measurements: "현황측정자료.xlsx"  # 현황 측정 자료 파일

  # 정온시설 엑셀 파일 구조
  facilities_structure:
    header_rows: 2  # 헤더 행 수
    name_column: 3  # 시설명 컬럼 인덱스 (0부터 시작)
    x_column: 7     # X 좌표 컬럼 인덱스
    y_column: 8     # Y 좌표 컬럼 인덱스

# 오염물질 설정
pollutants:
  - name: "PM2.5"
    display_name: "PM₂.₅"
    unit: "μg/m³"
    standard: 25.0  # 환경기준 (24시간 평균)

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
  default_level_min: null  # null이면 자동 설정
  default_level_max: null  # null이면 자동 설정
  default_level_interval: null  # null이면 level_count 사용

  grid_resolution: 200  # 보간 격자 해상도 (높을수록 부드러움)
  interpolation_method: "linear"  # linear, cubic, nearest

  # 시각화 설정
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
  api_key: "{api_key}"
  model: "gemini-2.5-flash"
  max_output_tokens: 16000
  temperature: 0.3
  max_retries: 3
  retry_delay: 35  # 초

# 보고서 설정
report:
  formats:
    - "txt"
    - "docx"

  # 샘플 보고서 로딩 제한
  sample_max_pages: 50
  sample_max_chars: 30000

  # 참조 자료 로딩 제한
  reference_max_pages: 10
  reference_max_chars: 5000

  # AERMOD 결과 로딩 제한
  aermod_result_max_chars: 20000

  # DOCX 설정
  docx:
    title: "환경영향평가서 - 대기질 부문"
    font_name: "맑은 고딕"
    font_size: 10
    line_spacing: 1.5
    image_width: 6.0  # 인치

# 한글 폰트 설정 (matplotlib)
font:
  family: "Malgun Gothic"
  unicode_minus: false
"""

    config_path = os.path.join(project_dir, "config.yaml")
    with open(config_path, 'w', encoding='utf-8') as f:
        f.write(config_template)
    print(f"\n  ✓ config.yaml 생성")

    # README.md 생성
    readme_content = f"""# {project_name}

{project_description}

**프로젝트 생성일**: {time.strftime('%Y-%m-%d')}

## 프로젝트 구조

```
{project_name}/
├── config.yaml              # 프로젝트 설정 파일 ⭐
├── t2_v2.py                # 메인 스크립트 ⭐
├── input/                  # AERMOD 입력 파일 폴더
│   ├── project.inp        # AERMOD 입력 파일 📝
│   └── project.out        # AERMOD 출력 파일 (자동 생성)
├── output/                # 결과 파일 폴더 (자동 생성)
│   ├── 대기질_환경영향평가서.txt
│   ├── 대기질_환경영향평가서.docx
│   └── 등농도곡선_*.png
├── sample/                # 샘플 보고서 폴더 (PDF)
├── data/                  # 프로젝트 데이터 폴더
│   ├── base.dxf          # 베이스맵 DXF 파일 📝
│   ├── 정온시설.xlsx      # 정온시설 위치 정보 📝
│   ├── 현황측정자료.xlsx   # 현황 측정 자료 📝
│   └── references/       # 참조 논문/보고서 폴더 (선택사항)
└── engine/               # AERMOD 엔진 폴더
    └── aermod.exe       # AERMOD 실행 파일 📝
```

**📝 = 사용자가 직접 준비해야 하는 파일**
**⭐ = 시스템 핵심 파일**

## 시작하기

### 1. 필수 라이브러리 설치

```bash
pip install pandas numpy matplotlib scipy ezdxf PyPDF2 python-docx pyyaml google-genai
```

### 2. 필요한 파일 준비

다음 파일들을 해당 폴더에 복사하세요:

#### 필수 파일
- **input/project.inp**: AERMOD 입력 파일
- **engine/aermod.exe**: AERMOD 실행 파일
- **data/base.dxf**: 베이스맵 DXF 파일
- **data/정온시설.xlsx**: 정온시설 위치 정보
- **data/현황측정자료.xlsx**: 현황 측정 자료

#### 선택 파일 (보고서 품질 향상)
- **sample/*.pdf**: 참조할 기존 환경영향평가서
- **data/references/*.pdf**: 관련 논문 및 전문 보고서

### 3. 설정 파일 수정

`config.yaml` 파일을 열어 다음 항목을 확인/수정하세요:

#### 반드시 확인할 항목
- **ai.api_key**: Google Gemini API 키 (https://aistudio.google.com/apikey 에서 발급)
- **paths.base_dir**: 프로젝트 폴더 경로

#### 필요시 수정할 항목
- **pollutants**: 분석할 오염물질 추가/수정
- **data_files**: 데이터 파일명이 다른 경우 수정
- **data_files.facilities_structure**: 정온시설 엑셀 구조가 다른 경우
- **isopleth.visualization**: 등농도곡선 색상, 크기 등 시각화 설정

### 4. 스크립트 실행

```bash
python t2_v2.py
```

### 5. 사용 가능한 명령어

1. **모델링**: AERMOD 시뮬레이션 실행
2. **분석**: 결과 분석 및 보고서 작성 (TXT)
3. **전체**: 모델링 + 분석 한번에 실행 (TXT)
4. **DOCX**: 보고서를 DOCX 형식으로 작성
5. **등농도선**: 등농도곡선만 생성 (레벨 범위/간격 조정 가능)
6. **종료**: 프로그램 종료

## 상세 설정 가이드

### 정온시설 엑셀 파일 구조

정온시설 엑셀 파일의 구조가 다른 경우, `config.yaml`의 `data_files.facilities_structure`를 수정하세요:

```yaml
facilities_structure:
  header_rows: 2      # 헤더 행 수 (데이터가 시작되기 전 행 수)
  name_column: 3      # 시설명이 있는 컬럼 인덱스 (0부터 시작)
  x_column: 7         # X 좌표 컬럼 인덱스
  y_column: 8         # Y 좌표 컬럼 인덱스
```

예: 엑셀 파일에서 시설명이 D열(4번째 컬럼)에 있다면 `name_column: 3`

### 오염물질 추가하기

새로운 오염물질을 추가하려면 `config.yaml`의 `pollutants`에 항목을 추가하세요:

```yaml
pollutants:
  - name: "CO"              # AERMOD에서 사용하는 이름
    display_name: "CO"      # 보고서에 표시될 이름
    unit: "ppm"             # 단위
    standard: 9.0           # 환경기준
```

### 등농도곡선 시각화 커스터마이징

`config.yaml`의 `isopleth.visualization`에서 등농도곡선의 외관을 조정할 수 있습니다:

```yaml
visualization:
  figure_width: 14          # 그림 너비 (인치)
  figure_height: 12         # 그림 높이 (인치)
  dpi: 300                  # 해상도
  contour_color: "blue"     # 등농도선 색상
  contour_linewidth: 1.2    # 등농도선 두께
  facility_color: "red"     # 정온시설 표시 색상
  facility_size: 8          # 정온시설 마커 크기
  label_fontsize: 6         # 정온시설 번호 폰트 크기
  base_map_color: "darkgray"  # 베이스맵 색상
  base_map_linewidth: 0.3   # 베이스맵 선 두께
  base_map_alpha: 0.5       # 베이스맵 투명도 (0~1)
```

## 사용 팁

### 📌 보고서 품질 향상 팁

1. **샘플 보고서 제공**: `sample` 폴더에 기존 환경영향평가서(PDF)를 추가하면 AI가 양식과 작성 스타일을 학습합니다.

2. **참조 문헌 추가**: `data/references` 폴더에 관련 논문, 정부 보고서, 유사 사업 평가서를 추가하면 더 전문적이고 근거 있는 보고서가 생성됩니다.

3. **현황 데이터 충실**: 정온시설 정보와 현황측정자료를 정확하게 제공할수록 보고서의 신뢰도가 높아집니다.

### 🎨 등농도곡선 활용

- **간격 모드**: 5 μg/m³ 간격으로 등농도선 생성
- **개수 모드**: 10개의 등농도선으로 전체 범위 분할
- 명령어 5번을 선택하면 대화형으로 설정 가능

### 💾 파일 형식

- **TXT**: 빠른 확인 및 내용 검토용
- **DOCX**: 편집 가능한 워드 문서 (한글에서 열어 HWP로 변환 가능)
- **PNG**: 고해상도(300 DPI) 등농도곡선 이미지

## 문제 해결

### API 키 오류
```
❌ 설정 초기화 실패
```
→ `config.yaml`의 `ai.api_key`에 유효한 Google Gemini API 키를 입력하세요.

### 파일을 찾을 수 없음
```
❌ AERMOD 출력 파일을 찾을 수 없습니다
```
→ 먼저 "1. 모델링" 명령을 실행하여 AERMOD 시뮬레이션을 수행하세요.

### DXF 파일 로드 실패
```
⚠️ DXF 파일을 찾을 수 없습니다
```
→ `data/base.dxf` 파일이 존재하는지 확인하고, `config.yaml`의 `data_files.base_map`이 올바른지 확인하세요.

### 정온시설이 표시되지 않음
```
⚠️ 정온시설 데이터 로드 실패
```
→ `config.yaml`의 `data_files.facilities_structure`에서 컬럼 인덱스를 확인하세요. 엑셀 파일을 열어 시설명, X좌표, Y좌표가 몇 번째 컬럼에 있는지 확인하세요.

### DOCX 파일 저장 실패
```
❌ DOCX 파일 저장 실패 - 권한 오류
```
→ 기존 DOCX 파일이 Word나 한글에서 열려있다면 닫아주세요.

## 다른 사업에 적용하기

이 시스템을 새로운 환경영향평가 사업에 적용하려면:

1. **새 프로젝트 초기화**
   ```bash
   python init_project.py
   ```

2. **파일 준비**: 새 사업의 AERMOD 입력 파일, DXF, 정온시설 정보 등을 준비

3. **설정 수정**: 새 프로젝트의 `config.yaml`을 사업에 맞게 수정

4. **실행**: `python t2_v2.py`로 스크립트 실행

## 라이선스 및 문의

프로젝트 생성일: {time.strftime('%Y-%m-%d')}
"""

    readme_path = os.path.join(project_dir, "README.md")
    with open(readme_path, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print(f"  ✓ README.md 생성")

    # .gitignore 생성
    gitignore_content = """# 출력 파일
output/
*.out
*.dat

# Python
__pycache__/
*.py[cod]
*.pyo
*.pyd
.Python
*.so

# IDE
.vscode/
.idea/
*.swp
*.swo

# 환경 파일
.env

# OS
.DS_Store
Thumbs.db

# 임시 파일
*.tmp
~*
"""

    gitignore_path = os.path.join(project_dir, ".gitignore")
    with open(gitignore_path, 'w', encoding='utf-8') as f:
        f.write(gitignore_content)
    print(f"  ✓ .gitignore 생성")

    # requirements.txt 생성
    requirements_content = """pandas>=2.0.0
numpy>=1.24.0
matplotlib>=3.7.0
scipy>=1.10.0
ezdxf>=1.0.0
PyPDF2>=3.0.0
python-docx>=0.8.11
PyYAML>=6.0
google-genai>=0.2.0
"""

    requirements_path = os.path.join(project_dir, "requirements.txt")
    with open(requirements_path, 'w', encoding='utf-8') as f:
        f.write(requirements_content)
    print(f"  ✓ requirements.txt 생성")

    # 완료 메시지
    print(f"\n" + "="*70)
    print(f"✅ 프로젝트 초기화 완료!")
    print("="*70)

    print(f"\n📁 프로젝트 폴더: {project_dir}")
    print(f"\n📌 다음 단계:")
    print(f"\n1️⃣ 필수 라이브러리 설치:")
    print(f"   cd {project_dir}")
    print(f"   pip install -r requirements.txt")

    print(f"\n2️⃣ 설정 파일 수정:")
    print(f"   - {os.path.basename(config_path)} 파일을 열어 API 키 등을 설정하세요")

    print(f"\n3️⃣ 필요한 파일 준비:")
    print(f"   - input/project.inp (AERMOD 입력 파일)")
    print(f"   - engine/aermod.exe (AERMOD 실행 파일)")
    print(f"   - data/base.dxf (베이스맵)")
    print(f"   - data/정온시설.xlsx (정온시설 위치)")
    print(f"   - data/현황측정자료.xlsx (현황 측정 자료)")

    print(f"\n4️⃣ 선택 파일 추가 (보고서 품질 향상):")
    print(f"   - sample/*.pdf (샘플 보고서)")
    print(f"   - data/references/*.pdf (참조 논문/보고서)")

    print(f"\n5️⃣ 스크립트 실행:")
    print(f"   python t2_v2.py")

    print(f"\n📖 자세한 사용법은 README.md를 참조하세요.")
    print("="*70 + "\n")


if __name__ == "__main__":
    try:
        init_new_project()
    except KeyboardInterrupt:
        print("\n\n❌ 사용자에 의해 취소되었습니다.")
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
