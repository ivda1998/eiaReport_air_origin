# AI 모델 비교 가이드

## 두 가지 버전 사용 가능

### 🤖 t2.py - Google Gemini 버전
```bash
python t2.py
```
- **AI 모델**: Google Gemini 2.5 Flash
- **API 키**: GOOGLE_API_KEY 환경변수 또는 코드 내 직접 입력
- **출력 파일**: `대기질_환경영향평가서.txt`, `대기질_환경영향평가서_YYYYMMDD_HHMMSS.docx`
- **특징**: 빠른 응답 속도, 할당량 초과 시 35초 대기 후 재시도

### 🧠 t3.py - Claude AI 버전
```bash
python t3.py
```
- **AI 모델**: Claude Sonnet 4 (claude-sonnet-4-20250514)
- **API 키**: ANTHROPIC_API_KEY 환경변수 또는 코드 내 직접 입력
- **출력 파일**: `대기질_환경영향평가서_Claude.txt`, `대기질_환경영향평가서_Claude_YYYYMMDD_HHMMSS.docx`
- **특징**: 높은 품질의 분석, 레이트 리미트 시 60초 대기 후 재시도

## 공통 기능

두 버전 모두 동일한 기능을 제공합니다:

1. **1. 모델링**: AERMOD 시뮬레이션만 실행
2. **2. 분석**: 기존 결과를 분석하여 TXT 보고서 작성
3. **3. 전체**: 모델링 + 분석 한번에 실행 (TXT)
4. **4. DOCX**: DOCX 형식으로 보고서 작성
5. **5. 종료**: 프로그램 종료

## 참조 자료 (공통)

- **샘플 보고서**: `sample/*.pdf` (보고서 양식)
- **정온시설 정보**: `data/정온시설.xlsx` (전체 38개 시설)
- **현황측정자료**: `data/현황측정자료.xlsx`
- **외부 전문자료**: `data/references/` (논문, 보고서 등 - 선택사항)

## API 키 설정 방법

### Google Gemini (t2.py)
```bash
# 환경변수로 설정
export GOOGLE_API_KEY='your-google-api-key'

# 또는 t2.py 파일 내 직접 입력
GOOGLE_API_KEY = "your-google-api-key"
```

### Claude AI (t3.py)
```bash
# 환경변수로 설정
export ANTHROPIC_API_KEY='your-anthropic-api-key'

# 또는 t3.py 파일 내 직접 입력
ANTHROPIC_API_KEY = "your-anthropic-api-key"
```

## 어떤 버전을 사용할까?

### Google Gemini (t2.py)를 선택하는 경우:
- 빠른 응답이 필요할 때
- 여러 번 테스트하고 싶을 때
- Google API 할당량이 충분할 때

### Claude AI (t3.py)를 선택하는 경우:
- 보고서 품질을 최우선으로 할 때
- 복잡한 분석과 해석이 필요할 때
- 더 상세하고 정확한 표현을 원할 때

## 추천 워크플로우

1. **첫 테스트**: t2.py로 빠르게 결과 확인
2. **품질 비교**: 같은 데이터로 t3.py 실행
3. **최종 제출**: 두 결과를 비교하여 더 나은 버전 선택

## 주의사항

- 두 버전을 동시에 실행하면 AERMOD 충돌 가능
- 출력 파일명이 다르므로 결과 비교 가능
- API 키가 없으면 해당 버전 실행 불가
- DOCX 파일 저장 시 기존 파일 닫아두기
