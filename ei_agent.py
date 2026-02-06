import os
import time
import google.generativeai as genai

# ==========================================
# 1. 환경 설정
# ==========================================
MY_API_KEY = "YOUR_API_KEY_HERE"

# 경로 설정
BASE_DIR = r"G:\eia"
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SAMPLE_DIR = os.path.join(BASE_DIR, "sample")
DATA_DIR = os.path.join(BASE_DIR, "data")

# API 인증
genai.configure(api_key=MY_API_KEY)

def get_working_model():
    """사용 가능한 모델을 순차적으로 시도하여 연결합니다."""
    model_candidates = [
        'gemini-2.5-flash',
        'gemini-2.0-flash',
        'gemini-2.5-pro',
        'gemini-flash-latest'
    ]
    
    for model_name in model_candidates:
        try:
            print(f"모델 연결 시도 중: {model_name}...")
            m = genai.GenerativeModel(model_name)
            m.generate_content("test", generation_config={"max_output_tokens": 1})
            print(f"✅ 성공: {model_name} 모델에 연결되었습니다.")
            return m
        except Exception as e:
            error_msg = str(e).splitlines()[0]
            print(f"❌ 실패: {model_name} ({error_msg})")
            continue
    
    raise Exception("모든 모델 연결에 실패했습니다. API 키나 할당량을 확인하세요.")

# 전역 모델 객체 생성
try:
    model = get_working_model()
except Exception as e:
    print(e)
    exit()

# ==========================================
# 2. 참조 자료 로드 함수
# ==========================================
def load_reference_files():
    """sample 폴더와 data 폴더의 파일들을 읽어옵니다."""
    reference_data = {
        'sample_reports': [],
        'facility_locations': [],
        'measurement_data': []
    }
    
    # 샘플 보고서 로드 (sample 폴더)
    if os.path.exists(SAMPLE_DIR):
        print(f"\n📂 샘플 보고서 로드 중 ({SAMPLE_DIR})...")
        for filename in os.listdir(SAMPLE_DIR):
            if filename.endswith(('.txt', '.hwp', '.pdf', '.docx')):
                try:
                    filepath = os.path.join(SAMPLE_DIR, filename)
                    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                        reference_data['sample_reports'].append({
                            'filename': filename,
                            'content': content[:30000]  # 최대 30,000자
                        })
                        print(f"  ✓ {filename} 로드 완료")
                except Exception as e:
                    print(f"  ✗ {filename} 로드 실패: {e}")
    
    # 데이터 파일 로드 (data 폴더)
    if os.path.exists(DATA_DIR):
        print(f"\n📂 참조 데이터 로드 중 ({DATA_DIR})...")
        for filename in os.listdir(DATA_DIR):
            try:
                filepath = os.path.join(DATA_DIR, filename)
                with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                    
                    # 파일명으로 분류
                    if '정온시설' in filename or 'facility' in filename.lower():
                        reference_data['facility_locations'].append({
                            'filename': filename,
                            'content': content[:20000]
                        })
                        print(f"  ✓ {filename} (정온시설 위치)")
                    elif '측정' in filename or 'measurement' in filename.lower():
                        reference_data['measurement_data'].append({
                            'filename': filename,
                            'content': content[:20000]
                        })
                        print(f"  ✓ {filename} (현황측정자료)")
                    else:
                        reference_data['measurement_data'].append({
                            'filename': filename,
                            'content': content[:20000]
                        })
                        print(f"  ✓ {filename} (기타 참조자료)")
            except Exception as e:
                print(f"  ✗ {filename} 로드 실패: {e}")
    
    return reference_data

# ==========================================
# 3. 분석 함수 (참조자료 활용)
# ==========================================
def analyze_output():
    file_path = os.path.join(INPUT_DIR, "project.out")
    
    if not os.path.exists(file_path):
        return "오류: 'project.out' 파일을 찾을 수 없습니다. 모델링이 성공했는지 확인하세요."
    
    print(f"\n[{file_path}] 결과 분석 중...")

    try:
        # AERMOD 결과 파일 읽기
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        aermod_data = content[-20000:] if len(content) > 20000 else content

        # 참조 자료 로드
        reference_data = load_reference_files()

        # 샘플 보고서 텍스트 결합
        sample_text = "\n\n".join([
            f"[샘플 보고서: {item['filename']}]\n{item['content']}"
            for item in reference_data['sample_reports']
        ])

        # 정온시설 위치 텍스트 결합
        facility_text = "\n\n".join([
            f"[정온시설 위치: {item['filename']}]\n{item['content']}"
            for item in reference_data['facility_locations']
        ])

        # 현황측정자료 텍스트 결합
        measurement_text = "\n\n".join([
            f"[현황측정자료: {item['filename']}]\n{item['content']}"
            for item in reference_data['measurement_data']
        ])

        prompt = f"""
당신은 환경영향평가 대기질 전문가입니다. 
제공된 AERMOD 모델링 결과와 참조자료를 바탕으로 환경영향평가서의 대기질 부분을 작성하세요.

[지시사항]
1. 샘플 보고서의 양식과 구성을 참조하여 동일한 형식으로 작성할 것
2. 정온시설 위치 정보를 활용하여 영향 예측 대상을 명시할 것
3. 현황측정자료를 참조하여 현황과 예측결과를 비교할 것
4. 오염물질별 최대 착지 농도와 발생 위치(좌표)를 명시할 것
5. 환경기준 대비 평가 결과를 명확히 제시할 것
6. 문체는 "~함", "~임" 체를 사용하고 전문용어를 정확히 사용할 것
7. 보고서로 바로 활용 가능하도록 체계적으로 작성할 것

[샘플 보고서 양식 참조]
{sample_text[:15000] if sample_text else "샘플 보고서 없음"}

[정온시설 위치 정보]
{facility_text[:5000] if facility_text else "정온시설 위치 정보 없음"}

[현황측정자료]
{measurement_text[:5000] if measurement_text else "현황측정자료 없음"}

[AERMOD 모델링 결과]
{aermod_data}

위 정보를 종합하여 환경영향평가서 대기질 부분을 작성하세요.
"""
        
        # 재시도 로직 (할당량 초과 대응)
        max_retries = 3
        for attempt in range(max_retries):
            try:
                print("\n🤖 AI 보고서 작성 중... (최대 1-2분 소요)")
                response = model.generate_content(
                    prompt,
                    generation_config={
                        "max_output_tokens": 8000,  # 긴 보고서 생성
                        "temperature": 0.3  # 일관성 있는 출력
                    }
                )
                return response.text
            except Exception as e:
                if "429" in str(e) and attempt < max_retries - 1:
                    wait_time = 35
                    print(f"⏳ 할당량 초과. {wait_time}초 대기 중... ({attempt+1}/{max_retries})")
                    time.sleep(wait_time)
                else:
                    raise

    except Exception as e:
        return f"분석 중 오류 발생: {str(e)}"

# ==========================================
# 4. 메인 실행 루프
# ==========================================
if __name__ == "__main__":
    print("\n" + "="*60)
    print("   AI 환경영향평가 대기질 보고서 작성 에이전트")
    print("="*60)
    
    while True:
        user_input = input("\n명령을 입력하세요 (분석 / 종료): ").strip()
        
        if user_input == "종료":
            print("\n프로그램을 종료합니다.")
            break
        elif user_input == "분석":
            result = analyze_output()
            print("\n" + "="*60)
            print("[AI 작성 환경영향평가서 - 대기질 부분]")
            print("="*60)
            print(result)
            
            # 결과를 output 폴더에 저장
            if not os.path.exists(OUTPUT_DIR): 
                os.makedirs(OUTPUT_DIR)
            report_file = os.path.join(OUTPUT_DIR, "대기질_환경영향평가서.txt")
            with open(report_file, "w", encoding="utf-8") as f:
                f.write(result)
            print(f"\n💾 결과가 {report_file}에 저장되었습니다.")
        else:
            print("⚠️ '분석' 또는 '종료'를 입력하세요.")