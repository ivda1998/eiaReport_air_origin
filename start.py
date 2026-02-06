import os
import subprocess

# 1. 경로 설정
BASE_DIR = r"G:\eia"
EXE_PATH = os.path.join(BASE_DIR, "engine", "aermod.exe")
INPUT_PATH = os.path.join(BASE_DIR, "input", "project.inp")

def run_modeling():
    # 에어모드 실행 명령
    print("에어모드 시뮬레이션을 시작합니다...")
    os.chdir(os.path.join(BASE_DIR, "input")) # 작업 디렉토리 변경
    result = subprocess.run([EXE_PATH, "project.inp"], capture_output=True)
    return "모델링 완료"

# 이후 제가 이 함수를 호출하여 실제 모델링을 제어하게 됩니다.