"""
AERMODAgent: AERMOD 대기확산 모델링 실행

역할:
- AERMOD 실행 파일 존재 확인
- 입력 파일 검증
- 모델링 실행 및 결과 반환
"""

import os
import subprocess


def run_aermod(config):
    """AERMOD 모델링 실행

    Args:
        config: ProjectConfig 인스턴스
    Returns:
        str: 실행 결과 메시지
    """
    if not os.path.exists(config.AERMOD_EXE):
        return f"❌ AERMOD 실행 파일을 찾을 수 없습니다: {config.AERMOD_EXE}"

    if not os.path.exists(config.INPUT_FILE):
        return f"❌ 입력 파일을 찾을 수 없습니다: {config.INPUT_FILE}"

    print(f"\n🚀 AERMOD 시뮬레이션 시작...")
    print(f"   입력 파일: {config.INPUT_FILE}")

    try:
        original_dir = os.getcwd()
        os.chdir(config.INPUT_DIR)

        aermod_config = config.config['aermod']
        timeout = aermod_config['timeout']

        result = subprocess.run(
            [config.AERMOD_EXE, aermod_config['input_file']],
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
