import os
import time
import subprocess
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from google.genai import Client

# 한글 폰트 설정
matplotlib.rcParams['font.family'] = 'Malgun Gothic'
matplotlib.rcParams['axes.unicode_minus'] = False

try:
    from PyPDF2 import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    Document = None

# ==========================================
# 환경 설정
# ==========================================
MY_API_KEY = "YOUR_API_KEY_HERE"
client = Client(api_key=MY_API_KEY)

BASE_DIR = r"G:\eia"
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SAMPLE_DIR = os.path.join(BASE_DIR, "sample")
DATA_DIR = os.path.join(BASE_DIR, "data")
ENGINE_DIR = os.path.join(BASE_DIR, "engine")

AERMOD_EXE = os.path.join(ENGINE_DIR, "aermod.exe")
INPUT_FILE = os.path.join(INPUT_DIR, "project.inp")
OUTPUT_FILE = os.path.join(INPUT_DIR, "project.out")


# ==========================================
# 1. 참조 자료 로딩 함수
# ==========================================
def load_sample_reports():
    """sample 폴더의 PDF 보고서를 텍스트로 추출"""
    samples = []

    if not os.path.exists(SAMPLE_DIR):
        print(f"⚠️ 샘플 폴더를 찾을 수 없습니다: {SAMPLE_DIR}")
        return samples

    print(f"\n📂 샘플 보고서 로딩 중...")
    for filename in os.listdir(SAMPLE_DIR):
        if filename.endswith('.pdf'):
            filepath = os.path.join(SAMPLE_DIR, filename)
            try:
                if PdfReader is None:
                    print(f"  ⚠️ PyPDF2가 설치되지 않아 {filename}을 읽을 수 없습니다.")
                    continue

                reader = PdfReader(filepath)
                text = ""
                # 최대 50페이지까지만 읽기 (토큰 제한)
                max_pages = min(50, len(reader.pages))
                for page_num in range(max_pages):
                    text += reader.pages[page_num].extract_text()

                samples.append({
                    'filename': filename,
                    'content': text[:30000]  # 최대 30,000자
                })
                print(f"  ✓ {filename} 로딩 완료 ({len(reader.pages)} 페이지)")
            except Exception as e:
                print(f"  ✗ {filename} 로딩 실패: {e}")

    return samples


def load_facility_data():
    """data 폴더의 정온시설 위치 정보 로딩"""
    facility_file = os.path.join(DATA_DIR, "정온시설.xlsx")

    if not os.path.exists(facility_file):
        print(f"⚠️ 정온시설 파일을 찾을 수 없습니다: {facility_file}")
        return None

    try:
        df = pd.read_excel(facility_file)
        # 헤더 행 제외 (보통 처음 2행)
        if len(df) > 2:
            df = df.iloc[2:]  # 3번째 행부터 데이터

        # 실제 데이터 개수 확인
        actual_count = len(df)
        print(f"✓ 정온시설 데이터 로딩 완료 ({actual_count}개 시설)")

        # 전체 데이터를 문자열로 변환 (max_colwidth 설정으로 전체 출력)
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
    measurement_file = os.path.join(DATA_DIR, "현황측정자료.xlsx")

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
    references_dir = os.path.join(DATA_DIR, "references")
    references = []

    if not os.path.exists(references_dir):
        print(f"ℹ️ references 폴더가 없습니다. 외부 전문자료 없이 진행합니다.")
        return None

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
                            'content': content[:5000]  # 최대 5,000자
                        })
                        print(f"  ✓ {filename} 로딩 완료")
                elif filename.endswith('.pdf') and PdfReader:
                    reader = PdfReader(filepath)
                    text = ""
                    max_pages = min(10, len(reader.pages))
                    for page_num in range(max_pages):
                        text += reader.pages[page_num].extract_text()
                    references.append({
                        'filename': filename,
                        'content': text[:5000]
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
        # 먼저 .dat 파일 찾기 (전체 격자 데이터)
        input_dir = os.path.dirname(output_file)
        dat_files = [f for f in os.listdir(input_dir) if f.endswith('.dat')]

        if dat_files:
            # .dat 파일에서 전체 격자 데이터 읽기
            dat_file = os.path.join(input_dir, dat_files[0])
            print(f"✓ 격자 데이터 파일 발견: {dat_files[0]}")

            concentrations = []
            x_coords = []
            y_coords = []

            with open(dat_file, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()

            # 헤더 스킵하고 데이터 읽기
            data_started = False
            for line in lines:
                # 데이터 시작 확인 (밑줄 라인 이후)
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

                            # 유효한 농도만 포함 (0이 아닌 값)
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

        # .dat 파일이 없으면 .out 파일에서 요약 데이터 사용
        print("⚠️ .dat 파일을 찾을 수 없습니다. .out 파일에서 요약 데이터를 사용합니다.")

        with open(output_file, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()

        concentrations = []
        x_coords = []
        y_coords = []

        # AERMOD 출력 형식: "VALUE IS 43.35212 AT (201895.00, 544385.00, ...)"
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
        dxf_file = os.path.join(DATA_DIR, "base.dxf")
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
        facility_file = os.path.join(DATA_DIR, "정온시설.xlsx")
        if not os.path.exists(facility_file):
            print(f"⚠️ 정온시설 파일을 찾을 수 없습니다.")
            return None

        df = pd.read_excel(facility_file)

        # 헤더가 있는 행 찾기 (x, y 컬럼)
        facilities = []
        for idx, row in df.iterrows():
            if idx < 2:  # 헤더 스킵
                continue
            try:
                # x, y 좌표 추출 (컬럼 인덱스로)
                x_val = row.iloc[7]  # 8번째 컬럼 (x)
                y_val = row.iloc[8]  # 9번째 컬럼 (y)
                name = row.iloc[3] if pd.notna(row.iloc[3]) else f"시설{idx-1}"

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
    """등농도곡선 생성

    Args:
        level_interval: 등농도선 간격 (None이면 level_count 사용)
    """
    try:
        x = data['x']
        y = data['y']
        conc = data['conc']

        # 데이터 검증
        if len(x) < 4:
            print(f"⚠️ 데이터 포인트가 너무 적습니다 ({len(x)}개). 최소 4개 필요.")
            return None

        # DXF 베이스맵 로드
        dxf_data = load_base_dxf()

        # 정온시설 위치 로드
        facilities = load_facility_locations()

        # 격자 생성 (더 촘촘하게)
        xi = np.linspace(x.min(), x.max(), 200)
        yi = np.linspace(y.min(), y.max(), 200)
        Xi, Yi = np.meshgrid(xi, yi)

        # 농도 보간 (linear 방식 사용 - 더 안정적)
        from scipy.interpolate import griddata
        try:
            Zi = griddata((x, y), conc, (Xi, Yi), method='linear', fill_value=0)
        except Exception as e:
            print(f"⚠️ linear 보간 실패, nearest 방법 시도 중...")
            Zi = griddata((x, y), conc, (Xi, Yi), method='nearest', fill_value=0)

        # 그래프 생성 (배경 흰색)
        fig, ax = plt.subplots(figsize=(14, 12), facecolor='white')
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
                               color='darkgray', linewidth=0.3, alpha=0.5, zorder=1)
                        entity_count += 1

                    elif entity.dxftype() == 'LWPOLYLINE':
                        points = list(entity.get_points())
                        if len(points) > 1:
                            xs = [p[0] for p in points]
                            ys = [p[1] for p in points]
                            # 닫힌 폴리라인인 경우 마지막 점 추가
                            if entity.closed:
                                xs.append(xs[0])
                                ys.append(ys[0])
                            ax.plot(xs, ys, color='darkgray', linewidth=0.3, alpha=0.5, zorder=1)
                            entity_count += 1

                    elif entity.dxftype() == 'POLYLINE':
                        points = list(entity.points())
                        if len(points) > 1:
                            xs = [p[0] for p in points]
                            ys = [p[1] for p in points]
                            ax.plot(xs, ys, color='darkgray', linewidth=0.3, alpha=0.5, zorder=1)
                            entity_count += 1

                    elif entity.dxftype() == 'CIRCLE':
                        center = entity.dxf.center
                        radius = entity.dxf.radius
                        circle = plt.Circle((center.x, center.y), radius,
                                          color='darkgray', fill=False,
                                          linewidth=0.3, alpha=0.5, zorder=1)
                        ax.add_patch(circle)
                        entity_count += 1

                except Exception as e:
                    continue

            print(f"  ✓ DXF 객체 {entity_count}개 렌더링 완료")

        # 등농도선 레벨 설정
        if level_min is None:
            level_min = max(conc.min(), 0.001)  # 최소값이 너무 작으면 0.001로
        if level_max is None:
            level_max = conc.max()

        # 레벨 생성 (간격 우선, 없으면 개수 사용)
        if level_interval is not None:
            # 간격으로 레벨 생성
            levels = np.arange(level_min, level_max + level_interval, level_interval)
            level_count = len(levels)
            print(f"  ✓ 간격 {level_interval} {unit}로 {level_count}개 레벨 생성")
        else:
            # 개수로 레벨 생성
            levels = np.linspace(level_min, level_max, level_count)

        # 등농도선 그리기 (선만, 색깔 채우기 없음)
        contour = ax.contour(Xi, Yi, Zi, levels=levels, colors='blue',
                            linewidths=1.2, alpha=0.8, zorder=2)

        # 등농도선 레이블 추가
        ax.clabel(contour, inline=True, fontsize=8, fmt=f'%.2f {unit}')

        # 최대 농도 지점 표시
        max_idx = np.argmax(conc)
        ax.plot(x[max_idx], y[max_idx], 'r*', markersize=18,
                label=f'최대농도: {conc[max_idx]:.3f} {unit}', zorder=4)

        # 정온시설 표시
        if facilities:
            for idx, facility in enumerate(facilities, start=1):
                # 정온시설 점 표시
                ax.plot(facility['x'], facility['y'], 'o',
                       color='red', markersize=8, markeredgecolor='black',
                       markeredgewidth=0.5, alpha=0.8, zorder=3)
                # 번호 표시
                ax.text(facility['x'], facility['y'], str(idx),
                       fontsize=6, fontweight='bold', color='white',
                       ha='center', va='center', zorder=4)
            # 범례에 정온시설 추가
            ax.plot([], [], 'o', color='red', markersize=8,
                   markeredgecolor='black', markeredgewidth=0.5,
                   label=f'정온시설 ({len(facilities)}개)')

        # 그래프 설정
        ax.set_xlabel('X 좌표 (m)', fontsize=12, fontweight='bold')
        ax.set_ylabel('Y 좌표 (m)', fontsize=12, fontweight='bold')

        # 제목 생성 (간격 정보 포함)
        if level_interval is not None:
            title_str = f'{pollutant_name} 등농도곡선\n({level_min:.3f} ~ {level_max:.3f} {unit}, 간격 {level_interval} {unit})'
        else:
            title_str = f'{pollutant_name} 등농도곡선\n({level_min:.3f} ~ {level_max:.3f} {unit}, {level_count}단계)'

        ax.set_title(title_str, fontsize=15, fontweight='bold', pad=20)
        ax.legend(loc='upper right', fontsize=10, framealpha=0.9)
        ax.grid(True, alpha=0.2, linestyle='--', color='gray')
        ax.set_aspect('equal')

        # 축 범위 설정 (약간 여유 추가)
        x_margin = (x.max() - x.min()) * 0.05
        y_margin = (y.max() - y.min()) * 0.05
        ax.set_xlim(x.min() - x_margin, x.max() + x_margin)
        ax.set_ylim(y.min() - y_margin, y.max() + y_margin)

        # 저장
        plt.tight_layout()
        plt.savefig(output_image_path, dpi=300, bbox_inches='tight', facecolor='white')
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
    """AERMOD 결과로부터 모든 오염물질의 등농도곡선 생성

    Args:
        level_min: 최소 농도 레벨 (None이면 자동)
        level_max: 최대 농도 레벨 (None이면 자동)
        level_count: 등농도선 개수 (기본 10개)
        level_interval: 등농도선 간격 (None이면 level_count 사용)
    """
    if not os.path.exists(OUTPUT_FILE):
        print("❌ AERMOD 출력 파일을 찾을 수 없습니다.")
        return []

    print("\n🎨 등농도곡선 생성 중...")

    # 오염물질 이름 추출
    pollutant_name = "PM2.5"  # 기본값
    try:
        with open(OUTPUT_FILE, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read(5000)  # 처음 5000자만 읽기
            import re
            match = re.search(r'POLLUTID\s+(\S+)', content)
            if match:
                pollutant_name = match.group(1)
                print(f"✓ 오염물질: {pollutant_name}")
    except:
        pass

    # 농도 데이터 파싱
    data = parse_aermod_concentrations(OUTPUT_FILE)
    if data is None:
        return []

    # 이미지 저장 폴더 생성
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    # 등농도곡선 생성
    image_paths = []

    # 오염물질에 따라 표시 이름 설정
    display_name = pollutant_name
    if pollutant_name == "PM2.5":
        display_name = "PM₂.₅"
    elif pollutant_name == "PM10":
        display_name = "PM₁₀"
    elif pollutant_name == "NO2":
        display_name = "NO₂"
    elif pollutant_name == "SO2":
        display_name = "SO₂"

    # 레벨 범위 자동 설정 (지정되지 않은 경우)
    if level_min is None and level_max is None:
        conc = data['conc']
        # 기준: PM2.5의 경우 환경기준(25 μg/m³, 35 μg/m³) 고려
        if pollutant_name == "PM2.5":
            level_max = max(50.0, conc.max())  # 최대 50 이상
            level_min = 0.1
        else:
            level_max = conc.max()
            level_min = max(0.1, conc.min())

        print(f"✓ 자동 레벨 범위: {level_min:.3f} ~ {level_max:.3f} μg/m³")

    image_path = os.path.join(OUTPUT_DIR, f"등농도곡선_{pollutant_name}.png")
    result = create_isopleth(data, image_path, pollutant_name=display_name, unit="μg/m³",
                            level_min=level_min, level_max=level_max, level_count=level_count,
                            level_interval=level_interval)
    if result:
        image_paths.append(result)

    return image_paths


# ==========================================
# 3. AERMOD 실행 함수
# ==========================================
def run_aermod():
    """AERMOD 모델링 실행"""
    if not os.path.exists(AERMOD_EXE):
        return f"❌ AERMOD 실행 파일을 찾을 수 없습니다: {AERMOD_EXE}"

    if not os.path.exists(INPUT_FILE):
        return f"❌ 입력 파일을 찾을 수 없습니다: {INPUT_FILE}"

    print(f"\n🚀 AERMOD 시뮬레이션 시작...")
    print(f"   입력 파일: {INPUT_FILE}")

    try:
        # 작업 디렉토리를 input 폴더로 변경
        original_dir = os.getcwd()
        os.chdir(INPUT_DIR)

        # AERMOD 실행
        result = subprocess.run(
            [AERMOD_EXE, "project.inp"],
            capture_output=True,
            text=True,
            timeout=300  # 5분 타임아웃
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
        return "❌ AERMOD 실행 시간 초과 (5분)"
    except Exception as e:
        os.chdir(original_dir)
        return f"❌ AERMOD 실행 오류: {str(e)}"


# ==========================================
# 4. 결과 분석 및 보고서 작성
# ==========================================
def analyze_output():
    """AERMOD 결과를 분석하여 환경영향평가서 작성"""

    if not os.path.exists(OUTPUT_FILE):
        return "❌ AERMOD 출력 파일을 찾을 수 없습니다. 먼저 모델링을 실행하세요."

    print(f"\n📊 AERMOD 결과 분석 중...")

    # 등농도곡선 생성
    isopleth_images = generate_all_isopleths()

    # AERMOD 결과 파일 읽기
    with open(OUTPUT_FILE, 'r', encoding='utf-8', errors='ignore') as f:
        aermod_data = f.read()[-20000:]  # 마지막 20,000자

    # 참조 자료 로딩
    print(f"\n📚 참조 자료 로딩 중...")
    sample_reports = load_sample_reports()
    facility_data = load_facility_data()
    measurement_data = load_measurement_data()
    reference_papers = load_reference_papers()

    # 샘플 보고서 텍스트 결합
    sample_text = "\n\n".join([
        f"[샘플 보고서: {item['filename']}]\n{item['content']}"
        for item in sample_reports
    ]) if sample_reports else "샘플 보고서 없음"

    # AI 프롬프트 구성
    prompt = f"""
당신은 환경영향평가 대기질 분야 전문가입니다.
제공된 AERMOD 모델링 결과와 참조자료를 바탕으로 환경영향평가서의 대기질 부분을 작성하세요.

[작성 지침]
1. 샘플 보고서의 양식과 구성을 참조하여 동일한 형식으로 작성
2. 정온시설 위치 정보를 활용하여 영향 예측 대상을 명시 (전체 38개 시설 모두 포함)
3. 현황측정자료를 참조하여 현황과 예측결과를 비교·분석
4. 오염물질별 최대 착지농도와 발생 위치(좌표)를 명확히 제시
5. 환경기준 대비 평가 결과 및 적부 판정
6. 문체는 "~함", "~임"을 사용하고 전문용어를 정확히 사용
7. 보고서로 바로 활용 가능하도록 체계적으로 작성
8. **중요**: 예측 결과 해석 시 다음 사항 반드시 포함:
   - 관련 선행연구나 논문의 연구결과를 인용하여 예측결과의 타당성 검증
   - 예: "이는 Kim et al.(2020)의 연구에서 제시한 도로변 대기질 영향범위와 유사한 결과임"
   - 예: "환경부(2021) 대기질 관리지침에 따르면..."
   - 예: "선행 유사사업(OO 도로건설, 2019)의 예측결과와 비교 시..."
   - 국내외 학술논문, 정부 보고서, 유사 환경영향평가서 등을 적극 인용
   - 인용 형식: 저자(연도) 또는 기관명(연도) 형식 사용

[참조 1: 샘플 보고서 양식]
{sample_text[:15000]}

[참조 2: 정온시설 위치 정보 (총 38개 시설)]
{facility_data[:10000] if facility_data else "정온시설 정보 없음"}

[참조 3: 현황측정자료]
{measurement_data[:5000] if measurement_data else "현황측정자료 없음"}

[참조 4: 외부 전문자료 및 선행연구]
{reference_papers[:5000] if reference_papers else "외부 전문자료가 없습니다. 대신 일반적인 대기질 관련 학술 연구와 정부 보고서를 인용하여 작성하세요. 예: 환경부(2022), 국립환경과학원(2021), Lee et al.(2020) 등"}

[AERMOD 모델링 결과]
{aermod_data}

위 정보를 종합하여 환경영향평가서 대기질 부분을 작성하세요.
**반드시 선행연구와 전문자료를 인용하여 예측결과의 신뢰성을 높이세요.**
"""

    # AI 보고서 생성 (재시도 로직 포함)
    print("\n🤖 AI 보고서 작성 중... (최대 1-2분 소요)")
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config={
                    "max_output_tokens": 16000,  # 토큰 수 증가
                    "temperature": 0.3
                }
            )

            # 응답 검증
            if response and hasattr(response, 'text') and response.text:
                result_text = response.text.strip()
                print(f"✓ AI 응답 생성 완료 ({len(result_text):,}자)")
                return result_text, isopleth_images
            else:
                print("⚠️ AI 응답이 비어있습니다.")
                if attempt < max_retries - 1:
                    print(f"   재시도 중... ({attempt+1}/{max_retries})")
                    time.sleep(10)
                    continue
                else:
                    raise Exception("AI 응답이 생성되지 않았습니다.")

        except Exception as e:
            error_msg = str(e)
            if "429" in error_msg and attempt < max_retries - 1:
                wait_time = 35
                print(f"⏳ API 할당량 초과. {wait_time}초 대기 중... ({attempt+1}/{max_retries})")
                time.sleep(wait_time)
            else:
                print(f"❌ AI 보고서 생성 실패: {error_msg}")
                raise Exception(f"AI 보고서 생성 실패: {error_msg}")


def save_report(content, format='txt', isopleth_images=None):
    """보고서를 파일로 저장"""
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    # 내용 검증
    if not content or not content.strip():
        print("⚠️ 저장할 내용이 비어있습니다.")
        return None

    if format == 'txt':
        report_file = os.path.join(OUTPUT_DIR, "대기질_환경영향평가서.txt")
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

        # 타임스탬프를 추가하여 파일명 충돌 방지
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        report_file = os.path.join(OUTPUT_DIR, f"대기질_환경영향평가서_{timestamp}.docx")

        # 기존 파일이 열려있는지 확인 및 안내
        base_report_file = os.path.join(OUTPUT_DIR, "대기질_환경영향평가서.docx")
        if os.path.exists(base_report_file):
            print(f"\n⚠️ 기존 파일이 있습니다. 새 파일명으로 저장합니다: {os.path.basename(report_file)}")

        try:
            doc = Document()

            # 문서 제목
            title = doc.add_heading('환경영향평가서 - 대기질 부문', level=1)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 내용을 줄 단위로 처리
            lines = content.split('\n')
            i = 0

            while i < len(lines):
                line = lines[i].strip()

                if not line:
                    i += 1
                    continue

                # 표 감지 (| 로 시작하는 라인)
                if line.startswith('|'):
                    table_lines = []
                    # 표 전체 수집
                    while i < len(lines) and lines[i].strip().startswith('|'):
                        table_lines.append(lines[i].strip())
                        i += 1

                    # 표 생성
                    if len(table_lines) > 1:
                        try:
                            # 첫 번째 행으로 열 개수 파악
                            header_cells = [cell.strip() for cell in table_lines[0].split('|')[1:-1]]
                            num_cols = len(header_cells)

                            # 구분선 제외
                            data_lines = [table_lines[0]]  # 헤더
                            for tl in table_lines[1:]:
                                if not tl.replace('|', '').replace('-', '').replace(' ', '').replace(':', ''):
                                    continue  # 구분선 스킵
                                data_lines.append(tl)

                            if len(data_lines) > 0:
                                # 테이블 생성
                                table = doc.add_table(rows=len(data_lines), cols=num_cols)
                                table.style = 'Light Grid Accent 1'

                                # 데이터 채우기
                                for row_idx, data_line in enumerate(data_lines):
                                    cells = [cell.strip() for cell in data_line.split('|')[1:-1]]
                                    for col_idx, cell_text in enumerate(cells):
                                        if col_idx < num_cols:
                                            cell = table.rows[row_idx].cells[col_idx]
                                            cell.text = cell_text
                                            # 첫 행은 굵게
                                            if row_idx == 0:
                                                for paragraph in cell.paragraphs:
                                                    for run in paragraph.runs:
                                                        run.bold = True
                        except Exception as e:
                            # 표 생성 실패 시 일반 텍스트로
                            p = doc.add_paragraph('\n'.join(table_lines))
                            p.paragraph_format.line_spacing = 1.0

                    continue

                # 마크다운 제목 처리
                if line.startswith('#'):
                    level = min(line.count('#', 0, 4), 3)
                    text = line.lstrip('#').strip()
                    if text:
                        doc.add_heading(text, level=level)
                    i += 1
                    continue

                # 숫자 제목 (예: 1., 가., (1) 등)
                if line[0:3].replace('.', '').replace(')', '').replace('(', '').replace('가', '').replace('나', '').strip():
                    first_chars = line.split()[0] if line.split() else ""
                    if any(c in first_chars for c in ['.', ')', '(']):
                        doc.add_heading(line, level=2)
                        i += 1
                        continue

                # 일반 문단 (다음 빈 줄까지 수집)
                para_lines = [line]
                i += 1
                while i < len(lines) and lines[i].strip() and not lines[i].strip().startswith('#') and not lines[i].strip().startswith('|'):
                    para_lines.append(lines[i].strip())
                    i += 1

                para_text = ' '.join(para_lines)
                if para_text:
                    p = doc.add_paragraph(para_text)
                    p.paragraph_format.line_spacing = 1.5

                    # 폰트 설정
                    for run in p.runs:
                        run.font.name = '맑은 고딕'
                        run.font.size = Pt(10)

            # 등농도곡선 이미지 추가
            if isopleth_images:
                doc.add_page_break()
                doc.add_heading('등농도곡선', level=1)
                for img_path in isopleth_images:
                    if os.path.exists(img_path):
                        try:
                            # 이미지 설명 추가
                            img_filename = os.path.basename(img_path)
                            doc.add_paragraph(f"[그림] {img_filename.replace('.png', '')}")
                            # 이미지 삽입 (너비 6인치)
                            doc.add_picture(img_path, width=Inches(6))
                            last_paragraph = doc.paragraphs[-1]
                            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            doc.add_paragraph()  # 간격 추가
                        except Exception as e:
                            print(f"⚠️ 이미지 삽입 실패 ({img_filename}): {e}")

            # 파일 저장 시도 (재시도 로직 추가)
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
                        # 다른 파일명으로 재시도
                        timestamp = time.strftime("%Y%m%d_%H%M%S") + f"_{attempt+1}"
                        report_file = os.path.join(OUTPUT_DIR, f"대기질_환경영향평가서_{timestamp}.docx")
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
            print("\n   TXT 파일은 정상적으로 저장되었습니다.")
            return None

        except Exception as e:
            print(f"\n⚠️ DOCX 생성 실패: {e}")
            import traceback
            traceback.print_exc()
            print("   TXT 파일로 대체 저장합니다.")
            return save_report(content, 'txt', isopleth_images)

    elif format == 'hwp':
        # HWP는 DOCX로 생성 후 안내
        print("\n📝 HWP 형식으로 저장하기 위해 먼저 DOCX를 생성합니다.")
        result = save_report(content, 'docx', isopleth_images)
        print("\n   📌 HWP 파일이 필요하면:")
        print("      1. 생성된 DOCX 파일을 한글(HWP)에서 열기")
        print("      2. '파일 > 다른 이름으로 저장 > HWP 형식' 선택")
        return result


# ==========================================
# 4. 메인 실행
# ==========================================
if __name__ == "__main__":
    print("\n" + "="*70)
    print("   환경영향평가 대기질 보고서 자동 작성 시스템")
    print("="*70)

    while True:
        print("\n[명령어]")
        print("  1. 모델링   - AERMOD 시뮬레이션 실행")
        print("  2. 분석     - 결과 분석 및 보고서 작성 (TXT)")
        print("  3. 전체     - 모델링 + 분석 한번에 실행 (TXT)")
        print("  4. DOCX     - 보고서를 DOCX 형식으로 작성")
        print("  5. 등농도선 - 등농도곡선만 생성")
        print("  6. 종료     - 프로그램 종료")

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
                    # 전체 내용이 너무 길면 앞부분만 출력
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
            # 모델링 실행
            result = run_aermod()
            print(result)

            if "SUCCESS" in result or "WARNING" in result:
                # 분석 및 보고서 작성
                try:
                    result_data = analyze_output()
                    if result_data:
                        report, isopleth_images = result_data
                        print("\n" + "="*70)
                        print("[AI 작성 환경영향평가서 - 대기질]")
                        print("="*70)
                        # 전체 내용이 너무 길면 앞부분만 출력
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
                # 기존 출력 결과가 있는지 확인
                if os.path.exists(OUTPUT_FILE):
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
                if os.path.exists(OUTPUT_FILE):
                    print("\n🎨 등농도곡선 생성을 시작합니다...")
                    print("\n[옵션 설정]")
                    print("  Enter를 누르면 자동으로 설정됩니다.")

                    # 최소 레벨 입력
                    level_min_str = input("  최소 농도 레벨 (μg/m³, 기본값: 자동): ").strip()
                    level_min = float(level_min_str) if level_min_str else None

                    # 최대 레벨 입력
                    level_max_str = input("  최대 농도 레벨 (μg/m³, 기본값: 자동): ").strip()
                    level_max = float(level_max_str) if level_max_str else None

                    # 간격 또는 개수 선택
                    print("\n  등농도선 설정 방법:")
                    print("    1. 간격으로 설정 (예: 5 μg/m³ 간격)")
                    print("    2. 개수로 설정 (예: 10개)")
                    interval_choice = input("  선택 (1 또는 2, 기본값: 2): ").strip()

                    level_interval = None
                    level_count = 10

                    if interval_choice == "1":
                        # 간격으로 설정
                        level_interval_str = input("  등농도선 간격 (μg/m³): ").strip()
                        level_interval = float(level_interval_str) if level_interval_str else None
                    else:
                        # 개수로 설정
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

        else:
            print("⚠️ 올바른 명령을 입력하세요.")