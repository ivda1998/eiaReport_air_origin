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
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH
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


# 전역 설정 객체
try:
    CONFIG = ProjectConfig()
    client = Client(api_key=CONFIG.config['ai']['api_key'])
except Exception as e:
    print(f"❌ 설정 초기화 실패: {e}")
    print("   config.yaml 파일을 확인하고 다시 시도하세요.")
    exit(1)


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
def analyze_output():
    """AERMOD 결과를 분석하여 환경영향평가서 작성"""

    if not os.path.exists(CONFIG.OUTPUT_FILE):
        return "❌ AERMOD 출력 파일을 찾을 수 없습니다. 먼저 모델링을 실행하세요.", []

    print(f"\n📊 AERMOD 결과 분석 중...")

    isopleth_images = generate_all_isopleths()

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
                print(f"⏳ API 할당량 초과. {retry_delay}초 대기 중... ({attempt+1}/{max_retries})")
                time.sleep(retry_delay)
            else:
                print(f"❌ AI 보고서 생성 실패: {error_msg}")
                raise Exception(f"AI 보고서 생성 실패: {error_msg}")


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
                                table.style = 'Light Grid Accent 1'

                                for row_idx, data_line in enumerate(data_lines):
                                    cells = [cell.strip() for cell in data_line.split('|')[1:-1]]
                                    for col_idx, cell_text in enumerate(cells):
                                        if col_idx < num_cols:
                                            cell = table.rows[row_idx].cells[col_idx]
                                            cell.text = cell_text
                                            if row_idx == 0:
                                                for paragraph in cell.paragraphs:
                                                    for run in paragraph.runs:
                                                        run.bold = True
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
    print("   환경영향평가 대기질 보고서 자동 작성 시스템 v2.0")
    print("="*70)
    print(f"   프로젝트: {CONFIG.config['project']['name']}")
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

        else:
            print("⚠️ 올바른 명령을 입력하세요.")
