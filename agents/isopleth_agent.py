"""
IsoplethAgent: 등농도곡선 생성

역할:
- AERMOD 출력 파일에서 농도 격자 데이터 파싱
- DXF 베이스맵 로드
- 정온시설 위치 로드
- 등농도곡선 시각화 생성
- 모든 오염물질의 등농도곡선 일괄 생성
"""

import os
import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


def parse_aermod_concentrations(output_file):
    """AERMOD 출력 파일에서 농도 격자 데이터 추출

    Args:
        output_file: AERMOD 출력 파일 경로
    Returns:
        dict or None: {'x': array, 'y': array, 'conc': array}
    """
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


def load_base_dxf(config):
    """DXF 베이스 맵 로드

    Args:
        config: ProjectConfig 인스턴스
    Returns:
        tuple or None: (doc, modelspace)
    """
    try:
        data_files = config.config['data_files']
        dxf_file = os.path.join(config.DATA_DIR, data_files['base_map'])

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


def load_facility_locations(config):
    """정온시설 위치 데이터 로드

    Args:
        config: ProjectConfig 인스턴스
    Returns:
        list or None: [{'name': str, 'x': float, 'y': float}, ...]
    """
    try:
        data_files = config.config['data_files']
        facility_file = os.path.join(config.DATA_DIR, data_files['facilities'])

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


def create_isopleth(data, output_image_path, config,
                    pollutant_name="오염물질", unit="μg/m³",
                    level_min=None, level_max=None, level_count=10, level_interval=None):
    """등농도곡선 생성

    Args:
        data: {'x': array, 'y': array, 'conc': array}
        output_image_path: 출력 이미지 경로
        config: ProjectConfig 인스턴스
        pollutant_name: 오염물질명
        unit: 농도 단위
        level_min/max/count/interval: 등농도선 레벨 설정
    Returns:
        str or None: 생성된 이미지 경로
    """
    try:
        x = data['x']
        y = data['y']
        conc = data['conc']

        if len(x) < 4:
            print(f"⚠️ 데이터 포인트가 너무 적습니다 ({len(x)}개). 최소 4개 필요.")
            return None

        dxf_data = load_base_dxf(config)
        facilities = load_facility_locations(config)

        # 설정에서 격자 해상도 가져오기
        isopleth_config = config.config['isopleth']
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


def generate_all_isopleths(config, level_min=None, level_max=None,
                           level_count=10, level_interval=None):
    """AERMOD 결과로부터 모든 오염물질의 등농도곡선 생성

    Args:
        config: ProjectConfig 인스턴스
        level_min/max/count/interval: 등농도선 레벨 설정
    Returns:
        list: 생성된 이미지 경로 목록
    """
    if not os.path.exists(config.OUTPUT_FILE):
        print("❌ AERMOD 출력 파일을 찾을 수 없습니다.")
        return []

    print("\n🎨 등농도곡선 생성 중...")

    pollutant_name = "PM2.5"
    try:
        with open(config.OUTPUT_FILE, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read(5000)
            match = re.search(r'POLLUTID\s+(\S+)', content)
            if match:
                pollutant_name = match.group(1)
                print(f"✓ 오염물질: {pollutant_name}")
    except:
        pass

    data = parse_aermod_concentrations(config.OUTPUT_FILE)
    if data is None:
        return []

    if not os.path.exists(config.OUTPUT_DIR):
        os.makedirs(config.OUTPUT_DIR)

    image_paths = []

    # 설정에서 오염물질 정보 가져오기
    pollutant_config = config.get_pollutant_config(pollutant_name)
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

    image_path = os.path.join(config.OUTPUT_DIR, f"등농도곡선_{pollutant_name}.png")
    result = create_isopleth(data, image_path, config,
                             pollutant_name=display_name, unit=unit,
                             level_min=level_min, level_max=level_max,
                             level_count=level_count, level_interval=level_interval)
    if result:
        image_paths.append(result)

    return image_paths
