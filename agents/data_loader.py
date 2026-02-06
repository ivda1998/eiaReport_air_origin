"""
DataLoaderAgent: 참조 자료 로딩

역할:
- 샘플 보고서 (PDF) 로딩
- 정온시설 위치 정보 (Excel) 로딩
- 현황측정자료 (Excel) 로딩
- 외부 전문자료/논문 로딩
"""

import os
import pandas as pd


def load_sample_reports(config):
    """sample 폴더의 PDF 보고서를 텍스트로 추출

    Args:
        config: ProjectConfig 인스턴스
    Returns:
        list: [{'filename': str, 'content': str}, ...]
    """
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

    samples = []

    if not os.path.exists(config.SAMPLE_DIR):
        print(f"⚠️ 샘플 폴더를 찾을 수 없습니다: {config.SAMPLE_DIR}")
        return samples

    report_config = config.config['report']
    max_pages = report_config['sample_max_pages']
    max_chars = report_config['sample_max_chars']

    print(f"\n📂 샘플 보고서 로딩 중...")
    for filename in os.listdir(config.SAMPLE_DIR):
        if filename.endswith('.pdf'):
            filepath = os.path.join(config.SAMPLE_DIR, filename)
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


def load_facility_data(config):
    """data 폴더의 정온시설 위치 정보 로딩

    Args:
        config: ProjectConfig 인스턴스
    Returns:
        str or None: 정온시설 데이터 문자열
    """
    data_files = config.config['data_files']
    facility_file = os.path.join(config.DATA_DIR, data_files['facilities'])

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


def load_measurement_data(config):
    """data 폴더의 현황측정자료 로딩

    Args:
        config: ProjectConfig 인스턴스
    Returns:
        str or None: 현황측정자료 문자열
    """
    data_files = config.config['data_files']
    measurement_file = os.path.join(config.DATA_DIR, data_files['measurements'])

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


def load_reference_papers(config):
    """외부 전문자료 폴더에서 논문/보고서 로딩

    Args:
        config: ProjectConfig 인스턴스
    Returns:
        str or None: 참고문헌 통합 텍스트
    """
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

    references_dir = os.path.join(config.DATA_DIR, "references")
    references = []

    if not os.path.exists(references_dir):
        print(f"ℹ️ references 폴더가 없습니다. 외부 전문자료 없이 진행합니다.")
        return None

    report_config = config.config['report']
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
