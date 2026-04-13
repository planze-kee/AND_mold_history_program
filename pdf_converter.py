"""
Word PDF 변환 및 병합 유틸리티

워드 (.docx) 파일을 PDF 로 변환하고, 여러 PDF 파일을 하나로 병합합니다.

사용 환경:
- Windows: Word COM 객체 사용 (Microsoft Word 설치 필요)
- Linux/WSL: LibreOffice headless mode 사용
"""

import os
import sys
import subprocess
import tempfile
from pathlib import Path
from typing import List, Union, Optional

try:
    from pypdf import PdfMerger
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False
    print("경고: pypdf 가 설치되지 않았습니다. PDF 병합 기능을 사용하려면 설치하세요.")
    print("pip install pypdf")


# ============================================================================
# Windows 환경: Word COM 을 사용한 변환
# ============================================================================
def docx_to_pdf_windows(docx_path: Union[str, Path], 
                        pdf_path: Optional[Union[str, Path]] = None,
                        save_as_pdf: bool = False) -> Optional[str]:
    """
    Word COM 을 사용하여 .docx 파일을 .pdf 로 변환합니다.
    
    Args:
        docx_path: 변환할 워드 파일 경로
        pdf_path: 출력할 PDF 파일 경로 (생략 시 원본 파일명과 동일한 경로에 저장)
        save_as_pdf: True 인 경우 PDF 로 저장하고 True, 아니면 다른 형식으로
    
    Returns:
        변환된 PDF 파일 경로, 실패 시 None
    """
    try:
        import comtypes.client
        
        # 경로 정규화
        docx_path = Path(docx_path).resolve()
        if not docx_path.exists():
            print(f"오류: 파일이 존재하지 않습니다: {docx_path}")
            return None
        
        # PDF 출력 경로 설정
        if pdf_path is None:
            pdf_path = docx_path.with_suffix('.pdf')
        else:
            pdf_path = Path(pdf_path).resolve()
        
        # Word 앱 인스턴스 생성
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False  # 워드 창 보이지 않게
        
        try:
            # 문서 열기
            doc = word.Documents.Open(str(docx_path))
            
            # PDF 로 저장 (wdFormatPDF = 17)
            doc.SaveAs(
                FileName=str(pdf_path),
                FileFormat=17  # wdFormatPDF
            )
            
            # 문서 닫기
            doc.Close()
            
            print(f"✓ 변환 완료: {docx_path} → {pdf_path}")
            return str(pdf_path)
            
        finally:
            # Word 앱 종료
            word.Quit()
            
    except ImportError:
        print("에러: comtypes 가 설치되어 있지 않습니다.")
        print("Windows 에서 사용 시: pip install comtypes pywin32")
        return None
    except Exception as e:
        print(f"에러: Word 변환 중 오류 발생 - {e}")
        return None


# ============================================================================
# Linux/WSL 환경: LibreOffice headless mode 사용
# ============================================================================
def docx_to_pdf_libreoffice(docx_path: Union[str, Path],
                            pdf_path: Optional[Union[str, Path]] = None,
                            headless: bool = True) -> Optional[str]:
    """
    LibreOffice headless 모드를 사용하여 .docx 파일을 .pdf 로 변환합니다.
    
    Args:
        docx_path: 변환할 워드 파일 경로
        pdf_path: 출력할 PDF 파일 경로 (생략 시 원본 파일명과 동일한 폴더에 저장)
        headless: True 인 경우 headless 모드 사용 (Linux/WSL)
    
    Returns:
        변환된 PDF 파일 경로, 실패 시 None
    """
    try:
        docx_path = Path(docx_path).resolve()
        if not docx_path.exists():
            print(f"오류: 파일이 존재하지 않습니다: {docx_path}")
            return None
        
        # PDF 출력 경로 설정
        if pdf_path is None:
            pdf_path = docx_path.with_suffix('.pdf')
        else:
            pdf_path = Path(pdf_path).resolve()
        
        # LibreOffice 명령어 구성
        cmd = [
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', str(pdf_path.parent),
            str(docx_path)
        ]
        
        print(f"실행: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        if result.returncode != 0:
            print(f"에러: LibreOffice 변환 실패 - {result.stderr}")
            return None
        
        # 생성된 PDF 파일 확인
        if pdf_path.exists():
            print(f"✓ 변환 완료: {docx_path} → {pdf_path}")
            return str(pdf_path)
        else:
            # 파일명이 다를 수 있으므로 같은 폴더에서 PDF 파일 찾기
            pdf_files = list(pdf_path.parent.glob(Path(docx_path).stem + '*.pdf'))
            if pdf_files:
                new_pdf_path = pdf_files[0]
                print(f"✓ 변환 완료: {docx_path} → {new_pdf_path}")
                return str(new_pdf_path)
            return None
            
    except FileNotFoundError:
        print("에러: LibreOffice 가 설치되어 있지 않습니다.")
        print("WSL 에서 설치: sudo apt update && sudo apt install -y libreoffice")
        return None
    except Exception as e:
        print(f"에러: LibreOffice 변환 중 오류 발생 - {e}")
        return None


# ============================================================================
# 자동 환경 감지 및 변환
# ============================================================================
def docx_to_pdf(docx_path: Union[str, Path],
                pdf_path: Optional[Union[str, Path]] = None) -> Optional[str]:
    """
    자동 환경 감지로 Word → PDF 변환을 수행합니다.
    
    Windows: com 사용
    Linux/WSL: LibreOffice headless mode 사용
    """
    # Windows 환경 확인
    if sys.platform == 'win32':
        try:
            return docx_to_pdf_windows(docx_path, pdf_path)
        except Exception:
            pass
    
    # Linux/WSL 환경
    return docx_to_pdf_libreoffice(docx_path, pdf_path)


# ============================================================================
# 여러 Word 파일을 한 번에 PDF 로 변환
# ============================================================================
def batch_docx_to_pdf(input_dir: Union[str, Path],
                      output_dir: Optional[Union[str, Path]] = None,
                      recursive: bool = False) -> List[str]:
    """
    디렉토리 내의 모든 .docx 파일을 PDF 로 변환합니다.
    
    Args:
        input_dir: 입력 디렉토리 경로
        output_dir: 출력 디렉토리 경로 (생략 시 원본과 동일)
        recursive: 하위 디렉토리 포함 여부
    
    Returns:
        변환된 PDF 파일 경로 목록
    """
    input_dir = Path(input_dir).resolve()
    if output_dir is None:
        output_dir = input_dir
    else:
        output_dir = Path(output_dir).resolve()
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # .docx 파일 찾기
    if recursive:
        docx_files = list(input_dir.rglob('*.docx'))
    else:
        docx_files = list(input_dir.glob('*.docx'))
    
    print(f"변환할 Word 파일 {len(docx_files)} 개 발견")
    
    pdf_paths = []
    for i, docx_file in enumerate(docx_files, 1):
        print(f"[{i}/{len(docx_files)}] 변환 중: {docx_file.name}")
        
        # 출력 경로 설정 (원래 폴더에 PDF 저장)
        pdf_path = output_dir / docx_file.with_suffix('.pdf').name
        
        pdf_path_str = docx_to_pdf(docx_file, pdf_path)
        if pdf_path_str:
            pdf_paths.append(pdf_path_str)
    
    print(f"✓ 총 {len(pdf_paths)} 개 파일 변환 완료")
    return pdf_paths


# ============================================================================
# PDF 병합 (pypdf 사용)
# ============================================================================
def merge_pdfs(pdf_files: List[Union[str, Path]],
               output_path: Union[str, Path],
               start_page: int = 0,
               end_page: Optional[int] = None) -> bool:
    """
    여러 PDF 파일을 하나의 PDF 로 병합합니다.
    
    Args:
        pdf_files: 병합할 PDF 파일 목록
        output_path: 출력할 PDF 파일 경로
        start_page: 각 PDF 의 시작 페이지 (0 기반)
        end_page: 각 PDF 의 종료 페이지 (None 일 경우 마지막 페이지까지)
    
    Returns:
        성공 여부
    """
    if not PYPDF_AVAILABLE:
        print("에러: pypdf 가 설치되어 있지 않습니다. PDF 병합을 사용할 수 없습니다.")
        print("pip install pypdf")
        return False
    
    if not pdf_files:
        print("에러: 병합할 PDF 파일이 없습니다.")
        return False
    
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    merger = PdfMerger()
    
    for i, pdf_file in enumerate(pdf_files, 1):
        pdf_file = Path(pdf_file).resolve()
        
        if not pdf_file.exists():
            print(f"경고: 파일이 존재하지 않음 - {pdf_file.name}")
            continue
        
        print(f"[{i}/{len(pdf_files)}] 추가 중: {pdf_file.name}")
        merger.append(str(pdf_file), pages=None if end_page is None else (start_page, end_page))
    
    try:
        with open(output_path, 'wb') as f:
            merger.write(f)
        print(f"✓ PDF 병합 완료: {output_path}")
        return True
    except Exception as e:
        print(f"에러: PDF 병합 실패 - {e}")
        return False
    finally:
        merger.close()


def merge_pdf_from_docx(docx_files: List[Union[str, Path]],
                        output_path: Union[str, Path]) -> bool:
    """
    여러 Word 파일을 PDF 로 변환 후 하나로 병합합니다.
    
    Args:
        docx_files: 변환할 Word 파일 목록
        output_path: 출력할 최종 PDF 파일 경로
    
    Returns:
        성공 여부
    """
    if not docx_files:
        print("에러: 변환할 Word 파일이 없습니다.")
        return False
    
    # 1. Word 파일을 PDF 로 변환
    pdf_files = []
    temp_pdfs = []
    
    for i, docx_file in enumerate(docx_files, 1):
        print(f"[{i}/{len(docx_files)}] Word → PDF 변환 중: {Path(docx_file).name}")
        pdf_path = docx_to_pdf(docx_file)
        if pdf_path:
            pdf_files.append(pdf_path)
            temp_pdfs.append(pdf_path)
    
    if not pdf_files:
        print("에러: 변환된 PDF 파일이 없습니다.")
        return False
    
    # 2. 변환된 PDF 파일들을 병합
    success = merge_pdfs(pdf_files, output_path)
    
    if not success:
        print("경고: 병합 실패. 개별 PDF 파일들은 남아있습니다.")
    
    return success


# ============================================================================
# 예제 코드
# ============================================================================
if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='Word PDF 변환 및 병합 유틸리티')
    subparsers = parser.add_subparsers(dest='command', help='사용할 명령어')
    
    # 단일 파일 변환
    convert_parser = subparsers.add_parser('convert', help='Word 파일을 PDF 로 변환')
    convert_parser.add_argument('input', help='변환할 Word 파일 경로')
    convert_parser.add_argument('-o', '--output', help='출력 PDF 파일 경로')
    
    # 배치 변환
    batch_parser = subparsers.add_parser('batch', help='디렉토리 내 모든 Word 파일 변환')
    batch_parser.add_argument('input_dir', help='입력 디렉토리')
    batch_parser.add_argument('-o', '--output_dir', help='출력 디렉토리')
    batch_parser.add_argument('-r', '--recursive', action='store_true', help='하위 디렉토리 포함')
    
    # PDF 병합
    merge_parser = subparsers.add_parser('merge', help='PDF 파일들을 병합')
    merge_parser.add_argument('inputs', nargs='+', help='병합할 PDF 파일들')
    merge_parser.add_argument('-o', '--output', required=True, help='출력 PDF 파일 경로')
    
    # Word → PDF 변환 후 병합
    docx_merge_parser = subparsers.add_parser('merge-docx', help='Word 파일들을 PDF 로 변환 후 병합')
    docx_merge_parser.add_argument('inputs', nargs='+', help='변환할 Word 파일들')
    docx_merge_parser.add_argument('-o', '--output', required=True, help='출력 PDF 파일 경로')
    
    args = parser.parse_args()
    
    if args.command == 'convert':
        result = docx_to_pdf(args.input, args.output)
        if result:
            print(f"완료: {result}")
        else:
            sys.exit(1)
    
    elif args.command == 'batch':
        results = batch_docx_to_pdf(args.input_dir, args.output_dir, args.recursive)
        if results:
            print(f"\n총 {len(results)} 개 파일 변환 완료")
        else:
            print("변환된 파일이 없습니다")
            sys.exit(1)
    
    elif args.command == 'merge':
        success = merge_pdfs(args.inputs, args.output)
        if not success:
            sys.exit(1)
    
    elif args.command == 'merge-docx':
        success = merge_pdf_from_docx(args.inputs, args.output)
        if not success:
            sys.exit(1)
    
    else:
        parser.print_help()
        sys.exit(1)
