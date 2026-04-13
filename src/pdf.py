# Word PDF 변환 및 병합 모듈 (함수 추가)

"""
Word (.docx) 파일을 PDF 로 변환하고, 여러 PDF 파일을 하나로 병합합니다.
"""

import os
import sys
import subprocess
from pathlib import Path
from typing import List, Optional

try:
    from pypdf import PdfMerger
    PYPDF_AVAILABLE = True
except ImportError:
    try:
        from PyPDF3 import PdfMerger
        PYPDF_AVAILABLE = True
    except ImportError:
        try:
            from pypdf2 import PdfMerger
            PYPDF_AVAILABLE = True
        except ImportError:
            PYPDF_AVAILABLE = False


class DocxToPdfConverter:
    """Word 파일을 PDF 로 변환하고 병합하는 클래스"""
    
    def __init__(self):
        self.pdf_files: List[str] = []
    
    def convert(self, docx_path: Path, output_path: Optional[Path] = None) -> Optional[str]:
        """
        단일 Word 파일을 PDF 로 변환합니다.
        
        Args:
            docx_path: 변환할 Word 파일 경로
            output_path: 출력할 PDF 파일 경로 (생략 시 원본 파일명과 동일)
        
        Returns:
            변환된 PDF 파일 경로, 실패 시 None
        """
        docx_path = Path(docx_path).resolve()
        
        if not docx_path.exists():
            print(f"오류: 파일이 존재하지 않습니다: {docx_path}")
            return None
        
        # 출력 경로 설정
        if output_path is None:
            output_path = docx_path.with_suffix('.pdf')
        else:
            output_path = Path(output_path).resolve()
            output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # 환경에 따라 다른 방식 사용
        if sys.platform == 'win32':
            return self._convert_windows(docx_path, output_path)
        else:
            return self._convert_libreoffice(docx_path, output_path)
    
    def _convert_windows(self, docx_path: Path, pdf_path: Path) -> Optional[str]:
        """Windows 환경: Word COM 사용"""
        import traceback
        try:
            import comtypes.client
            
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # 경고 창 안 보임
            
            try:
                doc = word.Documents.Open(str(docx_path), ReadOnly=True)
                # wdFormatPDF = 17
                doc.SaveAs2(str(pdf_path), FileFormat=17)  # SaveAs2 사용 (새 버전 호환)
                doc.Close()
                print(f"✓ 변환 완료: {docx_path.name} → {pdf_path.name}")
                print(f"  출력 파일: {pdf_path}")
                return str(pdf_path)
            except Exception as convert_error:
                print(f"✗ Word 문서 변환 중 오류 발생:")
                print(f"  파일: {docx_path}")
                print(f"  에러: {convert_error}")
                traceback.print_exc()
                return None
            finally:
                try:
                    word.Quit()
                except:
                    pass
                # Word 프로세스 강제 종료 (잔류 방지)
                import os
                os.system("taskkill /F /IM WINWORD.EXE >nul 2>&1")
                
        except ImportError as e:
            print(f"✗ comtypes 또는 pywin32 설치 필요:")
            print(f"  pip install comtypes pywin32")
            print(f"  세부 에러: {e}")
            return None
        except Exception as e:
            print(f"✗ Word COM 변환 중 전체 오류:")
            print(f"  에러: {e}")
            traceback.print_exc()
            return None
    
    def _convert_libreoffice(self, docx_path: Path, pdf_path: Path) -> Optional[str]:
        """Linux/WSL 환경: LibreOffice headless mode 사용"""
        try:
            cmd = [
                'libreoffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(pdf_path.parent),
                str(docx_path)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            if result.returncode != 0:
                print(f"에러: LibreOffice 변환 실패 - {result.stderr}")
                return None
            
            # 생성된 PDF 파일 확인
            if pdf_path.exists():
                print(f"✓ 변환 완료: {docx_path.name} → {pdf_path.name}")
                return str(pdf_path)
            
            # 다른 이름으로 생성된 경우 찾기
            pdf_files = list(pdf_path.parent.glob(Path(docx_path).stem + '*.pdf'))
            if pdf_files:
                new_path = pdf_files[0]
                print(f"✓ 변환 완료: {docx_path.name} → {new_path.name}")
                return str(new_path)
            
            return None
            
        except FileNotFoundError:
            print("에러: LibreOffice 가 설치되어 있지 않습니다.")
            print("WSL 에서 설치: sudo apt update && sudo apt install -y libreoffice")
            return None
        except Exception as e:
            print(f"에러: LibreOffice 변환 중 오류 - {e}")
            return None
    
    def convert_and_merge(self, docx_files: List[Path], output_path: Path) -> bool:
        """
        여러 Word 파일들을 PDF 로 변환하고 하나로 병합합니다.
        
        Args:
            docx_files: 변환할 Word 파일 목록
            output_path: 출력할 최종 PDF 파일 경로
        
        Returns:
            성공 여부
        """
        if not docx_files:
            print("에러: 변환할 Word 파일이 없습니다.")
            return False
        
        output_path = Path(output_path).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # 1. Word → PDF 변환
        print(f"Word 파일 {len(docx_files)} 개를 PDF 로 변환 중...")
        pdf_paths = []
        
        for i, docx_file in enumerate(docx_files, 1):
            print(f"[{i}/{len(docx_files)}] 변환 중: {docx_file.name}")
            pdf_path = self.convert(docx_file)
            if pdf_path:
                pdf_paths.append(pdf_path)
        
        if not pdf_paths:
            print("에러: 변환된 PDF 파일이 없습니다.")
            return False
        
        # 2. PDF 병합
        print(f"\n{len(pdf_paths)} 개 PDF 파일을 병합 중...")
        return self._merge_pdfs(pdf_paths, output_path)
    
    def _merge_pdfs(self, pdf_files: List[str], output_path: Path) -> bool:
        """
        PDF 파일들을 병합합니다.
        
        Args:
            pdf_files: 병합할 PDF 파일 목록
            output_path: 출력 파일 경로
        
        Returns:
            성공 여부
        """
        if not PYPDF_AVAILABLE:
            print("에러: pypdf 가 설치되어 있지 않습니다.")
            print("pip install pypdf")
            return False
        
        merger = PdfMerger()
        
        try:
            for i, pdf_file in enumerate(pdf_files, 1):
                print(f"[{i}/{len(pdf_files)}] 추가 중: {Path(pdf_file).name}")
                merger.append(pdf_file)
            
            with open(output_path, 'wb') as f:
                merger.write(f)
            
            print(f"✓ PDF 병합 완료: {output_path.name}")
            
            # 파일 정보 출력
            file_size = output_path.stat().st_size
            print(f"  - 크기: {file_size / 1024:.1f} KB")
            
            return True
            
        except Exception as e:
            print(f"에러: PDF 병합 실패 - {e}")
            return False
        finally:
            merger.close()
    
    def merge_only(self, pdf_files: List[Path], output_path: Path) -> bool:
        """
        이미 존재하는 PDF 파일들만 병합합니다.
        
        Args:
            pdf_files: 병합할 PDF 파일 목록
            output_path: 출력 파일 경로
        
        Returns:
            성공 여부
        """
        if not PYPDF_AVAILABLE:
            print("에러: pypdf 가 설치되어 있지 않습니다.")
            return False
        
        if not pdf_files:
            print("에러: 병합할 PDF 파일이 없습니다.")
            return False
        
        output_path = Path(output_path).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        merger = PdfMerger()
        
        try:
            for i, pdf_file in enumerate(pdf_files, 1):
                pdf_file = Path(pdf_file).resolve()
                if not pdf_file.exists():
                    print(f"경고: 파일이 없음 - {pdf_file.name}")
                    continue
                
                print(f"[{i}/{len(pdf_files)}] 추가 중: {pdf_file.name}")
                merger.append(str(pdf_file))
            
            with open(output_path, 'wb') as f:
                merger.write(f)
            
            print(f"✓ PDF 병합 완료: {output_path.name}")
            return True
            
        except Exception as e:
            print(f"에러: PDF 병합 실패 - {e}")
            return False
        finally:
            merger.close()


# ============================================================================
# 추가 함수들
# ============================================================================

def docx_to_pdf(docx_path: Path, output_path: Optional[Path] = None) -> Optional[str]:
    """단일 Word 파일을 PDF 로 변환"""
    converter = DocxToPdfConverter()
    return converter.convert(docx_path, output_path)


def convert_and_merge(docx_files: List[Path], output_path: Path) -> bool:
    """Word 파일들을 PDF 로 변환 후 병합"""
    converter = DocxToPdfConverter()
    return converter.convert_and_merge(docx_files, output_path)


def merge_pdfs(pdf_files: List[Path], output_path: Path) -> bool:
    """PDF 파일들만 병합"""
    converter = DocxToPdfConverter()
    return converter.merge_only(pdf_files, output_path)


def batch_docx_to_pdf(input_dir: Path, output_dir: Optional[Path] = None, recursive: bool = False) -> List[str]:
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
    converter = DocxToPdfConverter()
    
    for i, docx_file in enumerate(docx_files, 1):
        print(f"[{i}/{len(docx_files)}] 변환 중: {docx_file.name}")
        
        # 출력 경로 설정 (원래 폴더에 PDF 저장)
        pdf_path = output_dir / docx_file.with_suffix('.pdf').name
        
        pdf_path_str = converter.convert(docx_file, pdf_path)
        if pdf_path_str:
            pdf_paths.append(pdf_path_str)
    
    print(f"✓ 총 {len(pdf_paths)} 개 파일 변환 완료")
    return pdf_paths
