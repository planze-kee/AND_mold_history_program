# Word PDF 변환 및 병합 모듈

"""
Word (.docx) 파일을 PDF 로 변환하고, 여러 PDF 파일을 하나로 병합합니다.
"""

import os
import sys
import subprocess
from pathlib import Path
from typing import Callable, List, Optional

try:
    from pypdf import PdfWriter, PdfReader
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False


def _log(msg: str, callback: Optional[Callable[[str], None]]) -> None:
    if callback:
        callback(msg)
    else:
        print(msg)


class DocxToPdfConverter:
    """Word 파일을 PDF 로 변환하고 병합하는 클래스"""

    def __init__(self):
        self.pdf_files: List[str] = []

    def convert(self, docx_path: Path, output_path: Optional[Path] = None,
                callback: Optional[Callable[[str], None]] = None) -> Optional[str]:
        docx_path = Path(docx_path).resolve()

        if not docx_path.exists():
            _log(f"오류: 파일이 존재하지 않습니다: {docx_path}", callback)
            return None

        if output_path is None:
            output_path = docx_path.with_suffix('.pdf')
        else:
            output_path = Path(output_path).resolve()
            output_path.parent.mkdir(parents=True, exist_ok=True)

        if sys.platform == 'win32':
            return self._convert_windows(docx_path, output_path, callback)
        else:
            return self._convert_libreoffice(docx_path, output_path, callback)

    def _convert_windows(self, docx_path: Path, pdf_path: Path,
                         callback: Optional[Callable[[str], None]] = None) -> Optional[str]:
        import traceback
        try:
            import comtypes.client

            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            try:
                doc = word.Documents.Open(str(docx_path), ReadOnly=True)
                doc.SaveAs2(str(pdf_path), FileFormat=17)
                doc.Close()
                _log(f"✓ 변환 완료: {docx_path.name}", callback)
                return str(pdf_path)
            except Exception as convert_error:
                _log(f"✗ 변환 오류: {docx_path.name} — {convert_error}", callback)
                return None
            finally:
                try:
                    word.Quit()
                except Exception:
                    pass
                os.system("taskkill /F /IM WINWORD.EXE >nul 2>&1")

        except ImportError as e:
            _log(f"✗ comtypes/pywin32 설치 필요: {e}", callback)
            return None
        except Exception as e:
            _log(f"✗ Word COM 오류: {e}", callback)
            return None

    def _convert_libreoffice(self, docx_path: Path, pdf_path: Path,
                             callback: Optional[Callable[[str], None]] = None) -> Optional[str]:
        try:
            cmd = [
                'libreoffice', '--headless', '--convert-to', 'pdf',
                '--outdir', str(pdf_path.parent), str(docx_path)
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

            if result.returncode != 0:
                _log(f"✗ LibreOffice 변환 실패: {result.stderr}", callback)
                return None

            if pdf_path.exists():
                _log(f"✓ 변환 완료: {docx_path.name}", callback)
                return str(pdf_path)

            pdf_files = list(pdf_path.parent.glob(Path(docx_path).stem + '*.pdf'))
            if pdf_files:
                _log(f"✓ 변환 완료: {docx_path.name}", callback)
                return str(pdf_files[0])

            return None

        except FileNotFoundError:
            _log("✗ LibreOffice 가 설치되어 있지 않습니다.", callback)
            return None
        except Exception as e:
            _log(f"✗ LibreOffice 변환 오류: {e}", callback)
            return None

    def convert_and_merge(self, docx_files: List[Path], output_path: Path,
                          callback: Optional[Callable[[str], None]] = None,
                          cleanup: bool = False) -> bool:
        """Word 파일들을 PDF로 변환하고 병합.

        cleanup=True 시 각 DOCX를 임시 PDF로 변환 → 즉시 병합 → 임시 PDF 삭제.
        개별 PDF 파일이 디스크에 남지 않는다.
        """
        if not docx_files:
            _log("오류: 변환할 Word 파일이 없습니다.", callback)
            return False

        if not PYPDF_AVAILABLE and cleanup:
            _log("✗ pypdf 가 설치되어 있지 않습니다. pip install pypdf", callback)
            return False

        output_path = Path(output_path).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)

        total = len(docx_files)

        if cleanup:
            # 변환 즉시 병합 → 임시 PDF 삭제 (개별 파일 잔류 없음)
            import tempfile
            _log(f"Word 파일 {total}개를 PDF로 변환·병합 중 (개별 파일 미저장)...", callback)
            writer = PdfWriter()
            merged = 0

            with tempfile.TemporaryDirectory() as tmpdir:
                for i, docx_file in enumerate(docx_files, 1):
                    _log(f"[{i}/{total}] 변환 중: {docx_file.name}", callback)
                    tmp_pdf = Path(tmpdir) / (docx_file.stem + ".pdf")
                    result = self.convert(docx_file, tmp_pdf, callback=callback)
                    if result and Path(result).exists():
                        reader = PdfReader(result)
                        for page in reader.pages:
                            writer.add_page(page)
                        merged += 1
                    # TemporaryDirectory 소멸 시 자동 삭제되므로 별도 삭제 불필요

            if merged == 0:
                _log("✗ 변환된 PDF 파일이 없습니다.", callback)
                return False

            try:
                with open(output_path, "wb") as f:
                    writer.write(f)
                file_size = output_path.stat().st_size
                _log(f"✓ 병합 완료: {output_path.name}  ({file_size / 1024:.1f} KB)", callback)
                return True
            except Exception as e:
                _log(f"✗ PDF 저장 실패: {e}", callback)
                return False

        else:
            # 기존 방식: 변환 후 일괄 병합 (개별 PDF 파일 유지)
            _log(f"Word 파일 {total}개를 PDF로 변환 중...", callback)
            pdf_paths = []

            for i, docx_file in enumerate(docx_files, 1):
                _log(f"[{i}/{total}] 변환 중: {docx_file.name}", callback)
                pdf_path = self.convert(docx_file, callback=callback)
                if pdf_path:
                    pdf_paths.append(pdf_path)

            if not pdf_paths:
                _log("✗ 변환된 PDF 파일이 없습니다.", callback)
                return False

            _log(f"{len(pdf_paths)}개 PDF 파일을 병합 중...", callback)
            return self._merge_pdfs(pdf_paths, output_path, callback)

    def _merge_pdfs(self, pdf_files: List[str], output_path: Path,
                    callback: Optional[Callable[[str], None]] = None) -> bool:
        if not PYPDF_AVAILABLE:
            _log("✗ pypdf 가 설치되어 있지 않습니다. pip install pypdf", callback)
            return False

        writer = PdfWriter()
        total = len(pdf_files)

        try:
            for i, pdf_file in enumerate(pdf_files, 1):
                _log(f"[{i}/{total}] 병합 중: {Path(pdf_file).name}", callback)
                reader = PdfReader(str(pdf_file))
                for page in reader.pages:
                    writer.add_page(page)

            with open(output_path, 'wb') as f:
                writer.write(f)

            file_size = output_path.stat().st_size
            _log(f"✓ 병합 완료: {output_path.name}  ({file_size / 1024:.1f} KB)", callback)
            return True

        except Exception as e:
            _log(f"✗ PDF 병합 실패: {e}", callback)
            return False

    def merge_only(self, pdf_files: List[Path], output_path: Path,
                   callback: Optional[Callable[[str], None]] = None) -> bool:
        if not PYPDF_AVAILABLE:
            _log("✗ pypdf 가 설치되어 있지 않습니다. pip install pypdf", callback)
            return False

        if not pdf_files:
            _log("오류: 병합할 PDF 파일이 없습니다.", callback)
            return False

        output_path = Path(output_path).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)

        writer = PdfWriter()
        total = len(pdf_files)

        try:
            for i, pdf_file in enumerate(pdf_files, 1):
                pdf_file = Path(pdf_file).resolve()
                if not pdf_file.exists():
                    _log(f"경고: 파일이 없음 — {pdf_file.name}", callback)
                    continue
                _log(f"[{i}/{total}] 추가 중: {pdf_file.name}", callback)
                reader = PdfReader(str(pdf_file))
                for page in reader.pages:
                    writer.add_page(page)

            with open(output_path, 'wb') as f:
                writer.write(f)

            _log(f"✓ PDF 병합 완료: {output_path.name}", callback)
            return True

        except Exception as e:
            _log(f"✗ PDF 병합 실패: {e}", callback)
            return False


# ============================================================================
# 모듈 레벨 함수
# ============================================================================

def docx_to_pdf(docx_path: Path, output_path: Optional[Path] = None,
                callback: Optional[Callable[[str], None]] = None) -> Optional[str]:
    """단일 Word 파일을 PDF 로 변환"""
    converter = DocxToPdfConverter()
    return converter.convert(docx_path, output_path, callback)


def convert_and_merge(docx_files: List[Path], output_path: Path,
                      callback: Optional[Callable[[str], None]] = None,
                      cleanup: bool = False) -> bool:
    """Word 파일들을 PDF 로 변환 후 병합.
    cleanup=True 시 개별 PDF를 디스크에 남기지 않는다.
    """
    converter = DocxToPdfConverter()
    return converter.convert_and_merge(docx_files, output_path, callback, cleanup)


def merge_pdfs(pdf_files: List[Path], output_path: Path,
               callback: Optional[Callable[[str], None]] = None) -> bool:
    """PDF 파일들만 병합"""
    converter = DocxToPdfConverter()
    return converter.merge_only(pdf_files, output_path, callback)


def batch_docx_to_pdf(input_dir: Path, output_dir: Optional[Path] = None,
                      recursive: bool = False,
                      callback: Optional[Callable[[str], None]] = None) -> List[str]:
    """디렉토리 내의 모든 .docx 파일을 PDF 로 변환"""
    input_dir = Path(input_dir).resolve()
    if output_dir is None:
        output_dir = input_dir
    else:
        output_dir = Path(output_dir).resolve()

    output_dir.mkdir(parents=True, exist_ok=True)

    docx_files = list(input_dir.rglob('*.docx') if recursive else input_dir.glob('*.docx'))
    total = len(docx_files)
    _log(f"변환할 Word 파일 {total}개 발견", callback)

    pdf_paths = []
    converter = DocxToPdfConverter()

    for i, docx_file in enumerate(docx_files, 1):
        _log(f"[{i}/{total}] 변환 중: {docx_file.name}", callback)
        pdf_path = output_dir / docx_file.with_suffix('.pdf').name
        pdf_path_str = converter.convert(docx_file, pdf_path, callback)
        if pdf_path_str:
            pdf_paths.append(pdf_path_str)

    _log(f"✓ 총 {len(pdf_paths)}개 파일 변환 완료", callback)
    return pdf_paths
