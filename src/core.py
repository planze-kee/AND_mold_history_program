"""
통합 핵심 모듈: HWP 처리, 문서 생성, 이미지 처리
"""
import argparse
import csv
import json
import logging
import re
import shutil
import zlib
import struct
import olefile
from datetime import datetime, date
from pathlib import Path
from typing import Dict, List, Set, Tuple, Optional
from io import BytesIO

from docx import Document
from docx.shared import Cm, Pt
from openpyxl import Workbook, load_workbook
from PIL import Image
import sys

logger = logging.getLogger(__name__)


# ============================================================================
# 사용자 정의 예외
# ============================================================================
class HWPProcessingError(Exception):
    """HWP 파일 처리 중 발생하는 오류"""
    pass


class ImageExtractionError(Exception):
    """이미지 추출 중 발생하는 오류"""
    pass


class DocumentGenerationError(Exception):
    """Word 문서 생성 중 발생하는 오류"""
    pass


class XLSXProcessingError(Exception):
    """XLSX 파일 처리 중 발생하는 오류"""
    pass


# ============================================================================
# HWP 텍스트 추출
# ============================================================================
class HWPTextExtractor:
    """Extract text from HWP5 BodyText stream"""

    PARA_TEXT = 67  # 0x43

    def __init__(self, section_data: bytes):
        self.section_data = section_data

    def extract_all_text(self) -> List[str]:
        """Extract all text from decompressed section"""
        texts = []
        pos = 0

        while pos < len(self.section_data):
            if pos + 4 > len(self.section_data):
                break

            # Parse record header (4 bytes)
            hdr = struct.unpack('<I', self.section_data[pos:pos+4])[0]
            tag = hdr & 0x3FF
            level = (hdr >> 10) & 0x3FF
            size = (hdr >> 20) & 0xFFF

            pos += 4

            # Handle extended size
            if size == 0xFFF:
                if pos + 4 > len(self.section_data):
                    break
                size = struct.unpack('<I', self.section_data[pos:pos+4])[0]
                pos += 4

            if pos + size > len(self.section_data):
                break

            payload = self.section_data[pos:pos+size]

            # Extract PARA_TEXT records (tag 67)
            if tag == self.PARA_TEXT and size > 0:
                text = self._decode_text(payload)
                if text.strip():
                    texts.append(text)

            pos += size

        return texts

    def _decode_text(self, payload: bytes) -> str:
        """Decode text from PARA_TEXT payload (UTF-16LE)"""
        result = []

        for i in range(0, len(payload), 2):
            if i + 1 < len(payload):
                char_code = struct.unpack('<H', payload[i:i+2])[0]

                if char_code == 0x000D:
                    break
                elif char_code < 0x0020 and char_code != 0:
                    continue
                else:
                    try:
                        result.append(chr(char_code))
                    except (ValueError, OverflowError):
                        pass

        return ''.join(result).strip()


# ============================================================================
# HWP 데이터 추출
# ============================================================================
class HWPDataExtractor:
    """Extract form data from HWP using olefile"""

    def __init__(self, hwp_filepath: str):
        self.filepath = hwp_filepath
        self.filename = Path(hwp_filepath).stem

    def sanitize_filename(self, text):
        """파일명으로 사용할 수 없는 문자를 '_'로 치환"""
        return re.sub(r'[<>:"/\\|?*]', '_', text)

    def extract(self) -> Dict[str, str]:
        """Extract all form fields from HWP file"""
        row = {chr(65+i): '' for i in range(29)}  # A-AC
        row['A'] = self.filename

        try:
            ole = olefile.OleFileIO(self.filepath)
            all_streams = ole.listdir()

            section_data = None
            for stream_path in all_streams:
                stream_str = '/'.join(stream_path) if isinstance(stream_path, list) else str(stream_path)
                if 'BodyText' in stream_str and 'Section0' in stream_str:
                    try:
                        section_data = ole.openstream(stream_path).read()
                        break
                    except Exception:
                        pass

            if section_data is None:
                ole.close()
                return row

            try:
                decompressed = zlib.decompress(section_data, -15)
            except zlib.error:
                decompressed = section_data

            extractor = HWPTextExtractor(decompressed)
            texts = extractor.extract_all_text()
            self._extract_fields(row, texts)

            ole.close()

        except Exception as e:
            logger.warning(f"HWP 데이터 추출 실패: {self.filepath} - {e}")

        return row

    def _extract_fields(self, row: Dict, texts: List[str]):
        """Extract field values from text list with form-specific parsing"""

        form_start = 0
        for i, text in enumerate(texts):
            if '保管会社名' in text and ':' in text:
                form_start = i
                break

        if form_start >= len(texts):
            return

        form_texts = texts[form_start:]

        # Pattern 1: Label : Value (on same line)
        for text in form_texts:
            if ':' in text:
                parts = text.split(':', 1)
                if len(parts) == 2:
                    label, value = parts[0].strip(), parts[1].strip()
                    value = value.strip()

                    if '保管会社名' in label:
                        row['B'] = value
                    elif '作成日子' in label:
                        row['C'] = value
                    elif '管理番号' in label:
                        row['D'] = value
                    elif '金型価' in label or '金型価格' in label:
                        row['U'] = value

        # 다음 라벨 목록 (헤더 라벨들)
        next_labels = {
            '分 類', '分類', '現 保管処', '現保管処', '製作処', 'MODEL名', 'MODEL',
            '量産処', '品 名', '品名', '図番番号', '図番', '金型規格', '規格',
            '金型材質', 'BASE', 'CORE', 'CAVITY 数', 'CAVITY', '金型寿命',
            'GATE 型式', 'GATE', '使用機械', '機械', '契約日', '承認日'
        }

        # Pattern 2: Label on one line, value on next (개선된 로직)
        i = 0
        while i < len(form_texts):
            text = form_texts[i]

            if text == '分 類' or text == '分類':
                # 다음 라벨이 아닌 첫 번째 값 찾기
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['E'] = form_texts[j]
                        break

            elif text == '現 保管処' or text == '現保管処':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['F'] = form_texts[j]
                        break

            elif text == '製作処':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['G'] = form_texts[j]
                        break

            elif text == 'MODEL名' or text == 'MODEL':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['H'] = form_texts[j]
                        break

            elif text == '量産処':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['I'] = form_texts[j]
                        break

            elif text == '品 名' or text == '品名':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['J'] = form_texts[j]
                        break

            elif text == '図番番号' or text == '図番':
                # 숫자 패턴이 있으면 우선적으로 추출
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    next_text = form_texts[j]
                    if next_text.strip() and next_text not in next_labels:
                        row['K'] = next_text
                        break

            elif text == '金型規格' or text == '規格':
                # 사이즈 패턴 검사: XxXxX 형태 (예: 650x650x650)
                size_pattern = re.compile(r'\d{1,5}x\d{1,5}x\d{1,5}', re.IGNORECASE)
                size_value = ""
                # 다음 5개 항목까지 검사
                for j in range(i + 1, min(i + 6, len(form_texts))):
                    if size_pattern.search(form_texts[j]):
                        size_value = form_texts[j]
                        break
                row['L'] = size_value  # 패턴 없으면 빈 값

            elif text == '金型材質':
                # BASE와 CORE가 가로로 배치됨: BASE [값] | CORE [값]
                base_value = ""
                core_value = ""

                j = i + 1
                while j < len(form_texts) and j < i + 15:
                    current_text = form_texts[j]

                    if current_text == 'BASE':
                        for k in range(j + 1, len(form_texts)):
                            next_text = form_texts[k]
                            if next_text == 'CORE':
                                break
                            if next_text.strip() and next_text not in next_labels:
                                base_value = next_text
                                break

                    elif current_text == 'CORE':
                        for k in range(j + 1, len(form_texts)):
                            next_text = form_texts[k]
                            if next_text in next_labels and next_text not in ['BASE', 'CORE']:
                                break
                            if next_text.strip() and next_text not in next_labels:
                                core_value = next_text
                                break

                    j += 1

                row['M'] = base_value
                row['N'] = core_value

            elif text == 'CAVITY 数' or text == 'CAVITY':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['O'] = form_texts[j]
                        break

            elif text == '金型寿命':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['P'] = form_texts[j]
                        break

            elif text == 'GATE 型式' or text == 'GATE':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['Q'] = form_texts[j]
                        break

            elif text == '使用機械' or text == '機械':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['R'] = form_texts[j]
                        break

            elif text == '契約日':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    if form_texts[j].strip() and form_texts[j] not in next_labels:
                        row['S'] = form_texts[j]
                        break

            elif text == '承認日':
                for j in range(i + 1, min(i + 5, len(form_texts))):
                    next_text = form_texts[j]
                    if next_text.strip() and '年' in next_text:
                        row['T'] = next_text
                        break

            i += 1

        # 금형사진 파일명 자동 생성 및 저장 (품명_도면번호 형식)
        product_name = row['J'].strip()
        drawing_no = row['K'].strip()

        if product_name or drawing_no:
            product_name_safe = self.sanitize_filename(product_name)
            drawing_no_safe = self.sanitize_filename(drawing_no)

            image_filename = f"{product_name_safe}_{drawing_no_safe}"
            row[chr(65 + 28)] = image_filename


# ============================================================================
# HWP 처리기 (HWP → XLSX)
# ============================================================================
class HWPProcessor:
    """Process HWP files and convert to XLSX"""

    CANONICAL_HEADERS = [
        "File name", "保管会社名", "作成日子", "管理番号", "分 類", "現 保管処", "製作処", "MODEL名", "量産処", "品 名",
        "図番番号", "金型規格", "金型材質-BASE", "金型材質-CORE", "CAVITY 数", "金型寿命", "GATE 型式", "使用機械", "契約日", "承認日",
        "金型価", "新作", "増作", "二元化", "業者変更", "仕様変更", "修理内訳", "事 由", "金型写真",
    ]

    @classmethod
    def extract_rows_from_hwp(cls, input_dir: Path, callback=None) -> List[List[str]]:
        hwp_files = sorted(input_dir.glob("*.hwp"))
        if not hwp_files:
            raise FileNotFoundError(f"No HWP files found in: {input_dir}")

        rows: List[List[str]] = []
        total = len(hwp_files)
        for idx, hwp_file in enumerate(hwp_files, start=1):
            msg = f"[{idx}/{total}] 추출 중: {hwp_file.name}"
            if callback:
                callback(msg)
            else:
                print(msg)
            extractor = HWPDataExtractor(str(hwp_file))
            row_data = extractor.extract()
            row_values = [row_data.get(chr(65 + i), "") for i in range(29)]
            rows.append(row_values)

        return rows

    @classmethod
    def save_xlsx(cls, rows: List[List[str]], output_xlsx: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "data"

        ws.append(cls.CANONICAL_HEADERS)
        for row in rows:
            padded = list(row) + [""] * max(0, len(cls.CANONICAL_HEADERS) - len(row))
            ws.append(padded[:len(cls.CANONICAL_HEADERS)])

        output_xlsx.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_xlsx)

    @classmethod
    def process(cls, input_dir: Path, output_xlsx: Path, callback=None) -> None:
        """Main entry point: HWP folder → XLSX file"""
        if not input_dir.exists():
            raise FileNotFoundError(f"input-dir not found: {input_dir}")

        rows = cls.extract_rows_from_hwp(input_dir, callback=callback)
        cls.save_xlsx(rows, output_xlsx)

        msg = f"✓ 저장 완료: {output_xlsx} ({len(rows)}행)"
        if callback:
            callback(msg)
        else:
            print(msg)


# ============================================================================
# 이미지 추출기
# ============================================================================
class HWPImageExtractor:
    """Extract images from HWP files"""

    MAGIC_BYTES = {
        b'\xFF\xD8\xFF': 'jpg',
        b'\x89PNG\r\n\x1a\n': 'png',
        b'BM': 'bmp',
        b'GIF8': 'gif',
    }

    def __init__(self, hwp_path, output_dir='img'):
        self.hwp_path = hwp_path
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        self.file_name = Path(hwp_path).stem

        # HWP에서 품명과 도면번호 추출 및 파일명으로 안전한 형식으로 변환
        self.product_name = ""
        self.drawing_no = ""
        try:
            extractor = HWPDataExtractor(hwp_path)
            row_data = extractor.extract()
            self.product_name = self.sanitize_filename(row_data.get('J', '').strip())
            self.drawing_no = self.sanitize_filename(row_data.get('K', '').strip())
        except Exception as e:
            logger.warning(f"HWP 메타데이터 추출 실패 (파일명으로 대체): {hwp_path} - {e}")

    def sanitize_filename(self, text):
        """파일명으로 사용할 수 없는 문자를 '_'로 치환"""
        return re.sub(r'[<>:"/\\|?*]', '_', text)

    def detect_image_format(self, data):
        for magic, fmt in self.MAGIC_BYTES.items():
            if data.startswith(magic):
                return fmt
        return None

    def try_decompress(self, data):
        """Try to decompress zlib compressed data"""
        try:
            return zlib.decompress(data, -15)
        except zlib.error:
            try:
                return zlib.decompress(data)
            except zlib.error:
                return data

    def try_fix_image(self, data):
        try:
            img = Image.open(BytesIO(data))
            img.verify()
            return data, None
        except Exception:
            try:
                img = Image.open(BytesIO(data))
                output = BytesIO()
                img.convert('RGB').save(output, format='JPEG', quality=95)
                return output.getvalue(), 'jpg'
            except Exception as e:
                logger.debug(f"이미지 변환 실패 (스킵): {e}")
                return None, 'error'

    def extract_images(self):
        if not Path(self.hwp_path).exists():
            return 0

        if not olefile.isOleFile(self.hwp_path):
            return 0

        try:
            ole = olefile.OleFileIO(self.hwp_path)
        except (olefile.Error, OSError) as e:
            logger.warning(f"OLE 파일 열기 실패: {self.hwp_path} - {e}")
            return 0

        extracted_count = 0

        try:
            for entry in ole.listdir():
                entry_name = '/'.join(entry)

                if 'BinData' in entry_name and ole.get_type(entry) == olefile.STGTY_STREAM:
                    try:
                        raw_data = ole.openstream(entry).read()

                        if not raw_data:
                            continue

                        image_data = self.try_decompress(raw_data)

                        if not image_data:
                            continue

                        detected_fmt = self.detect_image_format(image_data)
                        fixed_data, conversion_fmt = self.try_fix_image(image_data)

                        if fixed_data is None:
                            continue

                        final_fmt = conversion_fmt or detected_fmt or 'jpg'
                        if final_fmt == 'pdf':
                            final_fmt = 'jpg'

                        image_index = extracted_count + 1
                        # 품명_도번 형식으로 생성 (예: VFD HOLDER_1072001450.jpg)
                        base_filename = f"{self.product_name}_{self.drawing_no}" if self.product_name and self.drawing_no else self.file_name
                        if image_index == 1:
                            output_filename = f"{base_filename}.{final_fmt}"
                        else:
                            output_filename = f"{base_filename}_{image_index}.{final_fmt}"

                        output_path = self.output_dir / output_filename

                        with open(output_path, 'wb') as f:
                            f.write(fixed_data)

                        extracted_count += 1

                    except Exception as e:
                        logger.debug(f"BinData 스트림 처리 실패 (스킵): {e}")
                        continue

            ole.close()

        except Exception as e:
            logger.error(f"이미지 추출 중 오류: {self.hwp_path} - {e}")
            return 0

        return extracted_count


# ============================================================================
# 문서 채우기 (XLSX → DOCX)
# ============================================================================
class DocumentFiller:
    """Fill DOCX from XLSX with placeholder tokens and images"""

    CHECKBOX_LABELS = {"新作", "増作", "二元化", "業者更変", "機種更変"}
    IMAGE_TOKEN = "{金型写真}"
    IMAGE_EXTS = [".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"]

    ALIASES = {
        "保管社名": ["保管会社名", "保管社名", "storage_company", "보관사명"],
        "作成日子": ["作成日子", "created_date", "작성일자"],
        "分  類": ["分 類", "분류"],
        "現保管処": ["現 保管処", "current_storage", "현보관처"],
        "製作処": ["製作処", "maker", "제작처"],
        "MODEL名": ["MODEL名", "model_name", "모델명"],
        "量産処": ["量産処", "mass_prod", "양산처"],
        "品  名": ["品 名", "product_name", "품명", "제품명"],
        "図番番号": ["図番番号", "drawing_no", "도번번호", "도번", "품번"],
        "金型規格": ["金型規格", "mold_size", "금형규격"],
        "BASE": ["金型材質-BASE", "BASE"],
        "CORE": ["金型材質-CORE", "CORE"],
        "CAVITY 数": ["CAVITY 数", "CAVITY수"],
        "GATE型式": ["GATE 型式", "GATE型式"],
        "使用機械": ["使用機械", "machine", "사용기계"],
        "契約日": ["契約日", "contract_date", "계약일"],
        "承認日": ["承認日", "approval_date", "승인일"],
        "修理訳内容": ["修理内訳", "修理訳内容", "repair_history"],
        "業者更変": ["業者変更", "業者更変"],
        "機種更変": ["仕様変更", "機種更変"],
        "金型": ["金型価", "금형가"],
        "金型命": ["金型寿命", "금형수명"],
    }

    @staticmethod
    def normalize(text: str) -> str:
        if text is None:
            return ""
        return re.sub(r"\s+", "", str(text)).strip().lower()

    @classmethod
    def row_norm_map(cls, row: Dict[str, str]) -> Dict[str, str]:
        return {cls.normalize(k): (v or "").strip() for k, v in row.items()}

    @classmethod
    def value_by_aliases(cls, norm_row: Dict[str, str], aliases: List[str]) -> str:
        for alias in aliases:
            key = cls.normalize(alias)
            if key in norm_row and norm_row[key] != "":
                return norm_row[key]
        return ""

    @classmethod
    def value_for_label(cls, label: str, norm_row: Dict[str, str]) -> str:
        if label == "金型写真":
            return cls.IMAGE_TOKEN

        if label in cls.CHECKBOX_LABELS:
            raw = cls.value_by_aliases(norm_row, [label] + cls.ALIASES.get(label, []))
            return "O" if raw == "1" else ""

        direct = cls.value_by_aliases(norm_row, [label])
        if direct != "":
            return direct

        if label in cls.ALIASES:
            return cls.value_by_aliases(norm_row, cls.ALIASES[label])

        return ""

    @classmethod
    def replace_placeholders(cls, text: str, norm_row: Dict[str, str]) -> Tuple[str, int, Set[str]]:
        if not text:
            return text, 0, set()

        replaced_count = 0
        replaced_labels: Set[str] = set()

        def repl(match):
            nonlocal replaced_count
            label = match.group(0)[1:-1].strip()
            value = cls.value_for_label(label, norm_row)
            replaced_count += 1
            replaced_labels.add(label)
            return value

        return re.sub(r"\{[^{}]+\}", repl, text), replaced_count, replaced_labels

    @staticmethod
    def apply_small_font_if_needed(paragraph, replaced_labels: Set[str]) -> None:
        if not ({"BASE", "CORE"} & replaced_labels):
            return
        for run in paragraph.runs:
            run.font.size = Pt(6)

    @classmethod
    def _copy_text_run_format(cls, src_run, dst_run) -> None:
        dst_run.bold = src_run.bold
        dst_run.italic = src_run.italic
        dst_run.underline = src_run.underline
        dst_run.style = src_run.style
        dst_run.font.name = src_run.font.name
        dst_run.font.size = src_run.font.size

    @classmethod
    def _insert_run_after(cls, paragraph, after_run, text: str = ""):
        new_run = paragraph.add_run(text)
        after_run._r.addnext(new_run._r)
        return new_run

    @classmethod
    def _insert_image_in_paragraph(cls, paragraph, image_path: Optional[Path]) -> int:
        para_text = paragraph.text or ""
        if cls.IMAGE_TOKEN not in para_text:
            return 0

        inserted = 0
        original_runs = list(paragraph.runs)

        # Case 1: token exists fully inside at least one run
        token_in_single_run = False
        for run in original_runs:
            run_text = run.text or ""
            if cls.IMAGE_TOKEN not in run_text:
                continue
            token_in_single_run = True

            parts = run_text.split(cls.IMAGE_TOKEN)
            run.text = parts[0]
            cursor = run

            for part in parts[1:]:
                if image_path and image_path.exists():
                    img_run = cls._insert_run_after(paragraph, cursor)
                    img_run.add_picture(str(image_path), width=Cm(18.5), height=Cm(13.87))
                    cursor = img_run
                    inserted += 1

                if part:
                    txt_run = cls._insert_run_after(paragraph, cursor, part)
                    cls._copy_text_run_format(run, txt_run)
                    cursor = txt_run

        if token_in_single_run:
            return inserted

        # Case 2: token is split across multiple runs
        if not original_runs:
            return 0

        base_run = original_runs[0]
        for r in original_runs:
            r.text = ""

        parts = para_text.split(cls.IMAGE_TOKEN)
        base_run.text = parts[0]
        cursor = base_run

        for part in parts[1:]:
            if image_path and image_path.exists():
                img_run = cls._insert_run_after(paragraph, cursor)
                img_run.add_picture(str(image_path), width=Cm(18.5), height=Cm(13.87))
                cursor = img_run
                inserted += 1

            if part:
                txt_run = cls._insert_run_after(paragraph, cursor, part)
                cls._copy_text_run_format(base_run, txt_run)
                cursor = txt_run

        return inserted

    @classmethod
    def insert_images_by_token(cls, doc: Document, image_path: Optional[Path]) -> int:
        count = 0
        for p in doc.paragraphs:
            count += cls._insert_image_in_paragraph(p, image_path)
        for t in doc.tables:
            for r in t.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        count += cls._insert_image_in_paragraph(p, image_path)
        return count

    @classmethod
    def find_image_for_output(cls, img_dir: Path, output_stem: str) -> Optional[Path]:
        """이미지 파일 찾기 (품명_도번 형식)

        1. 정확한 파일명으로 먼저 찾기
        2. sanitize 처리된 파일명으로 찾기
        3. output_stem에 포함된 도번으로도 확인
        """
        if not output_stem or not img_dir.exists():
            return None

        # 정확한 파일명으로 찾기
        for ext in cls.IMAGE_EXTS:
            p = img_dir / f"{output_stem}{ext}"
            if p.exists():
                return p

        # sanitize 처리한 이름으로 찾기 (파일명 사용 불가 문자 처리)
        sanitized_stem = re.sub(r'[<>:"/\\|?*]', '_', output_stem)
        if sanitized_stem != output_stem:
            for ext in cls.IMAGE_EXTS:
                p = img_dir / f"{sanitized_stem}{ext}"
                if p.exists():
                    return p

        # output_stem을 sanitize한 후 glob으로 찾기
        # 예: output_stem="1072001446 / 2001447"
        #     sanitized="1072001446 _ 2001447"
        #     파일: "FRONT POLE _ REAR POLE_1072001446 _ 2001447.jpg"
        # → glob: "*1072001446 _ 2001447*"로 검색
        matches = []
        search_stem = sanitized_stem  # 이미 sanitize된 버전 사용
        for ext in cls.IMAGE_EXTS:
            try:
                # 파일명에 공백이 있을 수 있으니 escape 처리
                matches.extend(img_dir.glob(f"*{search_stem}*{ext}"))
            except (OSError, ValueError):
                pass

        if matches:
            # 연번이 없는 파일 우선 (예: "xxx_yyy.jpg" > "xxx_yyy_2.jpg")
            matches_no_suffix = [m for m in matches if not m.stem.endswith(('_2', '_3', '_4', '_5'))]
            if matches_no_suffix:
                return matches_no_suffix[0]
            return matches[0]

        return None

    @classmethod
    def replace_in_doc(cls, doc: Document, norm_row: Dict[str, str]) -> int:
        total = 0
        for p in doc.paragraphs:
            new_text, cnt, labels = cls.replace_placeholders(p.text, norm_row)
            if new_text != p.text:
                p.text = new_text
                cls.apply_small_font_if_needed(p, labels)
            total += cnt
        for t in doc.tables:
            for r in t.rows:
                for c in r.cells:
                    for p in c.paragraphs:
                        new_text, cnt, labels = cls.replace_placeholders(p.text, norm_row)
                        if new_text != p.text:
                            p.text = new_text
                            cls.apply_small_font_if_needed(p, labels)
                        total += cnt
        return total

    @classmethod
    def pick_output_name(cls, row: Dict[str, str], idx: int) -> str:
        nrow = cls.row_norm_map(row)
        value = cls.value_by_aliases(nrow, ["File name", "file_name", "filename", "source_hwp"])
        if value:
            m = re.search(r"\d{2}-\d{3}", value)
            if m:
                return m.group(0)
            digits = re.findall(r"\d+", value)
            if digits:
                return "".join(digits)
            stem = Path(value).stem.strip()
            if stem:
                return stem
        return f"row_{idx:03d}"

    @staticmethod
    def unique_path(path: Path) -> Path:
        if not path.exists():
            return path
        n = 2
        while True:
            c = path.with_name(f"{path.stem}_{n}{path.suffix}")
            if not c.exists():
                return c
            n += 1

    @classmethod
    def load_rows_from_xlsx(cls, xlsx_path: Path) -> List[Dict[str, str]]:
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.active
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

        rows: List[Dict[str, str]] = []
        for row_values in ws.iter_rows(min_row=2, values_only=True):
            row_dict: Dict[str, str] = {}
            has_any = False
            for i, header in enumerate(headers):
                v = row_values[i] if i < len(row_values) else None
                text = "" if v is None else str(v).strip()
                if text:
                    has_any = True
                row_dict[header] = text
            if has_any:
                rows.append(row_dict)
        wb.close()
        return rows

    @classmethod
    def process(cls, xlsx_path: Path, template_path: Path, out_dir: Path, img_dir: Path,
                limit: int = 0, callback=None) -> None:
        """Main entry point: XLSX + Template → DOCX files with images"""
        def _log(msg):
            if callback:
                callback(msg)
            else:
                print(msg)

        if not xlsx_path.exists():
            raise FileNotFoundError(f"XLSX not found: {xlsx_path}")
        if not template_path.exists():
            raise FileNotFoundError(f"Template not found: {template_path}")

        rows = cls.load_rows_from_xlsx(xlsx_path)
        if limit > 0:
            rows = rows[:limit]
        if not rows:
            _log("XLSX에 처리할 행이 없습니다.")
            return

        out_dir.mkdir(parents=True, exist_ok=True)

        # 이미지 파일명 저장 (XLSX 업데이트용)
        image_updates = {}  # {row_idx: image_filename}

        total = len(rows)
        for idx, row in enumerate(rows, start=1):
            doc = Document(str(template_path))
            nrow = cls.row_norm_map(row)
            replaced = cls.replace_in_doc(doc, nrow)

            out_name = cls.pick_output_name(row, idx)

            # 이미지 파일 찾기: 우선순위
            image_path = None

            # 1. XLSX의 "金型写真"(품명_도번)으로 먼저 시도
            image_name = row.get("金型写真", "").strip()
            if image_name:
                image_path = cls.find_image_for_output(img_dir, image_name)

            # 2. "金型写真"으로 못 찾으면 XLSX의 도번으로 시도
            if not image_path:
                drawing_no = cls.value_by_aliases(nrow, ["図番番号", "drawing_no", "도번번호", "도번", "품번"])
                if drawing_no:
                    image_path = cls.find_image_for_output(img_dir, drawing_no)

            # 3. 도번으로도 못 찾으면 out_name(연번)으로 시도
            if not image_path:
                image_path = cls.find_image_for_output(img_dir, out_name)

            # 4. img_dir에 없으면 out_dir/.data/ 도 검색 (신규 발행 첨부 이미지)
            if not image_path:
                search_name = image_name or out_name
                image_path = cls.find_image_for_output(out_dir / ".data", search_name)

            inserted = cls.insert_images_by_token(doc, image_path)

            out_path = cls.unique_path(out_dir / f"{out_name}.docx")
            doc.save(str(out_path))

            # 실제 삽입된 이미지 파일명 저장 (XLSX 업데이트용)
            image_updates[idx] = image_path.name if image_path else ""

            img_status = "OK" if image_path else "NONE"
            _log(f"[{idx}/{total}] 저장: {out_name}.docx (치환={replaced}, 이미지={img_status})")

        # XLSX의 "金型写真" 컬럼 업데이트
        cls._update_xlsx_images(xlsx_path, image_updates)
        _log(f"✓ XLSX 이미지 컬럼 업데이트 완료")

    @classmethod
    def _update_xlsx_images(cls, xlsx_path: Path, image_updates: Dict[int, str]) -> None:
        """XLSX의 金型写真 컬럼에 실제 이미지 파일명 기록"""
        try:
            wb = load_workbook(xlsx_path)
            ws = wb.active

            # 金型写真 컬럼 정확하게 찾기 (전체 텍스트 매칭)
            photo_col_idx = None
            for col_idx in range(1, 40):
                h = ws.cell(1, col_idx).value
                if h:
                    h_str = str(h).strip()
                    # 정확하게 "金型写真" 찾기 (다른 컬럼과 혼동 방지)
                    if "金型写真" in h_str:
                        photo_col_idx = col_idx
                        break

            # 金型写真 컬럼이 없으면 마지막 컬럼 다음에 추가
            if photo_col_idx is None:
                # 마지막 비어있지 않은 컬럼 찾기
                last_col = 1
                for col_idx in range(1, 40):
                    if ws.cell(1, col_idx).value:
                        last_col = col_idx
                photo_col_idx = last_col + 1
                # 헤더 추가
                ws.cell(1, photo_col_idx).value = "金型写真"

            # 이미지 파일명 업데이트
            for row_idx, img_filename in image_updates.items():
                ws.cell(row_idx + 1, photo_col_idx).value = img_filename

            wb.save(xlsx_path)
            wb.close()
        except Exception as e:
            print(f"Warning: Could not update XLSX with image filenames: {e}")


# ============================================================================
# DocxSyncManager - XLSX-DOCX 동기화 (기능 1+2)
# ============================================================================
class DocxSyncManager:
    """XLSX 변경사항을 기존 DOCX 파일에 동기화 (파일명 변경 + 내용 갱신 + 이미지 rename)"""

    MANIFEST_FILENAME = "manifest.json"

    @classmethod
    def _manifest_path(cls, output_dir: Path) -> Path:
        return output_dir / cls.MANIFEST_FILENAME

    @classmethod
    def load_manifest(cls, output_dir: Path) -> Dict:
        """manifest.json 로드 (없으면 빈 딕셔너리)"""
        path = cls._manifest_path(output_dir)
        if path.exists():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    @classmethod
    def save_manifest(cls, output_dir: Path, manifest: Dict) -> None:
        """manifest.json 저장"""
        output_dir.mkdir(parents=True, exist_ok=True)
        path = cls._manifest_path(output_dir)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)

    @classmethod
    def extract_serial(cls, file_name: str) -> str:
        """파일명에서 연번 추출 (예: '19-001' → '001')"""
        m = re.search(r"-(\d+)$", file_name)
        if m:
            return m.group(1)
        m = re.search(r"(\d+)$", file_name)
        if m:
            return m.group(1)
        return ""

    @classmethod
    def rename_image_files(cls, img_dir: Path, old_serial: str, new_serial: str, callback=None) -> int:
        """연번 변경에 따른 이미지 파일명 변경"""
        if not img_dir.exists() or not old_serial or old_serial == new_serial:
            return 0
        renamed = 0
        for img_file in sorted(img_dir.iterdir()):
            if img_file.is_file() and img_file.name.startswith(f"{old_serial}_"):
                new_name = new_serial + img_file.name[len(old_serial):]
                new_path = img_dir / new_name
                if not new_path.exists():
                    img_file.rename(new_path)
                    renamed += 1
                    if callback:
                        callback(f"  이미지 rename: {img_file.name} → {new_name}")
        return renamed

    @classmethod
    def sync(cls, xlsx_path: Path, template_path: Path, output_dir: Path,
             img_dir: Path, callback=None) -> None:
        """XLSX 변경사항을 DOCX 파일에 동기화"""
        if callback:
            callback("동기화 시작...")

        rows = DocumentFiller.load_rows_from_xlsx(xlsx_path)
        manifest = cls.load_manifest(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        # 1단계: 파일명 변경 감지 및 rename
        renamed_files = 0
        for idx, row in enumerate(rows, start=1):
            mgmt_no = row.get("管理番号", "").strip()
            if not mgmt_no:
                continue
            current_name = DocumentFiller.pick_output_name(row, idx)
            if mgmt_no not in manifest:
                continue
            old_name = manifest[mgmt_no].get("file_name", "")
            if not old_name or old_name == current_name:
                continue

            # Word 파일 rename
            old_docx = output_dir / f"{old_name}.docx"
            new_docx = output_dir / f"{current_name}.docx"
            if old_docx.exists() and not new_docx.exists():
                old_docx.rename(new_docx)
                renamed_files += 1
                if callback:
                    callback(f"Word rename: {old_name}.docx → {current_name}.docx")

            # 이미지 파일 rename
            old_serial = cls.extract_serial(old_name)
            new_serial = cls.extract_serial(current_name)
            if old_serial and new_serial and old_serial != new_serial:
                img_count = cls.rename_image_files(img_dir, old_serial, new_serial, callback)
                if img_count > 0 and callback:
                    callback(f"  이미지 {img_count}개 rename 완료")

            # 이력 파일 rename (.data/ 하위 폴더)
            data_dir = output_dir / MaintenanceHistoryManager.DATA_DIR
            old_hist = data_dir / f"{old_name}_history.txt"
            new_hist = data_dir / f"{current_name}_history.txt"
            if old_hist.exists() and not new_hist.exists():
                old_hist.rename(new_hist)

        if renamed_files > 0 and callback:
            callback(f"파일명 변경: {renamed_files}건 처리됨")

        # 2단계: 전체 DOCX 재생성 (최신 XLSX 데이터로 덮어쓰기)
        if callback:
            callback("DOCX 내용 갱신 중...")

        total = len(rows)
        new_manifest = dict(manifest)  # 기존 manifest 복사

        for idx, row in enumerate(rows, start=1):
            try:
                doc = Document(str(template_path))
                nrow = DocumentFiller.row_norm_map(row)
                DocumentFiller.replace_in_doc(doc, nrow)

                out_name = DocumentFiller.pick_output_name(row, idx)
                image_name = row.get("金型写真", "").strip()
                search_name = image_name if image_name else out_name
                image_path = DocumentFiller.find_image_for_output(img_dir, search_name)
                if not image_path:
                    image_path = DocumentFiller.find_image_for_output(
                        output_dir / ".data", search_name
                    )

                DocumentFiller.insert_images_by_token(doc, image_path)
                out_path = output_dir / f"{out_name}.docx"
                doc.save(str(out_path))

                if callback:
                    callback(f"[{idx}/{total}] 갱신: {out_name}.docx")

                mgmt_no = row.get("管理番号", "").strip()
                if mgmt_no:
                    new_manifest[mgmt_no] = {
                        "file_name": out_name,
                        "serial": cls.extract_serial(out_name),
                        "品名": row.get("品 名", "").strip(),
                        "図番番号": row.get("図番番号", "").strip(),
                        "last_updated": datetime.now().isoformat(timespec="seconds"),
                    }
            except Exception as e:
                if callback:
                    callback(f"[{idx}/{total}] 오류: {e}")

        cls.save_manifest(output_dir, new_manifest)
        if callback:
            callback("✓ 동기화 완료 (manifest.json 저장됨)")


# ============================================================================
# MaintenanceHistoryManager - 유지보수 이력 관리 (기능 3)
# ============================================================================
class MaintenanceHistoryManager:
    """금형 유지보수 이력 관리 (.data/ 하위 폴더에 _history.txt 저장)"""

    DATA_DIR = ".data"

    @classmethod
    def get_history_path(cls, docx_path: Path) -> Path:
        data_dir = docx_path.parent / cls.DATA_DIR
        return data_dir / f"{docx_path.stem}_history.txt"

    @classmethod
    def read_history(cls, docx_path: Path) -> str:
        path = cls.get_history_path(docx_path)
        if path.exists():
            return path.read_text(encoding="utf-8")
        return f"# {docx_path.stem} 유지보수 이력\n\n"

    @classmethod
    def write_history(cls, docx_path: Path, content: str) -> None:
        path = cls.get_history_path(docx_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(content, encoding="utf-8")

    @classmethod
    def add_entry(cls, docx_path: Path, entry_date: str, entry_type: str,
                  content: str, person: str) -> None:
        """새 유지보수 이력 항목 추가"""
        current = cls.read_history(docx_path)
        entry = (
            f"\n## {entry_date}\n"
            f"- 유형: {entry_type}\n"
            f"- 내용: {content}\n"
            f"- 담당자: {person}\n"
        )
        cls.write_history(docx_path, current.rstrip() + "\n" + entry)

    @classmethod
    def _summarize_history(cls, history_content: str) -> Tuple[str, str]:
        """이력 내용에서 修理内訳/事由 요약 텍스트 생성 (최근 3개 항목)"""
        lines = history_content.splitlines()
        entries: List[Dict] = []
        current: Optional[Dict] = None

        for line in lines:
            if line.startswith("## "):
                if current:
                    entries.append(current)
                current = {"date": line[3:].strip(), "type": "", "content": "", "person": ""}
            elif current:
                if line.startswith("- 유형:"):
                    current["type"] = line[5:].strip()
                elif line.startswith("- 내용:"):
                    current["content"] = line[5:].strip()
                elif line.startswith("- 담당자:"):
                    current["person"] = line[6:].strip()
        if current:
            entries.append(current)

        recent = entries[-3:]
        repair_parts = [f"{e['date']}: {e['content']}" for e in recent if e.get("content")]
        reason_parts = [e["type"] for e in recent if e.get("type")]

        return " / ".join(repair_parts), " / ".join(reason_parts)

    @classmethod
    def update_xlsx_reason(cls, docx_path: Path, xlsx_path: Path, content: str,
                           callback=None) -> bool:
        """XLSX의 '事 由' 컬럼을 content로 갱신 (Word 재생성 없음)"""
        stem = docx_path.stem

        if not xlsx_path.exists():
            if callback:
                callback("오류: XLSX 파일을 찾을 수 없음")
            return False

        try:
            wb = load_workbook(xlsx_path)
            ws = wb.active
            headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

            reason_col_idx = None
            file_name_col_idx = None
            for i, h in enumerate(headers):
                if h == "File name":
                    file_name_col_idx = i
                elif h == "事 由":
                    reason_col_idx = i

            if file_name_col_idx is None:
                wb.close()
                if callback:
                    callback("오류: XLSX에 'File name' 컬럼 없음")
                return False

            target_row_idx = None
            for row_idx, row_values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                cell_val = str(row_values[file_name_col_idx] or "").strip()
                m = re.search(r"\d{2}-\d{3}", cell_val)
                row_stem = m.group(0) if m else Path(cell_val).stem
                if row_stem == stem:
                    target_row_idx = row_idx
                    break

            if target_row_idx is not None and reason_col_idx is not None:
                ws.cell(row=target_row_idx, column=reason_col_idx + 1).value = content

            wb.save(xlsx_path)
            wb.close()
            if callback:
                callback(f"XLSX 업데이트 완료 (行 {target_row_idx})")
            return True
        except PermissionError:
            if callback:
                callback(
                    f"XLSX 저장 실패: 파일이 다른 프로그램(Excel 등)에서 열려있습니다.\n"
                    f"  → {xlsx_path}\n"
                    f"  파일을 닫고 다시 시도하세요."
                )
            return False
        except Exception as e:
            if callback:
                callback(f"XLSX 업데이트 오류: {e}")
            return False

    @classmethod
    def apply_to_word(cls, docx_path: Path, xlsx_path: Path, template_path: Path,
                      img_dir: Path, callback=None) -> bool:
        """이력 내용을 XLSX에 반영하고 Word 파일 재생성"""
        if not docx_path.exists():
            if callback:
                callback("오류: Word 파일을 찾을 수 없음")
            return False

        history_content = cls.read_history(docx_path)
        stem = docx_path.stem

        if not xlsx_path.exists():
            if callback:
                callback("오류: XLSX 파일을 찾을 수 없음")
            return False

        # XLSX에서 해당 행 찾아 事 由 업데이트
        ok = cls.update_xlsx_reason(docx_path, xlsx_path, history_content, callback)
        if not ok:
            return False

        # Word 파일 재생성 (해당 행 데이터로)
        if not template_path.exists():
            if callback:
                callback("오류: 템플릿 파일을 찾을 수 없음")
            return False

        try:
            rows = DocumentFiller.load_rows_from_xlsx(xlsx_path)
            target_row = None
            for idx, row in enumerate(rows, start=1):
                row_name = DocumentFiller.pick_output_name(row, idx)
                if row_name == stem:
                    target_row = row
                    break

            if target_row is None:
                if callback:
                    callback("오류: XLSX에서 해당 행을 찾을 수 없음")
                return False

            doc = Document(str(template_path))
            nrow = DocumentFiller.row_norm_map(target_row)
            DocumentFiller.replace_in_doc(doc, nrow)

            image_name = target_row.get("金型写真", "").strip()
            image_path = None
            if image_name:
                image_path = DocumentFiller.find_image_for_output(img_dir, image_name)
                if not image_path:
                    # 신규 발행 첨부 이미지는 docx와 같은 디렉토리의 .data/ 에 저장됨
                    image_path = DocumentFiller.find_image_for_output(
                        docx_path.parent / ".data", image_name
                    )
            DocumentFiller.insert_images_by_token(doc, image_path)
            doc.save(str(docx_path))

            if callback:
                callback(f"✓ Word 파일 재생성 완료: {docx_path.name}")
            return True
        except Exception as e:
            if callback:
                callback(f"Word 재생성 오류: {e}")
            return False


# ============================================================================
# NewCardManager - 신규 이력카드 발행 (기능 4)
# ============================================================================
class NewCardManager:
    """신규 이력카드 발행: 자동 연번 생성 + XLSX 추가 + Word 파일 생성"""

    @classmethod
    def get_next_file_name(cls, xlsx_path: Path) -> str:
        """기존 XLSX에서 최대 연번을 찾아 다음 파일명 반환 (예: 19-006)"""
        rows = DocumentFiller.load_rows_from_xlsx(xlsx_path)
        max_serial = 0
        prefix = ""

        for row in rows:
            file_name = row.get("File name", "").strip()
            m = re.match(r"^(\d{2}-)(\d{3})$", file_name)
            if m:
                serial = int(m.group(2))
                if serial > max_serial:
                    max_serial = serial
                    prefix = m.group(1)

        if not prefix:
            prefix = date.today().strftime("%y") + "-"

        return f"{prefix}{max_serial + 1:03d}"

    @classmethod
    def sanitize(cls, text: str) -> str:
        return re.sub(r'[<>:"/\\|?*]', "_", text)

    @classmethod
    def add_to_xlsx(cls, xlsx_path: Path, row_dict: Dict[str, str]) -> None:
        """XLSX 맨 아래에 새 행 추가"""
        wb = load_workbook(xlsx_path)
        ws = wb.active
        headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
        new_row = [row_dict.get(h, "") for h in headers]
        ws.append(new_row)
        wb.save(xlsx_path)
        wb.close()

    @classmethod
    def generate_card(cls, xlsx_path: Path, template_path: Path, output_dir: Path,
                      img_dir: Path, row_dict: Dict[str, str],
                      image_source_path: Optional[Path] = None,
                      callback=None) -> Optional[Path]:
        """신규 이력카드 생성: XLSX 추가 → 이미지 복사(.data/) → Word 파일 생성"""
        try:
            product = cls.sanitize(row_dict.get("品 名", "").strip())
            drawing = cls.sanitize(row_dict.get("図番番号", "").strip())
            image_stem = f"{product}_{drawing}" if (product or drawing) else "image"
            row_dict.setdefault("金型写真", image_stem)

            # 첨부 이미지 처리: .data/ 에 품명_도번.ext 로 복사
            inserted_image_path: Optional[Path] = None
            if image_source_path and image_source_path.exists():
                data_dir = output_dir / ".data"
                data_dir.mkdir(parents=True, exist_ok=True)
                target_name = f"{image_stem}{image_source_path.suffix.lower()}"
                target_path = data_dir / target_name
                shutil.copy2(str(image_source_path), str(target_path))
                row_dict["金型写真"] = image_stem  # 확장자 제외한 스템
                inserted_image_path = target_path
                if callback:
                    callback(f"이미지 복사됨: .data/{target_name}")

            cls.add_to_xlsx(xlsx_path, row_dict)
            if callback:
                callback("XLSX에 새 행 추가 완료")

            file_name = row_dict.get("File name", "new_card").strip()

            doc = Document(str(template_path))
            nrow = DocumentFiller.row_norm_map(row_dict)
            DocumentFiller.replace_in_doc(doc, nrow)

            # 이미지 탐색: 직접 경로 → img_dir → .data/
            if inserted_image_path is None:
                search = row_dict.get("金型写真", "")
                if search:
                    inserted_image_path = DocumentFiller.find_image_for_output(img_dir, search)
                if not inserted_image_path and search:
                    inserted_image_path = DocumentFiller.find_image_for_output(
                        output_dir / ".data", search
                    )

            DocumentFiller.insert_images_by_token(doc, inserted_image_path)

            output_dir.mkdir(parents=True, exist_ok=True)
            out_path = DocumentFiller.unique_path(output_dir / f"{file_name}.docx")
            doc.save(str(out_path))

            if callback:
                callback(f"✓ 신규 이력카드 생성 완료: {out_path.name}")
            return out_path
        except Exception as e:
            if callback:
                callback(f"✗ 생성 오류: {e}")
            return None


if __name__ == "__main__":
    # CLI 사용 예제
    parser = argparse.ArgumentParser(description="Integrated HWP/XLSX processing tool")
    subparsers = parser.add_subparsers(dest='command', help='Commands')

    # HWP → XLSX
    hwp_parser = subparsers.add_parser('hwp2xlsx', help='Convert HWP files to XLSX')
    hwp_parser.add_argument('--input-dir', default='YES', help='Input directory with HWP files')
    hwp_parser.add_argument('--output', default='output_from_hwp.xlsx', help='Output XLSX path')

    # Extract images
    img_parser = subparsers.add_parser('extract-images', help='Extract images from HWP files')
    img_parser.add_argument('--input-dir', default='YES', help='Input directory with HWP files')
    img_parser.add_argument('--output-dir', default='img', help='Output directory for images')

    # XLSX → DOCX
    docx_parser = subparsers.add_parser('xlsx2docx', help='Fill DOCX files from XLSX')
    docx_parser.add_argument('--xlsx', default='output_from_hwp.xlsx', help='Input XLSX path')
    docx_parser.add_argument('--template', default='Word_양식.docx', help='Template DOCX path')
    docx_parser.add_argument('--out-dir', default='output', help='Output directory')
    docx_parser.add_argument('--img-dir', default='img', help='Image directory')
    docx_parser.add_argument('--limit', type=int, default=0, help='Process first N rows only')

    args = parser.parse_args()

    if args.command == 'hwp2xlsx':
        HWPProcessor.process(Path(args.input_dir), Path(args.output))
    elif args.command == 'extract-images':
        input_dir = Path(args.input_dir)
        hwp_files = sorted(input_dir.glob('*.hwp'))
        total_extracted = 0
        for hwp_file in hwp_files:
            extractor = HWPImageExtractor(str(hwp_file), args.output_dir)
            extracted = extractor.extract_images()
            if extracted > 0:
                print(f"[+] {hwp_file.stem}: {extracted} image(s)")
                total_extracted += extracted
        print(f"\n[*] Total extracted: {total_extracted}")
    elif args.command == 'xlsx2docx':
        DocumentFiller.process(
            Path(args.xlsx),
            Path(args.template),
            Path(args.out_dir),
            Path(args.img_dir),
            args.limit
        )
    elif args.command == 'docx2pdf':
        from pdf import DocxToPdfConverter
        converter = DocxToPdfConverter()
        converter.convert(Path(args.input), Path(args.output))
    elif args.command == 'merge-docx-pdf':
        from pdf import DocxToPdfConverter
        converter = DocxToPdfConverter()
        input_files = [Path(f) for f in args.inputs]
        output_path = Path(args.output)
        converter.convert_and_merge(input_files, output_path)

