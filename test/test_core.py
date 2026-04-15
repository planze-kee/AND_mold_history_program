"""
단위 테스트: src/core.py 핵심 로직 검증
pytest 로 실행: pytest test/test_core.py -v
"""
import sys
import re
from pathlib import Path

import pytest

# 프로젝트 루트를 sys.path에 추가
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.core import (
    HWPTextExtractor,
    HWPFieldExtractor,
    HWPDataExtractor,
    DocumentFiller,
    ImageCache,
    MoldHistoryCard,
)
from src.constants import (
    HWPConstants,
    DocumentConstants,
    PathConstants,
    HWPFormConstants,
)


# ===========================================================================
# HWPTextExtractor
# ===========================================================================
class TestHWPTextExtractor:

    def _make_para_text_record(self, text: str) -> bytes:
        """단순 PARA_TEXT 레코드 바이트를 생성 (tag=67, UTF-16LE)"""
        payload = text.encode("utf-16-le")
        size = len(payload)
        # tag=67(0x43), level=0, size (20비트)
        hdr = (67 & 0x3FF) | (0 << 10) | ((size & 0xFFF) << 20)
        import struct
        return struct.pack('<I', hdr) + payload

    def test_extract_single_text(self):
        data = self._make_para_text_record("Hello")
        extractor = HWPTextExtractor(data)
        result = extractor.extract_all_text()
        assert result == ["Hello"]

    def test_extract_multiple_records(self):
        data = b""
        for word in ["One", "Two", "Three"]:
            data += self._make_para_text_record(word)
        extractor = HWPTextExtractor(data)
        result = extractor.extract_all_text()
        assert result == ["One", "Two", "Three"]

    def test_skip_empty_text(self):
        data = self._make_para_text_record("   ")
        extractor = HWPTextExtractor(data)
        result = extractor.extract_all_text()
        assert result == []

    def test_empty_data(self):
        extractor = HWPTextExtractor(b"")
        assert extractor.extract_all_text() == []

    def test_para_text_tag_constant(self):
        assert HWPConstants.PARA_TEXT_TAG == 0x43 == 67


# ===========================================================================
# HWPFieldExtractor
# ===========================================================================
class TestHWPFieldExtractor:
    """폼 텍스트 파싱 로직 단위 테스트"""

    def _row(self):
        return {chr(65 + i): '' for i in range(HWPConstants.FIELD_COUNT)}

    # ----- Pattern 1: 동일 라인 "Label : Value" -----

    def test_pattern1_storage_company(self):
        texts = ["保管会社名: ACME Corp"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['B'] == "ACME Corp"

    def test_pattern1_created_date(self):
        texts = ["作成日子: 2024.03.15"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['C'] == "2024.03.15"

    def test_pattern1_management_no_inline(self):
        texts = ["管理番号: 26-001"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['D'] == "26-001"

    def test_pattern1_mold_price(self):
        texts = ["金型価: 1500000"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['U'] == "1500000"

    # ----- Pattern 2: 다음 줄 값 -----

    def test_pattern2_management_no_nextline(self):
        texts = ["管理番号", "26-005", "分 類"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['D'] == "26-005"

    def test_pattern2_classification(self):
        texts = ["分 類", "プレス型"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['E'] == "プレス型"

    def test_pattern2_classification_alt_label(self):
        texts = ["分類", "射出型"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['E'] == "射出型"

    def test_pattern2_product_name(self):
        texts = ["品 名", "VFD HOLDER"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['J'] == "VFD HOLDER"

    def test_pattern2_drawing_no(self):
        texts = ["図番番号", "1072001450"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['K'] == "1072001450"

    def test_pattern2_skips_boundary(self):
        """경계 라벨이 나오면 필드값을 비워둔다"""
        texts = ["分 類", "管理番号"]  # 바로 다음이 경계 라벨
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['E'] == ""

    # ----- GATE 型式 / 使用機械 (break on boundary) -----

    def test_gate_type_extracted(self):
        texts = ["GATE 型式", "サイドゲート"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['Q'] == "サイドゲート"

    def test_gate_type_empty_when_boundary(self):
        texts = ["GATE 型式", "使用機械"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['Q'] == ""

    def test_machine_extracted(self):
        texts = ["使用機械", "日精NS-350"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['R'] == "日精NS-350"

    # ----- 承認日 (年 포함 검사) -----

    def test_approval_date_with_year(self):
        texts = ["承認日", "2024年3月"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['T'] == "2024年3月"

    def test_approval_date_without_year_skipped(self):
        texts = ["承認日", "2024.03.15"]  # 年 없음
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['T'] == ""

    # ----- 金型規格 -----

    def test_mold_size_pattern(self):
        texts = ["金型規格", "650x450x350"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['L'] == "650x450x350"

    def test_mold_size_no_match(self):
        texts = ["金型規格", "不明"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['L'] == ""

    # ----- 金型材質 (BASE/CORE) -----

    def test_mold_material_base_core(self):
        texts = ["金型材質", "BASE", "S55C", "CORE", "SKD11"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['M'] == "S55C"
        assert row['N'] == "SKD11"

    def test_mold_material_base_only(self):
        texts = ["金型材質", "BASE", "S45C"]
        row = self._row()
        HWPFieldExtractor(texts).extract(row)
        assert row['M'] == "S45C"
        assert row['N'] == ""


# ===========================================================================
# DocumentFiller 헬퍼 메서드
# ===========================================================================
class TestDocumentFiller:

    def test_normalize_whitespace(self):
        assert DocumentFiller.normalize("品  名") == "品名"
        assert DocumentFiller.normalize("  ABC  ") == "abc"

    def test_normalize_none(self):
        assert DocumentFiller.normalize(None) == ""

    def test_row_norm_map(self):
        row = {"品 名": "TestPart", "図番番号": "12345"}
        norm = DocumentFiller.row_norm_map(row)
        assert "品名" in norm
        assert norm["品名"] == "TestPart"

    def test_value_by_aliases_found(self):
        norm_row = {"품명": "HOLDER", "管理番号": "26-001"}
        val = DocumentFiller.value_by_aliases(norm_row, ["品 名", "품명"])
        assert val == "HOLDER"

    def test_value_by_aliases_not_found(self):
        norm_row = {}
        val = DocumentFiller.value_by_aliases(norm_row, ["品 名", "품명"])
        assert val == ""

    def test_replace_placeholders_basic(self):
        norm_row = {"관리번호": "26-001"}
        text = "{管理番号}"
        # 직접 치환 확인 (alias 없는 단순 케이스)
        replaced, cnt, labels = DocumentFiller.replace_placeholders(text, {"관리번호": "26-001"})
        assert cnt == 1

    def test_replace_placeholders_image_token(self):
        norm_row = {}
        text = "{金型写真}"
        replaced, cnt, labels = DocumentFiller.replace_placeholders(text, norm_row)
        assert replaced == DocumentConstants.IMAGE_TOKEN
        assert "金型写真" in labels

    def test_pick_output_name_serial(self):
        row = {"File name": "26-003"}
        name = DocumentFiller.pick_output_name(row, 1)
        assert name == "26-003"

    def test_pick_output_name_fallback(self):
        row = {"File name": ""}
        name = DocumentFiller.pick_output_name(row, 5)
        assert name == "row_005"

    def test_checkbox_value_true(self):
        norm_row = DocumentFiller.row_norm_map({"新作": "1"})
        val = DocumentFiller.value_for_label("新作", norm_row)
        assert val == "O"

    def test_checkbox_value_false(self):
        norm_row = DocumentFiller.row_norm_map({"新作": "0"})
        val = DocumentFiller.value_for_label("新作", norm_row)
        assert val == ""

    def test_find_image_exact_match(self, tmp_path):
        img = tmp_path / "VFD_HOLDER.png"
        img.write_bytes(b"fake")
        result = DocumentFiller.find_image_for_output(tmp_path, "VFD_HOLDER")
        assert result == img

    def test_find_image_not_found(self, tmp_path):
        result = DocumentFiller.find_image_for_output(tmp_path, "NONEXISTENT")
        assert result is None

    def test_find_image_glob_match(self, tmp_path):
        img = tmp_path / "VFD HOLDER_1072001450.png"
        img.write_bytes(b"fake")
        result = DocumentFiller.find_image_for_output(tmp_path, "1072001450")
        assert result == img


# ===========================================================================
# MoldHistoryCard 데이터 모델 및 검증
# ===========================================================================
class TestMoldHistoryCard:

    def _valid_dict(self):
        return {
            "File name": "26-001",
            "管理番号": "M001",
            "品 名": "TEST PART",
            "図番番号": "1234567890",
        }

    def test_from_dict_valid(self):
        card = MoldHistoryCard.from_dict(self._valid_dict())
        assert card.file_name == "26-001"
        assert card.management_no == "M001"
        assert card.product_name == "TEST PART"
        assert card.drawing_no == "1234567890"

    def test_validate_no_errors(self):
        card = MoldHistoryCard.from_dict(self._valid_dict())
        assert card.validate() == []

    def test_validate_missing_management_no(self):
        d = self._valid_dict()
        d["管理番号"] = ""
        errors = MoldHistoryCard.from_dict(d).validate()
        assert any("管理番号" in e for e in errors)

    def test_validate_missing_product_name(self):
        d = self._valid_dict()
        d["品 名"] = ""
        errors = MoldHistoryCard.from_dict(d).validate()
        assert len(errors) > 0

    def test_validate_missing_drawing_no(self):
        d = self._valid_dict()
        d["図番番号"] = ""
        errors = MoldHistoryCard.from_dict(d).validate()
        assert len(errors) > 0

    def test_validate_invalid_file_name_format(self):
        d = self._valid_dict()
        d["File name"] = "invalid_name"
        errors = MoldHistoryCard.from_dict(d).validate()
        assert any("파일명" in e for e in errors)

    def test_validate_invalid_created_date(self):
        d = self._valid_dict()
        d["作成日子"] = "2024/01/15"  # 잘못된 형식
        errors = MoldHistoryCard.from_dict(d).validate()
        assert any("작성일자" in e for e in errors)

    def test_validate_valid_created_date(self):
        d = self._valid_dict()
        d["作成日子"] = "2024.01.15"
        errors = MoldHistoryCard.from_dict(d).validate()
        assert errors == []

    def test_checkbox_true(self):
        d = self._valid_dict()
        d["新作"] = "1"
        d["二元化"] = "1"
        card = MoldHistoryCard.from_dict(d)
        assert card.is_new is True
        assert card.is_dual is True
        assert card.is_increase is False

    def test_to_dict_roundtrip(self):
        d = self._valid_dict()
        d["新作"] = "1"
        card = MoldHistoryCard.from_dict(d)
        result = card.to_dict()
        assert result["File name"] == "26-001"
        assert result["新作"] == "1"
        assert result["増作"] == ""

    def test_from_dict_optional_fields_none(self):
        card = MoldHistoryCard.from_dict(self._valid_dict())
        assert card.storage_company is None
        assert card.classification is None


# ===========================================================================
# constants 모듈 무결성 확인
# ===========================================================================
class TestConstants:

    def test_canonical_headers_count(self):
        assert len(DocumentConstants.CANONICAL_HEADERS) == HWPConstants.FIELD_COUNT

    def test_image_exts_have_dots(self):
        for ext in DocumentConstants.IMAGE_EXTS:
            assert ext.startswith(".")

    def test_boundary_labels_nonempty(self):
        assert len(HWPFormConstants.BOUNDARY_LABELS) > 10

    def test_nextline_field_map_keys_in_boundary(self):
        """NEXTLINE_FIELD_MAP의 모든 키는 BOUNDARY_LABELS에 포함되어야 한다"""
        for label in HWPFormConstants.NEXTLINE_FIELD_MAP:
            assert label in HWPFormConstants.BOUNDARY_LABELS, f"{label!r} not in BOUNDARY_LABELS"

    def test_image_field_index(self):
        # A=0, B=1, ... Z=25, AA=26, AB=27, AC=28
        # chr(65 + 28) = chr(93) = ']' — 이건 단일문자 인코딩이 아님
        # 프로그램에서는 chr(65+28) = '[' 방식이 아니라 row key로 사용
        # IMAGE_FIELD_INDEX == 28 이어야 "AC" 컬럼 (0-indexed 28번째)
        assert HWPFormConstants.IMAGE_FIELD_INDEX == 28
        # 실제로 사용하는 방식: chr(65 + 28) → 프로그램 내부에서 row 키로 사용
        assert chr(65 + HWPFormConstants.IMAGE_FIELD_INDEX) == chr(65 + 28)

    def test_path_constants(self):
        assert PathConstants.DATA_DIR == ".data"
        assert PathConstants.MANIFEST_FILE == "manifest.json"


# ===========================================================================
# ImageCache
# ===========================================================================
class TestImageCache:

    def test_empty_dir(self, tmp_path):
        cache = ImageCache(tmp_path)
        assert len(cache) == 0
        assert cache.find("anything") is None

    def test_nonexistent_dir(self, tmp_path):
        cache = ImageCache(tmp_path / "no_such_dir")
        assert cache.find("x") is None

    def test_exact_match(self, tmp_path):
        img = tmp_path / "VFD_HOLDER.png"
        img.write_bytes(b"fake")
        cache = ImageCache(tmp_path)
        assert cache.find("VFD_HOLDER") == img

    def test_sanitized_stem_match(self, tmp_path):
        # 파일명의 특수문자가 '_'로 치환된 경우
        img = tmp_path / "VFD_HOLDER_1072001450.png"
        img.write_bytes(b"fake")
        cache = ImageCache(tmp_path)
        # 슬래시 포함 도번 → find() 내부에서 sanitize('/'→'_') 후 인덱스 매칭
        assert cache.find("VFD/HOLDER/1072001450") == img
        # 이미 sanitize된 버전으로도 동일하게 조회됨
        assert cache.find("VFD_HOLDER_1072001450") == img

    def test_numbered_suffix_ignored(self, tmp_path):
        # _2 붙은 파일이 있을 때 base stem으로 찾기
        img2 = tmp_path / "PART_2.png"
        img2.write_bytes(b"fake")
        cache = ImageCache(tmp_path)
        result = cache.find("PART")
        assert result == img2

    def test_unnumbered_preferred_over_numbered(self, tmp_path):
        img_base = tmp_path / "PART.png"
        img_num = tmp_path / "PART_2.png"
        img_base.write_bytes(b"base")
        img_num.write_bytes(b"numbered")
        cache = ImageCache(tmp_path)
        assert cache.find("PART") == img_base

    def test_multiple_extensions(self, tmp_path):
        jpg = tmp_path / "PART.jpg"
        jpg.write_bytes(b"fake")
        cache = ImageCache(tmp_path)
        assert cache.find("PART") == jpg

    def test_non_image_ignored(self, tmp_path):
        txt = tmp_path / "readme.txt"
        txt.write_bytes(b"text")
        cache = ImageCache(tmp_path)
        assert cache.find("readme") is None

    def test_invalidate_rebuilds(self, tmp_path):
        cache = ImageCache(tmp_path)
        assert cache.find("NEW") is None

        new_img = tmp_path / "NEW.png"
        new_img.write_bytes(b"fake")
        cache.invalidate()
        assert cache.find("NEW") == new_img

    def test_trailing_underscore_stripped(self, tmp_path):
        img = tmp_path / "MR14_CASE.png"
        img.write_bytes(b"fake")
        cache = ImageCache(tmp_path)
        # trailing '_' 을 rstrip한 키로 검색
        assert cache.find("MR14_CASE_") == img
