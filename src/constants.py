"""
상수 정의 모듈 — 금형이력카드 프로그램
매직 넘버, 하드코딩 문자열, 필드 정의 등을 한 곳에서 관리
"""
from typing import Dict, List, Set


# ============================================================================
# HWP 파일 포맷 상수
# ============================================================================
class HWPConstants:
    """HWP5 파일 포맷 관련 상수"""
    PARA_TEXT_TAG: int = 0x43          # 67: PARA_TEXT 레코드 태그
    FIELD_COUNT: int = 29              # 추출 필드 수 (A~AC)
    SECTION_STREAM_KEYWORD = "BodyText"
    SECTION_STREAM_NAME = "Section0"

    # 이미지 스트림 식별
    BIN_DATA_KEYWORD = "BinData"

    # 이미지 매직 바이트
    IMAGE_MAGIC_BYTES: Dict[bytes, str] = {
        b'\xFF\xD8\xFF': 'jpg',
        b'\x89PNG\r\n\x1a\n': 'png',
        b'BM': 'bmp',
        b'GIF8': 'gif',
    }


# ============================================================================
# 문서 처리 상수
# ============================================================================
class DocumentConstants:
    """DOCX/XLSX 문서 처리 관련 상수"""

    # 이미지 토큰 (Word 템플릿에서 이미지 삽입 위치 표시)
    IMAGE_TOKEN = "{金型写真}"

    # 이미지 삽입 크기 (cm)
    IMAGE_WIDTH_CM: float = 18.5
    IMAGE_HEIGHT_CM: float = 13.87

    # 지원 이미지 확장자 (우선순위 순)
    IMAGE_EXTS: List[str] = [
        ".png", ".jpg", ".jpeg", ".bmp",
        ".gif", ".tif", ".tiff", ".webp",
    ]

    # 이미지 변환 포맷 (PDF → JPG)
    IMAGE_CONVERT_FMT = "JPEG"
    IMAGE_CONVERT_QUALITY: int = 95

    # 체크박스 필드 (값이 "1"이면 "O", 아니면 "")
    CHECKBOX_LABELS: Set[str] = {"新作", "増作", "二元化", "業者更変", "機種更変"}

    # 금형규격 패턴 (예: 650x650x650)
    SIZE_PATTERN = r'\d{1,5}x\d{1,5}x\d{1,5}'

    # 파일명 불가 문자 패턴
    INVALID_FILENAME_CHARS = r'[<>:"/\\|?*]'
    INVALID_FILENAME_REPLACEMENT = '_'

    # XLSX 헤더 (HWP → XLSX 변환 시 사용)
    CANONICAL_HEADERS: List[str] = [
        "File name", "保管会社名", "作成日子", "管理番号", "分 類",
        "現 保管処", "製作処", "MODEL名", "量産処", "品 名",
        "図番番号", "金型規格", "金型材質-BASE", "金型材質-CORE", "CAVITY 数",
        "金型寿命", "GATE 型式", "使用機械", "契約日", "承認日",
        "金型価", "新作", "増作", "二元化", "業者変更",
        "仕様変更", "修理内訳", "事 由", "金型写真",
    ]

    # XLSX 워크시트 이름
    XLSX_SHEET_NAME = "data"

    # XLSX "金型写真" 컬럼 검색 범위
    XLSX_MAX_COL_SEARCH: int = 40

    # 필드 별칭 매핑 (플레이스홀더 치환에 사용)
    ALIASES: Dict[str, List[str]] = {
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

    # BASE/CORE 글자 크기 (금형재질 필드)
    MOLD_MATERIAL_FONT_SIZE_PT: int = 6

    # 이미지 검색 시 제외할 연번 접미사
    IMAGE_NUMBERED_SUFFIXES = ('_2', '_3', '_4', '_5')

    # pick_output_name 기본값 포맷
    OUTPUT_NAME_FALLBACK = "row_{:03d}"


# ============================================================================
# 경로 관련 상수
# ============================================================================
class PathConstants:
    """파일 경로 및 디렉터리 관련 상수"""
    DATA_DIR = ".data"
    HISTORY_SUFFIX = "_history.txt"
    MANIFEST_FILE = "manifest.json"
    DEFAULT_IMG_DIR = "img"
    DEFAULT_OUTPUT_DIR = "output"


# ============================================================================
# HWP 폼 필드 파싱 상수
# ============================================================================
class HWPFormConstants:
    """HWP 폼 필드 파싱에 사용되는 레이블/패턴 상수"""

    # 폼 시작 감지 레이블
    FORM_START_LABEL = '保管会社名'

    # 필드 탐색 최대 거리 (다음 몇 개 항목까지 볼지)
    FIELD_LOOKAHEAD: int = 5
    MATERIAL_LOOKAHEAD: int = 15

    # 필드 경계 레이블 (이 레이블이 나오면 현재 필드 탐색 중단)
    BOUNDARY_LABELS: Set[str] = {
        '保管会社名', '作成日子', '管理番号',
        '分 類', '分類', '現 保管処', '現保管処', '製作処', 'MODEL名', 'MODEL',
        '量産処', '品 名', '品名', '図番番号', '図番', '金型規格', '規格',
        '金型材質', 'BASE', 'CORE', 'CAVITY 数', 'CAVITY', '金型寿命',
        'GATE 型式', 'GATE', '使用機械', '機械', '契約日', '承認日',
    }

    # Pattern 1: "레이블 : 값" 형식 필드 매핑 (레이블 → row 키)
    INLINE_FIELD_MAP: Dict[str, str] = {
        '保管会社名': 'B',
        '作成日子': 'C',
        '管理番号': 'D',
        '金型価': 'U',
        '金型価格': 'U',
    }

    # Pattern 2: 다음 줄 값 형식 필드 매핑 (레이블 목록 → row 키)
    NEXTLINE_FIELD_MAP: Dict[str, str] = {
        '管理番号': 'D',
        '分 類': 'E',
        '分類': 'E',
        '現 保管処': 'F',
        '現保管処': 'F',
        '製作処': 'G',
        'MODEL名': 'H',
        'MODEL': 'H',
        '量産処': 'I',
        '品 名': 'J',
        '品名': 'J',
        '図番番号': 'K',
        '図番': 'K',
        'CAVITY 数': 'O',
        'CAVITY': 'O',
        '金型寿命': 'P',
    }

    # GATE/機械/契約日 — 경계 라벨에서 break 처리하는 필드
    BREAK_ON_BOUNDARY_FIELD_MAP: Dict[str, str] = {
        'GATE 型式': 'Q',
        'GATE': 'Q',
        '使用機械': 'R',
        '機械': 'R',
        '契約日': 'S',
    }

    # 承認日 — 연도(年) 포함 값만 취하는 필드
    APPROVAL_DATE_LABEL = '承認日'
    APPROVAL_DATE_KEY = 'T'
    APPROVAL_DATE_MARKER = '年'

    # 金型材質 관련
    MOLD_MATERIAL_LABEL = '金型材質'
    MOLD_MATERIAL_BASE_KEY = 'M'
    MOLD_MATERIAL_CORE_KEY = 'N'

    # 金型規格 관련
    MOLD_SIZE_LABEL = '金型規格'
    MOLD_SIZE_ALT_LABEL = '規格'
    MOLD_SIZE_KEY = 'L'
    MOLD_SIZE_LOOKAHEAD: int = 6

    # 금형사진 이미지 파일명 컬럼 (0-indexed: 28 → chr(65+28) = AC)
    IMAGE_FIELD_INDEX: int = 28
