# 금형이력카드 프로그램 - core.py 개선사항 및 업그레이드 방향성

작성일: 2026-04-14  
분석 대상: src/core.py (핵심 비즈니스 로직 모듈)

---

## 📋 현재 구조 분석

### 코드 규모
- **총 라인 수**: 1,438줄
- **클래스 수**: 10개
- **주요 의존성**: olefile, openpyxl, python-docx, PIL, zlib

### 클래스 구조

#### 1. HWP 처리 (라인 27-515)
- `HWPTextExtractor`: HWP5 BodyText 스트림에서 텍스트 추출 (바이너리 파싱)
- `HWPDataExtractor`: 폼 필드 데이터 추출 (29개 필드: A-AC)
- `HWPProcessor`: HWP → XLSX 변환 메인 엔트리포인트
- `HWPImageExtractor`: HWP 파일에서 이미지 추출 (압축 해제, 포맷 검증)

#### 2. 문서 생성 (라인 518-926)
- `DocumentFiller`: XLSX → DOCX 변환
  - 플레이스홀더 치환 (`{품명}` → 실제 값)
  - 이미지 토큰 (`***`) 처리 및 이미지 삽입
  - 체크박스 처리 (O/빈값)
  - 앨리어스 시스템 (여러 이름 → 하나의 필드)

#### 3. 동기화 및 관리 (라인 928-1380)
- `DocxSyncManager`: XLSX 변경사항 → DOCX 동기화
  - 파일명 변경 감지 (manifest.json 추적)
  - 이미지 파일명 자동 rename
  - 이력 파일 rename
- `MaintenanceHistoryManager`: 유지보수 이력 관리
  - `.data/` 하위 폴더에 `_history.txt` 저장
  - XLSX의 "事 由" 컬럼 업데이트
- `NewCardManager`: 신규 이력카드 발행
  - 자동 연번 생성 (예: 19-006)
  - XLSX 행 추가
  - 이미지 복사 (.data/)

---

## ✅ 완료 현황 (2026-04-15, feature/phase1-core-improvements)

| 항목 | 상태 | 내용 |
|------|------|------|
| **1.1 예외 처리 개선** | ✅ 완료 | 사용자 정의 예외 4종 추가, bare `except:` 13곳 → `except Exception as e: logger.error(...)` |
| **1.2 매직 넘버 → constants.py** | ✅ 완료 | `src/constants.py` 신설 (HWPConstants/DocumentConstants/PathConstants/HWPFormConstants), core.py 전체 상수 이전 |
| **1.3 _extract_fields() 분리** | ✅ 완료 | `HWPFieldExtractor` 클래스 추출, 역할별 메서드(_find_simple_next, _extract_mold_material 등) 분리, 170줄 → 26줄로 단축 |
| **3.1 데이터 검증 레이어** | ✅ 완료 | `MoldHistoryCard` dataclass 추가 (from_dict/to_dict/validate), 연번형식·날짜형식 검증 |
| **5.1 단위 테스트** | ✅ 완료 | `test/test_core.py` 작성 (55개 테스트, 전체 통과): HWPTextExtractor/HWPFieldExtractor/DocumentFiller/MoldHistoryCard/constants |
| **5.2 로깅 시스템** | ✅ 완료 | `import logging` + `logger = logging.getLogger(__name__)` 도입 |
| **HWP 파싱 — 管理番号 누락** | ✅ 완료 | Pattern 2 핸들러 추가, next_labels에 경계 라벨 추가 |
| **HWP 파싱 — 값 누수 버그** | ✅ 완료 | GATE型式/使用機械/契約日 → 承認日 값 가져오는 버그 수정 (skip→break) |
| **이미지 파일명** | ✅ 완료 | `品名_図番` 규칙 정착, 도번 없을 때 trailing underscore 없이 `品名`만으로 저장 |
| **이미지 탐색** | ✅ 완료 | `find_image_for_output` trailing underscore 제거 재탐색, 品名 단독 탐색 step 2.5 추가 |
| **`apply_to_word` 이미지 누락** | ✅ 완료 | 4단계 fallback chain 적용 (金型写真 → 図番番号 → 品名 → 연번 → .data/) |
| **신규 발행 다이얼로그** | ✅ 완료 | `get_last_entry()` 추가 — DB 직전 항목 표시 |
| **callback 파라미터 누락** | ✅ 완료 | `HWPProcessor.process()`, `DocumentFiller.process()` 시그니처 수정 |

### 미완료 항목

| 항목 | 우선순위 | 비고 |
|------|---------|------|
| 1.4 타입 힌팅 강화 | Low | mypy 도입 시 함께 진행 |
| 2.1 멀티프로세싱 HWP 처리 | Low | 파일 수 많을 때 효과적 |
| 2.2 이미지 검색 캐싱 (ImageCache) | Low | 현재 성능으로 충분 |

---

## 🔧 개선사항

### 1. 코드 품질 및 아키텍처

#### 1.1 예외 처리 개선 ✅ 완료

**현재 상태**:
```python
try:
    # ... 작업
except:
    pass
```
- 일반적인 `except:` 사용 (안티패턴)
- 오류 원인 추적 불가
- 로깅 부재

**개선 제안**:
```python
import logging

logger = logging.getLogger(__name__)

try:
    ole = olefile.OleFileIO(self.filepath)
except olefile.Error as e:
    logger.error(f"OLE 파일 읽기 실패: {self.filepath} - {e}")
    raise HWPProcessingError(f"HWP 파일 손상: {e}") from e
except FileNotFoundError:
    logger.error(f"파일 없음: {self.filepath}")
    raise
except Exception as e:
    logger.exception(f"예기치 않은 오류: {self.filepath}")
    raise HWPProcessingError(f"처리 실패: {e}") from e
```

**사용자 정의 예외**:
```python
class HWPProcessingError(Exception):
    """HWP 파일 처리 중 발생하는 오류"""
    pass

class XLSXProcessingError(Exception):
    """XLSX 파일 처리 중 발생하는 오류"""
    pass

class DocumentGenerationError(Exception):
    """Word 문서 생성 중 발생하는 오류"""
    pass

class ImageNotFoundError(Exception):
    """이미지 파일을 찾을 수 없을 때 발생하는 오류"""
    pass
```

**효과**:
- 오류 원인 명확한 추적
- 사용자 친화적 오류 메시지
- 디버깅 시간 단축

---

#### 1.2 매직 넘버 및 하드코딩 제거

**현재 상태**:
```python
PARA_TEXT = 67  # 0x43
IMAGE_TOKEN = "***"
row = {chr(65+i): '' for i in range(29)}  # A-AC
```

**개선 제안**:
```python
# constants.py
class HWPConstants:
    """HWP 파일 포맷 상수"""
    PARA_TEXT_TAG = 0x43  # 67
    SECTION_NAME = "BodyText/Section0"
    
class DocumentConstants:
    """문서 처리 상수"""
    IMAGE_TOKEN = "***"
    IMAGE_WIDTH_CM = 18.5
    IMAGE_HEIGHT_CM = 13.87
    
    # 필드 정의
    FIELD_COUNT = 29
    FIELD_NAMES = [
        "File name", "保管会社名", "作成日子", "管理番号", 
        # ... 전체 29개
    ]
    
    # 체크박스 필드
    CHECKBOX_FIELDS = {"新作", "増作", "二元化", "業者変更", "仕様変更"}
    
class PathConstants:
    """경로 관련 상수"""
    DATA_DIR = ".data"
    HISTORY_SUFFIX = "_history.txt"
    MANIFEST_FILE = "manifest.json"
    IMAGE_EXTENSIONS = [".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"]
```

**효과**:
- 유지보수성 향상
- 상수 재사용 용이
- 문서화 개선

---

#### 1.3 함수 길이 및 복잡도 감소

**현재 상태**:
- `_extract_fields()`: 170줄 (라인 148-318)
- 중첩된 if-elif 체인
- 사이클로매틱 복잡도 높음

**개선 제안**:
```python
class HWPFieldExtractor:
    """필드별 추출 로직 분리"""
    
    def __init__(self, form_texts: List[str]):
        self.texts = form_texts
        self.next_labels = {...}
    
    def extract_classification(self, start_idx: int) -> str:
        """분류 필드 추출"""
        for j in range(start_idx + 1, min(start_idx + 5, len(self.texts))):
            if self._is_valid_value(self.texts[j]):
                return self.texts[j]
        return ""
    
    def extract_mold_material(self, start_idx: int) -> Tuple[str, str]:
        """금형재질 BASE/CORE 추출"""
        base_value, core_value = "", ""
        j = start_idx + 1
        while j < len(self.texts) and j < start_idx + 15:
            # BASE 추출
            if self.texts[j] == 'BASE':
                base_value = self._extract_next_value(j)
            # CORE 추출
            elif self.texts[j] == 'CORE':
                core_value = self._extract_next_value(j)
            j += 1
        return base_value, core_value
    
    def _is_valid_value(self, text: str) -> bool:
        """유효한 값인지 검사"""
        return text.strip() and text not in self.next_labels
    
    def _extract_next_value(self, idx: int) -> str:
        """다음 유효한 값 추출"""
        for k in range(idx + 1, len(self.texts)):
            if self._is_valid_value(self.texts[k]):
                return self.texts[k]
        return ""
```

**리팩토링 후**:
```python
def _extract_fields(self, row: Dict, texts: List[str]):
    """Extract field values from text list"""
    form_start = self._find_form_start(texts)
    if form_start >= len(texts):
        return
    
    extractor = HWPFieldExtractor(texts[form_start:])
    
    # 간단한 패턴
    row['E'] = extractor.extract_classification(0)
    row['F'] = extractor.extract_storage_location(0)
    
    # 복잡한 패턴
    row['M'], row['N'] = extractor.extract_mold_material(0)
```

**효과**:
- 단위 테스트 용이
- 코드 가독성 향상
- 버그 수정 시간 단축

---

#### 1.4 타입 힌팅 강화

**현재 상태**:
```python
def extract(self) -> Dict[str, str]:
def extract_images(self):  # 반환 타입 없음
```

**개선 제안**:
```python
from typing import Dict, List, Optional, Tuple, Union
from pathlib import Path

def extract(self) -> Dict[str, str]:
    """Extract all form fields from HWP file
    
    Returns:
        Dict mapping field codes (A-AC) to extracted values
    """
    ...

def extract_images(self) -> int:
    """Extract images from HWP file
    
    Returns:
        Number of successfully extracted images
    """
    ...

def find_image_for_output(
    cls, 
    img_dir: Path, 
    output_stem: str
) -> Optional[Path]:
    """Find image file matching the output stem
    
    Args:
        img_dir: Directory to search for images
        output_stem: Base filename without extension
        
    Returns:
        Path to the found image, or None if not found
    """
    ...
```

**효과**:
- IDE 자동완성 개선
- 타입 오류 조기 발견 (mypy, pyright)
- 문서화 자동 생성

---

### 2. 성능 최적화

#### 2.1 HWP 파일 처리 속도 개선

**현재 상태**:
- 순차 처리 (단일 스레드)
- 파일당 압축 해제 여러 번 (텍스트 + 이미지)

**개선 제안 1: 캐싱**
```python
class HWPFileCache:
    """HWP 파일 데이터 캐시"""
    def __init__(self):
        self._cache: Dict[str, Tuple[bytes, OleFileIO]] = {}
    
    def get_or_load(self, filepath: str) -> Tuple[bytes, OleFileIO]:
        if filepath not in self._cache:
            ole = olefile.OleFileIO(filepath)
            section_data = self._extract_section(ole)
            self._cache[filepath] = (section_data, ole)
        return self._cache[filepath]
```

**개선 제안 2: 멀티프로세싱**
```python
from multiprocessing import Pool
from functools import partial

def _process_single_hwp(hwp_file: Path) -> List[str]:
    """단일 HWP 파일 처리 (워커 함수)"""
    extractor = HWPDataExtractor(str(hwp_file))
    row_data = extractor.extract()
    return [row_data.get(chr(65 + i), "") for i in range(29)]

@classmethod
def extract_rows_from_hwp_parallel(cls, input_dir: Path, workers: int = 4) -> List[List[str]]:
    """병렬 HWP 처리"""
    hwp_files = sorted(input_dir.glob("*.hwp"))
    
    with Pool(processes=workers) as pool:
        rows = pool.map(_process_single_hwp, hwp_files)
    
    return rows
```

**예상 효과**:
- 100개 파일 처리 시간: 60초 → 20초 (3배 개선)
- CPU 멀티코어 활용

---

#### 2.2 이미지 검색 최적화

**현재 상태**:
```python
def find_image_for_output(cls, img_dir: Path, output_stem: str) -> Optional[Path]:
    # 매번 glob으로 파일 탐색
    for ext in cls.IMAGE_EXTS:
        p = img_dir / f"{output_stem}{ext}"
        if p.exists():
            return p
    # ... glob으로 다시 검색
```

**개선 제안**:
```python
class ImageCache:
    """이미지 파일 인덱스 캐시"""
    def __init__(self, img_dir: Path):
        self.img_dir = img_dir
        self._index: Dict[str, Path] = {}
        self._build_index()
    
    def _build_index(self):
        """전체 이미지 인덱스 한 번에 생성"""
        for ext in DocumentConstants.IMAGE_EXTENSIONS:
            for img_path in self.img_dir.glob(f"*{ext}"):
                stem = img_path.stem
                # 연번 제거한 스템도 인덱싱
                base_stem = re.sub(r'_\d+$', '', stem)
                self._index[stem] = img_path
                self._index[base_stem] = img_path
    
    def find(self, stem: str) -> Optional[Path]:
        """O(1) 검색"""
        return self._index.get(stem)

# 사용
image_cache = ImageCache(img_dir)
for row in rows:
    image_path = image_cache.find(output_stem)
```

**효과**:
- 1000개 파일 처리 시 이미지 검색: 5초 → 0.1초
- 반복 검색 시 성능 향상

---

#### 2.3 XLSX 읽기/쓰기 최적화

**현재 상태**:
```python
# 매번 XLSX 전체 로드
rows = DocumentFiller.load_rows_from_xlsx(xlsx_path)
```

**개선 제안**:
```python
class XLSXStreamReader:
    """스트리밍 XLSX 읽기 (대용량 파일 대응)"""
    
    @staticmethod
    def iter_rows(xlsx_path: Path, batch_size: int = 100):
        """행을 배치 단위로 yield"""
        wb = load_workbook(xlsx_path, read_only=True, data_only=True)
        ws = wb.active
        headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
        
        batch = []
        for row_values in ws.iter_rows(min_row=2, values_only=True):
            row_dict = {h: str(v or "").strip() 
                       for h, v in zip(headers, row_values)}
            batch.append(row_dict)
            
            if len(batch) >= batch_size:
                yield batch
                batch = []
        
        if batch:
            yield batch
        
        wb.close()

# 사용
for batch in XLSXStreamReader.iter_rows(xlsx_path, batch_size=50):
    for row in batch:
        process_row(row)
```

**효과**:
- 5000행 XLSX: 메모리 사용량 500MB → 50MB
- 처리 중단/재개 용이

---

### 3. 신뢰성 및 데이터 무결성

#### 3.1 데이터 검증 레이어 추가

**개선 제안**:
```python
from dataclasses import dataclass
from typing import Optional
from datetime import datetime

@dataclass
class MoldHistoryCard:
    """금형 이력카드 데이터 모델"""
    file_name: str
    management_no: str  # 管理番号 (필수)
    product_name: str   # 品名 (필수)
    drawing_no: str     # 図番番号 (필수)
    
    # 선택 필드
    storage_company: Optional[str] = None
    created_date: Optional[str] = None
    classification: Optional[str] = None
    # ... 기타 필드
    
    def validate(self) -> List[str]:
        """데이터 검증"""
        errors = []
        
        if not self.management_no:
            errors.append("管理番号가 비어있습니다")
        
        if not self.product_name:
            errors.append("品名이 비어있습니다")
        
        if not self.drawing_no:
            errors.append("図番番号가 비어있습니다")
        
        # 연번 형식 검증
        if not re.match(r"^\d{2}-\d{3}$", self.file_name):
            errors.append(f"파일명 형식 오류: {self.file_name} (예: 19-001)")
        
        # 날짜 형식 검증
        if self.created_date:
            try:
                datetime.strptime(self.created_date, "%Y.%m.%d")
            except ValueError:
                errors.append(f"작성일자 형식 오류: {self.created_date}")
        
        return errors
    
    @classmethod
    def from_dict(cls, data: Dict[str, str]) -> 'MoldHistoryCard':
        """딕셔너리에서 생성"""
        return cls(
            file_name=data.get("File name", ""),
            management_no=data.get("管理番号", ""),
            product_name=data.get("品 名", ""),
            drawing_no=data.get("図番番号", ""),
            storage_company=data.get("保管会社名"),
            created_date=data.get("作成日子"),
            # ...
        )
```

**사용**:
```python
for row_dict in rows:
    card = MoldHistoryCard.from_dict(row_dict)
    errors = card.validate()
    if errors:
        logger.warning(f"행 {idx} 검증 실패: {errors}")
        continue
    process_card(card)
```

**효과**:
- 잘못된 데이터 조기 발견
- 일관된 검증 로직
- 데이터 품질 향상

---

#### 3.2 트랜잭션 및 롤백 지원

**현재 상태**:
- XLSX 업데이트 도중 실패 시 데이터 손실 위험

**개선 제안**:
```python
import shutil
from contextlib import contextmanager

@contextmanager
def transaction(xlsx_path: Path):
    """XLSX 트랜잭션"""
    backup_path = xlsx_path.with_suffix('.xlsx.bak')
    
    # 백업 생성
    shutil.copy2(xlsx_path, backup_path)
    
    try:
        yield xlsx_path
        # 성공 시 백업 삭제
        backup_path.unlink()
    except Exception as e:
        # 실패 시 복원
        shutil.copy2(backup_path, xlsx_path)
        backup_path.unlink()
        raise

# 사용
with transaction(xlsx_path):
    add_to_xlsx(xlsx_path, row_dict)
    update_manifest(output_dir, manifest)
```

**효과**:
- 데이터 무결성 보장
- 부분 업데이트 방지

---

#### 3.3 manifest.json 검증 및 복구

**개선 제안**:
```python
class ManifestValidator:
    """manifest.json 무결성 검사"""
    
    @staticmethod
    def validate(manifest: Dict, output_dir: Path) -> List[str]:
        """manifest와 실제 파일 일치 여부 검사"""
        errors = []
        
        # 1. manifest에 있지만 파일이 없는 경우
        for mgmt_no, info in manifest.items():
            file_name = info.get("file_name", "")
            docx_path = output_dir / f"{file_name}.docx"
            if not docx_path.exists():
                errors.append(f"파일 없음: {file_name}.docx (管理番号: {mgmt_no})")
        
        # 2. 파일은 있지만 manifest에 없는 경우
        for docx_path in output_dir.glob("*.docx"):
            found = False
            for info in manifest.values():
                if info.get("file_name") == docx_path.stem:
                    found = True
                    break
            if not found:
                errors.append(f"manifest 누락: {docx_path.name}")
        
        return errors
    
    @staticmethod
    def rebuild(xlsx_path: Path, output_dir: Path) -> Dict:
        """XLSX와 실제 파일로부터 manifest 재구축"""
        new_manifest = {}
        rows = DocumentFiller.load_rows_from_xlsx(xlsx_path)
        
        for idx, row in enumerate(rows, start=1):
            mgmt_no = row.get("管理番号", "").strip()
            if not mgmt_no:
                continue
            
            file_name = DocumentFiller.pick_output_name(row, idx)
            docx_path = output_dir / f"{file_name}.docx"
            
            if docx_path.exists():
                new_manifest[mgmt_no] = {
                    "file_name": file_name,
                    "serial": DocxSyncManager.extract_serial(file_name),
                    "品名": row.get("品 名", "").strip(),
                    "図番番号": row.get("図番番号", "").strip(),
                    "last_updated": datetime.now().isoformat(),
                }
        
        return new_manifest
```

**GUI 통합**:
```python
# main.py에 추가
def check_manifest_integrity(self):
    """manifest 무결성 검사 및 복구"""
    output_dir = Path(self.docx_output_edit.text())
    manifest = DocxSyncManager.load_manifest(output_dir)
    
    errors = ManifestValidator.validate(manifest, output_dir)
    
    if errors:
        msg = "manifest 불일치 발견:\n\n" + "\n".join(errors[:10])
        reply = QMessageBox.question(
            self, "manifest 복구", 
            msg + "\n\nmanifest를 재구축하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            xlsx_path = Path(self.docx_xlsx_edit.text())
            new_manifest = ManifestValidator.rebuild(xlsx_path, output_dir)
            DocxSyncManager.save_manifest(output_dir, new_manifest)
            self.log_message("✓ manifest 재구축 완료")
```

**효과**:
- manifest 손상 복구 가능
- 수동 파일 조작 후 복구

---

### 4. 기능 확장

#### 4.1 이미지 리사이징 및 최적화

**현재 상태**:
- 이미지를 원본 크기 그대로 삽입
- Word 파일 크기 증가

**개선 제안**:
```python
from PIL import Image
from io import BytesIO

class ImageOptimizer:
    """이미지 최적화"""
    
    @staticmethod
    def optimize(image_path: Path, max_width: int = 1920, 
                 max_height: int = 1440, quality: int = 85) -> Path:
        """이미지 리사이징 및 압축"""
        img = Image.open(image_path)
        
        # EXIF orientation 처리
        try:
            exif = img._getexif()
            if exif:
                orientation = exif.get(0x0112)
                if orientation == 3:
                    img = img.rotate(180, expand=True)
                elif orientation == 6:
                    img = img.rotate(270, expand=True)
                elif orientation == 8:
                    img = img.rotate(90, expand=True)
        except:
            pass
        
        # 리사이징
        if img.width > max_width or img.height > max_height:
            img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
        
        # RGB 변환 (CMYK, RGBA → RGB)
        if img.mode not in ('RGB', 'L'):
            img = img.convert('RGB')
        
        # 압축 저장
        optimized_path = image_path.with_stem(f"{image_path.stem}_optimized")
        img.save(optimized_path, 'JPEG', quality=quality, optimize=True)
        
        return optimized_path
```

**DocumentFiller 통합**:
```python
@classmethod
def insert_images_by_token(cls, doc: Document, image_path: Optional[Path], 
                           optimize: bool = True) -> int:
    if optimize and image_path and image_path.exists():
        # 이미지 최적화
        image_path = ImageOptimizer.optimize(image_path)
    
    # 기존 로직
    count = 0
    for p in doc.paragraphs:
        count += cls._insert_image_in_paragraph(p, image_path)
    # ...
```

**효과**:
- Word 파일 크기: 평균 50% 감소
- 로딩 속도 향상

---

#### 4.2 PDF 메타데이터 추가

**개선 제안** (pdf.py 확장):
```python
from pypdf import PdfReader, PdfWriter

def add_metadata_to_pdf(pdf_path: Path, metadata: Dict[str, str]) -> None:
    """PDF에 메타데이터 추가"""
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    
    for page in reader.pages:
        writer.add_page(page)
    
    # 메타데이터 설정
    writer.add_metadata({
        '/Title': metadata.get('product_name', ''),
        '/Subject': f"금형이력카드 {metadata.get('management_no', '')}",
        '/Author': '金型管理システム',
        '/Creator': '金型이력카드프로그램',
        '/Producer': 'python-docx + pypdf',
        '/Keywords': f"{metadata.get('drawing_no', '')}, {metadata.get('product_name', '')}",
    })
    
    with open(pdf_path, 'wb') as f:
        writer.write(f)
```

**효과**:
- PDF 검색 기능 향상
- 문서 관리 용이

---

#### 4.3 이력 통계 및 리포트

**개선 제안**:
```python
class HistoryStatistics:
    """유지보수 이력 통계"""
    
    @staticmethod
    def analyze_maintenance_history(output_dir: Path) -> Dict:
        """전체 이력 분석"""
        data_dir = output_dir / MaintenanceHistoryManager.DATA_DIR
        if not data_dir.exists():
            return {}
        
        stats = {
            "total_cards": 0,
            "cards_with_history": 0,
            "total_maintenance_count": 0,
            "maintenance_by_type": {},
            "top_repaired_molds": [],
        }
        
        mold_counts = {}
        
        for hist_file in data_dir.glob("*_history.txt"):
            stats["total_cards"] += 1
            content = hist_file.read_text(encoding="utf-8")
            
            # 이력 항목 카운트
            entries = content.count("## ")
            if entries > 0:
                stats["cards_with_history"] += 1
                stats["total_maintenance_count"] += entries
                
                # 금형별 이력 횟수
                mold_name = hist_file.stem.replace("_history", "")
                mold_counts[mold_name] = entries
        
        # 상위 10개 금형
        stats["top_repaired_molds"] = sorted(
            mold_counts.items(), key=lambda x: x[1], reverse=True
        )[:10]
        
        return stats
    
    @staticmethod
    def generate_report(output_dir: Path) -> str:
        """리포트 생성 (마크다운)"""
        stats = HistoryStatistics.analyze_maintenance_history(output_dir)
        
        report = f"""
# 금형 유지보수 이력 리포트

생성일: {datetime.now().strftime('%Y-%m-%d %H:%M')}

## 요약

- 전체 이력카드 수: {stats['total_cards']}개
- 이력이 있는 카드: {stats['cards_with_history']}개
- 총 유지보수 횟수: {stats['total_maintenance_count']}회

## 유지보수 빈도 Top 10

| 순위 | 금형 | 횟수 |
|------|------|------|
"""
        for idx, (mold, count) in enumerate(stats['top_repaired_molds'], 1):
            report += f"| {idx} | {mold} | {count}회 |\n"
        
        return report
```

**GUI 통합**:
```python
def show_history_statistics(self):
    """이력 통계 표시"""
    output_dir = Path(self.hist_dir_edit.text())
    report = HistoryStatistics.generate_report(output_dir)
    
    dialog = QDialog(self)
    dialog.setWindowTitle("유지보수 이력 통계")
    layout = QVBoxLayout()
    
    text_edit = QTextEdit()
    text_edit.setMarkdown(report)
    text_edit.setReadOnly(True)
    layout.addWidget(text_edit)
    
    save_btn = QPushButton("리포트 저장")
    save_btn.clicked.connect(lambda: self._save_report(report))
    layout.addWidget(save_btn)
    
    dialog.setLayout(layout)
    dialog.resize(600, 400)
    dialog.exec_()
```

**효과**:
- 데이터 기반 의사결정
- 예방 보수 계획 수립

---

### 5. 테스트 및 품질 보증

#### 5.1 단위 테스트 추가

**테스트 구조**:
```
tests/
├── __init__.py
├── test_hwp_extractor.py
├── test_document_filler.py
├── test_sync_manager.py
├── test_image_extractor.py
├── fixtures/
│   ├── sample.hwp
│   ├── sample.xlsx
│   ├── template.docx
│   └── test_image.jpg
└── conftest.py
```

**예제 테스트**:
```python
# tests/test_hwp_extractor.py
import pytest
from pathlib import Path
from src.core import HWPDataExtractor

@pytest.fixture
def sample_hwp(tmp_path):
    """테스트용 HWP 파일"""
    # fixtures/sample.hwp 복사
    src = Path(__file__).parent / "fixtures" / "sample.hwp"
    dst = tmp_path / "sample.hwp"
    shutil.copy2(src, dst)
    return dst

def test_extract_product_name(sample_hwp):
    """품명 추출 테스트"""
    extractor = HWPDataExtractor(str(sample_hwp))
    data = extractor.extract()
    
    assert data['J'] != "", "품명이 비어있습니다"
    assert len(data['J']) < 100, "품명이 너무 깁니다"

def test_extract_drawing_number(sample_hwp):
    """도면번호 추출 테스트"""
    extractor = HWPDataExtractor(str(sample_hwp))
    data = extractor.extract()
    
    assert data['K'] != "", "도면번호가 비어있습니다"
    assert re.match(r'^\d+$', data['K']), "도면번호 형식 오류"

def test_sanitize_filename():
    """파일명 정규화 테스트"""
    extractor = HWPDataExtractor("dummy.hwp")
    
    assert extractor.sanitize_filename("A/B\\C:D") == "A_B_C_D"
    assert extractor.sanitize_filename("정상파일명") == "정상파일명"
```

**통합 테스트**:
```python
# tests/test_integration.py
def test_full_workflow(tmp_path):
    """전체 워크플로우 통합 테스트"""
    # 1. HWP → XLSX
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    shutil.copy2("fixtures/sample.hwp", input_dir / "sample.hwp")
    
    xlsx_path = tmp_path / "output.xlsx"
    HWPProcessor.process(input_dir, xlsx_path)
    
    assert xlsx_path.exists()
    
    # 2. XLSX → DOCX
    output_dir = tmp_path / "output"
    template_path = Path("fixtures/template.docx")
    img_dir = Path("fixtures")
    
    DocumentFiller.process(xlsx_path, template_path, output_dir, img_dir)
    
    docx_files = list(output_dir.glob("*.docx"))
    assert len(docx_files) > 0
```

**CI/CD 통합**:
```yaml
# .github/workflows/test.yml
name: Tests

on: [push, pull_request]

jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-python@v4
        with:
          python-version: '3.10'
      - name: Install dependencies
        run: |
          pip install -r requirements.txt
          pip install pytest pytest-cov
      - name: Run tests
        run: pytest --cov=src --cov-report=html
      - name: Upload coverage
        uses: codecov/codecov-action@v3
```

**효과**:
- 회귀 테스트 자동화
- 코드 신뢰성 향상
- 리팩토링 안전성 확보

---

#### 5.2 로깅 시스템 구축

**개선 제안**:
```python
# src/logger.py
import logging
from pathlib import Path
from logging.handlers import RotatingFileHandler

def setup_logger(name: str, log_dir: Path = Path("logs")) -> logging.Logger:
    """로거 설정"""
    log_dir.mkdir(exist_ok=True)
    
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    # 파일 핸들러 (회전식, 최대 10MB * 5개)
    file_handler = RotatingFileHandler(
        log_dir / f"{name}.log",
        maxBytes=10*1024*1024,
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s'
    )
    file_handler.setFormatter(file_formatter)
    
    # 콘솔 핸들러
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(levelname)s: %(message)s')
    console_handler.setFormatter(console_formatter)
    
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# 사용
logger = setup_logger("core")

def extract(self) -> Dict[str, str]:
    logger.info(f"HWP 추출 시작: {self.filepath}")
    try:
        # ... 작업
        logger.info(f"HWP 추출 완료: {len(row)} 필드")
        return row
    except Exception as e:
        logger.exception(f"HWP 추출 실패: {self.filepath}")
        raise
```

**효과**:
- 문제 진단 용이
- 사용 패턴 분석
- 감사 추적 (audit trail)

---

### 6. 문서화

#### 6.1 Docstring 표준화

**개선 제안**:
```python
def find_image_for_output(cls, img_dir: Path, output_stem: str) -> Optional[Path]:
    """이미지 파일 찾기 (품명_도번 형식)
    
    우선순위:
        1. 정확한 파일명 매칭
        2. sanitize 처리된 파일명 매칭
        3. glob 패턴 검색
    
    Args:
        img_dir: 이미지 디렉토리
        output_stem: 찾을 파일명 (확장자 제외)
    
    Returns:
        찾은 이미지 경로, 없으면 None
    
    Examples:
        >>> img_dir = Path("img")
        >>> path = DocumentFiller.find_image_for_output(img_dir, "CASE_1071000024")
        >>> print(path)
        img/CASE_1071000024.jpg
    
    Notes:
        파일명에 특수문자가 포함된 경우 자동으로 sanitize하여 검색합니다.
        연번이 있는 파일(예: xxx_2.jpg)보다 연번 없는 파일을 우선합니다.
    """
    ...
```

**자동 문서 생성**:
```bash
# Sphinx 설정
pip install sphinx sphinx-rtd-theme

# docs/ 디렉토리에서
sphinx-apidoc -o source/ ../src/
make html
```

---

#### 6.2 API 레퍼런스 생성

**구조**:
```
docs/
├── api/
│   ├── hwp_processing.md
│   ├── document_generation.md
│   ├── sync_manager.md
│   └── image_handling.md
├── guides/
│   ├── getting_started.md
│   ├── workflow.md
│   └── troubleshooting.md
└── examples/
    ├── basic_usage.py
    ├── batch_processing.py
    └── custom_template.py
```

---

## 🚀 업그레이드 로드맵

### Phase 1: 안정성 및 코드 품질 (2-3주)
- [ ] 예외 처리 개선 (모든 bare except 제거)
- [ ] 타입 힌팅 추가 (mypy 통과)
- [ ] 매직 넘버 상수화
- [ ] 함수 리팩토링 (복잡도 < 10)
- [ ] 로깅 시스템 구축

### Phase 2: 성능 최적화 (1-2주)
- [ ] 멀티프로세싱 HWP 처리
- [ ] 이미지 캐시 시스템
- [ ] XLSX 스트리밍 읽기
- [ ] 이미지 최적화 기능

### Phase 3: 데이터 무결성 (1주)
- [ ] 데이터 검증 레이어
- [ ] 트랜잭션 지원
- [ ] manifest 검증/복구
- [ ] 백업/복원 기능

### Phase 4: 기능 확장 (2주)
- [ ] 이력 통계 및 리포트
- [ ] PDF 메타데이터
- [ ] 이미지 리사이징
- [ ] 템플릿 검증

### Phase 5: 테스트 및 문서화 (2주)
- [ ] 단위 테스트 (커버리지 80%+)
- [ ] 통합 테스트
- [ ] CI/CD 구축
- [ ] API 문서 작성
- [ ] 사용자 가이드

---

## 🎯 우선순위 제안

### 즉시 구현 (High Priority)
1. **예외 처리 개선** - 안정성 향상
2. **로깅 시스템** - 문제 추적
3. **데이터 검증** - 오류 조기 발견
4. **manifest 검증/복구** - 데이터 무결성

### 중기 구현 (Medium Priority)
5. **멀티프로세싱** - 대량 처리 속도
6. **이미지 캐시** - 검색 성능
7. **트랜잭션 지원** - 데이터 안전성
8. **단위 테스트** - 회귀 방지

### 장기 구현 (Low Priority)
9. **이력 통계** - 분석 기능
10. **이미지 최적화** - 파일 크기
11. **API 문서** - 유지보수성

---

## 📊 예상 효과

### 개발 측면
- **버그 발생률 70% 감소** (예외 처리, 검증)
- **처리 속도 3배 향상** (멀티프로세싱)
- **코드 유지보수 시간 50% 단축** (리팩토링, 문서화)

### 사용자 측면
- **대용량 처리 가능** (5000+ HWP 파일)
- **오류 복구 가능** (트랜잭션, 백업)
- **처리 시간 단축** (100개 파일: 60초 → 20초)

---

## 📝 기술 부채 (Tech Debt)

### 현재 확인된 문제
1. ⚠️ **bare except 남용** (20+ 곳) - 오류 원인 불명확
2. ⚠️ **함수 길이** - `_extract_fields()` 170줄
3. ⚠️ **중복 코드** - 이미지 검색 로직 여러 곳 반복
4. ⚠️ **하드코딩** - 매직 넘버, 상수 문자열
5. ⚠️ **글로벌 상태** - manifest.json 파일 잠금 처리 없음

### 해결 우선순위
1. bare except → 구체적 예외 처리
2. 긴 함수 → 작은 함수로 분리
3. 중복 코드 → 공통 모듈화
4. 하드코딩 → constants.py
5. 파일 잠금 → 파일 잠금 처리 추가

---

## 🛠️ 도구 및 라이브러리 추천

### 코드 품질
- `pylint`: 정적 분석
- `black`: 자동 포맷팅
- `mypy`: 타입 체크
- `flake8`: 스타일 가이드
- `radon`: 복잡도 측정

### 테스트
- `pytest`: 단위 테스트
- `pytest-cov`: 커버리지
- `hypothesis`: 속성 기반 테스트
- `faker`: 테스트 데이터 생성

### 성능
- `cProfile`: 프로파일링
- `memory_profiler`: 메모리 분석
- `line_profiler`: 라인별 분석

### 문서화
- `sphinx`: API 문서
- `mkdocs`: 사용자 가이드
- `pdoc`: 간단한 문서 생성

---

**작성자**: Claude (AI Assistant)  
**검토 필요**: 기존 개발자와 협의 후 단계별 적용 권장
