# 금형이력카드프로그램 — 개발 이력

> **작업 방식**: 전 구간 [Claude Code](https://claude.ai/claude-code) (Anthropic) **Plan 모드**를 활용한 AI 페어 프로그래밍으로 진행.  
> Plan 모드에서 설계 → 구현 순서를 사전에 합의한 뒤 코드 작성, 덕분에 요구사항 누락 없이 단계별로 안정적으로 진행할 수 있었다.

---

## Phase 1 — 기초 파이프라인 구축

**목표**: HWP 파일에서 데이터를 추출해 Word 문서를 자동 생성

| 순서 | 내용 |
|------|------|
| 1 | HWP → CSV 추출 시도 → 한자 헤더 깨짐 문제 발생 |
| 2 | CSV 대신 **XLSX** 기반으로 전환 (안정성 확보) |
| 3 | `hwp_to_xlsx.py` 작성: `YES/*.hwp` → `output_from_hwp.xlsx` |
| 4 | `xlsx_to_doc.py` 작성: XLSX 데이터 → DOCX 텍스트 치환 |
| 5 | 체크박스 필드(`新作`, `増作`, `二元化`, `業者更変`, `機種更変`) → 값 `1`이면 `O`, 아니면 공란 처리 |
| 6 | `{BASE}`, `{CORE}` 치환값에 폰트 6pt 적용 |

---

## Phase 2 — 이미지 삽입

**목표**: HWP 내부 금형 사진을 추출해 Word 문서에 자동 삽입

| 순서 | 내용 |
|------|------|
| 1 | `HWPImageExtractor` 구현: HWP 압축 해제 → PNG/JPG/BMP 추출 |
| 2 | `xlsx_to_doc_img.py` 작성: `{金型写真}` 위치에 이미지 삽입 |
| 3 | 이미지 크기 고정: **12.4 cm × 9.3 cm** |
| 4 | 토큰이 여러 XML run으로 분리된 경우(예: `{`, `金型`, `写真`, `}`) 처리하는 run-level 비파괴 삽입 방식 적용 |
| 5 | 이미지 매칭 규칙: `img/<연번>_<품명>_<도번>.<확장자>` |

---

## Phase 3 — 통합 GUI 및 단일 실행 파일

**목표**: 별개 스크립트들을 하나의 GUI 프로그램으로 통합

| 순서 | 내용 |
|------|------|
| 1 | 기존 스크립트들을 `src/core.py`로 통합 (클래스 기반 리팩터링) |
| 2 | PyQt5 기반 `main.py` GUI 작성 (탭 구성: HWP 처리 / 이미지 추출 / DOCX 생성) |
| 3 | PyInstaller `.spec` 파일로 EXE 빌드 구성 |
| 4 | `data/`, `YES/`, `img/` 폴더를 EXE 내부에 번들링 |

---

## Phase 4 — XLSX DB 및 이미지 파일명 체계 개선

**목표**: 이미지 관리 편의성 향상 + XLSX를 단일 DB로 운용

| 순서 | 내용 |
|------|------|
| 1 | 이미지 파일명 규칙 변경: `연번_品名_図番番号.ext` 형식으로 통일 |
| 2 | 파일명 정제 로직 추가 (특수문자/공백 처리) |
| 3 | XLSX → 이미지 자동 연동: 연번·품명·도번으로 이미지 검색 |
| 4 | `00.DB_19-000.xlsx` 파일을 마스터 DB로 지정 |

---

## Phase 5 — EXE 빌드 및 첫 배포

**목표**: 현장 배포 가능한 단독 실행 EXE 완성

| 순서 | 내용 |
|------|------|
| 1 | `금형이력카드프로그램.spec` 완성 및 PyInstaller 빌드 |
| 2 | `금형이력카드프로그램_배포/` 폴더에 EXE + 필수 폴더 구조 패키징 |
| 3 | `BUILD_GUIDE.md`, `USAGE.md` 작성 |
| 4 | 배포 후 현장 검증 완료 |

---

## v2 — 기능 확장 (Plan 모드 설계)

**설계 문서**: `C:\Users\UKY\.claude\plans\zany-dreaming-lollipop.md`  
**작업 기간**: 2026-04-13

### v2 기능 1 — XLSX-DOCX 동기화 (`DocxSyncManager`)

- `data/output/manifest.json` 도입: `管理番号 → {file_name, serial, ...}` 매핑 추적
- DOCX 탭에 "동기화" 버튼 추가
- 동기화 시 XLSX 최신 데이터로 Word 파일 재생성 + 파일명 rename 처리

### v2 기능 2 — 연번 변경 시 이미지 파일명 자동 변경

- 동기화 중 `file_name` 변경 감지 시 `img/` 폴더 이미지도 `{old_serial}_` → `{new_serial}_` 자동 rename
- `re.search(r'(\d{3})', file_name)`으로 연번 추출

### v2 기능 3 — 금형 유지보수 이력 관리 (`MaintenanceHistoryManager`)

- "이력 관리" 탭 추가 (4번째 탭)
- 이력 파일: `data/output/.data/<file_name>_history.txt` (`.md` → `.txt`로 변경)
- 저장 시 `.txt` 파일 + XLSX DB(`事 由` 열) 동시 갱신
- "Word에 반영" 버튼: 이력 전문을 `事 由` 필드에 기록 후 Word 재생성
- 이력 탭 독립 파일 설정 (XLSX/템플릿/이미지 폴더 별도 지정)

### v2 기능 4 — 신규 이력카드 발행 (`NewCardDialog` + `NewCardManager`)

- 메인 화면 상단 "★ 신규 이력카드 발행" 버튼
- 다이얼로그 구성:
  - DB 엑셀 파일 선택 (찾기 버튼 포함, 변경 시 File name 자동 재계산)
  - 필수 항목: 품명 / 도번번호 / 관리번호
  - 선택 항목: 보관회사명 / 작성일자 / **승인일** / 분류 / 제작처 / MODEL명 / 양산처 / 금형규격 / CAVITY수
  - 체크박스 이중언어 레이아웃:
    ```
    新作      増作      二元化    業者更変   機種更変
    신작      증작      이원화    업자변경   기종변경
    [ ]       [ ]       [ ]       [ ]        [ ]
    ```
  - 금형 사진 첨부 (선택)
- File name 자동 계산: XLSX 최대 연번 + 1
- 확인 시 XLSX에 신규 행 append + Word 파일 자동 생성

### v2 기능 5 — PDF 변환/병합

- "PDF 변환/병합" 탭 추가
- 3가지 모드: 단일 변환 / 일괄 변환 / 변환 후 병합
- `src/pdf.py`: `DocxToPdfConverter` 클래스 (Word COM / LibreOffice headless)
- pypdf 6.x 호환: `PdfMerger` 제거 → `PdfWriter` + `PdfReader` 방식으로 수정
- 병합 결과 파일 크기 표시

---

## v2 — 탭 순서 및 독립성 재구성

| 탭 순서 | 탭명 | 설명 |
|---------|------|------|
| 1 | 문서 생성/동기화 | DOCX 일괄 생성 + 동기화 |
| 2 | 이력 관리 | 유지보수 이력 기록·조회·Word 반영 |
| 3 | PDF 변환/병합 | DOCX → PDF 변환 및 병합 |
| 4 | HWP 처리 | HWP → XLSX 변환 |
| 5 | 이미지 추출 | HWP → 이미지 추출 |

- 각 탭이 **독립적**으로 파일 경로를 설정할 수 있도록 리팩터링
- 단계 번호 레이블 제거 (순차 강제 없애기)

---

## 버그 수정 이력

| 날짜 | 현상 | 원인 | 수정 |
|------|------|------|------|
| 2026-04-13 | 다른 AI(Qwen)가 코드 손상 | `save_history`, `run_sync` 등 핵심 메서드를 `pass`로 덮어씀 | `main.py` 전체 재작성으로 복구 |
| 2026-04-13 | PDF 병합 "pypdf 미설치" 오류 | pypdf 6.x에서 `PdfMerger` API 제거됨 | `PdfWriter`+`PdfReader` 방식으로 교체 |
| 2026-04-13 | 19-003.docx 금형 사진 누락 | 구버전 코드로 생성된 파일 | 최신 코드로 재생성 |
| 2026-04-13 | 이력 저장 시 XLSX 기존값 미삭제 | `update_xlsx_reason()`에서 셀 초기화 누락 | 쓰기 전 셀 값 `None`으로 초기화 후 갱신 |

---

## Git 저장소

```
git@github.com:planze-kee/AND_mold_history_program.git
```

- `.gitignore`로 대용량 데이터 제외: `YES/`, `img/`, `data/output/`, `data/input/`, `dist/`, `build/`
- 소스 코드(`src/`, `main.py`, `*.spec`, `requirements.txt`)만 추적

---

## 기술 스택

| 분류 | 라이브러리 |
|------|-----------|
| GUI | PyQt5 |
| Word 처리 | python-docx |
| Excel 처리 | openpyxl |
| HWP 파싱 | olefile (직접 파싱) |
| 이미지 | Pillow |
| PDF 변환 | comtypes (Windows Word COM) |
| PDF 병합 | pypdf 6.x (`PdfWriter`+`PdfReader`) |
| 빌드 | PyInstaller |
| AI 도구 | Claude Code — Plan 모드 (Anthropic) |

---

CLAUDE 토큰 소진으로 인한 CODEX 사용.


## v2 추가 작업 로그 (2026-04-13)

| 항목 | 내용 |
|------|------|
| 앱 아이콘 교체 | 타이틀바 좌측 상단 아이콘을 AND-LOGO-1.png로 적용 |
| 비율 유지 처리 | 정사각 아이콘 캔버스에 원본 비율로 중앙 배치하여 찌그러짐 방지 |
| 표시성 개선 | 타이틀바 아이콘 크기 한계(Windows OS 고정)로 인해, UI 내부 좌측 상단에 대형 AND 로고 추가 |
| 경로 처리 | Downloads 경로 우선 + 프로젝트 img/AND-LOGO-1.png fallback 적용 |

### 이번 작업에서 Codex-1 역할
- CSV 한자 헤더 깨짐 원인 분석 및 우회 전략 수립
: CSV 컬럼명 매칭 실패를 확인하고, 헤더 의존도를 낮춘 인덱스/alias 기반 치환 로직으로 보완.
- HWP -> XLSX 직결 파이프라인 구현
: `hwp_to_xlsx.py`를 만들어 CSV 단계를 건너뛰고 안정적인 XLSX 헤더로 데이터 생성하도록 정리.
- DOCX 생성 파이프라인 분리
: `xlsx_to_doc.py`(텍스트 전용)와 `xlsx_to_doc_img.py`(이미지 삽입 전용)로 책임 분리.
- 이미지 삽입 안정화
: `{金型写真}` 토큰 삽입 시 문단 전체 재작성 방식 제거, run-level 비파괴 삽입 방식으로 개선.
: 토큰이 여러 run으로 분리된 케이스도 처리하도록 fallback 로직 추가.
- 체크박스/서식 규칙 반영
: `新作/增作/二元化/業者更変/機種更変`의 `1 -> O` 처리, `{BASE}/{CORE}` 6pt 서식 유지.
- 경로 정책 정리
: 주요 실행 스크립트에서 절대경로 하드코딩 제거, 상대경로 기반 실행으로 통일.
- 검증 실행
: 요구사항에 맞춰 전체 실행 대신 샘플 3건 우선 검증 반복 수행 후 결과 로그 확인.

### 이번 작업에서 Codex-2 역할
- main.py에서 아이콘 로딩/표시 로직 분석 및 원인 진단
- 타이틀바 아이콘 비율 유지 로직 구현 (_set_window_icon)
- UI 내부 대형 로고 표시 함수 추가 (_set_top_logo, _logo_candidates)
- 코드 적용 후 구문 검증(AST parse)으로 안전성 확인
## v2 추가 작업 로그 (2026-04-13, Codex)