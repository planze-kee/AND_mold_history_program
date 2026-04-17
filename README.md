# 금형이력카드 처리 프로그램

> **버전**: 2.3 &nbsp;|&nbsp; **최종 수정**: 2026-04-17 &nbsp;|&nbsp; **브랜치**: main

HWP 파일에서 데이터 및 이미지를 추출하고, XLSX 데이터를 기반으로 워드 문서를 자동 생성하는 통합 프로그램입니다.

## 기능

### 탭 1 — 문서 생성/동기화
- XLSX DB → Word 템플릿 자동 채우기 (일괄 생성)
- XLSX 변경 시 기존 Word 파일 **증분 동기화** (변경된 행만 재생성, 미변경 행 스킵)
  - 행 내용 해시(SHA1) + 이미지 서명(mtime/size)으로 변경 감지
  - 템플릿 파일 교체 시 자동 전체 재생성
  - "강제 전체 재생성" 체크박스로 수동 전체 재생성 가능
- 이미지 자동 삽입: `金型写真` → 図番番号 → 品名 → 연번 순서로 fallback 탐색

### 탭 2 — 이력 관리
- 유지보수 이력 기록·조회 (파일별)
- "저장": `.data/<연번>_history.txt` + XLSX `事 由` 컬럼 동시 갱신
- "Word에 반영": 이력 내용을 포함해 Word 파일 재생성 (이미지 fallback 탐색 동일하게 적용)

### 탭 3 — PDF 변환/병합
- 단일 DOCX → PDF 변환
- 일괄 변환 (폴더 내 전체)
- 변환 후 PDF 병합 (pypdf 6.x 호환)
- 변환 진행 상황 실시간 로그 출력

### 탭 4 — HWP 처리
- HWP 파일 → XLSX DB 일괄 변환 (서브폴더 내 HWP 파일 자동 탐색)
- 양식 필드 자동 추출 (품명, 도번, 규격, 재질, 계약일 등)

### 탭 5 — 이미지 추출
- HWP 내부 이미지 자동 추출
- 이미지 파일명 규칙: `品名_図番.ext` (도번 없으면 `品名.ext`)

### 신규 이력카드 발행 (★ 버튼)
- DB 연번 자동 계산 + 직접 입력 가능
- DB 직전 항목(파일명 / 품명) 표시
- XLSX 행 추가 + Word 파일 자동 생성

## 설치

```bash
pip install -r requirements.txt
python main.py
```

### 시스템 요구사항
- Python 3.7+
- Windows (PDF 변환은 Microsoft Word COM 사용)

## 디렉토리 구조

```
project/
├── src/
│   ├── core.py          # 핵심 로직 (HWP파싱/이미지/DOCX생성/이력관리)
│   ├── config.py        # YAML 설정 관리
│   ├── pdf.py           # PDF 변환/병합
│   └── __init__.py
├── data/
│   ├── templates/       # Word 템플릿 (Word_양식.docx)
│   └── output/          # 생성된 Word 파일 및 이력 데이터
├── img/                 # 추출된 이미지
├── main.py              # PyQt5 GUI
├── config.yaml          # 경로 및 UI 설정 (자동 저장)
├── requirements.txt
└── BUILD_GUIDE.md
```

## 설정 파일 (config.yaml)

경로 및 창 위치/크기가 자동 저장됩니다. 프로그램 종료 시 현재 값으로 갱신됩니다.

```yaml
paths:
  hist_xlsx:    data/output/00.DB_19-000.xlsx   # 이력관리 DB
  hist_template: data/templates/Word_양식.docx  # 이력관리 템플릿
  hist_img:     img                              # 이력관리 이미지 폴더
  docx_xlsx:    ...                              # 문서생성 DB
  # ... 탭별 경로 설정
ui:
  window_x: 100
  window_y: 100
  window_width: 560
  window_height: 529
```

## 주요 클래스

| 클래스 | 역할 |
|--------|------|
| `HWPProcessor` | HWP → XLSX 변환 (서브폴더 rglob, EXE 환경 단일 프로세스) |
| `HWPImageExtractor` | HWP → 이미지 추출 (品名_図番 파일명) |
| `DocumentFiller` | XLSX → DOCX 생성, 이미지 삽입 |
| `DocxSyncManager` | XLSX 증분 동기화 (해시/이미지서명 비교, 템플릿 변경 자동 감지) |
| `MaintenanceHistoryManager` | 유지보수 이력 저장/반영 |
| `NewCardManager` | 신규 이력카드 발행 |
| `DocxToPdfConverter` | DOCX → PDF 변환/병합 (callback 실시간 로그) |
| `Config` | config.yaml 읽기/쓰기 |

## 템플릿 placeholder

```
{品 名}  {図番番号}  {管理番号}  {保管会社名}  {作成日子}  {承認日}
{金型規格}  {金型材質-BASE}  {金型材質-CORE}  {CAVITY 数}  {金型寿命}
{GATE 型式}  {使用機械}  {契約日}  {金型価}
{新作}  {増作}  {二元化}  {業者変更}  {仕様変更}
{修理内訳}  {事 由}
{金型写真}   ← 이미지 자동 삽입 위치
```

## 이미지 탐색 우선순위

Word 생성/재생성 시 이미지를 다음 순서로 탐색합니다:

1. XLSX `金型写真` 컬럼값
2. XLSX `図番番号` 컬럼값
3. XLSX `品名` 컬럼값
4. 연번(out_name)
5. `.data/` 폴더 (신규 발행 첨부 이미지)

trailing underscore/hyphen 자동 제거 후 재탐색 포함.

## 최근 변경 이력

| 버전 | 날짜 | 내용 |
|------|------|------|
| 2.3 | 2026-04-17 | HWP 변환 버그 수정 2건 + PDF 로그 미출력 수정 |
| 2.2 | 2026-04-16 | EXE 용량 감축 (1.2GB→54MB) + 동기화 증분 재생성 |
| 2.1 | 2026-04-14~15 | Phase 1 코드 품질 개선 + 버그 수정 6건 + 신규 발행 UX 개선 |
| 2.0 | 2026-04-13 | v2 기능 추가 (이력관리, 동기화, PDF변환, 신규발행, 5탭 구조) |
| 1.0 | 2026-04-10 | 초기 배포 (HWP→XLSX→DOCX 3단계 파이프라인) |

### v2.3 주요 수정 내용

- **HWP 변환 "No HWP files found" 오류**: `glob` → `rglob` 전환으로 서브폴더(`HWP 원본/`) 내 파일 탐색
- **HWP 변환 "process pool terminated" 오류**: PyInstaller EXE 환경에서 `ProcessPoolExecutor` 자동 비활성화 (`sys.frozen` 감지)
- **PDF 변환 작업 로그 미출력**: `src/pdf.py` 전체 `print()` → `callback()` 패턴 교체, 3개 모드 모두 실시간 로그 전달

### v2.2 주요 수정 내용

- **EXE 용량 감축**: `.spec`에서 `YES/`, `img/` 번들 제거 + `--onedir` 전환 + excludes 14개 추가 → 1.2GB → ~54MB(ZIP)
- **동기화 증분 재생성**: 행 해시·이미지 서명 비교로 미변경 행 스킵, 템플릿 변경 자동 감지, `force_all` 옵션

### v2.1 주요 수정 내용

- **QThread 전환**: `threading.Thread` → `Worker(QThread)` (UI 스레드 안전성 확보)
- **config.yaml**: 탭별 경로 15개 + 창 위치/크기 자동 저장
- **HWP 파싱 버그 수정**: `管理番号` 값 누락, `GATE型式/使用機械/契約日` 값 누수
- **이미지 탐색 버그 수정**: trailing underscore 제거, `apply_to_word()` 이미지 누락
- **신규 발행 다이얼로그**: DB 직전 항목 표시 + 파일명 직접 입력 가능

## EXE 빌드

자세한 빌드 방법: [BUILD_GUIDE.md](BUILD_GUIDE.md)

```bat
build.bat
```

## 저장소

```
git@github.com:planze-kee/AND_mold_history_program.git
```

## 라이선스 및 기술 스택

| 분류 | 라이브러리 |
|------|-----------|
| GUI | PyQt5 5.15.7 |
| Word | python-docx 0.8.11 |
| Excel | openpyxl 3.1.5 |
| HWP 파싱 | olefile 0.46 (직접 파싱) |
| 이미지 | Pillow 12.2.0 |
| PDF | comtypes (Word COM), pypdf 6.x |
| 설정 | PyYAML 6.0.2 |
| 빌드 | PyInstaller 6.19.0 |
