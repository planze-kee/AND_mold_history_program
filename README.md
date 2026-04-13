# 금형이력카드 처리 프로그램

HWP 파일에서 데이터 및 이미지를 추출하고, XLSX 데이터를 기반으로 워드 문서를 자동 생성하는 통합 프로그램입니다.

## 기능

### 1단계: HWP → XLSX 변환
- HWP 파일에서 양식 데이터 추출
- XLSX 파일로 일괄 변환

### 2단계: 이미지 추출
- HWP 파일 내 이미지 자동 추출
- PNG, JPG, BMP 등 지원

### 3단계: DOCX 생성
- XLSX의 데이터로 템플릿 문서 자동 채우기
- 추출된 이미지 자동 삽입
- 일괄 생성 및 개별 처리 지원

## 설치

```bash
# 필수 패키지 설치
pip install -r requirements.txt
```

## 사용법

### GUI 실행 (Python)

```bash
python main.py
```

### EXE 파일로 실행 (배포용)

1. `build.bat` 파일을 더블클릭하여 EXE 빌드
2. `dist/금형이력카드프로그램.exe` 파일이 생성됨
3. EXE 파일을 더블클릭하여 실행

**자세한 빌드 방법**: [BUILD_GUIDE.md](BUILD_GUIDE.md) 참조

### CLI 사용 (고급)

```bash
# HWP → XLSX 변환
python -m src.core hwp2xlsx --input-dir YES --output output_from_hwp.xlsx

# 이미지 추출
python -m src.core extract-images --input-dir YES --output-dir img

# XLSX → DOCX 생성
python -m src.core xlsx2docx --xlsx output_from_hwp.xlsx --template "11-000_양식.docx" --out-dir output --img-dir img
```

## 디렉토리 구조

```
project/
├── src/
│   ├── __init__.py          # 모듈 초기화
│   └── core.py              # 핵심 로직 (통합)
├── data/
│   ├── templates/           # 문서 템플릿
│   ├── input/               # 입력 데이터
│   └── output/              # 생성된 출력물
├── YES/                     # 테스트 데이터 (HWP 파일)
├── img/                     # 추출된 이미지
├── main.py                  # PyQt GUI 프로그램
├── requirements.txt         # 필수 패키지
└── README.md                # 이 파일
```

## 주요 클래스

### `HWPProcessor`
- HWP 파일 → XLSX 변환
- 데이터 추출 및 정규화

### `HWPImageExtractor`
- HWP 파일에서 이미지 추출
- 압축 해제 및 형식 감지

### `DocumentFiller`
- XLSX 데이터로 DOCX 채우기
- placeholder 기반 치환 (`{필드명}`)
- 이미지 자동 삽입

## 설정 파일

### 템플릿 문서
- `data/templates/11-000_양식.docx` - 기본 템플릿

템플릿에서 다음과 같이 placeholder를 사용합니다:
```
{保管会社名}        → 보관 회사명
{作成日子}          → 작성 날짜
{管理番号}          → 관리 번호
{金型写真}          → 이미지 자동 삽입
...
```

## 테스트 데이터

- `YES/` - HWP 샘플 파일들 (19-001.hwp ~ 19-038.hwp)
- `data/input/DB_form.xlsx` - 참조 DB

## 지원 포맷

- **입력**: HWP (한글 문서 형식)
- **중간**: XLSX (엑셀 스프레드시트)
- **출력**: DOCX (워드 문서), PNG/JPG (이미지)

## 주의사항

1. 템플릿 문서(`11-000_양식.docx`)를 먼저 준비해야 합니다.
2. 이미지 추출 전에 HWP → XLSX 변환을 완료해야 합니다.
3. DOCX 생성 시 원본 템플릿이 변경되지 않습니다.

## 라이선스 및 요구사항

- Python 3.7+
- PyQt5
- openpyxl, olefile, python-docx, Pillow
