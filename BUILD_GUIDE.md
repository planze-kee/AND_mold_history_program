# 금형이력카드 프로그램 - EXE 빌드 가이드

이 문서는 Python 프로그램을 독립 실행형 EXE 파일로 변환하여 배포하는 방법을 설명합니다.

---

## 📋 요구사항

### 빌드를 위한 최소 요구사항
- **Python**: 3.7 이상 (설치 시 PATH에 추가되어 있어야 함)
- **Windows**: Windows 7 이상
- **디스크 공간**: 500MB 이상 (빌드 과정 포함)

### 확인 방법
터미널에서 다음 명령어로 Python이 설치되었는지 확인하세요:
```bash
python --version
pip --version
```

---

## 🚀 빠른 시작 (권장)

### 방법 1: 자동 빌드 스크립트 사용 (가장 간단)

1. **프로젝트 폴더 열기**
   - 탐색기에서 `금형이력카드프로그램` 폴더로 이동

2. **build.bat 파일 실행**
   - `build.bat` 파일을 더블클릭하면 자동으로 빌드 시작
   - 진행 상황이 콘솔 창에 표시됨 (약 30~60초 소요)

3. **EXE 파일 확인**
   - 빌드 완료 후 `dist` 폴더에 `금형이력카드프로그램.exe` 생성됨

```
dist/
└── 금형이력카드프로그램.exe  ← 배포할 파일
```

---

## 🔧 상세 수동 빌드 (고급 사용자)

### Step 1: 필수 패키지 설치

프로젝트 폴더에서 다음 명령어 실행:
```bash
pip install -r requirements-build.txt
```

### Step 2: PyInstaller로 EXE 생성

```bash
pyinstaller --name "금형이력카드프로그램" ^
    --onefile ^
    --windowed ^
    --add-data "data:data" ^
    --add-data "YES:YES" ^
    --add-data "img:img" ^
    --hidden-import=openpyxl ^
    --hidden-import=olefile ^
    --hidden-import=docx ^
    --hidden-import=PIL ^
    main.py
```

**옵션 설명**:
| 옵션 | 설명 |
|------|------|
| `--onefile` | 모든 파일을 하나의 EXE로 통합 |
| `--windowed` | 콘솔 창 없이 GUI만 표시 |
| `--name` | EXE 파일 이름 지정 |
| `--add-data` | 필요한 폴더/파일 포함 (형식: source:dest) |
| `--hidden-import` | PyInstaller가 감지하지 못한 모듈 명시적 포함 |

### Step 3: 빌드 결과 확인

빌드 완료 후 다음 폴더 구조 생성:
```
프로젝트폴더/
├── build/          (임시 빌드 파일 - 무시해도 됨)
├── dist/           (배포용 폴더)
│   └── 금형이력카드프로그램.exe
├── main.spec       (PyInstaller 설정 파일)
└── ...
```

---

## 📦 배포 준비

### 배포 폴더 구성

사용자에게 배포할 폴더 구조:

```
금형이력카드프로그램/
├── 금형이력카드프로그램.exe   (실행 파일)
├── YES/                      (테스트 HWP 파일들)
├── data/                     (템플릿 및 출력)
│   ├── templates/
│   │   └── 11-000_양식.docx
│   └── output/
├── img/                      (이미지 저장 폴더)
└── USAGE.md                  (사용 설명서)
```

### 배포 단계

1. **EXE 파일 추출**
   ```
   dist/금형이력카드프로그램.exe → 배포 폴더
   ```

2. **필요한 폴더 복사**
   ```
   YES/          → 배포 폴더/YES/
   data/         → 배포 폴더/data/
   USAGE.md      → 배포 폴더/USAGE.md
   ```

3. **패키징**
   - 배포 폴더를 ZIP으로 압축하여 사용자에게 전달

---

## 🔍 빌드 최적화 옵션

### 더 작은 EXE 파일 생성

```bash
pyinstaller --onefile --windowed --strip main.py
```

### 더 빠른 시작 속도

```bash
pyinstaller --onefile --windowed -y main.py
```

### 커스텀 아이콘 추가

```bash
pyinstaller --onefile --windowed --icon=app.ico main.py
```

(프로젝트 폴더에 `app.ico` 파일 필요)

---

## ⚠️ 문제 해결

### Q: "pyinstaller 명령을 찾을 수 없음" 오류
**원인**: PyInstaller가 설치되지 않았음
**해결**:
```bash
pip install PyInstaller==6.19.0
```

### Q: "EXE 파일이 실행되지 않음"
**원인**: 필수 폴더가 없거나 경로 오류
**해결**:
1. `data/`, `YES/` 폴더가 EXE와 같은 폴더에 있는지 확인
2. 템플릿 파일 경로 확인: `data/templates/11-000_양식.docx`
3. 콘솔 창에서 EXE 실행하여 오류 메시지 확인:
   ```bash
   cd dist
   금형이력카드프로그램.exe
   ```

### Q: EXE 파일 크기가 너무 큼 (1GB+)
**원인**: PyInstaller가 불필요한 라이브러리도 포함함
**해결**: 기본 설정으로 정상 (첫 실행 시 임시 압축 해제)

### Q: 안티바이러스 프로그램에서 경고 발생
**원인**: PyInstaller로 생성된 EXE는 패킹 방식 때문에 일부 안티바이러스에서 오탐
**해결**:
1. 신뢰할 수 있는 프로그램으로 등록
2. 또는 Python 환경에서 직접 실행: `python main.py`

---

## 💡 팁

### 1. 자동 빌드 배치 파일
`build.bat`를 수정하여 추가 옵션 설정 가능:

```batch
pyinstaller --name "프로그램명" ^
    --onefile ^
    --windowed ^
    --icon=app.ico ^
    main.py
```

### 2. 버전 정보 추가
EXE 파일의 속성에 버전 정보를 추가하려면 `--version-file` 옵션 사용

### 3. CI/CD 통합
GitHub Actions 등에서 자동 빌드 설정 가능

---

## 📊 빌드 결과 예상

| 항목 | 예상값 |
|------|--------|
| EXE 파일 크기 | ~1.2GB |
| 빌드 시간 | 30~60초 |
| 첫 실행 시간 | 3~5초 (압축 해제) |
| 이후 실행 시간 | 1~2초 |

---

## 🎯 완성 체크리스트

빌드 완료 후 확인 사항:

- [ ] `dist/금형이력카드프로그램.exe` 파일 생성됨
- [ ] EXE 파일이 실행됨 (더블클릭)
- [ ] GUI 창이 정상 표시됨
- [ ] 필수 데이터 폴더가 같은 디렉토리에 있음
- [ ] 1단계, 2단계, 3단계 모두 정상 작동 (테스트용 YES 폴더 사용)
- [ ] 도움말 버튼(?) 클릭 시 팝업 표시됨

---

## 📝 참고사항

- **Python 의존성**: EXE 파일은 독립 실행형이므로 사용자는 Python을 설치할 필요 없음
- **바이러스 검사**: Windows Defender 등에서 경고할 수 있음 (오탐) - 신뢰할 수 있는 파일로 등록하면 됨
- **파일 구조**: `--add-data` 옵션으로 폴더 구조를 자동 포함시킬 수 있음

---

**마지막 수정**: 2026-04-15  
**버전**: 2.1
