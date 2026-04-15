# 금형이력카드 프로그램 - GUI 개선사항 및 업그레이드 방향성

작성일: 2026-04-14  
분석 대상: main.py (PyQt5 기반 GUI 애플리케이션)

---

## 📋 현재 구조 분석

### 주요 기능 (탭 구성)
1. **문서 생성/동기화**: XLSX → DOCX 변환 및 기존 파일 동기화
2. **이력 관리**: 유지보수 이력 편집 및 Word 반영
3. **PDF 변환/병합**: Word → PDF 변환 (단일/일괄/병합)
4. **HWP to 엑셀**: HWP 파일 → XLSX 변환
5. **이미지 추출**: HWP 파일에서 이미지 추출
6. **신규 이력카드 발행**: 다이얼로그 기반 신규 카드 생성

### 코드 구조
- **총 라인 수**: 1,269줄
- **클래스 구성**:
  - `WorkerSignals`: Qt 신호 관리
  - `NewCardDialog`: 신규 이력카드 발행 다이얼로그
  - `MainWindow`: 메인 GUI 창 (560×520 픽셀 기본 크기)
- **스레딩**: `Thread(daemon=True)` 사용으로 백그라운드 작업 처리

---

## ✅ 완료 현황 (2026-04-14, feature/phase1-core-improvements)

| 항목 | 상태 | 내용 |
|------|------|------|
| **1.2 QThread 전환** | ✅ 완료 | `threading.Thread` 9곳 → `Worker(QThread)` 전환, `_start_worker()` 헬퍼 추가 |
| **1.3 설정 파일 시스템 (config.yaml)** | ✅ 완료 | `src/config.py` 신규 작성, 탭별 경로 15개 + 창 위치/크기 자동 저장/복원 |
| **2.3 로그 기능 강화** | ✅ 완료 | 타임스탬프 추가, 성공(녹색)/오류(빨간색)/기본(흰색) HTML 색상 구분 |
| **2.4 창 위치/크기 저장** | ✅ 완료 | `closeEvent()` 오버라이드로 종료 시 config.yaml에 자동 저장, 다음 실행 시 복원 |
| **신규 이력카드 발행 UX 개선** | ✅ 완료 | DB 직전 항목(파일명/품명) 표시 바 추가, File name 편집 가능 QLineEdit 전환 |

### 미완료 항목

| 항목 | 우선순위 | 비고 |
|------|---------|------|
| 1.1 비즈니스 로직/UI 분리 (MVC) | Low | main.py 1,400줄+, 기능 안정 후 진행 권장 |
| 2.1 진행률 표시 (progress_bar 실활용) | Medium | core.py callback에 progress % 전달 필요 |
| 2.5 드래그 앤 드롭 | Medium | 파일 찾기 UX 개선 |
| 3.1 일괄 작업 마법사 | Low | 단계별 자동 실행 |
| 3.5 DB 검색/필터링 | Low | XLSX 직접 조회 UI |
| 6.2 키보드 단축키 | Low | Ctrl+R (실행), Ctrl+L (로그 지우기) 등 |

---

## 🔧 개선사항

### 1. 아키텍처 & 코드 품질

#### 1.1 비즈니스 로직과 UI 분리
**현재 상태**:
- `MainWindow` 클래스에 UI, 비즈니스 로직, 이벤트 핸들러가 혼재
- 1,269줄의 단일 파일로 유지보수성 저하

**개선 제안**:
```
main.py              → GUI 레이아웃 정의
controllers.py       → 이벤트 핸들러 및 비즈니스 로직
workers.py           → QThread 기반 백그라운드 작업자
models.py            → 데이터 모델 및 상태 관리
utils/ui_helpers.py  → UI 유틸리티 함수
```

**효과**:
- 각 모듈의 책임 명확화
- 단위 테스트 용이성 증가
- 코드 재사용성 향상

---

#### 1.2 스레딩 개선 (QThread 활용)
**현재 상태**:
```python
Thread(target=task, daemon=True).start()
```
- Python `threading.Thread` 직접 사용
- 신호 전달은 `WorkerSignals` 활용

**개선 제안**:
```python
class Worker(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool)
    
    def __init__(self, task_func, *args):
        super().__init__()
        self.task_func = task_func
        self.args = args
    
    def run(self):
        try:
            self.task_func(*self.args, callback=self.log_signal.emit)
            self.finished_signal.emit(True)
        except Exception as e:
            self.log_signal.emit(f"✗ 오류: {e}")
            self.finished_signal.emit(False)
```

**효과**:
- Qt 이벤트 루프와 안전한 통합
- 작업 취소, 일시정지 기능 구현 가능
- 메모리 관리 개선

---

#### 1.3 설정 관리 (config.yaml 또는 .ini 파일)
**현재 상태**:
- 기본 경로들이 코드에 하드코딩
```python
self.hwp_input_edit = QLineEdit("YES")
self.hwp_output_edit = QLineEdit("data/output/output_from_hwp.xlsx")
```

**개선 제안**:
```yaml
# config.yaml
paths:
  hwp_input: "YES"
  hwp_output: "data/output/output_from_hwp.xlsx"
  templates: "data/templates"
  output: "data/output"
  images: "img"
  
database:
  default_xlsx: "data/output/00.DB_19-000.xlsx"
  
ui:
  window_size: [560, 520]
  log_max_lines: 1000
```

**효과**:
- 사용자별 설정 저장 가능
- 배포 환경별 설정 분리 용이
- 코드 재컴파일 없이 설정 변경 가능

---

### 2. UI/UX 개선

#### 2.1 작업 진행률 표시 개선
**현재 상태**:
```python
self.progress_bar = QProgressBar()
self.progress_bar.setVisible(False)
```
- 진행률 바는 존재하나 실제 사용 안 됨

**개선 제안**:
- HWP 파일 처리 시 진행률 실시간 표시
```python
# src/core.py 내부
for i, hwp_file in enumerate(hwp_files):
    callback(f"처리 중: {hwp_file.name}", progress=int((i/total)*100))
```

- 작업별 예상 시간 표시
- 중단 버튼 추가

**효과**:
- 사용자 경험 향상
- 장시간 작업 시 응답성 체감 개선

---

#### 2.2 오류 처리 및 피드백 강화
**현재 상태**:
```python
except Exception as e:
    self.signals.log.emit(f"✗ 오류: {e}")
```
- 일반적인 예외 처리만 존재
- 사용자에게 해결 방법 제시 부족

**개선 제안**:
```python
class ErrorHandler:
    ERROR_MESSAGES = {
        FileNotFoundError: "파일을 찾을 수 없습니다.\n경로: {path}\n\n해결방법:\n1. 파일 경로 확인\n2. '찾기' 버튼으로 재선택",
        PermissionError: "파일 접근 권한이 없습니다.\n\n해결방법:\n1. 파일이 다른 프로그램에서 열려있는지 확인\n2. 관리자 권한으로 실행",
        # ... 기타 오류 타입
    }
    
    @staticmethod
    def handle(exception, context=""):
        error_type = type(exception)
        message = ErrorHandler.ERROR_MESSAGES.get(
            error_type, 
            f"예기치 않은 오류가 발생했습니다.\n{str(exception)}"
        )
        return message.format(path=getattr(exception, 'filename', ''))
```

**효과**:
- 사용자 친화적 오류 메시지
- 문제 해결 가능성 증가
- 지원 요청 감소

---

#### 2.3 로그 기능 강화
**현재 상태**:
```python
self.log_text = QTextEdit()
self.log_text.setReadOnly(True)
```
- 단순 텍스트 추가만 가능
- 로그 저장/검색 기능 없음

**개선 제안**:
- **로그 필터링**: 오류만, 경고만, 전체
- **로그 내보내기**: .txt, .csv 저장
- **로그 검색**: Ctrl+F 단축키로 검색
- **타임스탬프 추가**: `[2026-04-14 14:32:15] ✓ HWP 변환 완료`
- **색상 구분**: 오류(빨강), 경고(주황), 성공(초록)

```python
def log_message(self, msg, level="info"):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    colors = {"error": "#E53935", "warning": "#FB8C00", "success": "#43A047", "info": "#1E88E5"}
    color = colors.get(level, colors["info"])
    
    html = f'<span style="color:{color}">[{timestamp}] {msg}</span>'
    self.log_text.append(html)
```

**효과**:
- 문제 추적 용이성 증가
- 로그 분석 가능
- 디버깅 시간 단축

---

#### 2.4 반응형 레이아웃
**현재 상태**:
```python
self.setGeometry(100, 100, 560, 520)
```
- 고정된 창 크기

**개선 제안**:
- `QSplitter` 추가로 패널 크기 조절 가능하게
- 로그 창 접기/펼치기 버튼
- 최소/최대 창 크기 제한 설정
- 창 크기 및 위치 저장 (다음 실행 시 복원)

```python
def closeEvent(self, event):
    settings = QSettings("AND", "MoldHistory")
    settings.setValue("geometry", self.saveGeometry())
    settings.setValue("windowState", self.saveState())
    super().closeEvent(event)

def __init__(self):
    # ... 기존 코드
    settings = QSettings("AND", "MoldHistory")
    if settings.value("geometry"):
        self.restoreGeometry(settings.value("geometry"))
```

**효과**:
- 다양한 모니터 해상도 대응
- 사용자 맞춤형 레이아웃
- UX 일관성 유지

---

#### 2.5 드래그 앤 드롭 지원
**개선 제안**:
- HWP 파일을 GUI에 드래그 앤 드롭하면 자동으로 입력 폴더 설정
- Word 파일을 PDF 탭에 드래그하면 변환 목록에 추가
- Excel 파일을 문서 생성 탭에 드래그하면 자동 경로 설정

```python
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
    
    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        # 현재 탭에 따라 처리
        current_tab = self.tabs.currentIndex()
        if current_tab == 0 and files[0].endswith('.hwp'):
            self.hwp_input_edit.setText(Path(files[0]).parent)
```

**효과**:
- 파일 찾기 클릭 횟수 감소
- 워크플로우 속도 향상
- 직관적인 사용성

---

### 3. 기능 추가

#### 3.1 일괄 작업 (Batch Wizard)
**개선 제안**:
새로운 탭 또는 메뉴: **"전체 워크플로우 실행"**

```
[ 체크박스 ] 1단계: HWP → XLSX 변환
[ 체크박스 ] 2단계: 이미지 추출
[ 체크박스 ] 3단계: DOCX 생성
[ 체크박스 ] 4단계: PDF 변환
[ 체크박스 ] 5단계: PDF 병합

[전체 실행] 버튼
```

- 체크된 단계를 순차적으로 자동 실행
- 각 단계 완료 후 다음 단계로 자동 진행
- 실패 시 중단 or 계속 진행 옵션

**효과**:
- 반복 작업 자동화
- 사용자 개입 최소화
- 처리 시간 단축

---

#### 3.2 프리뷰 기능
**개선 제안**:
- **DOCX 미리보기**: 생성 전 템플릿에 샘플 데이터 적용하여 미리보기
- **PDF 미리보기**: 변환 후 내장 뷰어로 확인 (PyMuPDF 활용)
- **이미지 미리보기**: 추출된 이미지 썸네일 표시

```python
from PyQt5.QtWidgets import QGraphicsView, QGraphicsScene
from PyQt5.QtGui import QPixmap

class PreviewDialog(QDialog):
    def __init__(self, image_path):
        super().__init__()
        scene = QGraphicsScene()
        pixmap = QPixmap(str(image_path))
        scene.addPixmap(pixmap)
        
        view = QGraphicsView(scene)
        layout = QVBoxLayout()
        layout.addWidget(view)
        self.setLayout(layout)
```

**효과**:
- 결과 확인 시간 단축
- 오류 조기 발견
- 사용자 만족도 향상

---

#### 3.3 실행 이력 (History) 기능
**개선 제안**:
- 최근 실행한 작업 목록 표시
- 같은 설정으로 재실행 가능
- 실행 시간, 성공/실패 여부 기록

```python
# history.json
{
  "executions": [
    {
      "timestamp": "2026-04-14 14:30:00",
      "task": "HWP → XLSX 변환",
      "input": "YES/",
      "output": "data/output/output.xlsx",
      "status": "success",
      "duration": "12.5s"
    }
  ]
}
```

**효과**:
- 반복 작업 효율화
- 작업 이력 추적 가능
- 설정 재입력 불필요

---

#### 3.4 템플릿 관리
**개선 제안**:
- 템플릿 등록/관리 UI
- 여러 템플릿 중 선택 (드롭다운)
- 템플릿 미리보기
- 기본 템플릿 설정

```
[템플릿 관리] 버튼
  → 다이얼로그 팝업
     - 등록된 템플릿 목록
     - [추가] [삭제] [기본값 설정] 버튼
```

**효과**:
- 다양한 양식 지원
- 템플릿 전환 용이
- 비전문가도 템플릿 변경 가능

---

#### 3.5 DB 검색 및 필터링
**개선 제안**:
- XLSX DB를 GUI에서 직접 조회/검색
- 품명, 도번번호, 관리번호로 검색
- 검색 결과에서 바로 Word 생성/PDF 변환

```
[DB 검색] 탭 추가
  검색창: [_____________] [검색]
  
  결과 테이블:
  | 品名 | 図番番号 | 管理番号 | File name | 작업 |
  |------|----------|----------|-----------|------|
  | CASE | 1071000024 | 19-001 | 19-001... | [Word생성] [PDF변환] |
```

**효과**:
- DB 관리 편의성 증가
- 특정 항목만 처리 가능
- Excel 직접 열기 불필요

---

### 4. 성능 및 안정성

#### 4.1 멀티프로세싱 지원
**현재 상태**:
- 단일 스레드로 순차 처리

**개선 제안**:
```python
from multiprocessing import Pool

def batch_convert(files, output_dir, workers=4):
    with Pool(processes=workers) as pool:
        results = pool.map(convert_single, files)
    return results
```

**효과**:
- 대량 파일 처리 속도 향상 (최대 4배)
- CPU 코어 효율적 활용

---

#### 4.2 캐싱 및 중복 작업 방지
**개선 제안**:
- 이미 변환된 파일은 건너뛰기 (체크섬 비교)
- "강제 재변환" 체크박스로 선택 가능

```python
import hashlib

def get_file_hash(path):
    with open(path, 'rb') as f:
        return hashlib.md5(f.read()).hexdigest()

# cache.json에 파일명: 해시값 저장
```

**효과**:
- 불필요한 작업 방지
- 처리 시간 단축
- 리소스 절약

---

#### 4.3 자동 저장 및 복구
**개선 제안**:
- 작업 중 중단 시 진행 상태 저장
- 재실행 시 "이전 작업 계속하기" 옵션

```python
# state.json
{
  "task": "docx_generation",
  "completed": ["19-001.docx", "19-002.docx"],
  "total": 100,
  "timestamp": "2026-04-14 14:30:00"
}
```

**효과**:
- 예기치 않은 종료 대응
- 대량 작업 안정성 확보

---

### 5. 배포 및 유지보수

#### 5.1 자동 업데이트 기능
**개선 제안**:
- GitHub Releases 또는 내부 서버에서 최신 버전 확인
- 업데이트 알림 및 자동 다운로드

```python
import requests

def check_update():
    response = requests.get("https://api.github.com/repos/user/repo/releases/latest")
    latest = response.json()["tag_name"]
    current = "v1.0.0"
    return latest > current
```

**효과**:
- 사용자 최신 버전 유지
- 버그 수정 신속 배포

---

#### 5.2 사용자 매뉴얼 내장
**현재 상태**:
- 각 탭에 `?` 버튼으로 간단한 도움말 표시

**개선 제안**:
- `F1` 키로 상세 매뉴얼 열기
- HTML 또는 PDF 형식의 매뉴얼 내장
- 비디오 튜토리얼 링크

```python
def show_manual(self):
    manual_path = Path(__file__).parent / "docs" / "manual.html"
    QDesktopServices.openUrl(QUrl.fromLocalFile(str(manual_path)))
```

**효과**:
- 사용자 학습 곡선 단축
- 지원 요청 감소

---

#### 5.3 로깅 및 디버그 모드
**개선 제안**:
- 명령줄 인자로 디버그 모드 활성화
```bash
python main.py --debug
```
- 상세 로그를 파일에 기록 (`logs/app.log`)
- 원격 로그 전송 (선택적)

```python
import logging

logging.basicConfig(
    filename='logs/app.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
```

**효과**:
- 문제 진단 용이
- 원격 지원 가능

---

### 6. 접근성 및 국제화

#### 6.1 다국어 지원 (i18n)
**개선 제안**:
- 한국어/일본어/영어 전환 가능
- Qt Linguist 활용

```python
translator = QTranslator()
translator.load(f"translations/app_{locale}.qm")
app.installTranslator(translator)
```

**효과**:
- 글로벌 사용자 지원
- 일본 본사 사용 가능성

---

#### 6.2 키보드 단축키
**개선 제안**:
```
Ctrl+O : 파일 열기
Ctrl+S : 저장
Ctrl+R : 실행
Ctrl+Q : 종료
F5     : 새로고침
Ctrl+L : 로그 지우기
```

```python
QShortcut(QKeySequence("Ctrl+R"), self, self.run_current_task)
```

**효과**:
- 파워 유저 생산성 향상
- 마우스 사용 최소화

---

## 🚀 업그레이드 로드맵

### Phase 1: 안정성 및 코드 품질 (1-2주)
- [ ] 비즈니스 로직 분리 (MVC 패턴)
- [x] QThread 전환 ✅ 2026-04-14
- [x] 설정 파일 시스템 구축 ✅ 2026-04-14 (config.yaml)
- [x] 오류 처리 강화 ✅ 2026-04-14 (사용자 정의 예외 + logging)
- [x] 로깅 시스템 개선 ✅ 2026-04-14 (타임스탬프 + 색상)

### Phase 2: UX 개선 (1-2주)
- [ ] 진행률 표시 개선
- [ ] 드래그 앤 드롭 지원
- [ ] 프리뷰 기능 추가
- [ ] 반응형 레이아웃

### Phase 3: 기능 확장 (2-3주)
- [ ] 일괄 작업 마법사
- [ ] DB 검색/필터링
- [ ] 템플릿 관리
- [ ] 실행 이력 기능

### Phase 4: 성능 최적화 (1주)
- [ ] 멀티프로세싱 지원
- [ ] 캐싱 시스템
- [ ] 자동 저장/복구

### Phase 5: 배포 준비 (1주)
- [ ] 자동 업데이트 기능
- [ ] 사용자 매뉴얼 작성
- [ ] 인스톨러 제작 (PyInstaller/Nuitka)
- [ ] 다국어 지원

---

## 🎯 우선순위 제안

### 즉시 구현 (High Priority)
1. **오류 처리 강화** - 사용자 불만 최소화
2. **로깅 개선** - 문제 추적 용이성
3. **진행률 표시** - UX 향상
4. **설정 파일 시스템** - 유지보수성

### 중기 구현 (Medium Priority)
5. **드래그 앤 드롭** - 편의성
6. **일괄 작업 마법사** - 자동화
7. **프리뷰 기능** - 오류 조기 발견

### 장기 구현 (Low Priority)
8. **DB 검색 기능** - 고급 사용자용
9. **다국어 지원** - 확장성
10. **자동 업데이트** - 운영 편의성

---

## 📊 예상 효과

### 개발 측면
- **유지보수 시간 50% 감소** (코드 분리로 버그 수정 용이)
- **테스트 커버리지 80%+** (단위 테스트 가능한 구조)
- **신규 기능 추가 시간 30% 단축** (모듈화)

### 사용자 측면
- **작업 시간 40% 단축** (일괄 작업, 캐싱)
- **오류 발생률 60% 감소** (검증 강화)
- **학습 시간 50% 단축** (직관적 UI, 매뉴얼)

---

## 📝 참고사항

### 현재 코드 강점
- ✅ 모든 핵심 기능 구현 완료
- ✅ PyQt5 기반 안정적 GUI
- ✅ 스레드 분리로 UI 블로킹 방지
- ✅ 로고 및 아이콘 브랜딩

### 개선 필요 영역
- ⚠️ 1,269줄 단일 파일 (리팩토링 필요)
- ⚠️ 하드코딩된 경로
- ⚠️ 진행률 표시 미활용
- ⚠️ 예외 처리 일반화

---

## 🛠️ 기술 스택 제안

### 추가 라이브러리
- `PyYAML`: 설정 파일 관리
- `loguru`: 고급 로깅
- `PyMuPDF`: PDF 미리보기
- `requests`: 자동 업데이트
- `pytest`: 단위 테스트

### 개발 도구
- `black`: 코드 포맷팅
- `pylint`: 코드 품질 검사
- `pyinstaller`: 실행 파일 빌드

---

**작성자**: Claude (AI Assistant)  
**검토 필요**: 기존 개발자와 협의 후 우선순위 확정
