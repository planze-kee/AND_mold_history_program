"""
다크테마 모듈 — 색상 팔레트 상수 + 전역 QSS

main()에서 app.setStyleSheet(DARK_QSS) 1회 적용한다.
버튼은 property 기반으로 역할을 구분한다:
    btn.setProperty("class", "primary")  # 주 실행 버튼 (액센트색)
    btn.setProperty("class", "danger")   # 취소/위험 버튼
"""

# ============================================================================
# 색상 팔레트
# ============================================================================
BG_WINDOW   = "#1e1e1e"   # 최상위 배경
BG_PANEL    = "#252526"   # 그룹박스/탭 페이지 배경
BG_INPUT    = "#2d2d30"   # 입력 위젯 배경
BG_HOVER    = "#3e3e42"   # hover 배경
BORDER      = "#3f3f46"   # 테두리
TEXT        = "#d4d4d4"   # 본문 텍스트
TEXT_DIM    = "#9e9e9e"   # 보조 텍스트
TEXT_DISABLED = "#6e6e6e" # 비활성 텍스트
ACCENT      = "#0a84d0"   # 액센트 (파랑)
ACCENT_HOVER = "#1494e0"
ACCENT_DOWN = "#0874b8"
SUCCESS     = "#4ec9b0"   # 성공 (청록)
ERROR       = "#f48771"   # 오류 (연빨강)
WARNING     = "#dcdcaa"   # 경고 (연노랑)
TIMESTAMP   = "#6e6e6e"   # 로그 타임스탬프

# 로그 색상 (log_message에서 사용)
LOG_DEFAULT = TEXT
LOG_SUCCESS = SUCCESS
LOG_ERROR   = ERROR
LOG_WARNING = WARNING


# ============================================================================
# 전역 QSS
# ============================================================================
DARK_QSS = f"""
/* ── 기본 ─────────────────────────────────────────────── */
QMainWindow, QDialog, QMessageBox {{
    background-color: {BG_WINDOW};
}}
QWidget {{
    color: {TEXT};
    font-size: 12px;
}}
QLabel {{
    background: transparent;
}}
QToolTip {{
    background-color: {BG_INPUT};
    color: {TEXT};
    border: 1px solid {BORDER};
    padding: 3px 6px;
}}

/* ── 탭 ───────────────────────────────────────────────── */
QTabWidget::pane {{
    border: 1px solid {BORDER};
    background-color: {BG_PANEL};
    top: -1px;
}}
QTabBar::tab {{
    background-color: {BG_WINDOW};
    color: {TEXT_DIM};
    border: 1px solid {BORDER};
    border-bottom: none;
    padding: 6px 9px;
    margin-right: 2px;
}}
QTabBar::tab:selected {{
    background-color: {BG_PANEL};
    color: {TEXT};
    border-top: 2px solid {ACCENT};
}}
QTabBar::tab:hover:!selected {{
    background-color: {BG_HOVER};
    color: {TEXT};
}}

/* ── 그룹박스 ─────────────────────────────────────────── */
QGroupBox {{
    background-color: {BG_PANEL};
    border: 1px solid {BORDER};
    border-radius: 4px;
    margin-top: 8px;
    padding-top: 8px;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 8px;
    padding: 0 4px;
    color: {TEXT_DIM};
}}

/* ── 입력 위젯 ────────────────────────────────────────── */
QLineEdit, QTextEdit, QSpinBox {{
    background-color: {BG_INPUT};
    color: {TEXT};
    border: 1px solid {BORDER};
    border-radius: 3px;
    padding: 3px 6px;
    selection-background-color: {ACCENT};
}}
QLineEdit:focus, QTextEdit:focus, QSpinBox:focus {{
    border-color: {ACCENT};
}}
QLineEdit:disabled, QTextEdit:disabled, QSpinBox:disabled {{
    color: {TEXT_DISABLED};
    background-color: {BG_WINDOW};
}}
QLineEdit[dragOver="true"] {{
    border: 1px dashed {ACCENT};
}}
QSpinBox::up-button, QSpinBox::down-button {{
    background-color: {BG_HOVER};
    border: none;
    width: 16px;
}}
QSpinBox::up-arrow {{
    image: none; border-left: 4px solid transparent; border-right: 4px solid transparent;
    border-bottom: 5px solid {TEXT_DIM}; width: 0; height: 0;
}}
QSpinBox::down-arrow {{
    image: none; border-left: 4px solid transparent; border-right: 4px solid transparent;
    border-top: 5px solid {TEXT_DIM}; width: 0; height: 0;
}}

/* ── 버튼 ─────────────────────────────────────────────── */
QPushButton {{
    background-color: {BG_HOVER};
    color: {TEXT};
    border: 1px solid {BORDER};
    border-radius: 3px;
    padding: 5px 12px;
}}
QPushButton:hover {{
    background-color: #4a4a50;
}}
QPushButton:pressed {{
    background-color: {BG_INPUT};
}}
QPushButton:disabled {{
    color: {TEXT_DISABLED};
    background-color: {BG_PANEL};
    border-color: {BORDER};
}}
QPushButton[class="primary"] {{
    background-color: {ACCENT};
    color: #ffffff;
    border: none;
    font-weight: bold;
    padding: 8px 12px;
}}
QPushButton[class="primary"]:hover {{
    background-color: {ACCENT_HOVER};
}}
QPushButton[class="primary"]:pressed {{
    background-color: {ACCENT_DOWN};
}}
QPushButton[class="primary"]:disabled {{
    background-color: {BG_HOVER};
    color: {TEXT_DISABLED};
}}
QPushButton[class="danger"] {{
    background-color: transparent;
    color: {ERROR};
    border: 1px solid {ERROR};
    font-weight: bold;
    padding: 0px 6px;
}}
QPushButton[class="danger"]:hover {{
    background-color: #4a2a25;
}}
QPushButton[class="danger"]:disabled {{
    color: {TEXT_DISABLED};
    border-color: {TEXT_DISABLED};
    background-color: transparent;
}}

/* ── 리스트 ───────────────────────────────────────────── */
QListWidget {{
    background-color: {BG_INPUT};
    border: 1px solid {BORDER};
    border-radius: 3px;
}}
QListWidget::item {{
    padding: 3px 6px;
}}
QListWidget::item:selected {{
    background-color: {ACCENT};
    color: #ffffff;
}}
QListWidget::item:hover:!selected {{
    background-color: {BG_HOVER};
}}

/* ── 체크박스 / 라디오 ────────────────────────────────── */
QCheckBox, QRadioButton {{
    spacing: 6px;
}}
QCheckBox::indicator, QRadioButton::indicator {{
    width: 14px;
    height: 14px;
    border: 1px solid {BORDER};
    background-color: {BG_INPUT};
}}
QRadioButton::indicator {{
    border-radius: 7px;
}}
QCheckBox::indicator {{
    border-radius: 3px;
}}
QCheckBox::indicator:checked, QRadioButton::indicator:checked {{
    background-color: {ACCENT};
    border-color: {ACCENT};
}}
QCheckBox::indicator:hover, QRadioButton::indicator:hover {{
    border-color: {ACCENT};
}}

/* ── 진행률 바 ────────────────────────────────────────── */
QProgressBar {{
    background-color: {BG_INPUT};
    border: 1px solid {BORDER};
    border-radius: 3px;
    text-align: center;
    color: {TEXT};
}}
QProgressBar::chunk {{
    background-color: {ACCENT};
    border-radius: 2px;
}}

/* ── 스플리터 ─────────────────────────────────────────── */
QSplitter::handle {{
    background-color: {BORDER};
}}
QSplitter::handle:horizontal {{
    width: 2px;
}}

/* ── 스크롤바 ─────────────────────────────────────────── */
QScrollBar:vertical {{
    background-color: {BG_WINDOW};
    width: 10px;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background-color: {BG_HOVER};
    border-radius: 5px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{
    background-color: #55555c;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar:horizontal {{
    background-color: {BG_WINDOW};
    height: 10px;
    margin: 0;
}}
QScrollBar::handle:horizontal {{
    background-color: {BG_HOVER};
    border-radius: 5px;
    min-width: 24px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
}}

/* ── 메뉴/메시지박스 버튼 ─────────────────────────────── */
QMessageBox QPushButton {{
    min-width: 64px;
}}
"""


def apply_dark_titlebar(window) -> None:
    """Windows 10/11 타이틀바를 다크 모드로 전환 (실패해도 무시)."""
    try:
        import ctypes
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        hwnd = int(window.winId())
        value = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value), ctypes.sizeof(value))
    except Exception:
        pass
