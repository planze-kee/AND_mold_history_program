"""
금형이력카드 처리 프로그램 - PyQt GUI
"""
import sys
from datetime import date
from pathlib import Path
from threading import Thread
from typing import Optional

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QTextEdit, QTabWidget,
    QGroupBox, QRadioButton, QButtonGroup, QSpinBox, QMessageBox,
    QProgressBar, QDialog, QCheckBox, QListWidget, QSplitter,
    QFormLayout, QDialogButtonBox
)
from PyQt5.QtCore import Qt, pyqtSignal, QObject
from PyQt5.QtGui import QFont, QIcon, QPixmap, QPainter

from src.core import (
    HWPProcessor, HWPImageExtractor, DocumentFiller,
    DocxSyncManager, MaintenanceHistoryManager, NewCardManager
)


class WorkerSignals(QObject):
    """신호 객체"""
    finished = pyqtSignal()
    error = pyqtSignal(str)
    log = pyqtSignal(str)
    progress = pyqtSignal(int)


# ============================================================================
# 신규 이력카드 발행 다이얼로그
# ============================================================================
class NewCardDialog(QDialog):
    """신규 이력카드 발행 다이얼로그"""

    def __init__(self, parent=None, xlsx_path=None):
        super().__init__(parent)
        self._init_xlsx_path = str(xlsx_path) if xlsx_path else "data/output/00.DB_19-000.xlsx"
        self._auto_file_name = ""
        self.result_data = None
        self._image_path: Optional[Path] = None
        self.setWindowTitle("신규 이력카드 발행")
        self.setMinimumWidth(480)
        self._build_ui()
        self._calc_next_file_name()

    def _build_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(12, 12, 12, 12)

        # DB 엑셀 파일 선택
        xlsx_group = QGroupBox("DB 엑셀 파일")
        xlsx_row = QHBoxLayout()
        self.dlg_xlsx_edit = QLineEdit(self._init_xlsx_path)
        self.dlg_xlsx_edit.textChanged.connect(self._calc_next_file_name)
        xlsx_browse_btn = QPushButton("찾기...")
        xlsx_browse_btn.setMaximumWidth(55)
        xlsx_browse_btn.clicked.connect(self._browse_xlsx)
        xlsx_row.addWidget(self.dlg_xlsx_edit, 1)
        xlsx_row.addWidget(xlsx_browse_btn)
        xlsx_group.setLayout(xlsx_row)
        main_layout.addWidget(xlsx_group)

        fn_row = QHBoxLayout()
        fn_row.addWidget(QLabel("File name (자동생성):"))
        self.file_name_label = QLabel("계산 중...")
        self.file_name_label.setStyleSheet("font-weight: bold; color: #1976D2;")
        fn_row.addWidget(self.file_name_label)
        fn_row.addStretch()
        main_layout.addLayout(fn_row)

        req_group = QGroupBox("필수 항목")
        req_form = QFormLayout()
        req_form.setSpacing(6)
        self.fields = {}
        for key, label in [("品 名", "품명"), ("図番番号", "도번번호"), ("管理番号", "관리번호")]:
            edit = QLineEdit()
            edit.setPlaceholderText(f"{label} 입력 (필수)")
            self.fields[key] = edit
            req_form.addRow(f"{label}:", edit)
        req_group.setLayout(req_form)
        main_layout.addWidget(req_group)

        opt_group = QGroupBox("선택 항목")
        opt_form = QFormLayout()
        opt_form.setSpacing(6)
        optional = [
            ("保管会社名", "보관회사명"), ("作成日子", "작성일자"), ("承認日", "승인일"),
            ("分 類", "분류"), ("製作処", "제작처"),
            ("MODEL 명", "MODEL 명"), ("量産処", "양산처"),
            ("金型規格", "금형규격"), ("CAVITY 数", "CAVITY 수"),
        ]
        for key, label in optional:
            edit = QLineEdit()
            if key == "作成日子":
                edit.setText(date.today().strftime("%Y.%m.%d"))
            self.fields[key] = edit
            opt_form.addRow(f"{label}:", edit)

        # 체크박스: 이중언어 라벨(위) + 체크박스(아래) 세로 배치
        check_layout = QHBoxLayout()
        self.checkboxes = {}
        checkbox_items = [
            ("新作",    "新作\n신규제작"),
            ("増作",    "増作\n증작"),
            ("二元化",  "二元化\n이원화"),
            ("業者変更","業者更変\n업체변경"),
            ("仕様変更","機種更変\n기종변경"),
        ]
        for key, display in checkbox_items:
            vbox = QVBoxLayout()
            vbox.setAlignment(Qt.AlignHCenter)
            lbl = QLabel(display)
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet("font-size: 12px;")
            cb = QCheckBox()
            cb_row = QHBoxLayout()
            cb_row.addStretch()
            cb_row.addWidget(cb)
            cb_row.addStretch()
            vbox.addWidget(lbl)
            vbox.addLayout(cb_row)
            self.checkboxes[key] = cb
            check_layout.addLayout(vbox)
        check_layout.addStretch()

        img_row = QHBoxLayout()
        img_row.addWidget(QLabel("금형 사진:"))
        self.image_name_label = QLabel("(선택 없음)")
        self.image_name_label.setStyleSheet("color: gray;")
        self.image_name_label.setMaximumWidth(200)
        img_row.addWidget(self.image_name_label, 1)
        img_browse_btn = QPushButton("찾기...")
        img_browse_btn.setMaximumWidth(55)
        img_browse_btn.clicked.connect(self._browse_image)
        img_row.addWidget(img_browse_btn)
        img_clear_btn = QPushButton("✕")
        img_clear_btn.setMaximumWidth(26)
        img_clear_btn.clicked.connect(self._clear_image)
        img_row.addWidget(img_clear_btn)

        opt_inner = QVBoxLayout()
        opt_inner.addLayout(opt_form)
        opt_inner.addLayout(check_layout)
        opt_inner.addLayout(img_row)
        opt_group.setLayout(opt_inner)
        main_layout.addWidget(opt_group)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        btn_box.button(QDialogButtonBox.Ok).setText("발행")
        btn_box.button(QDialogButtonBox.Cancel).setText("취소")
        btn_box.button(QDialogButtonBox.Ok).setStyleSheet(
            "background-color: #4CAF50; color: white; padding: 4px 12px;")
        main_layout.addWidget(btn_box)

        self.setLayout(main_layout)

    def _browse_xlsx(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "DB 엑셀 파일 선택", ".", "Excel 파일 (*.xlsx *.xls)")
        if path:
            self.dlg_xlsx_edit.setText(path)

    def _calc_next_file_name(self):
        xlsx_path = Path(self.dlg_xlsx_edit.text())
        if xlsx_path.exists():
            try:
                self._auto_file_name = NewCardManager.get_next_file_name(xlsx_path)
                self.file_name_label.setText(self._auto_file_name)
            except Exception:
                self.file_name_label.setText("오류 (수동 입력 필요)")
        else:
            self.file_name_label.setText("XLSX 없음")

    def _browse_image(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "금형 사진 선택", ".",
            "이미지 파일 (*.jpg *.jpeg *.png *.bmp *.gif *.tif *.tiff)")
        if path:
            self._image_path = Path(path)
            self.image_name_label.setText(self._image_path.name)
            self.image_name_label.setStyleSheet("color: black;")

    def _clear_image(self):
        self._image_path = None
        self.image_name_label.setText("(선택 없음)")
        self.image_name_label.setStyleSheet("color: gray;")

    def get_image_path(self) -> Optional[Path]:
        return self._image_path

    def _on_accept(self):
        missing = []
        for key, label in [("品 名", "품명"), ("図番番号", "도번번호"), ("管理番号", "관리번호")]:
            if not self.fields[key].text().strip():
                missing.append(label)
        if missing:
            QMessageBox.warning(self, "입력 오류", f"필수 항목 누락: {', '.join(missing)}")
            return
        self.result_data = {"File name": self._auto_file_name or "new"}
        for key, edit in self.fields.items():
            self.result_data[key] = edit.text().strip()
        for key, cb in self.checkboxes.items():
            self.result_data[key] = "1" if cb.isChecked() else ""
        self.accept()

    def get_data(self):
        return self.result_data


# ============================================================================
# 메인 윈도우
# ============================================================================
class MainWindow(QMainWindow):
    """메인 GUI 창"""

    def __init__(self):
        super().__init__()
        self.signals = WorkerSignals()
        self.signals.log.connect(self.log_message)
        self.signals.error.connect(self.show_error)
        self.signals.finished.connect(self.on_task_finished)
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("금형이력카드 처리 프로그램")
        self._set_window_icon()
        self.setGeometry(100, 100, 560, 520)
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout()
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(10, 10, 10, 10)

        top_layout = QHBoxLayout()
        self.logo_label = QLabel()
        self._set_top_logo()
        top_layout.addWidget(self.logo_label)
        self.new_card_btn = QPushButton("★ 신규 이력카드 발행")
        self.new_card_btn.clicked.connect(self.show_new_card_dialog)
        self.new_card_btn.setStyleSheet(
            "background-color: #9C27B0; color: white; padding: 6px 16px; "
            "font-weight: bold; font-size: 12px;")
        top_layout.addWidget(self.new_card_btn)
        top_layout.addStretch()
        main_layout.addLayout(top_layout)

        tabs = QTabWidget()
        tabs.addTab(self.create_docx_tab(), "문서 생성/동기화")
        tabs.addTab(self.create_history_tab(), "이력 관리")
        tabs.addTab(self.create_pdf_tab(), "PDF 변환/병합")
        tabs.addTab(self.create_hwp_tab(), "HWP to 엑셀")
        tabs.addTab(self.create_image_tab(), "이미지 추출")
        main_layout.addWidget(tabs)

        log_label = QLabel("작업 로그:")
        log_label.setFont(QFont("Arial", 9, QFont.Bold))
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setMinimumHeight(100)
        main_layout.addWidget(log_label)
        main_layout.addWidget(self.log_text, 1)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumHeight(20)
        main_layout.addWidget(self.progress_bar)

        central.setLayout(main_layout)

    def _logo_candidates(self):
        return [
            Path("c:/Users/UKY/Downloads/AND-LOGO-1.png"),
            Path("img/AND-LOGO-1.png"),
            Path(__file__).resolve().parent / "img" / "AND-LOGO-1.png",
        ]

    def _set_top_logo(self):
        """UI 내부 좌측 상단에 로고를 크게 표시"""
        target_h = 25
        for logo_path in self._logo_candidates():
            if not logo_path.exists():
                continue
            pixmap = QPixmap(str(logo_path))
            if pixmap.isNull():
                continue
            scaled = pixmap.scaledToHeight(target_h, Qt.SmoothTransformation)
            self.logo_label.setPixmap(scaled)
            self.logo_label.setFixedSize(scaled.size())
            self.logo_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            return
        self.logo_label.setText("AND")
        self.logo_label.setStyleSheet("font-weight: bold; color: #1976D2;")

    def _set_window_icon(self):
        """윈도우 좌측 상단 아이콘 설정 (AND 로고)"""
        for icon_path in self._logo_candidates():
            if icon_path.exists():
                original = QPixmap(str(icon_path))
                if original.isNull():
                    continue
                # 윈도우 아이콘은 정사각형으로 소비되므로, 투명 정사각 캔버스에
                # 원본 비율로 중앙 배치해 찌그러짐을 방지한다.
                side = max(original.width(), original.height())
                canvas = QPixmap(side, side)
                canvas.fill(Qt.transparent)
                painter = QPainter(canvas)
                x = (side - original.width()) // 2
                y = (side - original.height()) // 2
                painter.drawPixmap(x, y, original)
                painter.end()
                self.setWindowIcon(QIcon(canvas))
                return

    # -----------------------------------------------------------------------
    # 탭 생성
    # -----------------------------------------------------------------------
    def create_hwp_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(10, 10, 10, 10)

        title_layout = QHBoxLayout()
        title_label = QLabel("HWP → XLSX 변환")
        title_label.setFont(QFont("Arial", 11, QFont.Bold))
        title_layout.addWidget(title_label)
        help_btn = QPushButton("?")
        help_btn.setMaximumWidth(30)
        help_btn.clicked.connect(self.show_hwp_help)
        title_layout.addWidget(help_btn)
        layout.addLayout(title_layout)

        group = QGroupBox("한글파일 처리")
        vbox = QVBoxLayout()
        vbox.setSpacing(6)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("입력 폴더:"))
        self.hwp_input_edit = QLineEdit("YES")
        hbox.addWidget(self.hwp_input_edit)
        btn = QPushButton("찾기...")
        btn.clicked.connect(lambda: self.browse_folder(self.hwp_input_edit))
        hbox.addWidget(btn)
        vbox.addLayout(hbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("출력 엑셀 파일:"))
        self.hwp_output_edit = QLineEdit("data/output/output_from_hwp.xlsx")
        hbox.addWidget(self.hwp_output_edit)
        btn = QPushButton("찾기...")
        btn.clicked.connect(lambda: self.browse_save_file(
            self.hwp_output_edit, "Excel Files (*.xlsx)"))
        hbox.addWidget(btn)
        vbox.addLayout(hbox)

        self.hwp_run_btn = QPushButton("HWP → XLSX 변환")
        self.hwp_run_btn.clicked.connect(self.run_hwp_conversion)
        self.hwp_run_btn.setStyleSheet(
            "background-color: #4CAF50; color: white; padding: 8px; font-weight: bold;")
        vbox.addWidget(self.hwp_run_btn)

        group.setLayout(vbox)
        layout.addWidget(group)
        widget.setLayout(layout)
        return widget

    def create_image_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(10, 10, 10, 10)

        title_layout = QHBoxLayout()
        title_label = QLabel("이미지 추출")
        title_label.setFont(QFont("Arial", 11, QFont.Bold))
        title_layout.addWidget(title_label)
        help_btn = QPushButton("?")
        help_btn.setMaximumWidth(30)
        help_btn.clicked.connect(self.show_image_help)
        title_layout.addWidget(help_btn)
        layout.addLayout(title_layout)

        group = QGroupBox("한글파일에서 이미지 추출")
        vbox = QVBoxLayout()
        vbox.setSpacing(6)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("입력 폴더:"))
        self.img_input_edit = QLineEdit("YES")
        hbox.addWidget(self.img_input_edit)
        btn = QPushButton("찾기...")
        btn.clicked.connect(lambda: self.browse_folder(self.img_input_edit))
        hbox.addWidget(btn)
        vbox.addLayout(hbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("출력 폴더:"))
        self.img_output_edit = QLineEdit("img")
        hbox.addWidget(self.img_output_edit)
        btn = QPushButton("찾기...")
        btn.clicked.connect(lambda: self.browse_folder(self.img_output_edit))
        hbox.addWidget(btn)
        vbox.addLayout(hbox)

        self.img_run_btn = QPushButton("이미지 추출")
        self.img_run_btn.clicked.connect(self.run_image_extraction)
        self.img_run_btn.setStyleSheet(
            "background-color: #2196F3; color: white; padding: 8px; font-weight: bold;")
        vbox.addWidget(self.img_run_btn)

        group.setLayout(vbox)
        layout.addWidget(group)
        widget.setLayout(layout)
        return widget

    def create_docx_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(10, 10, 10, 10)

        title_layout = QHBoxLayout()
        title_label = QLabel("워드 문서 생성 / 동기화")
        title_label.setFont(QFont("Arial", 11, QFont.Bold))
        title_layout.addWidget(title_label)
        help_btn = QPushButton("?")
        help_btn.setMaximumWidth(30)
        help_btn.clicked.connect(self.show_docx_help)
        title_layout.addWidget(help_btn)
        layout.addLayout(title_layout)

        group = QGroupBox("엑셀 파일로부터 워드파일 생성")
        vbox = QVBoxLayout()
        vbox.setSpacing(5)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("엑셀 파일:"))
        self.docx_xlsx_edit = QLineEdit("data/output/output_from_hwp.xlsx")
        hbox.addWidget(self.docx_xlsx_edit)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(60)
        btn.clicked.connect(lambda: self.browse_file(
            self.docx_xlsx_edit, "Excel Files (*.xlsx)"))
        hbox.addWidget(btn)
        vbox.addLayout(hbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("템플릿:"))
        self.docx_template_edit = QLineEdit("data/templates/Word_양식.docx")
        hbox.addWidget(self.docx_template_edit)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(60)
        btn.clicked.connect(lambda: self.browse_file(
            self.docx_template_edit, "Word Files (*.docx)"))
        hbox.addWidget(btn)
        vbox.addLayout(hbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("이미지 폴더:"))
        self.docx_img_edit = QLineEdit("img")
        hbox.addWidget(self.docx_img_edit)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(60)
        btn.clicked.connect(lambda: self.browse_folder(self.docx_img_edit))
        hbox.addWidget(btn)
        vbox.addLayout(hbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("출력 폴더:"))
        self.docx_output_edit = QLineEdit("data/output")
        hbox.addWidget(self.docx_output_edit)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(60)
        btn.clicked.connect(lambda: self.browse_folder(self.docx_output_edit))
        hbox.addWidget(btn)
        vbox.addLayout(hbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("행 제한 (0=전체):"))
        self.docx_limit_spin = QSpinBox()
        self.docx_limit_spin.setMinimum(0)
        self.docx_limit_spin.setMaximum(10000)
        self.docx_limit_spin.setValue(0)
        self.docx_limit_spin.setMaximumWidth(80)
        hbox.addWidget(self.docx_limit_spin)
        hbox.addStretch()
        vbox.addLayout(hbox)

        btn_row = QHBoxLayout()
        self.docx_run_btn = QPushButton("DOCX 생성")
        self.docx_run_btn.clicked.connect(self.run_docx_generation)
        self.docx_run_btn.setStyleSheet(
            "background-color: #FF9800; color: white; padding: 8px; font-weight: bold;")
        btn_row.addWidget(self.docx_run_btn)

        self.sync_run_btn = QPushButton("동기화 (파일명+내용 갱신)")
        self.sync_run_btn.clicked.connect(self.run_sync)
        self.sync_run_btn.setStyleSheet(
            "background-color: #607D8B; color: white; padding: 8px; font-weight: bold;")
        btn_row.addWidget(self.sync_run_btn)
        vbox.addLayout(btn_row)

        group.setLayout(vbox)
        layout.addWidget(group)
        widget.setLayout(layout)
        return widget

    def create_history_tab(self):
        """이력 관리 탭"""
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(10, 10, 10, 10)

        title_layout = QHBoxLayout()
        title_label = QLabel("유지보수 이력 관리")
        title_label.setFont(QFont("Arial", 11, QFont.Bold))
        title_layout.addWidget(title_label)
        help_btn = QPushButton("?")
        help_btn.setMaximumWidth(30)
        help_btn.clicked.connect(self.show_history_help)
        title_layout.addWidget(help_btn)
        layout.addLayout(title_layout)

        # DB / 파일 설정 그룹
        file_group = QGroupBox("파일 설정")
        fg = QVBoxLayout()
        fg.setSpacing(4)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("DB 엑셀 파일:"))
        self.hist_xlsx_edit = QLineEdit("data/output/00.DB_19-000.xlsx")
        hbox.addWidget(self.hist_xlsx_edit)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(55)
        btn.clicked.connect(lambda: self.browse_file(
            self.hist_xlsx_edit, "Excel Files (*.xlsx)"))
        hbox.addWidget(btn)
        fg.addLayout(hbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("템플릿 파일:"))
        self.hist_template_edit = QLineEdit("data/templates/Word_양식.docx")
        hbox.addWidget(self.hist_template_edit)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(55)
        btn.clicked.connect(lambda: self.browse_file(
            self.hist_template_edit, "Word Files (*.docx)"))
        hbox.addWidget(btn)
        fg.addLayout(hbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("이미지 폴더:"))
        self.hist_img_edit = QLineEdit("img")
        hbox.addWidget(self.hist_img_edit)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(55)
        btn.clicked.connect(lambda: self.browse_folder(self.hist_img_edit))
        hbox.addWidget(btn)
        fg.addLayout(hbox)

        file_group.setLayout(fg)
        layout.addWidget(file_group)

        # 출력 폴더 선택
        folder_row = QHBoxLayout()
        folder_row.addWidget(QLabel("Word 파일 폴더:"))
        self.hist_dir_edit = QLineEdit("data/output")
        self.hist_dir_edit.textChanged.connect(self.refresh_history_list)
        folder_row.addWidget(self.hist_dir_edit)
        browse_btn = QPushButton("찾기...")
        browse_btn.setMaximumWidth(60)
        browse_btn.clicked.connect(self._browse_history_folder)
        folder_row.addWidget(browse_btn)
        refresh_btn = QPushButton("새로고침")
        refresh_btn.setMaximumWidth(70)
        refresh_btn.clicked.connect(self.refresh_history_list)
        folder_row.addWidget(refresh_btn)
        layout.addLayout(folder_row)

        # 스플리터: 파일 목록 | 이력 편집기
        splitter = QSplitter(Qt.Horizontal)

        # 왼쪽: 파일 목록
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(0, 0, 4, 0)
        left_layout.addWidget(QLabel("이력카드 목록:"))
        self.history_list = QListWidget()
        self.history_list.currentItemChanged.connect(self.on_history_file_selected)
        left_layout.addWidget(self.history_list)
        left_widget.setLayout(left_layout)

        # 오른쪽: 이력 편집기
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(4, 0, 0, 0)
        right_layout.addWidget(QLabel("유지보수 이력 (직접 편집 가능):"))
        self.history_editor = QTextEdit()
        self.history_editor.setPlaceholderText("왼쪽에서 이력카드 파일을 선택하세요.")
        right_layout.addWidget(self.history_editor)

        # 버튼 행
        hist_btn_row = QHBoxLayout()

        self.save_history_btn = QPushButton("저장")
        self.save_history_btn.clicked.connect(self.save_history)
        self.save_history_btn.setStyleSheet(
            "background-color: #4CAF50; color: white; padding: 5px;")
        self.save_history_btn.setEnabled(False)
        hist_btn_row.addWidget(self.save_history_btn)

        self.apply_history_btn = QPushButton("Word에 반영")
        self.apply_history_btn.clicked.connect(self.apply_history_to_word)
        self.apply_history_btn.setStyleSheet(
            "background-color: #FF9800; color: white; padding: 5px;")
        self.apply_history_btn.setEnabled(False)
        hist_btn_row.addWidget(self.apply_history_btn)

        right_layout.addLayout(hist_btn_row)
        right_widget.setLayout(right_layout)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([160, 340])
        layout.addWidget(splitter, 1)

        widget.setLayout(layout)
        return widget

    # -----------------------------------------------------------------------
    # 공통 유틸리티
    # -----------------------------------------------------------------------
    def browse_folder(self, line_edit):
        folder = QFileDialog.getExistingDirectory(self, "폴더 선택", line_edit.text() or ".")
        if folder:
            line_edit.setText(folder)

    def browse_file(self, line_edit, file_filter):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "파일 선택", line_edit.text() or ".", file_filter)
        if file_path:
            line_edit.setText(file_path)

    def browse_save_file(self, line_edit, file_filter):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "파일 저장", line_edit.text() or ".", file_filter)
        if file_path:
            line_edit.setText(file_path)

    def disable_buttons(self):
        self.hwp_run_btn.setEnabled(False)
        self.img_run_btn.setEnabled(False)
        self.docx_run_btn.setEnabled(False)
        self.sync_run_btn.setEnabled(False)
        self.new_card_btn.setEnabled(False)
        self.pdf_run_btn.setEnabled(False)

    def enable_buttons(self):
        self.hwp_run_btn.setEnabled(True)
        self.img_run_btn.setEnabled(True)
        self.docx_run_btn.setEnabled(True)
        self.sync_run_btn.setEnabled(True)
        self.new_card_btn.setEnabled(True)
        self.pdf_run_btn.setEnabled(True)

    def log_message(self, msg):
        self.log_text.append(msg)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum())

    def show_error(self, error_msg):
        self.log_message(f"✗ 오류: {error_msg}")
        QMessageBox.critical(self, "오류", error_msg)

    def on_task_finished(self):
        self.enable_buttons()

    # -----------------------------------------------------------------------
    # 작업 실행
    # -----------------------------------------------------------------------
    def run_hwp_conversion(self):
        input_dir = Path(self.hwp_input_edit.text())
        output_file = Path(self.hwp_output_edit.text())
        if not input_dir.exists():
            self.show_error("입력 폴더가 존재하지 않습니다.")
            return
        self.disable_buttons()

        def task():
            try:
                HWPProcessor.process(
                    input_dir, output_file,
                    callback=lambda msg: self.signals.log.emit(msg))
                self.signals.log.emit(f"✓ HWP → XLSX 변환 완료: {output_file.name}")
            except Exception as e:
                self.signals.log.emit(f"✗ 오류: {e}")
            self.signals.finished.emit()

        Thread(target=task, daemon=True).start()

    def run_image_extraction(self):
        input_dir = Path(self.img_input_edit.text())
        output_dir = Path(self.img_output_edit.text())
        if not input_dir.exists():
            self.show_error("입력 폴더가 존재하지 않습니다.")
            return
        self.disable_buttons()

        def task():
            try:
                hwp_files = sorted(input_dir.glob("*.hwp"))
                total = 0
                for hwp_file in hwp_files:
                    extractor = HWPImageExtractor(str(hwp_file), str(output_dir))
                    n = extractor.extract_images()
                    if n:
                        total += n
                        self.signals.log.emit(f"  {hwp_file.name}: {n}개 추출")
                self.signals.log.emit(f"✓ 이미지 추출 완료: 총 {total}개")
            except Exception as e:
                self.signals.log.emit(f"✗ 오류: {e}")
            self.signals.finished.emit()

        Thread(target=task, daemon=True).start()

    def run_docx_generation(self):
        xlsx_file = Path(self.docx_xlsx_edit.text())
        template = Path(self.docx_template_edit.text())
        output_dir = Path(self.docx_output_edit.text())
        img_dir = Path(self.docx_img_edit.text())
        limit = self.docx_limit_spin.value()
        if not xlsx_file.exists():
            self.show_error("엑셀 파일이 존재하지 않습니다.")
            return
        if not template.exists():
            self.show_error("템플릿 파일이 존재하지 않습니다.")
            return
        self.disable_buttons()

        def task():
            try:
                DocumentFiller.process(
                    xlsx_file, template, output_dir, img_dir, limit,
                    callback=lambda msg: self.signals.log.emit(msg))
                self.signals.log.emit("✓ DOCX 생성 완료")
            except Exception as e:
                self.signals.log.emit(f"✗ 오류: {e}")
            self.signals.finished.emit()

        Thread(target=task, daemon=True).start()

    def run_sync(self):
        xlsx_file = Path(self.docx_xlsx_edit.text())
        template = Path(self.docx_template_edit.text())
        output_dir = Path(self.docx_output_edit.text())
        img_dir = Path(self.docx_img_edit.text())
        if not xlsx_file.exists():
            self.show_error("엑셀 파일이 존재하지 않습니다.")
            return
        if not template.exists():
            self.show_error("템플릿 파일이 존재하지 않습니다.")
            return
        self.disable_buttons()

        def task():
            try:
                DocxSyncManager.sync(
                    xlsx_file, template, output_dir, img_dir,
                    callback=lambda msg: self.signals.log.emit(msg))
                self.signals.log.emit("✓ 동기화 완료")
            except Exception as e:
                self.signals.log.emit(f"✗ 오류: {e}")
            self.signals.finished.emit()

        Thread(target=task, daemon=True).start()

    # -----------------------------------------------------------------------
    # 이력 관리
    # -----------------------------------------------------------------------
    def _browse_history_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self, "폴더 선택", self.hist_dir_edit.text() or ".")
        if folder:
            self.hist_dir_edit.setText(folder)
            # textChanged 신호로 refresh_history_list 자동 호출됨

    def refresh_history_list(self):
        self.history_list.clear()
        self.history_editor.clear()
        self.save_history_btn.setEnabled(False)
        self.apply_history_btn.setEnabled(False)
        output_dir = Path(self.hist_dir_edit.text())
        if not output_dir.exists():
            return
        docx_files = sorted(output_dir.glob("*.docx"))
        for f in docx_files:
            if not f.stem.endswith("_history"):
                self.history_list.addItem(f.name)
        self.log_message(f"{len(docx_files)}개 파일 로드됨 ({output_dir})")

    def on_history_file_selected(self, current, _):
        """목록에서 파일 선택 시 이력 내용 표시"""
        if current is None:
            return
        output_dir = Path(self.hist_dir_edit.text())
        docx_path = output_dir / current.text()
        content = MaintenanceHistoryManager.read_history(docx_path)
        self.history_editor.setPlainText(content)
        self.save_history_btn.setEnabled(True)
        self.apply_history_btn.setEnabled(True)

    def save_history(self):
        """편집기 내용을 _history.txt 파일에 저장하고 XLSX DB 갱신"""
        current = self.history_list.currentItem()
        if current is None:
            return
        output_dir = Path(self.hist_dir_edit.text())
        docx_path = output_dir / current.text()
        content = self.history_editor.toPlainText()
        MaintenanceHistoryManager.write_history(docx_path, content)
        self.log_message(f"✓ 이력 저장됨: {docx_path.stem}_history.txt")

        xlsx_path = Path(self.hist_xlsx_edit.text())
        if xlsx_path.exists():
            MaintenanceHistoryManager.update_xlsx_reason(
                docx_path, xlsx_path, content,
                callback=lambda msg: self.log_message(msg))
        else:
            self.log_message(
                "  (XLSX 미설정 — 위 'DB 엑셀 파일' 경로를 지정하면 DB도 함께 갱신됩니다)")

    def apply_history_to_word(self):
        """이력 내용을 Word 파일에 반영 (XLSX 업데이트 + Word 재생성)"""
        current = self.history_list.currentItem()
        if current is None:
            return
        output_dir = Path(self.hist_dir_edit.text())
        docx_path = output_dir / current.text()
        xlsx_path = Path(self.hist_xlsx_edit.text())
        template_path = Path(self.hist_template_edit.text())
        img_dir = Path(self.hist_img_edit.text())

        if not xlsx_path.exists():
            self.show_error("이력 관리 탭의 'DB 엑셀 파일'을 먼저 설정하세요.")
            return
        if not template_path.exists():
            self.show_error("이력 관리 탭의 '템플릿 파일'을 먼저 설정하세요.")
            return

        self.disable_buttons()

        def task():
            try:
                ok = MaintenanceHistoryManager.apply_to_word(
                    docx_path, xlsx_path, template_path, img_dir,
                    callback=lambda msg: self.signals.log.emit(msg))
                if ok:
                    self.signals.log.emit(f"✓ Word 반영 완료: {docx_path.name}")
                else:
                    self.signals.log.emit("✗ 오류: Word 반영 실패 (로그 확인)")
            except Exception as e:
                self.signals.log.emit(f"✗ 오류: {e}")
            self.signals.finished.emit()

        Thread(target=task, daemon=True).start()

    # -----------------------------------------------------------------------
    # 신규 이력카드 발행
    # -----------------------------------------------------------------------
    def show_new_card_dialog(self):
        init_xlsx = self.docx_xlsx_edit.text() if hasattr(self, 'docx_xlsx_edit') else ""
        dlg = NewCardDialog(self, xlsx_path=init_xlsx if init_xlsx else None)
        if dlg.exec_() != QDialog.Accepted:
            return
        row_data = dlg.get_data()
        image_path = dlg.get_image_path()
        if not row_data:
            return

        xlsx_path = Path(dlg.dlg_xlsx_edit.text())
        template_path = Path(self.docx_template_edit.text())
        output_dir = Path(self.docx_output_edit.text())
        img_dir = Path(self.docx_img_edit.text())

        if not xlsx_path.exists():
            self.show_error("신규 발행 다이얼로그에서 DB 엑셀 파일을 올바르게 설정하세요.")
            return
        if not template_path.exists():
            self.show_error("문서 생성/동기화 탭에서 템플릿 파일을 먼저 설정하세요.")
            return

        self.disable_buttons()

        def task():
            try:
                out_path = NewCardManager.generate_card(
                    xlsx_path, template_path, output_dir, img_dir, row_data,
                    image_source_path=image_path,
                    callback=lambda msg: self.signals.log.emit(msg))
                if out_path:
                    self.signals.log.emit(
                        f"✓ 신규 이력카드 발행 완료: {Path(out_path).name}")
                else:
                    self.signals.log.emit("✗ 오류: 신규 이력카드 발행 실패")
            except Exception as e:
                self.signals.log.emit(f"✗ 오류: {e}")
            self.signals.finished.emit()

        Thread(target=task, daemon=True).start()

    # -----------------------------------------------------------------------
    # 도움말
    # -----------------------------------------------------------------------
    def show_hwp_help(self):
        help_text = """
1단계: HWP → XLSX 변환

HWP 파일들이 있는 폴더를 선택하여 XLSX 로 일괄 변환합니다.

【사용 방법】
1. 입력 폴더: HWP 파일들이 있는 폴더 선택 (기본: YES/)
2. 출력 엑셀 파일: 결과 XLSX 저장 경로 설정
3. 'HWP → XLSX 변환' 버튼 클릭
        """
        QMessageBox.information(self, "1단계 도움말", help_text.strip())

    def show_image_help(self):
        help_text = """
2단계: 이미지 추출

HWP 파일에 포함된 이미지를 추출하여 img/ 폴더에 저장합니다.

【저장 형식】
파일명: 品名_図番.확장자
(예: CASE_1071000024.jpg)

【사용 방법】
1. 입력 폴더: HWP 파일들이 있는 폴더 선택
2. 출력 폴더: 이미지를 저장할 폴더 선택
3. '이미지 추출' 버튼 클릭
        """
        QMessageBox.information(self, "2단계 도움말", help_text.strip())

    def show_docx_help(self):
        help_text = """
3단계: 워드 문서 생성 / 동기화

【DOCX 생성】
엑셀 데이터를 템플릿에 입력하여 Word 파일 일괄 생성.
- 행 제한 0 = 전체 처리

【동기화 (파일명+내용 갱신)】
엑셀에서 데이터를 수정한 후 기존 Word 파일에 반영합니다.

동기화 처리 내용:
1. File name 변경 감지 → Word 파일 rename
2. 연번 변경 시 이미지 파일도 rename
3. 모든 Word 파일 최신 엑셀 데이터로 내용 갱신
4. manifest.json 저장 (관리번호 기반 추적)

【주의】
동기화는 管理番号 컬럼을 기준으로 파일을 추적합니다.
        """
        QMessageBox.information(self, "3단계 도움말", help_text.strip())

    # -----------------------------------------------------------------------
    # 5단계: PDF 변환/병합
    # -----------------------------------------------------------------------
    def create_pdf_tab(self):
        """PDF 변환/병합 탭"""
        widget = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(10, 10, 10, 10)

        title_layout = QHBoxLayout()
        title_label = QLabel("5단계: Word → PDF 변환 / 병합")
        title_label.setFont(QFont("Arial", 11, QFont.Bold))
        title_layout.addWidget(title_label)
        help_btn = QPushButton("?")
        help_btn.setMaximumWidth(30)
        help_btn.clicked.connect(self.show_pdf_help)
        title_layout.addWidget(help_btn)
        layout.addLayout(title_layout)

        # 모드 선택
        mode_group = QGroupBox("변환 모드")
        mode_layout = QHBoxLayout()
        self.pdf_mode_group = QButtonGroup(self)
        for idx, label in enumerate(["단일 파일 변환", "일괄 변환 (폴더)", "변환 후 병합"]):
            rb = QRadioButton(label)
            self.pdf_mode_group.addButton(rb, idx)
            mode_layout.addWidget(rb)
        self.pdf_mode_group.button(0).setChecked(True)
        self.pdf_mode_group.buttonClicked.connect(self._on_pdf_mode_changed)
        mode_group.setLayout(mode_layout)
        layout.addWidget(mode_group)

        # ── 모드 0: 단일 변환 ──────────────────────────────────
        self.pdf_single_group = QGroupBox("단일 파일 변환")
        sg = QVBoxLayout()
        sg.setSpacing(5)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("Word 파일:"))
        self.pdf_single_input = QLineEdit()
        self.pdf_single_input.setPlaceholderText("변환할 .docx 파일 선택")
        hbox.addWidget(self.pdf_single_input)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(55)
        btn.clicked.connect(lambda: self.browse_file(
            self.pdf_single_input, "Word Files (*.docx)"))
        hbox.addWidget(btn)
        sg.addLayout(hbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("출력 PDF:"))
        self.pdf_single_output = QLineEdit()
        self.pdf_single_output.setPlaceholderText("저장할 .pdf 경로 (비우면 원본 폴더)")
        hbox.addWidget(self.pdf_single_output)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(55)
        btn.clicked.connect(lambda: self.browse_save_file(
            self.pdf_single_output, "PDF Files (*.pdf)"))
        hbox.addWidget(btn)
        sg.addLayout(hbox)

        self.pdf_single_group.setLayout(sg)
        layout.addWidget(self.pdf_single_group)

        # ── 모드 1: 일괄 변환 ──────────────────────────────────
        self.pdf_batch_group = QGroupBox("일괄 변환 (폴더 내 모든 .docx)")
        bg = QVBoxLayout()
        bg.setSpacing(5)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("입력 폴더:"))
        self.pdf_batch_input = QLineEdit("data/output")
        hbox.addWidget(self.pdf_batch_input)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(55)
        btn.clicked.connect(lambda: self.browse_folder(self.pdf_batch_input))
        hbox.addWidget(btn)
        bg.addLayout(hbox)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("출력 폴더:"))
        self.pdf_batch_output = QLineEdit("data/output_pdf")
        hbox.addWidget(self.pdf_batch_output)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(55)
        btn.clicked.connect(lambda: self.browse_folder(self.pdf_batch_output))
        hbox.addWidget(btn)
        bg.addLayout(hbox)

        self.pdf_recursive_chk = QCheckBox("하위 폴더 포함")
        bg.addWidget(self.pdf_recursive_chk)
        self.pdf_batch_group.setLayout(bg)
        layout.addWidget(self.pdf_batch_group)

        # ── 모드 2: 변환 후 병합 ───────────────────────────────
        self.pdf_merge_group = QGroupBox("Word 파일들을 PDF로 변환 후 하나로 병합")
        mg = QVBoxLayout()
        mg.setSpacing(5)

        list_btn_row = QHBoxLayout()
        add_btn = QPushButton("파일 추가...")
        add_btn.clicked.connect(self._pdf_add_files)
        list_btn_row.addWidget(add_btn)
        remove_btn = QPushButton("선택 제거")
        remove_btn.clicked.connect(self._pdf_remove_files)
        list_btn_row.addWidget(remove_btn)
        clear_btn = QPushButton("전체 지우기")
        clear_btn.clicked.connect(lambda: self.pdf_merge_list.clear())
        list_btn_row.addWidget(clear_btn)
        list_btn_row.addStretch()
        mg.addLayout(list_btn_row)

        self.pdf_merge_list = QListWidget()
        self.pdf_merge_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.pdf_merge_list.setMaximumHeight(110)
        mg.addWidget(self.pdf_merge_list)

        hbox = QHBoxLayout()
        hbox.addWidget(QLabel("출력 PDF:"))
        self.pdf_merge_output = QLineEdit("data/output_pdf/merged.pdf")
        hbox.addWidget(self.pdf_merge_output)
        btn = QPushButton("찾기...")
        btn.setMaximumWidth(55)
        btn.clicked.connect(lambda: self.browse_save_file(
            self.pdf_merge_output, "PDF Files (*.pdf)"))
        hbox.addWidget(btn)
        mg.addLayout(hbox)

        self.pdf_merge_group.setLayout(mg)
        layout.addWidget(self.pdf_merge_group)

        # 실행 버튼
        self.pdf_run_btn = QPushButton("▶ 실행")
        self.pdf_run_btn.clicked.connect(self.run_pdf)
        self.pdf_run_btn.setStyleSheet(
            "background-color: #E91E63; color: white; padding: 8px; font-weight: bold;")
        layout.addWidget(self.pdf_run_btn)

        layout.addStretch()
        widget.setLayout(layout)

        # 초기 모드 표시
        self._on_pdf_mode_changed(None)
        return widget

    def _on_pdf_mode_changed(self, _):
        mode = self.pdf_mode_group.checkedId()
        self.pdf_single_group.setVisible(mode == 0)
        self.pdf_batch_group.setVisible(mode == 1)
        self.pdf_merge_group.setVisible(mode == 2)

    def _pdf_add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Word 파일 선택", ".", "Word Files (*.docx)")
        for f in files:
            self.pdf_merge_list.addItem(f)

    def _pdf_remove_files(self):
        for item in self.pdf_merge_list.selectedItems():
            self.pdf_merge_list.takeItem(self.pdf_merge_list.row(item))

    def run_pdf(self):
        mode = self.pdf_mode_group.checkedId()
        if mode == 0:
            self._run_pdf_single()
        elif mode == 1:
            self._run_pdf_batch()
        else:
            self._run_pdf_merge()

    def _run_pdf_single(self):
        input_path = Path(self.pdf_single_input.text())
        if not input_path.exists():
            self.show_error("Word 파일이 존재하지 않습니다.")
            return
        output_text = self.pdf_single_output.text().strip()
        output_path = Path(output_text) if output_text else None
        self.disable_buttons()

        def task():
            try:
                from src.pdf import docx_to_pdf
                result = docx_to_pdf(input_path, output_path)
                if result:
                    self.signals.log.emit(f"✓ PDF 변환 완료: {Path(result).name}")
                else:
                    self.signals.log.emit("✗ PDF 변환 실패 (Word가 설치되어 있는지 확인)")
            except Exception as e:
                self.signals.log.emit(f"✗ 오류: {e}")
            self.signals.finished.emit()

        Thread(target=task, daemon=True).start()

    def _run_pdf_batch(self):
        input_dir = Path(self.pdf_batch_input.text())
        if not input_dir.exists():
            self.show_error("입력 폴더가 존재하지 않습니다.")
            return
        output_dir = Path(self.pdf_batch_output.text())
        recursive = self.pdf_recursive_chk.isChecked()
        self.disable_buttons()

        def task():
            try:
                from src.pdf import batch_docx_to_pdf
                results = batch_docx_to_pdf(input_dir, output_dir, recursive)
                self.signals.log.emit(f"✓ 일괄 변환 완료: {len(results)}개 → {output_dir}")
            except Exception as e:
                self.signals.log.emit(f"✗ 오류: {e}")
            self.signals.finished.emit()

        Thread(target=task, daemon=True).start()

    def _run_pdf_merge(self):
        count = self.pdf_merge_list.count()
        if count == 0:
            self.show_error("'파일 추가...' 버튼으로 Word 파일을 목록에 추가하세요.")
            return
        output_path = Path(self.pdf_merge_output.text())
        docx_files = [
            Path(self.pdf_merge_list.item(i).text()) for i in range(count)
        ]
        self.disable_buttons()

        def task():
            try:
                from src.pdf import convert_and_merge
                ok = convert_and_merge(docx_files, output_path)
                if ok:
                    self.signals.log.emit(f"✓ 변환·병합 완료: {output_path.name}")
                else:
                    self.signals.log.emit("✗ 변환·병합 실패 (로그 확인)")
            except Exception as e:
                self.signals.log.emit(f"✗ 오류: {e}")
            self.signals.finished.emit()

        Thread(target=task, daemon=True).start()

    def show_history_help(self):
        help_text = """
4단계: 유지보수 이력 관리

【목표】
각 이력카드에 대한 금형 유지보수 이력을 txt 파일로 관리합니다.
이력 파일은 Word 파일 폴더 내 .data/ 서브폴더에 저장됩니다.

【사용 방법】
1. Word 파일 폴더 선택 후 목록에서 이력카드 파일 선택
2. 오른쪽 편집기에서 이력 직접 작성/수정
3. '저장': _history.txt 저장 + XLSX 事由 컬럼 즉시 갱신
4. 'Word에 반영': XLSX 갱신 후 Word 파일 재생성

【버튼 설명】
- 저장: 편집기 내용을 _history.txt 파일에 저장 + DB 갱신
- Word에 반영: 이력 내용을 XLSX의 {事  由} 필드에 반영 후 Word 재생성
  (3단계 탭의 엑셀 파일/템플릿 설정이 필요합니다)
        """
        QMessageBox.information(self, "4단계 도움말", help_text.strip())


    def show_pdf_help(self):
        help_text = """
5단계: Word → PDF 변환 / 병합

【모드 설명】
① 단일 파일 변환
   Word 파일 1개를 PDF로 변환합니다.
   출력 경로를 비워두면 원본과 같은 폴더에 저장됩니다.

② 일괄 변환 (폴더)
   지정한 폴더 내의 모든 .docx 파일을 PDF로 일괄 변환합니다.
   '하위 폴더 포함' 체크 시 재귀적으로 처리합니다.

③ 변환 후 병합
   여러 Word 파일을 PDF로 변환한 뒤 하나의 PDF로 합칩니다.
   목록 순서대로 병합됩니다.

【사전 요구사항】
- Windows: Microsoft Word 설치 필요 (comtypes 사용)
- PDF 병합: pip install pypdf
        """
        QMessageBox.information(self, "5단계 도움말", help_text.strip())


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
