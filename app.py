"""PyQt6 기반 PPT -> PDF 변환 GUI.

- 드래그앤드롭 / 파일 선택
- 1280 가로 고정 해상도 (옵션 없음)
- JPEG 재압축 체크박스
- QThread로 변환 분리, 진행률 바
- 완료 후 최종 용량 표시 및 폴더 열기

보안 메모:
    - 예외 trackback은 UI에 노출하지 않고 %APPDATA%\\ppt2pdf\\error.log 로 기록.
      사용자는 "로그 보기" 버튼으로만 전체 내용을 확인.
    - 출력 PDF가 이미 존재하면 덮어쓰기 전에 반드시 확인을 받는다.

디버깅:
    - GUI 가 뜨기도 전에 죽는 startup 크래시는 %APPDATA%\\ppt2pdf\\startup_crash.log
      에 기록된다 (import·QApplication 생성 포함 전역 try 로 감쌌다).
    - 콘솔을 보고 싶으면 `build_debug.bat` 로 `--console` 모드 exe 를 빌드해
      cmd.exe 에서 실행하면 stderr 가 그대로 출력된다.
    - 네이티브 크래시(세그폴트 류)는 faulthandler 가 stderr 및 로그에 남긴다.
"""
from __future__ import annotations

import faulthandler
import logging
import os
import subprocess
import sys
import traceback
from logging.handlers import RotatingFileHandler


def _log_dir() -> str:
    """플랫폼별 로그 디렉터리 (Windows: %APPDATA%\\ppt2pdf, 기타: ~/.ppt2pdf)."""
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, "ppt2pdf")
    return os.path.join(os.path.expanduser("~"), ".ppt2pdf")


LOG_PATH = os.path.join(_log_dir(), "error.log")
STARTUP_CRASH_LOG = os.path.join(_log_dir(), "startup_crash.log")

# faulthandler 에 넘긴 파일은 프로세스 전 생애 동안 열려 있어야 네이티브
# 크래시가 기록되므로, GC 로 닫히지 않게 모듈 전역에서 참조를 유지한다.
_CRASH_FH = None


def _install_crash_logging() -> None:
    """PyQt6 / converter import 전에 먼저 호출해야 import 실패도 기록된다."""
    global _CRASH_FH
    try:
        os.makedirs(_log_dir(), exist_ok=True)
    except Exception:
        pass

    try:
        _CRASH_FH = open(STARTUP_CRASH_LOG, "a", encoding="utf-8", buffering=1)
        _CRASH_FH.write("\n===== fault handler armed =====\n")
        faulthandler.enable(file=_CRASH_FH)
    except Exception:
        _CRASH_FH = None

    def _hook(exc_type, exc, tb):
        try:
            with open(STARTUP_CRASH_LOG, "a", encoding="utf-8") as f:
                f.write("\n===== uncaught exception =====\n")
                traceback.print_exception(exc_type, exc, tb, file=f)
        except Exception:
            pass
        sys.__excepthook__(exc_type, exc, tb)

    sys.excepthook = _hook


# PyQt6 / converter import 자체가 실패하는 "GUI 가 아예 안 뜨는" 시나리오도
# 잡기 위해 가장 먼저 크래시 로깅을 설치한다.
_install_crash_logging()

from PyQt6.QtCore import Qt, QThread, pyqtSignal  # noqa: E402
from PyQt6.QtGui import QDragEnterEvent, QDropEvent  # noqa: E402
from PyQt6.QtWidgets import (  # noqa: E402
    QApplication,
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from converter import convert_ppt_to_pdf  # noqa: E402


def _init_logger() -> logging.Logger:
    os.makedirs(_log_dir(), exist_ok=True)
    logger = logging.getLogger("ppt2pdf")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    handler = RotatingFileHandler(
        LOG_PATH, maxBytes=512 * 1024, backupCount=2, encoding="utf-8"
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    )
    logger.addHandler(handler)
    return logger


_LOG = _init_logger()


def _summarize_error(exc: BaseException) -> str:
    """사용자에게 보일 한 줄 요약. 내부 경로/스택은 노출하지 않는다."""
    name = type(exc).__name__
    msg = str(exc).strip() or "(메시지 없음)"
    # 메시지 자체에 경로가 섞이는 경우가 있어 160자로 제한.
    if len(msg) > 160:
        msg = msg[:157] + "..."
    return f"{name}: {msg}"


class ConvertWorker(QThread):
    progress = pyqtSignal(int, int)
    finished_ok = pyqtSignal(str, int)  # (pdf_path, size_bytes)
    failed = pyqtSignal(str)

    def __init__(self, ppt: str, pdf: str, jpeg: bool) -> None:
        super().__init__()
        self._ppt = ppt
        self._pdf = pdf
        self._jpeg = jpeg

    def run(self) -> None:
        try:
            size = convert_ppt_to_pdf(
                self._ppt,
                self._pdf,
                jpeg_fallback=self._jpeg,
                progress_cb=lambda c, t: self.progress.emit(c, t),
            )
            self.finished_ok.emit(self._pdf, size)
        except Exception as e:
            _LOG.error("변환 실패: %s", traceback.format_exc())
            self.failed.emit(_summarize_error(e))


class DropLineEdit(QLineEdit):
    """PPT 파일을 드롭 받아 경로를 채우는 입력란."""

    def __init__(self, placeholder: str) -> None:
        super().__init__()
        self.setPlaceholderText(placeholder)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, e: QDragEnterEvent) -> None:
        if e.mimeData().hasUrls():
            for u in e.mimeData().urls():
                if u.toLocalFile().lower().endswith((".ppt", ".pptx")):
                    e.acceptProposedAction()
                    return
        e.ignore()

    def dropEvent(self, e: QDropEvent) -> None:
        for u in e.mimeData().urls():
            path = u.toLocalFile()
            if path.lower().endswith((".ppt", ".pptx")):
                self.setText(path)
                e.acceptProposedAction()
                return


class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("PPT -> PDF 변환기 (1280x720, 폰트 보존)")
        self.resize(560, 260)
        self._worker: ConvertWorker | None = None
        self._last_pdf: str | None = None
        self._build_ui()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("PPT 파일 (드래그앤드롭 가능):"))
        row1 = QHBoxLayout()
        self.ppt_edit = DropLineEdit("변환할 .ppt 또는 .pptx 파일 경로")
        btn_browse = QPushButton("찾아보기…")
        btn_browse.clicked.connect(self._pick_ppt)
        row1.addWidget(self.ppt_edit)
        row1.addWidget(btn_browse)
        layout.addLayout(row1)

        layout.addWidget(QLabel("출력 PDF 경로:"))
        row2 = QHBoxLayout()
        self.pdf_edit = QLineEdit()
        self.pdf_edit.setPlaceholderText("비워두면 입력 파일과 같은 폴더에 저장")
        btn_out = QPushButton("저장 위치…")
        btn_out.clicked.connect(self._pick_pdf)
        row2.addWidget(self.pdf_edit)
        row2.addWidget(btn_out)
        layout.addLayout(row2)

        self.jpeg_chk = QCheckBox("용량 최소화 (JPEG 품질 85로 재압축)")
        layout.addWidget(self.jpeg_chk)

        info = QLabel(
            "슬라이드 이미지는 가로 1280px 고정입니다 (16:9 -> 720, 4:3 -> 960)."
        )
        info.setStyleSheet("color: #666;")
        layout.addWidget(info)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.status = QLabel("대기 중")
        layout.addWidget(self.status)

        row3 = QHBoxLayout()
        self.btn_convert = QPushButton("변환 시작")
        self.btn_convert.clicked.connect(self._start)
        self.btn_open = QPushButton("폴더 열기")
        self.btn_open.setEnabled(False)
        self.btn_open.clicked.connect(self._open_folder)
        self.btn_log = QPushButton("로그 보기")
        self.btn_log.clicked.connect(self._open_log)
        row3.addWidget(self.btn_convert)
        row3.addWidget(self.btn_open)
        row3.addWidget(self.btn_log)
        layout.addLayout(row3)

    def _pick_ppt(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "PPT 파일 선택", "", "PowerPoint (*.ppt *.pptx)"
        )
        if path:
            self.ppt_edit.setText(path)

    def _pick_pdf(self) -> None:
        src = self.ppt_edit.text().strip()
        default = os.path.splitext(src)[0] + ".pdf" if src else ""
        path, _ = QFileDialog.getSaveFileName(
            self, "PDF 저장 위치", default, "PDF (*.pdf)"
        )
        if path:
            self.pdf_edit.setText(path)

    def _start(self) -> None:
        ppt = self.ppt_edit.text().strip()
        if not ppt or not os.path.isfile(ppt):
            QMessageBox.warning(self, "입력 필요", "유효한 PPT 파일을 선택하세요.")
            return

        pdf = self.pdf_edit.text().strip()
        if not pdf:
            pdf = os.path.splitext(ppt)[0] + ".pdf"

        if os.path.exists(pdf):
            reply = QMessageBox.question(
                self,
                "덮어쓰기 확인",
                f"이미 존재하는 파일입니다:\n{pdf}\n\n덮어쓸까요?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if reply != QMessageBox.StandardButton.Yes:
                return

        self.btn_convert.setEnabled(False)
        self.btn_open.setEnabled(False)
        self.progress.setValue(0)
        self.status.setText("PowerPoint 실행 중…")

        self._worker = ConvertWorker(ppt, pdf, self.jpeg_chk.isChecked())
        self._worker.progress.connect(self._on_progress)
        self._worker.finished_ok.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._worker.start()

    def _on_progress(self, cur: int, total: int) -> None:
        self.progress.setMaximum(total)
        self.progress.setValue(cur)
        self.status.setText(f"슬라이드 내보내는 중… {cur}/{total}")

    def _on_done(self, pdf: str, size: int) -> None:
        self._last_pdf = pdf
        mb = size / (1024 * 1024)
        self.status.setText(f"완료 — {os.path.basename(pdf)}  ({mb:.2f} MB)")
        self.btn_convert.setEnabled(True)
        self.btn_open.setEnabled(True)

        if mb > 10 and not self.jpeg_chk.isChecked():
            QMessageBox.information(
                self,
                "용량 안내",
                f"PDF 용량이 {mb:.1f} MB 로 목표(10MB 이하)를 초과합니다.\n"
                "'용량 최소화' 옵션을 켜고 다시 변환하면 더 작아집니다.",
            )

    def _on_failed(self, summary: str) -> None:
        self.btn_convert.setEnabled(True)
        self.status.setText("실패")
        QMessageBox.critical(
            self,
            "변환 실패",
            f"{summary}\n\n자세한 내용은 '로그 보기'를 확인하세요.",
        )

    def _open_folder(self) -> None:
        if not self._last_pdf:
            return
        folder = os.path.dirname(os.path.abspath(self._last_pdf))
        self._open_path(folder)

    def _open_log(self) -> None:
        if not os.path.isfile(LOG_PATH):
            QMessageBox.information(self, "로그", "아직 기록된 로그가 없습니다.")
            return
        self._open_path(LOG_PATH)

    @staticmethod
    def _open_path(path: str) -> None:
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])


def main() -> int:
    try:
        app = QApplication(sys.argv)
        w = MainWindow()
        w.show()
        return app.exec()
    except BaseException:
        # 여기까지 왔으면 QApplication/MainWindow 생성 전후에서 치명적 오류.
        try:
            with open(STARTUP_CRASH_LOG, "a", encoding="utf-8") as f:
                f.write("\n===== main() crash =====\n")
                traceback.print_exc(file=f)
        except Exception:
            pass
        raise


if __name__ == "__main__":
    sys.exit(main())
