"""PyQt6 기반 PPT → PDF 변환 GUI.

- 드래그앤드롭 / 파일 선택
- 1280 가로 고정 해상도 (옵션 없음)
- JPEG 재압축 체크박스
- QThread로 변환 분리, 진행률 바
- 완료 후 최종 용량 표시 및 폴더 열기
"""
from __future__ import annotations

import os
import subprocess
import sys
import traceback

from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtWidgets import (
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

from converter import convert_ppt_to_pdf


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
            self.failed.emit(f"{e}\n\n{traceback.format_exc()}")


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
        self.setWindowTitle("PPT → PDF 변환기 (1280×720, 폰트 보존)")
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
            "슬라이드 이미지는 가로 1280px 고정입니다 (16:9 → 720, 4:3 → 960)."
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
        row3.addWidget(self.btn_convert)
        row3.addWidget(self.btn_open)
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

    def _on_failed(self, msg: str) -> None:
        self.btn_convert.setEnabled(True)
        self.status.setText("실패")
        QMessageBox.critical(
            self,
            "변환 실패",
            "PowerPoint 미설치이거나 파일이 손상되었을 수 있습니다.\n\n" + msg,
        )

    def _open_folder(self) -> None:
        if not self._last_pdf:
            return
        folder = os.path.dirname(os.path.abspath(self._last_pdf))
        if sys.platform == "win32":
            subprocess.Popen(["explorer", folder])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])


def main() -> int:
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
