"""ppt2pdf 환경 진단 도구 (doctor).

사용자 PC 에서 변환이 실패할 때 원인을 빠르게 좁히기 위한 CLI 점검기.
점검 항목:
    1. Python / OS / CPU 아키텍처
    2. 주요 의존 패키지 (pywin32, PyQt6, img2pdf, Pillow) import + 버전
    3. PowerPoint COM Dispatch / 버전 / AutomationSecurity 설정 가능 여부
    4. %APPDATA%\\ppt2pdf 쓰기 권한
    5. img2pdf 최소 변환 smoke test (1x1 PNG -> PDF)

모든 결과를 콘솔에 출력하고, 동시에 %APPDATA%\\ppt2pdf\\doctor_report.txt
에 저장한다. 사용자로부터 파일을 받으면 원인 파악이 빨라진다.
"""
from __future__ import annotations

import io
import os
import platform
import sys
import tempfile
import traceback
from datetime import datetime
from typing import Callable

OK = "[ OK ]"
WARN = "[WARN]"
FAIL = "[FAIL]"


def _log_dir() -> str:
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, "ppt2pdf")
    return os.path.join(os.path.expanduser("~"), ".ppt2pdf")


REPORT_PATH = os.path.join(_log_dir(), "doctor_report.txt")


class Reporter:
    """콘솔 + 파일 동시 출력."""

    def __init__(self) -> None:
        self.buf = io.StringIO()
        self.fail_count = 0
        self.warn_count = 0

    def line(self, s: str = "") -> None:
        print(s)
        self.buf.write(s + "\n")

    def section(self, title: str) -> None:
        self.line("")
        self.line(f"--- {title} ---")

    def result(self, status: str, label: str, detail: str = "") -> None:
        line = f"{status} {label}"
        if detail:
            line += f"  | {detail}"
        self.line(line)
        if status == FAIL:
            self.fail_count += 1
        elif status == WARN:
            self.warn_count += 1

    def exc(self, label: str, e: BaseException) -> None:
        self.result(FAIL, label, f"{type(e).__name__}: {e}")
        self.buf.write(traceback.format_exc() + "\n")

    def save(self, path: str) -> None:
        try:
            os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.buf.getvalue())
        except Exception as e:
            print(f"(리포트 저장 실패: {e})")


def check_system(r: Reporter) -> None:
    r.section("System")
    r.line(f"Python        : {sys.version.splitlines()[0]}")
    r.line(f"Executable    : {sys.executable}")
    r.line(f"Platform      : {platform.platform()}")
    r.line(f"Architecture  : {platform.machine()}  ({platform.architecture()[0]})")
    r.line(f"Frozen (exe?) : {getattr(sys, 'frozen', False)}")
    if sys.platform != "win32":
        r.result(WARN, "Windows 여부", f"현재 플랫폼은 {sys.platform}. 이 앱은 Windows 전용")
    else:
        r.result(OK, "Windows 플랫폼")


def check_packages(r: Reporter) -> None:
    """의존 패키지 import 상태 + 버전을 보고.

    중요: 각 `import X` 는 정적 문장으로 적어야 PyInstaller 가 번들에 포함한다.
    `__import__("X")` 같은 동적 호출은 frozen 실행 파일에 모듈이 빠져
    healthy 빌드에서도 [FAIL] 로 잘못 보고된다.
    """
    r.section("Dependencies")

    # pywin32
    try:
        import pywintypes  # noqa: F401
        import win32com.client  # noqa: F401
        version = getattr(pywintypes, "__version__", "unknown")
        r.result(OK, "import pywin32", f"pywintypes {version}")
    except Exception as e:
        r.exc("import pywin32", e)

    # PyQt6 (Qt 버전까지 함께 보고)
    try:
        import PyQt6  # noqa: F401
        from PyQt6.QtCore import QT_VERSION_STR, PYQT_VERSION_STR
        r.result(
            OK,
            "import PyQt6",
            f"PyQt6 {PYQT_VERSION_STR} / Qt {QT_VERSION_STR}",
        )
    except Exception as e:
        r.exc("import PyQt6", e)

    # img2pdf
    try:
        import img2pdf
        r.result(OK, "import img2pdf", f"version {getattr(img2pdf, '__version__', 'unknown')}")
    except Exception as e:
        r.exc("import img2pdf", e)

    # Pillow
    try:
        import PIL
        r.result(OK, "import Pillow", f"version {getattr(PIL, '__version__', 'unknown')}")
    except Exception as e:
        r.exc("import Pillow", e)

    # psutil (선택 의존성)
    try:
        import psutil
        r.result(OK, "import psutil", f"version {getattr(psutil, '__version__', 'unknown')}")
    except Exception as e:
        r.exc("import psutil", e)


def check_appdata(r: Reporter) -> None:
    r.section("AppData writable")
    d = _log_dir()
    try:
        os.makedirs(d, exist_ok=True)
        probe = os.path.join(d, ".doctor_probe")
        with open(probe, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(probe)
        r.result(OK, "쓰기 가능", d)
    except Exception as e:
        r.exc(f"쓰기 불가 ({d})", e)


def check_powerpoint(r: Reporter) -> None:
    r.section("PowerPoint COM")
    if sys.platform != "win32":
        r.result(WARN, "PowerPoint COM", "Windows 외 플랫폼이라 건너뜀")
        return
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception as e:
        r.exc("pywin32 import", e)
        return

    pythoncom.CoInitialize()
    app = None
    try:
        # DispatchEx 로 격리된 새 인스턴스 사용. Dispatch 는 사용자가 이미 열어둔
        # PowerPoint 세션에 붙을 수 있는데, 그 경우 아래 Quit() 가 사용자의
        # 작업 창을 통째로 닫아 저장 안 된 데이터가 손실될 수 있다.
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
        except Exception as e:
            r.exc("PowerPoint.Application DispatchEx 실패", e)
            r.line("   -> PowerPoint 가 설치되지 않았거나 COM 등록이 손상됐을 수 있음.")
            return

        try:
            version = app.Version
            r.result(OK, "PowerPoint 실행", f"version {version}")
        except Exception as e:
            r.exc("PowerPoint.Version 조회 실패", e)

        try:
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            r.result(OK, "AutomationSecurity 설정", "매크로 자동 실행 차단 OK")
        except Exception as e:
            r.exc("AutomationSecurity 설정 실패", e)
    finally:
        # DispatchEx 로 직접 띄운 인스턴스만 닫으므로 사용자의 기존 세션에는
        # 영향이 없다.
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def check_img2pdf_smoke(r: Reporter) -> None:
    r.section("img2pdf smoke test")
    try:
        from PIL import Image
        import img2pdf
    except Exception as e:
        r.exc("의존 모듈 import 실패", e)
        return

    try:
        with tempfile.TemporaryDirectory() as tmp:
            png = os.path.join(tmp, "probe.png")
            pdf = os.path.join(tmp, "probe.pdf")
            Image.new("RGB", (1, 1), (255, 255, 255)).save(png, "PNG")
            with open(pdf, "wb") as f:
                f.write(img2pdf.convert([png]))
            size = os.path.getsize(pdf)
            r.result(OK, "1x1 PNG -> PDF 변환", f"{size} bytes")
    except Exception as e:
        r.exc("img2pdf 변환 실패", e)


def _wait_for_enter() -> None:
    """결과를 사용자가 읽을 수 있도록 Enter 대기. 리디렉션된 환경은 건너뜀."""
    # stdin 이 pipe 로 연결된 경우(자동화·CI)는 input() 이 즉시 EOF 로 끝나므로
    # 아무 효과가 없다. isatty 로 사람이 보는 터미널인지 확인한다.
    try:
        if not sys.stdin.isatty():
            return
    except Exception:
        pass
    print()
    print("Enter 키를 누르면 창이 닫힙니다.")
    try:
        input()
    except Exception:
        pass


def main() -> int:
    # --no-pause: CI / 자동화 호출시 Enter 대기 없이 즉시 종료.
    no_pause = "--no-pause" in sys.argv

    r = Reporter()
    r.line("ppt2pdf doctor — 환경 진단 리포트")
    r.line(f"생성 시각: {datetime.now().isoformat(timespec='seconds')}")

    checks: list[Callable[[Reporter], None]] = [
        check_system,
        check_packages,
        check_appdata,
        check_powerpoint,
        check_img2pdf_smoke,
    ]
    for fn in checks:
        try:
            fn(r)
        except Exception as e:
            r.exc(f"점검 중 예상치 못한 오류: {fn.__name__}", e)

    r.section("Summary")
    r.result(OK if r.fail_count == 0 else FAIL, f"FAIL: {r.fail_count}")
    r.result(OK if r.warn_count == 0 else WARN, f"WARN: {r.warn_count}")

    r.save(REPORT_PATH)
    r.line("")
    r.line(f"리포트 저장: {REPORT_PATH}")

    # 사용자가 결과를 읽을 수 있도록 기본적으로 Enter 를 기다린다.
    # 자동화에서는 `doctor.exe --no-pause` 로 호출하면 즉시 종료.
    if not no_pause:
        _wait_for_enter()
    return 0 if r.fail_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
