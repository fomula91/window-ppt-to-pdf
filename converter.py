"""PPT -> PNG(1280 가로 고정) -> PDF 변환 로직.

Windows에 설치된 PowerPoint를 COM으로 호출해 슬라이드를 PNG로 내보낸 뒤
img2pdf로 결합한다. 폰트는 사용자 PC에 존재하는 그대로 렌더링되므로
변환된 PDF에서 글꼴이 치환·깨질 일이 없다.

보안 메모:
    - VBA 매크로·ActiveX 자동 실행을 막기 위해
      `Application.AutomationSecurity = msoAutomationSecurityForceDisable (3)` 설정.
    - 확장자 + 파일 시그니처(매직 바이트) 양쪽 검증으로 엉뚱한 파일이 PowerPoint
      파서에 전달되지 않게 한다.
    - 변환 실패 시 잔류한 POWERPNT.EXE를 선택적으로 정리한다(psutil 있을 때).
"""
from __future__ import annotations

import argparse
import os
import sys
import tempfile
from typing import Callable, Optional

TARGET_WIDTH = 1280

ALLOWED_EXTS = (".ppt", ".pptx", ".pps", ".ppsx", ".pptm", ".ppsm")

# PPTX/PPSX/PPTM/PPSM 은 ZIP 기반 OOXML, PPT/PPS 는 OLE CFB.
_ZIP_MAGIC = b"PK\x03\x04"
_OLE_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

# msoAutomationSecurity enum
_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def _ensure_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("이 프로그램은 Windows + PowerPoint 환경에서만 동작합니다.")


def _validate_input(ppt_abs: str) -> None:
    """확장자 화이트리스트 + 매직 바이트 검증."""
    if not os.path.isfile(ppt_abs):
        raise FileNotFoundError(f"PPT 파일을 찾을 수 없습니다: {ppt_abs}")

    ext = os.path.splitext(ppt_abs)[1].lower()
    if ext not in ALLOWED_EXTS:
        raise ValueError(
            f"허용되지 않는 확장자입니다: {ext or '(없음)'}. "
            f"허용: {', '.join(ALLOWED_EXTS)}"
        )

    with open(ppt_abs, "rb") as f:
        head = f.read(8)
    if ext in (".pptx", ".ppsx", ".pptm", ".ppsm"):
        if not head.startswith(_ZIP_MAGIC):
            raise ValueError(
                "파일 시그니처가 OOXML(ZIP)이 아닙니다. 손상되었거나 위조된 파일일 수 있습니다."
            )
    else:  # .ppt, .pps
        if not head.startswith(_OLE_MAGIC):
            raise ValueError(
                "파일 시그니처가 OLE CFB가 아닙니다. 손상되었거나 위조된 파일일 수 있습니다."
            )


def _kill_stale_powerpoint() -> None:
    """Close/Quit 실패 후 잔류한 POWERPNT.EXE 를 최후의 수단으로 정리."""
    try:
        import psutil  # type: ignore
    except ImportError:
        return
    for proc in psutil.process_iter(attrs=["name"]):
        try:
            if (proc.info.get("name") or "").lower() == "powerpnt.exe":
                proc.kill()
        except Exception:
            pass


def convert_ppt_to_pdf(
    ppt_path: str,
    pdf_path: str,
    jpeg_fallback: bool = False,
    progress_cb: Optional[Callable[[int, int], None]] = None,
) -> int:
    """PPT/PPTX를 PDF로 변환하고 최종 PDF 바이트 크기를 반환한다.

    progress_cb(current, total) 은 슬라이드 한 장을 내보낼 때마다 호출된다.
    jpeg_fallback=True 면 PNG를 품질 85 JPEG로 재저장해 용량을 낮춘다.
    """
    _ensure_windows()

    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    from PIL import Image
    import img2pdf

    ppt_abs = os.path.abspath(ppt_path)
    pdf_abs = os.path.abspath(pdf_path)
    _validate_input(ppt_abs)

    pythoncom.CoInitialize()
    powerpoint = None
    presentation = None
    clean_exit = False
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        # VBA 매크로 / ActiveX / 외부 링크 자동 실행 차단.
        try:
            powerpoint.AutomationSecurity = _AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception:
            pass
        try:
            powerpoint.DisplayAlerts = 0  # ppAlertsNone
        except Exception:
            pass
        try:
            powerpoint.Visible = 1
        except Exception:
            pass

        presentation = powerpoint.Presentations.Open(
            ppt_abs, ReadOnly=True, Untitled=False, WithWindow=False
        )

        slide_w = float(presentation.PageSetup.SlideWidth)
        slide_h = float(presentation.PageSetup.SlideHeight)
        width_px = TARGET_WIDTH
        height_px = int(round(TARGET_WIDTH * slide_h / slide_w))

        total = int(presentation.Slides.Count)
        if total == 0:
            raise ValueError("슬라이드가 비어 있습니다.")

        with tempfile.TemporaryDirectory(prefix="ppt2pdf_") as tmp:
            img_paths = []
            for idx in range(1, total + 1):
                slide = presentation.Slides.Item(idx)
                png = os.path.join(tmp, f"slide_{idx:04d}.png")
                slide.Export(png, "PNG", width_px, height_px)

                if jpeg_fallback:
                    jpg = os.path.join(tmp, f"slide_{idx:04d}.jpg")
                    with Image.open(png) as im:
                        im.convert("RGB").save(
                            jpg, "JPEG", quality=85, optimize=True
                        )
                    img_paths.append(jpg)
                else:
                    img_paths.append(png)

                if progress_cb:
                    progress_cb(idx, total)

            os.makedirs(os.path.dirname(pdf_abs) or ".", exist_ok=True)
            with open(pdf_abs, "wb") as f:
                f.write(img2pdf.convert(img_paths))

        clean_exit = True
        return os.path.getsize(pdf_abs)
    finally:
        closed_ok = True
        try:
            if presentation is not None:
                presentation.Close()
        except Exception:
            closed_ok = False
        try:
            if powerpoint is not None:
                powerpoint.Quit()
        except Exception:
            closed_ok = False
        if not closed_ok or not clean_exit:
            _kill_stale_powerpoint()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _cli() -> int:
    p = argparse.ArgumentParser(description="PPT -> 이미지 기반 PDF 변환기")
    p.add_argument("input", help="입력 .ppt 또는 .pptx 경로")
    p.add_argument(
        "output",
        nargs="?",
        help="출력 .pdf 경로 (생략 시 입력과 같은 폴더에 생성)",
    )
    p.add_argument(
        "--jpeg",
        action="store_true",
        help="PNG 대신 JPEG(품질 85) 재압축으로 용량 최소화",
    )
    args = p.parse_args()

    out = args.output or os.path.splitext(args.input)[0] + ".pdf"

    def progress(cur: int, total: int) -> None:
        print(f"  [{cur}/{total}] 내보내는 중...", flush=True)

    size = convert_ppt_to_pdf(args.input, out, args.jpeg, progress)
    mb = size / (1024 * 1024)
    print(f"완료: {out}  ({mb:.2f} MB)")
    return 0


if __name__ == "__main__":
    sys.exit(_cli())
