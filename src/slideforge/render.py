"""PDF and PNG from a .pptx with PowerPoint (Windows) or LibreOffice.

PowerPoint renders Office math and SVG pictures as they look on screen and embeds the fonts in
the PDF. It runs without a window through COM, driven by a PowerShell script so no Python COM
package is needed. The deck is copied first: opening the very file the user has open in
PowerPoint would hand over their window, and closing it would close theirs.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import warnings
from pathlib import Path

POWERSHELL = r"""
param([string]$Deck, [string]$Pdf, [string]$PngDir, [int]$Width, [int]$Height)
$ErrorActionPreference = 'Stop'
$running = Get-Process POWERPNT -ErrorAction SilentlyContinue
$app = New-Object -ComObject PowerPoint.Application
try {
    $prs = $app.Presentations.Open($Deck, -1, 0, 0)
    try {
        if ($Pdf) { $prs.SaveAs($Pdf, 32) }
        if ($PngDir) {
            foreach ($s in $prs.Slides) {
                $s.Export((Join-Path $PngDir ('slide{0:d3}.png' -f $s.SlideIndex)), 'PNG', $Width, $Height)
            }
        }
    } finally { $prs.Close() }
} finally {
    if (-not $running) { $app.Quit() }
}
"""


def powerpoint_available() -> bool:
    if sys.platform != "win32":
        return False
    import winreg

    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID").Close()
    except OSError:
        return False
    return True


def powerpoint_running() -> bool:
    """Whether a PowerPoint process exists, which an export would share."""
    if sys.platform != "win32":
        return False
    done = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/NH"], check=False, capture_output=True, text=True
    )
    return "POWERPNT.EXE" in done.stdout.upper()


def soffice() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def _shell() -> str:
    return shutil.which("pwsh") or "powershell"


def export(
    pptx: str | Path,
    *,
    pdf: str | Path | None = None,
    png_dir: str | Path | None = None,
    png_width: int = 2667,
    engine: str = "auto",
) -> list[Path]:
    """Write a PDF and/or one PNG per slide; returns the written files. engine is "powerpoint",
    "libreoffice" or "auto" (PowerPoint when installed). LibreOffice writes the PDF only."""
    pptx = Path(pptx).resolve()
    if engine == "auto":
        engine = "powerpoint" if powerpoint_available() else "libreoffice"
    if engine == "powerpoint":
        return _powerpoint(pptx, pdf, png_dir, png_width)
    if engine == "libreoffice":
        if png_dir:
            warnings.warn("slide PNGs need PowerPoint; writing the PDF only", stacklevel=2)
        return _libreoffice(pptx, pdf) if pdf else []
    raise ValueError(f"engine must be 'powerpoint', 'libreoffice' or 'auto', not {engine!r}")


def _powerpoint(pptx: Path, pdf, png_dir, png_width: int) -> list[Path]:
    from pptx import Presentation

    prs = Presentation(str(pptx))
    height = round(png_width * prs.slide_height / prs.slide_width)
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        copy = work / pptx.name
        shutil.copy2(pptx, copy)
        script = work / "export.ps1"
        script.write_text(POWERSHELL, encoding="utf-8-sig")
        target = Path(pdf).resolve() if pdf else None
        folder = Path(png_dir).resolve() if png_dir else None
        if folder:
            folder.mkdir(parents=True, exist_ok=True)
            # A longer deck exported here before would leave slides this one does not have.
            for old in folder.glob("slide*.png"):
                if old.stem.removeprefix("slide").isdigit():
                    old.unlink()
        if target:
            target.parent.mkdir(parents=True, exist_ok=True)
        # Windows PowerShell 5.1 drops empty arguments, which shifts the ones after them.
        args = ["-Deck", str(copy), "-Width", str(png_width), "-Height", str(height)]
        if target:
            args += ["-Pdf", str(target)]
        if folder:
            args += ["-PngDir", str(folder)]
        done = subprocess.run(
            [_shell(), "-NoProfile", "-NonInteractive", "-File", str(script), *args],
            check=False,
            capture_output=True,
            text=True,
            errors="replace",
        )
        if done.returncode:
            raise RuntimeError(f"PowerPoint export failed:\n{done.stderr[-2000:]}")
    written = [target] if target else []
    if folder:
        written += [folder / f"slide{i:03d}.png" for i in range(1, len(prs.slides) + 1)]
    return written


def _libreoffice(pptx: Path, pdf) -> list[Path]:
    office = soffice()
    if office is None:
        raise RuntimeError("neither PowerPoint nor LibreOffice is installed")
    target = Path(pdf).resolve()
    with tempfile.TemporaryDirectory() as tmp:
        done = subprocess.run(
            [office, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(pptx)],
            check=False,
            capture_output=True,
            text=True,
            errors="replace",
        )
        made = Path(tmp) / pptx.with_suffix(".pdf").name
        if done.returncode or not made.exists():
            raise RuntimeError(f"LibreOffice export failed:\n{done.stderr[-2000:]}")
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(made, target)
    return [target]
