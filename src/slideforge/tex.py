"""LaTeX to an SVG picture through xelatex: any OpenType math font TeX can find, TeX's spacing,
not editable in PowerPoint."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from lxml import etree
from pptx.util import Pt

from .picture import add_svg, fit

# xelatex finds system fonts by name, but the math fonts of a TeX distribution only by file name.
FONT_FILES = {
    "Latin Modern Math": "latinmodern-math.otf",
    "STIX Two Math": "STIXTwoMath-Regular.otf",
    "TeX Gyre Pagella Math": "texgyrepagella-math.otf",
    "Fira Math": "FiraMath-Regular.otf",
    "GFS Neohellenic Math": "GFSNeohellenicMath.otf",
    "Lete Sans Math": "LeteSansMath.otf",
}
TEX_PT = 10


def available() -> bool:
    return bool(shutil.which("xelatex") and shutil.which("pdftocairo"))


def _run(*cmd: str, cwd: Path) -> None:
    done = subprocess.run(cmd, check=False, cwd=cwd, capture_output=True, text=True, errors="replace")
    if done.returncode:
        log = next(cwd.glob("*.log"), None)
        detail = log.read_text(errors="replace")[-2000:] if log else (done.stdout + done.stderr)[-2000:]
        raise RuntimeError(f"{cmd[0]} failed:\n{detail}")


def render(latex: str, *, font: str = "STIX Two Math") -> tuple[bytes, bytes, float, float]:
    """LaTeX -> (svg, png fallback, width pt, height pt) at TeX's 10 pt. MiKTeX is told not to
    install missing packages, so a missing font or package is an error, not a prompt."""
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        (work / "eq.tex").write_text(
            "\\documentclass[border=0.5pt]{standalone}\n\\usepackage{unicode-math}\n"
            f"\\setmathfont{{{FONT_FILES.get(font, font)}}}\n"
            f"\\begin{{document}}$\\displaystyle {latex}$\\end{{document}}\n",
            encoding="utf-8",
        )
        tex = ["xelatex", "-interaction=nonstopmode", "-halt-on-error"]
        if "miktex" in (shutil.which("xelatex") or "").lower():
            tex.append("-disable-installer")
        _run(*tex, "eq.tex", cwd=work)
        _run("pdftocairo", "-svg", "eq.pdf", "eq.svg", cwd=work)
        _run("pdftocairo", "-png", "-r", "600", "-transp", "-singlefile", "eq.pdf", "eq", cwd=work)
        root = etree.parse(str(work / "eq.svg")).getroot()
        width, height = (float(root.get(k).removesuffix("pt")) for k in ("width", "height"))
        return (work / "eq.svg").read_bytes(), (work / "eq.png").read_bytes(), width, height


def add_equation(slide, latex: str, *, font: str, size_pt: float, left, top, width):
    data, png, w_pt, h_pt = render(latex, font=font)
    scale = size_pt / TEX_PT
    picture = add_svg(slide, data, png, left=left, top=top, width=Pt(w_pt * scale), height=Pt(h_pt * scale))
    fit(picture, width)
    return picture
