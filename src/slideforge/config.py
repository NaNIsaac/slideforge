"""Deck settings from TOML. Every key has a default, so an empty file is a valid configuration.

[deck]
width = 13.333        # inches
height = 7.5
font = "Aptos"
size = 20             # body text, pt
title_size = 32
template = ""         # a .pptx whose masters and theme the deck uses

[math]
route = "omml"        # "omml": editable Office math; "tex": SVG picture through xelatex
font = "Cambria Math" # omml: a font installed in Windows
tex_font = "STIX Two Math"
size = 24

[export]
pdf = true
png = false
png_width = 2667      # pixels; the height follows the slide's aspect ratio
"""

from __future__ import annotations

import tomllib
from dataclasses import dataclass, field, fields
from pathlib import Path
from typing import Any


@dataclass
class DeckSettings:
    width: float = 13.333
    height: float = 7.5
    font: str = "Aptos"
    size: float = 20
    title_size: float = 32
    template: str = ""


@dataclass
class MathSettings:
    route: str = "omml"
    font: str = "Cambria Math"
    tex_font: str = "STIX Two Math"
    size: float = 24

    def __post_init__(self) -> None:
        if self.route not in ("omml", "tex"):
            raise ValueError(f"math.route must be 'omml' or 'tex', not {self.route!r}")


@dataclass
class ExportSettings:
    pdf: bool = True
    png: bool = False
    png_width: int = 2667


@dataclass
class Config:
    deck: DeckSettings = field(default_factory=DeckSettings)
    math: MathSettings = field(default_factory=MathSettings)
    export: ExportSettings = field(default_factory=ExportSettings)
    slides: list[dict[str, Any]] = field(default_factory=list)
    root: Path = field(default_factory=Path.cwd)

    @classmethod
    def from_dict(cls, data: dict[str, Any], root: Path | None = None) -> Config:
        sections = {"deck": DeckSettings, "math": MathSettings, "export": ExportSettings}
        unknown = set(data) - set(sections) - {"slide"}
        if unknown:
            raise ValueError(f"unknown sections: {', '.join(sorted(unknown))}")
        parts = {}
        for name, kind in sections.items():
            values = data.get(name, {})
            known = {f.name for f in fields(kind)}
            if extra := set(values) - known:
                raise ValueError(f"unknown keys in [{name}]: {', '.join(sorted(extra))}")
            parts[name] = kind(**values)
        return cls(**parts, slides=data.get("slide", []), root=root or Path.cwd())

    @classmethod
    def load(cls, path: str | Path) -> Config:
        path = Path(path)
        return cls.from_dict(tomllib.loads(path.read_text(encoding="utf-8")), root=path.parent)
