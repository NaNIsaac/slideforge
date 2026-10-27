from __future__ import annotations

import argparse
from pathlib import Path

from .config import Config
from .deck import Deck, build
from .export import export


def _build(args) -> list[Path]:
    config = Config.load(args.config)
    out = args.out or args.config.with_suffix(".pptx")
    written = [build(config, out)]
    if config.export.pdf or config.export.png:
        written += export(
            out,
            pdf=out.with_suffix(".pdf") if config.export.pdf else None,
            png_dir=out.parent / f"{out.stem}-png" if config.export.png else None,
            png_width=config.export.png_width,
            engine=args.engine,
        )
    return written


def _math(args) -> list[Path]:
    config = Config.load(args.config) if args.config else Config()
    if args.route:
        config.math.route = args.route
    if args.font:
        config.math.font = config.math.tex_font = args.font
    if args.size:
        config.math.size = args.size
    deck = Deck(config)
    slide = deck.slide()
    for latex in args.latex:
        deck.math(slide, latex, width=deck.width - 1.6, left=0.8)
    return [deck.save(args.out)]


def _export(args) -> list[Path]:
    return export(
        args.deck,
        pdf=args.pdf or (None if args.png else args.deck.with_suffix(".pdf")),
        png_dir=args.png,
        png_width=args.width,
        engine=args.engine,
    )


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog=__package__, description="Build PowerPoint decks from TOML and Python."
    )
    commands = parser.add_subparsers(required=True)

    p = commands.add_parser("build", help="build the deck a TOML file describes, then export it")
    p.add_argument("config", type=Path)
    p.add_argument("-o", "--out", type=Path, help="default: the TOML file's name with .pptx")
    p.add_argument("--engine", choices=("auto", "powerpoint", "libreoffice"), default="auto")
    p.set_defaults(run=_build)

    p = commands.add_parser("math", help="one slide with the given LaTeX equations, one per line")
    p.add_argument("latex", nargs="+")
    p.add_argument("--route", choices=("omml", "tex"))
    p.add_argument("--font")
    p.add_argument("--size", type=float)
    p.add_argument("--config", type=Path, help="take [deck] and [math] from this TOML file")
    p.add_argument("-o", "--out", type=Path, default=Path("equations.pptx"))
    p.set_defaults(run=_math)

    p = commands.add_parser("export", help="PDF and/or slide PNGs from a .pptx")
    p.add_argument("deck", type=Path)
    p.add_argument("--pdf", type=Path, help="default: the deck's name with .pdf, unless --png is given")
    p.add_argument("--png", type=Path, metavar="DIR", help="one PNG per slide into DIR (PowerPoint only)")
    p.add_argument("--width", type=int, default=2667, help="PNG width in pixels")
    p.add_argument("--engine", choices=("auto", "powerpoint", "libreoffice"), default="auto")
    p.set_defaults(run=_export)

    args = parser.parse_args(argv)
    for path in args.run(args):
        print(path)
