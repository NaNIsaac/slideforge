import zipfile

import pytest
from pptx.util import Inches

from slideforge import Config, Deck, build

from .conftest import PNG, SVG, needs_office_math, needs_tex


def parts(path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as z:
        return {name: z.read(name) for name in z.namelist()}


def test_svg_figure_keeps_svg_and_fallback(tmp_path):
    (tmp_path / "f.svg").write_bytes(SVG)
    (tmp_path / "f.png").write_bytes(PNG)
    config = Config.from_dict(
        {"slide": [{"title": "T", "item": [{"figure": "f.svg", "width": 3}]}]}, root=tmp_path
    )
    files = parts(build(config, tmp_path / "out.pptx"))
    assert any(name.endswith(".svg") for name in files)
    assert b"svgBlip" in files["ppt/slides/slide1.xml"]


def test_svg_without_fallback_is_an_error(tmp_path):
    (tmp_path / "f.svg").write_bytes(SVG)
    deck = Deck(Config(root=tmp_path))
    with pytest.raises(FileNotFoundError):
        deck.figure(deck.slide(), "f.svg")


def test_items_flow_down_the_slide():
    deck = Deck()
    slide = deck.slide("Title")
    first = deck.text(slide, "one")
    second = deck.text(slide, "two\nthree")
    assert second.top > first.top + first.height - 1
    assert deck.prs.slide_width == Inches(13.333)


def test_unknown_item_is_an_error():
    with pytest.raises(ValueError):
        Deck().add({"item": [{"chart": "x"}]})


@needs_office_math
def test_omml_equation_is_office_math(tmp_path):
    config = Config.from_dict({"slide": [{"item": [{"math": r"\sum_n J_n(S)"}]}]})
    slide = parts(build(config, tmp_path / "m.pptx"))["ppt/slides/slide1.xml"]
    assert b"oMathPara" in slide and b"mc:Fallback" in slide


@needs_tex
def test_tex_equation_is_svg(tmp_path):
    config = Config.from_dict({"math": {"route": "tex"}, "slide": [{"item": [{"math": "E=mc^2"}]}]})
    files = parts(build(config, tmp_path / "t.pptx"))
    assert any(name.endswith(".svg") for name in files)
