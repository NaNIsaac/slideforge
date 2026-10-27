from slideforge import Config, build, export

from .conftest import needs_office_math, needs_powerpoint


@needs_powerpoint
@needs_office_math
def test_pdf_and_png(tmp_path):
    config = Config.from_dict({"slide": [{"title": "Export", "item": [{"math": r"e^{i\pi}+1=0"}]}]})
    deck = build(config, tmp_path / "e.pptx")
    written = export(
        deck, pdf=tmp_path / "e.pdf", png_dir=tmp_path / "png", png_width=800, engine="powerpoint"
    )
    assert (tmp_path / "e.pdf").read_bytes().startswith(b"%PDF")
    assert [p.name for p in written[1:]] == ["slide001.png"]
