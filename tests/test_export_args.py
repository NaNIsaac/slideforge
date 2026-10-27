"""export() without PowerPoint: the PowerShell call is replaced by a fake that records it."""

import pytest

from slideforge import Config, build, render


@pytest.fixture
def deck(tmp_path):
    return build(Config.from_dict({"slide": [{"title": "one"}, {"title": "two"}]}), tmp_path / "d.pptx")


@pytest.fixture
def calls(monkeypatch):
    seen = []

    def fake_run(cmd, **kwargs):
        seen.append(cmd)
        args = dict(zip(cmd[5::2], cmd[6::2], strict=False))
        if "-PngDir" in args:
            for i in (1, 2):
                (render.Path(args["-PngDir"]) / f"slide{i:03d}.png").write_bytes(b"png")
        return render.subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(render.subprocess, "run", fake_run)
    return seen


def test_stale_slides_are_removed_and_names_sort(deck, tmp_path, calls):
    folder = tmp_path / "png"
    folder.mkdir()
    for name in ("slide003.png", "slide12.png", "notes.png"):
        (folder / name).write_bytes(b"old")
    written = render.export(deck, png_dir=folder, engine="powerpoint")
    assert [p.name for p in written] == ["slide001.png", "slide002.png"]
    assert sorted(p.name for p in folder.iterdir()) == ["notes.png", "slide001.png", "slide002.png"]


def test_no_empty_arguments(deck, tmp_path, calls):
    render.export(deck, pdf=tmp_path / "d.pdf", engine="powerpoint")
    assert "" not in calls[0] and "-PngDir" not in calls[0]


def test_libreoffice_writes_the_pdf_and_skips_pngs(deck, tmp_path, monkeypatch):
    monkeypatch.setattr(render, "_libreoffice", lambda pptx, pdf: [pdf])
    with pytest.warns(UserWarning, match="PDF only"):
        written = render.export(deck, pdf=tmp_path / "d.pdf", png_dir=tmp_path / "png", engine="libreoffice")
    assert written == [tmp_path / "d.pdf"]
