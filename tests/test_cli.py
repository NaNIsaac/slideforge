from slideforge.cli import main


def test_build_without_export(tmp_path, capsys):
    toml = tmp_path / "deck.toml"
    toml.write_text('[export]\npdf = false\n\n[[slide]]\ntitle = "Hello"\n')
    main(["build", str(toml)])
    assert (tmp_path / "deck.pptx").exists()
    assert capsys.readouterr().out.strip().endswith("deck.pptx")
