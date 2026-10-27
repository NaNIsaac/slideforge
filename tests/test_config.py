import pytest

from slideforge import Config


def test_empty_file_gives_defaults():
    config = Config.from_dict({})
    assert config.math.route == "omml"
    assert config.deck.font == "Aptos"
    assert config.slides == []


def test_sections_and_slides():
    config = Config.from_dict({"math": {"route": "tex", "size": 20}, "slide": [{"title": "A"}]})
    assert (config.math.route, config.math.size) == ("tex", 20)
    assert config.slides[0]["title"] == "A"


@pytest.mark.parametrize("data", [{"deck": {"colour": "red"}}, {"maths": {}}, {"math": {"route": "png"}}])
def test_mistakes_are_errors(data):
    with pytest.raises(ValueError):
        Config.from_dict(data)


def test_load_resolves_paths_from_the_file(tmp_path):
    (tmp_path / "deck.toml").write_text('[deck]\nfont = "Arial"\n')
    config = Config.load(tmp_path / "deck.toml")
    assert config.root == tmp_path
    assert config.deck.font == "Arial"
