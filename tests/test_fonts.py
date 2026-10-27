"""fonts.ps1 against a made-up font in a temp folder: no download, no registry, no real fonts."""

import hashlib
import json
import re
import shutil
import subprocess
import zipfile
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[1] / "fonts" / "fonts.ps1"
SHELL = shutil.which("pwsh")

pytestmark = pytest.mark.skipif(SHELL is None, reason="PowerShell 7 (pwsh) not on PATH")


@pytest.fixture
def env(tmp_path):
    source = tmp_path / "source"
    source.mkdir()
    font = b"not really a font"
    (source / "Plain-Regular.otf").write_bytes(font)
    with zipfile.ZipFile(source / "zipped.zip", "w") as z:
        z.writestr("inside/Zipped-Regular.otf", font)

    def sha(p):
        return hashlib.sha256(p.read_bytes()).hexdigest()

    manifest = {
        "fonts": [
            {"name": "Plain Math", "version": "1.0", "license": "OFL-1.1",
             "url": "https://example.invalid/Plain-Regular.otf", "sha256": sha(source / "Plain-Regular.otf"),
             "file": "Plain-Regular.otf", "face": "Plain Math"},
            {"name": "Zipped Math", "version": "2.0", "license": "OFL-1.1",
             "url": "https://example.invalid/zipped.zip", "sha256": sha(source / "zipped.zip"),
             "member": "inside/Zipped-Regular.otf", "member_sha256": hashlib.sha256(font).hexdigest(),
             "file": "Zipped-Regular.otf", "face": "Zipped Math"},
        ]
    }  # fmt: skip
    (tmp_path / "fonts.json").write_text(json.dumps(manifest))
    return tmp_path


def run(env, *args):
    cmd = [SHELL, "-NoProfile", "-NonInteractive", "-File", str(SCRIPT), *args,
           "-Manifest", str(env / "fonts.json"), "-FontDir", str(env / "Fonts"),
           "-RegistryFile", str(env / "registry.json"), "-State", str(env / "state.json"),
           "-Source", str(env / "source")]  # fmt: skip
    done = subprocess.run(cmd, check=False, capture_output=True, text=True)
    # PowerShell colours and wraps error messages to the console width.
    out = re.sub(r"\x1b\[[0-9;]*m", "", done.stdout + done.stderr)
    return done.returncode, re.sub(r"\s*\n\s*\|\s*", " ", out)


def registry(env) -> dict:
    path = env / "registry.json"
    return json.loads(path.read_text(encoding="utf-8-sig")) if path.exists() else {}


def test_whatif_writes_nothing(env):
    code, out = run(env, "install", "-WhatIf")
    assert code == 0, out
    assert not (env / "Fonts").exists() and registry(env) == {}


def test_install_then_remove_restores(env):
    code, out = run(env, "install")
    assert code == 0, out
    assert sorted(p.name for p in (env / "Fonts").iterdir()) == ["Plain-Regular.otf", "Zipped-Regular.otf"]
    assert set(registry(env)) == {"Plain Math (OpenType)", "Zipped Math (OpenType)"}
    code, out = run(env, "remove")
    assert code == 0, out
    assert list((env / "Fonts").iterdir()) == []
    assert registry(env) == {}


def test_font_installed_by_someone_else_is_left_alone(env):
    (env / "Fonts").mkdir()
    (env / "Fonts" / "Plain-Regular.otf").write_bytes(b"theirs")
    run(env, "install", "-Font", "Plain Math")
    run(env, "remove", "-Font", "Plain Math")
    assert (env / "Fonts" / "Plain-Regular.otf").read_bytes() == b"theirs"


def test_wrong_hash_is_refused(env):
    (env / "source" / "Plain-Regular.otf").write_bytes(b"changed")
    code, out = run(env, "install", "-Font", "Plain Math")
    assert code != 0 and "check the new SHA-256" in out
    assert not (env / "Fonts" / "Plain-Regular.otf").exists()


def test_unknown_font_is_an_error(env):
    code, out = run(env, "list", "-Font", "Comic Math")
    assert code != 0 and "Comic Math" in out


def test_list_shows_state(env):
    run(env, "install", "-Font", "Plain Math")
    code, out = run(env, "list")
    assert code == 0, out
    assert "Plain Math" in out and "Zipped Math" in out


def test_manifest_pins_every_font():
    fonts = json.loads((SCRIPT.parent / "fonts.json").read_text())["fonts"]
    for entry in fonts:
        assert len(entry["sha256"]) == 64
        assert entry["version"] in entry["url"] or "ctan.org" in entry["url"]
        assert ("member" in entry) == ("member_sha256" in entry)
