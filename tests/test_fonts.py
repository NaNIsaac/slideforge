"""fonts.ps1 against a made-up font in a temp folder: no download and no real fonts.

On Windows the registry values go to a temporary key under HKCU that already holds values of its
own, so the code path that writes the real font key is the one under test (#5); elsewhere they go
to a JSON file."""

import hashlib
import json
import re
import shutil
import subprocess
import sys
import uuid
import zipfile
from pathlib import Path

import pytest

SCRIPT = Path(__file__).resolve().parents[1] / "fonts" / "fonts.ps1"
SHELL = shutil.which("pwsh")

pytestmark = pytest.mark.skipif(SHELL is None, reason="PowerShell 7 (pwsh) not on PATH")
WINDOWS = sys.platform == "win32"
# Values another program put in the key before the script ran; every test checks they survive.
OTHERS = {"Someone Sans (TrueType)": r"C:\fonts\SomeoneSans.ttf", "Other Serif (OpenType)": "OtherSerif.otf"}


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
    if not WINDOWS:
        yield tmp_path
        return
    import winreg

    sub = rf"Software\slideforge-test-{uuid.uuid4().hex}"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, sub) as key:
        for name, value in OTHERS.items():
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
    (tmp_path / "key.txt").write_text(sub)
    try:
        yield tmp_path
    finally:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, sub)


def run(env, *args):
    cmd = [SHELL, "-NoProfile", "-NonInteractive", "-File", str(SCRIPT), *args,
           "-Manifest", str(env / "fonts.json"), "-FontDir", str(env / "Fonts"),
           "-State", str(env / "state.json"), "-Source", str(env / "source")]  # fmt: skip
    if WINDOWS:
        cmd += ["-RegistryKey", "HKCU:\\" + (env / "key.txt").read_text()]
    else:
        cmd += ["-RegistryFile", str(env / "registry.json")]
    done = subprocess.run(cmd, check=False, capture_output=True, text=True)
    # PowerShell colours and wraps error messages to the console width.
    out = re.sub(r"\x1b\[[0-9;]*m", "", done.stdout + done.stderr)
    return done.returncode, re.sub(r"\s*\n\s*\|\s*", " ", out)


def registry(env) -> dict:
    """Values the script wrote, without the ones that were there before it."""
    if not WINDOWS:
        path = env / "registry.json"
        return json.loads(path.read_text(encoding="utf-8-sig")) if path.exists() else {}
    import winreg

    values = {}
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, (env / "key.txt").read_text()) as key:
        i = 0
        while True:
            try:
                name, value, _ = winreg.EnumValue(key, i)
            except OSError:
                break
            values[name] = value
            i += 1
    for name, value in OTHERS.items():
        assert values.pop(name) == value, f"{name} was changed or removed"
    return values


def test_whatif_writes_nothing_and_names_the_registry_value(env):
    code, out = run(env, "install", "-WhatIf")
    assert code == 0, out
    assert not (env / "Fonts").exists() and registry(env) == {}
    assert "registry value 'Plain Math (OpenType)'" in out


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
