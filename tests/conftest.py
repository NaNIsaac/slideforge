import base64
import os

import pytest

from slideforge import omml, tex
from slideforge.export import powerpoint_available, powerpoint_running

# 1x1 transparent PNG
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
)
SVG = b'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>'

needs_office_math = pytest.mark.skipif(
    (path := omml.stylesheet()) is None or not path.exists(),
    reason="MML2OMML.XSL (Microsoft Office) not found",
)
needs_tex = pytest.mark.skipif(not tex.available(), reason="xelatex or pdftocairo not on PATH")

# PowerPoint runs one instance per user, so a test would share the window of someone working in
# PowerPoint; the tests drive it only when asked to and when no one else has it open.
needs_powerpoint = pytest.mark.skipif(
    os.environ.get("SLIDEFORGE_COM") != "1" or not powerpoint_available() or powerpoint_running(),
    reason="PowerPoint tests run with SLIDEFORGE_COM=1, PowerPoint installed and not already running",
)
