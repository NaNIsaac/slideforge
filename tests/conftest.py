import base64

import pytest

from slideforge import omml, tex
from slideforge.export import powerpoint_available

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
needs_powerpoint = pytest.mark.skipif(not powerpoint_available(), reason="PowerPoint not installed")
