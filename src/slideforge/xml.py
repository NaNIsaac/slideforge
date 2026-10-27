NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "asvg": "http://schemas.microsoft.com/office/drawing/2016/SVG/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def q(tag: str) -> str:
    """'a:blip' -> '{namespace}blip'."""
    prefix, name = tag.split(":")
    return f"{{{NS[prefix]}}}{name}"
