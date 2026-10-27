"""SVG pictures in PowerPoint, with a PNG fallback for viewers older than Office 2016.

python-pptx takes raster images only. PowerPoint stores an SVG picture as an ordinary picture
whose blip carries an extension pointing at the SVG part; this module writes that extension.
"""

from __future__ import annotations

import io

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.util import Emu

from .xml import NS, q

SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"


def add_svg(slide, svg: bytes, png: bytes, *, left, top, width=None, height=None):
    """Add an SVG picture showing `png` where SVG is not supported. Without width or height the
    picture takes the PNG's size; with one of them the other keeps the aspect ratio."""
    picture = slide.shapes.add_picture(io.BytesIO(png), left, top, width, height)
    package = slide.part.package
    part = Part(package.next_partname("/ppt/media/svg%d.svg"), "image/svg+xml", package, svg)
    rid = slide.part.relate_to(part, RT.IMAGE)
    blip = picture._element.find(f".//{q('a:blip')}")
    ext = etree.SubElement(etree.SubElement(blip, q("a:extLst")), q("a:ext"), uri=SVG_EXT)
    etree.SubElement(ext, q("asvg:svgBlip"), {q("r:embed"): rid}, nsmap={"asvg": NS["asvg"]})
    return picture


def fit(picture, width) -> None:
    """Shrink `picture` to `width` if it is wider, keeping the aspect ratio."""
    if width is not None and picture.width > width:
        picture.height = Emu(round(picture.height * width / picture.width))
        picture.width = width
