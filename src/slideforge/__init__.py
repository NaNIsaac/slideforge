"""Build PowerPoint decks from Python and TOML: native equations, vector figures, PDF export."""

from .config import Config
from .deck import Deck, add_equation, add_figure, build
from .picture import add_svg
from .render import export

__all__ = ["Config", "Deck", "add_equation", "add_figure", "add_svg", "build", "export"]
__version__ = "0.1.0"
