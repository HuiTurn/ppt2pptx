"""Public API for ppt2pptx."""

from .converter import ConversionResult, convert, inspect_ppt

__all__ = ["ConversionResult", "convert", "inspect_ppt"]
__version__ = "0.3.4"
