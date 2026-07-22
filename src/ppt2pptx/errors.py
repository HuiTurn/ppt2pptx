class Ppt2PptxError(Exception):
    """Base error for expected conversion failures."""


class InvalidPpt(Ppt2PptxError):
    """The input is not a supported PowerPoint binary presentation."""


class UnsafeOutputPathError(Ppt2PptxError):
    """An output path could overwrite an input or has the wrong extension."""


class EncryptedPresentationError(Ppt2PptxError):
    """The presentation is encrypted and needs password support."""


class UnsupportedPptVersionError(Ppt2PptxError):
    """The file predates the PowerPoint 97 binary record format."""
