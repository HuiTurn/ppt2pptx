# ppt2pptx

`ppt2pptx` converts Microsoft PowerPoint 97–2003 binary `.ppt` presentations
to Office Open XML `.pptx` using only the Python standard library. It reads
the CFB/OLE container and PowerPoint record stream directly; it never launches
PowerPoint, LibreOffice, COM, or a conversion service.

The project is useful for a broad set of real presentations, but it is not a
complete implementation of every legacy PowerPoint feature. Unsupported or
lossy content is diagnosed instead of being silently presented as fully
faithful.

## Highlights

- Opens RC4 CryptoAPI password-protected presentations when a password is supplied.
- Preserves normal slide order and dimensions without mistaking masters or notes for slides.
- Restores positioned editable text, fonts, sizes, colors, bold/italic/underline,
  paragraph alignment, bullets, rotation, flips, and safe external hyperlinks.
- Preserves PNG, JPEG, GIF, TIFF, EMF, WMF, and PICT media, including cropping,
  position, rotation, and flips when present.
- Recreates common editable shapes, solid/gradient backgrounds, comments, speaker
  notes, slide numbers, dates, headers, and footers.
- Copies common core properties such as title, author, keywords, and timestamps.
- Supports atomic single-file and recursive batch conversion with JSON diagnostics.

## Usage

```console
python -m pip install -e .
ppt2pptx presentation.ppt
ppt2pptx presentation.ppt -o presentation.pptx --report report.json
ppt2pptx protected.ppt --password-file password.txt
ppt2pptx inspect presentation.ppt --json
ppt2pptx batch input-directory -o output-directory --recursive --report batch.json
```

The input is always read-only and output files are written atomically.  The
converter refuses to overwrite its input.

The Python API accepts the password directly:

```python
from ppt2pptx import convert

result = convert("protected.ppt", "protected.pptx", password="secret")
print(result.report.to_dict())
```

## Current limitations

Charts, SmartArt, animation timelines, audio/video playback, embedded OLE
objects, and complex freeform or grouped master geometry remain incomplete.
Solid and common gradient backgrounds are retained, while advanced fills and
effects may be approximated. PICT data is preserved, but rendering depends on
the PPTX consumer. PowerPoint 95 and earlier files use a different record
format and are deliberately rejected with a clear error.

## Development

```console
PYTHONPATH=src python -m unittest discover -s tests -v
python -m build
```

The converter itself does not depend on LibreOffice. The real-file regression
script may use LibreOffice only to verify that generated packages render:

```console
PYTHONPATH=src python scripts/validate_real_files.py tests/real_samples \
  -o tests/real_output --render --password hello
```

## Specification

- [MS-PPT: PowerPoint Binary File Format](https://learn.microsoft.com/openspecs/office_file_formats/ms-ppt/)
- [MS-CFB: Compound File Binary File Format](https://learn.microsoft.com/openspecs/windows_protocols/ms-cfb/)
- [MS-ODRAW: Office Drawing Binary File Format](https://learn.microsoft.com/openspecs/office_file_formats/ms-odraw/)
