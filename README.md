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
- Preserves normal slide order and dimensions without mistaking masters or notes
  for slides, including append-only incremental saves with older document
  containers still present in the record stream.
- Resolves each slide's `masterIdRef` through the current master persist list,
  so layout/title-master backgrounds and decorations are taken from the
  referenced master rather than the first master in the file.
- Preserves hidden-slide state so excluded slides remain hidden during slide shows.
- Respects each slide's `fMasterObjects` setting when flattening master
  decorations, including PowerPoint's "Hide Background Graphics" behavior.
- Inherits vertical text anchoring from corresponding master placeholders when
  the slide placeholder does not store its own OfficeArt `anchorText`.
- Restores positioned editable text, fonts, sizes, colors, bold/italic/underline,
  paragraph alignment, bullets, rotation, flips, and safe external hyperlinks.
- Preserves PNG, JPEG, GIF, TIFF, EMF, WMF, and PICT media, including cropping,
  position, rotation, and flips when present.
- Recreates common editable shapes, including zero-axis horizontal/vertical
  lines without introducing a diagonal drift, plus solid/gradient/picture
  backgrounds, comments, speaker notes, slide numbers, dates, headers, and
  footers.
- Reconstructs legacy native tables (regular grids of rectangle cells) as
  editable DrawingML `<a:tbl>` graphic frames with preserved cell fills and
  1-pt borders, instead of flattening them into scattered shapes.
- Copies common core properties such as title, author, keywords, and timestamps.
- Supports atomic single-file and recursive batch conversion with JSON diagnostics.

## Usage

```console
python -m pip install ppt2pptx
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
Solid, common gradient, and stretched picture backgrounds are retained, while
advanced fills and effects may be approximated. PICT data is preserved, but
rendering depends on the PPTX consumer. PowerPoint 95 and earlier files use a
different record format and are deliberately rejected with a clear error.

Unsupported or approximated advanced objects are reported with object-backed
warning codes such as `ANIMATION_OMITTED`, `AUDIO_OMITTED`, `VIDEO_OMITTED`,
`MEDIA_ACTION_OMITTED`,
`EMBEDDED_OLE_OMITTED`, `LINKED_OLE_OMITTED`,
`ACTIVEX_CONTROL_OMITTED`, `CHART_OMITTED`,
`DIAGRAM_OR_SMARTART_OMITTED`, and `COMPLEX_FREEFORM_OMITTED`. Each warning
includes `count`, `record_types`, and `locations` (`slide_index`,
`record_offset`, `object_kind`) when a matching object is found.
OLE diagnostics count
external objects rather than their container/storage records and use
`ExObjRefAtom` to identify the slide that owns each object. Legacy MS Graph and
Excel chart ProgIDs inside an OLE container are classified as one slide-bound
`CHART_OMITTED` object instead of also producing a generic OLE warning. Files
with incremental-save history use only the latest `ExternalOleObjectAtom` per
object ID and references inside current slide persist ranges, so stale chart or
OLE revisions do not inflate warning counts. Files with those objects do not
receive a blanket advanced-feature warning. A legacy
`AnimationInfo` container and its child atom likewise count as one slide-bound
animation object, not two record-level losses. PowerPoint 2002+ `___PPT10`
timing trees are inspected for actual effect nodes and matched to the same
legacy fallback by OfficeArt shape ID, avoiding both silent timeline loss and
duplicate warnings for one effect. Text `ParaBuildContainer` settings are also
bound to their timing effect by shape/build ID, so paragraph sequencing,
direction, and automatic-delay loss is represented by that same object warning.
Legacy chart build settings are linked the same way, while chart editability
loss remains a separate `CHART_OMITTED` diagnostic.
PowerPoint-saved SmartArt is detected from the owning shape's `metroBlob`
DrawingML package and produces one slide-bound
`DIAGRAM_OR_SMARTART_OMITTED` diagnostic when only its preview is retained.
PowerPoint-saved media that reopens as a preview picture is detected through
the owning shape's `InteractiveInfoAtom` `II_MediaAction` and produces one
slide-bound `MEDIA_ACTION_OMITTED` diagnostic when playback behavior is lost.

## Development

```console
python -m pip install -e .
PYTHONPATH=src python -m unittest discover -s tests -v
python -m build
```

The converter itself does not depend on LibreOffice or PowerPoint. The real-file
regression script may use LibreOffice only to verify that generated packages
render:

```console
PYTHONPATH=src python scripts/validate_real_files.py tests/real_samples \
  -o tests/real_output --render --password hello
```

On Windows with Microsoft PowerPoint installed, bilateral visual regression
exports per-slide PNGs from the source `.ppt` and converted `.pptx`, then writes
metrics and diffs. Each side runs in an isolated `DispatchEx` instance; the
report records the owned process IDs, and cleanup never targets unrelated
PowerPoint processes. It also records PowerPoint's source/output object counts
and per-slide structure differences for text, pictures, tables, charts,
SmartArt, groups, OLE/media, comments, and speaker notes:

```console
PYTHONPATH=src python scripts/make_visual_fixture.py -o tests/fixtures/visual_minimal.ppt
PYTHONPATH=src python scripts/compare_powerpoint_visual.py \
  tests/fixtures/visual_minimal.ppt -o tests/visual_evidence/visual_minimal
```

## Specification

- [MS-PPT: PowerPoint Binary File Format](https://learn.microsoft.com/openspecs/office_file_formats/ms-ppt/)
- [MS-CFB: Compound File Binary File Format](https://learn.microsoft.com/openspecs/windows_protocols/ms-cfb/)
- [MS-ODRAW: Office Drawing Binary File Format](https://learn.microsoft.com/openspecs/office_file_formats/ms-odraw/)
