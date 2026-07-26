# Real-file regression corpus

Real presentations are downloaded into the ignored `tests/real_samples/`
directory and are never redistributed by this project. The primary corpus uses
the Apache POI project's public PowerPoint test data, including:

- `basic_test_ppt_file.ppt`
- `SampleShow.ppt`
- `54880_chinese.ppt`
- `54541_cropped_bitmap.ppt`
- `Single_Coloured_Page_With_Fonts_and_Alignments.ppt`
- `WithComments.ppt` and `WithLinks.ppt`
- `pictures.ppt`, table, text-shape, background, sound, and header/footer samples
- `ole2-embedding-2003.ppt`, `ppt_with_embeded.ppt`, and
  `testPPT_oleWorkbook.ppt`
- password-protected and PowerPoint 95 compatibility probes

An additional historical PowerPoint 7.0 presentation comes from the public
`j3-fortran/j3-papers` archive (`years/2006/06-144.ppt`).

These files exercise different producers, text encodings, embedded pictures,
slide dimensions, drawing records, and historical PowerPoint versions. They
are used for local compatibility and rendering tests only. Run the corpus with:

```console
PYTHONPATH=src python scripts/validate_real_files.py tests/real_samples \
  -o tests/real_output --render --password hello \
  --report tests/real_output/report.json
```

## PowerPoint bilateral visual regression

On Windows hosts with Microsoft PowerPoint registered for COM automation, use
the Office PNG pipeline instead of LibreOffice when measuring visual fidelity:

```console
PYTHONPATH=src python scripts/make_visual_fixture.py -o tests/fixtures/visual_minimal.ppt
PYTHONPATH=src python scripts/compare_powerpoint_visual.py \
  tests/fixtures/visual_minimal.ppt -o tests/visual_evidence/visual_minimal \
  --width 960 --height 720
```

The evidence directory contains `reference/`, `actual/`, `diff/`, and
`report.json` with provider `office`, PowerPoint version, SHA-256 digests,
hidden-slide manifests, structure counts, conversion warnings, and per-slide
MAE/RMSE/SSIM metrics. Source and output exports use separate `DispatchEx`
instances whose owned process IDs are recorded in the report; the tool never
attaches to or terminates unrelated PowerPoint processes. LibreOffice
`--render` remains a package-render smoke check only; it does not compare
against the source `.ppt`.

`office_structure` records PowerPoint's own per-slide and total counts for
shapes, text-bearing shapes, pictures, tables, charts, groups, OLE objects,
media, comments, and non-empty speaker-note bodies on both sides. Its
`differences` list identifies the slide and object category for every mismatch.
The controlled `tests/fixtures/visual_ole.ppt` fixture additionally verifies
that an OLE-to-preview-picture structural difference has one slide-bound
`EMBEDDED_OLE_OMITTED` diagnostic rather than one warning count per storage
record.
The controlled `tests/fixtures/visual_chart.ppt` fixture verifies the same
object linkage for a legacy `MSGraph.Chart.8` chart and rejects duplicate
`EMBEDDED_OLE_OMITTED` reporting when `CHART_OMITTED` is the precise category.
Apache POI `test-data/slideshow/37625.ppt` additionally exercises incremental
save history: 189 stored `ExternalOleObjectAtom` records collapse to the 10
objects referenced by current slide persist ranges (8 charts and 2 other OLE
objects), matching PowerPoint's source structure census.
The same file contains 14 saved `DocumentContainer` revisions; the latest one
must produce all 29 current slides rather than the 28 listed by the first
historical revision.
The controlled `tests/fixtures/visual_animation_object.ppt` fixture contains
one legacy animated textbox with no audio. Its static source/output rendering
must remain identical while the report contains exactly one slide-bound
`ANIMATION_OMITTED` object that combines the `AnimationInfo` fallback with the
matching `___PPT10` `ExtTimeNodeContainer` effect instead of missing or
double-counting either representation. Its `ParaBuildContainer` must also be
included in the same warning rather than counted as another animation object.
The controlled `tests/fixtures/visual_chart_animation.ppt` fixture verifies
that a by-series MS Graph build contributes its `ChartBuildContainer` records
to the three real animation effects while the chart itself still produces one
independent, slide-bound `CHART_OMITTED` editability diagnostic.
The controlled `tests/fixtures/visual_smartart.ppt` fixture verifies that a
PowerPoint-saved SmartArt `metroBlob` produces exactly one slide-bound
`DIAGRAM_OR_SMARTART_OMITTED` warning. The Office census must show one source
SmartArt object becoming one output picture, and the bilateral metrics keep
that preview conversion within the recorded visual bounds.
The controlled `tests/fixtures/visual_video.ppt` fixture verifies that a
PowerPoint-saved poster picture carrying `II_MediaAction` produces exactly one
slide-bound `MEDIA_ACTION_OMITTED` warning. The Office census must show media
`1 -> 0` while picture count remains `1 -> 1`; the static poster rendering must
remain pixel-identical.
The controlled `tests/fixtures/visual_background_picture.ppt` fixture contains
one embedded PNG used through the slide background's OfficeArt
`fillType=picture`/`fillBlip` properties plus one editable foreground textbox.
It verifies that the converter writes a DrawingML background relationship,
does not turn the background into a picture shape, emits no lossy warning, and
matches PowerPoint's source rendering pixel-for-pixel.
The controlled `tests/fixtures/visual_zero_extent_lines.ppt` fixture recreates
54 individual master lines from the grouped grid in pinned Apache POI
`37625.ppt` test data while removing all unrelated slide/master objects and
group coordinators. It verifies that zero-width vertical and zero-height
horizontal anchors remain zero in the parsed model and DrawingML transform
instead of becoming slightly diagonal one-unit boxes. The source has no
slide-local shapes; the output's 54 editable shapes are the intentionally
flattened master grid and produce no conversion warning.
The controlled `tests/fixtures/visual_master_objects_disabled.ppt` fixture has
one large red master rectangle and one blue slide-local rectangle, with
`SlideAtom.fMasterObjects` cleared through PowerPoint's
`DisplayMasterShapes=False` property. The output must retain exactly the one
slide-local editable shape, emit no warning, and match the PowerPoint source
render pixel-for-pixel.
The controlled `tests/fixtures/visual_placeholder_anchor.ppt` fixture contains
one PowerPoint-generated title whose master uses `anchorText=bottom`. Its
equal-length fixture rewrite removes only the slide title's OfficeArt PID 135
entry, leaving the corresponding master property intact. It verifies that the
editable DrawingML text box inherits `anchor="b"`, produces no unrelated
warning, and matches PowerPoint pixel-for-pixel.
The controlled `tests/fixtures/visual_master_selection.ppt` fixture contains
one blank slide whose selected layout master supplies a solid blue background
while the first master does not. It verifies the
`SlideAtom.masterIdRef` -> `MasterPersistAtom.persistIdRef` mapping, emits no
warning, keeps PowerPoint's zero-shape structure census, and matches the source
pixel-for-pixel. Apache POI `37625.ppt` additionally verifies selection of its
legacy title master on slides 1 and 29 without regressing the other 27 slides.
The controlled `tests/fixtures/visual_pattern_line.ppt` fixture copies one
patterned grid line from pinned Apache POI `37625.ppt` test data into a
PowerPoint-generated slide. It verifies that OfficeArt
`lineType=msolinePattern` plus its 8x8 DIB `lineFillBlip` becomes an editable
DrawingML `pct30` line fill with matching foreground/background colors, no
warning, identical one-shape structure, and pixel-identical rendering.
The controlled `tests/fixtures/visual_header_footer_master.ppt` fixture has one
blank slide and one fixed date field whose selected master places a blue
14-point Tahoma placeholder at the bottom edge. It verifies editable field
inheritance for geometry, alignment, vertical anchoring, font, and color with
no warning, identical one-text-shape structure, and pixel-identical rendering.
Apache POI `37625.ppt` additionally exercises document-level date/footer values
across 29 slides, including `fMasterObjects` suppression and
`DocumentAtom.fOmitTitlePlace` on title-master slides.
