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
The controlled `tests/fixtures/visual_animation_object.ppt` fixture contains
one legacy animated textbox with no audio. Its static source/output rendering
must remain identical while the report contains exactly one slide-bound
`ANIMATION_OMITTED` object that combines the `AnimationInfo` fallback with the
matching `___PPT10` `ExtTimeNodeContainer` effect instead of missing or
double-counting either representation. Its `ParaBuildContainer` must also be
included in the same warning rather than counted as another animation object.
