# Changelog

## Unreleased

- Record PowerPoint's own source/output structure census in bilateral reports,
  including per-slide text, picture, table, chart, group, OLE, media, comment,
  and speaker-note counts with object-category mismatch locations.
- Bind sparse speaker notes by `NotesAtom.slideId`, and emit schema-valid Notes
  Masters with `notesStyle`, complete group transforms, and a dedicated theme
  part so PowerPoint reopens presentations containing notes.
- Collapse OLE container, object, and storage records into one diagnostic per
  external object; bind it to the owning slide through `ExObjRefAtom`, and
  distinguish embedded, linked, and ActiveX/OLE-control loss.
- Classify legacy MS Graph/Excel chart ProgIDs within their OLE container so
  one chart produces one slide-bound `CHART_OMITTED` diagnostic rather than a
  document-level chart warning plus a duplicate generic OLE warning.
- Collapse incremental-save OLE/chart revisions by external object ID, retain
  only references in current slide persist ranges, and classify the latest
  object atom so stale saved revisions cannot inflate warning counts.
- Collapse each legacy `AnimationInfo` container and child atom into one
  slide-bound `ANIMATION_OMITTED` diagnostic, so one animated shape is counted
  once rather than once per binary record.
- Parse `___PPT10` binary-tag timing trees, identify actual animation effects
  through `TimePropertyList` effect metadata, and merge each effect with its
  matching legacy `AnimationInfo` fallback by OfficeArt shape ID.
- Parse `BuildListContainer` text builds, bind `ParaBuildContainer` settings
  to the matching timing effect by shape/build ID, and include the paragraph
  build records in that one object diagnostic instead of silently omitting them.
- Bind `ChartBuildContainer` and `ChartBuildAtom` settings to the matching
  chart animation effects by shape/build ID, while retaining the separate
  chart-editability diagnostic for the same legacy MS Graph object.
- Detect PowerPoint-saved SmartArt through the shape's `metroBlob` DrawingML
  package, report one slide-bound object instead of silently preserving only
  its preview, and include SmartArt in the PowerPoint structure census.
- Detect PowerPoint-saved media preview shapes through their `II_MediaAction`,
  report one slide-bound `MEDIA_ACTION_OMITTED` object when playback behavior
  is lost, and count those legacy preview pictures as media in Office reports.
- Isolate every bilateral PowerPoint export with `DispatchEx`, record the
  exact owned process IDs, retry transient COM disconnects, and disable macros
  and link-update prompts without attaching to or terminating unrelated
  PowerPoint instances.
- Detect legacy PowerPoint tables (regular grids of rectangle autoshape
  cells + thin border lines) and emit them as editable DrawingML
  `<a:tbl>` graphic frames instead of flattening them into scattered
  shapes and text boxes. Cell fill colours and 1-pt black borders
  (`<a:tblBorders>`) are preserved, and the absorbed cell + border
  offsets are tracked in `Presentation.excluded_offsets` so they no
  longer trigger misleading `COMPLEX_FREEFORM_OMITTED` warnings or get
  re-emitted as duplicate shapes.
- Attach `locations` (`slide_index`, `record_offset`, `object_kind`) to
  object-backed lossy-feature warnings, and add a COM-generated animated
  fixture that asserts `ANIMATION_OMITTED` with slide-scoped locations.
- Replace the unconditional `ADVANCED_FEATURES_APPROXIMATED` warning with
  object-backed diagnostics (`ANIMATION_OMITTED`, `AUDIO_OMITTED`,
  `VIDEO_OMITTED`, `EMBEDDED_OLE_OMITTED`, `CHART_OMITTED`,
  `DIAGRAM_OR_SMARTART_OMITTED`, `COMPLEX_FREEFORM_OMITTED`) that only appear
  when matching records or unparsed freeforms are present.
- Make generated PPTX packages openable in Microsoft PowerPoint by emitting a
  complete theme `fmtScheme` (three fill/line/effect/background styles),
  non-empty master text styles, valid `sldLayoutId` values, and group transform
  stubs on masters/layouts.
- Add a Windows PowerPoint COM bilateral visual regression tool that exports
  per-slide reference/actual PNGs, writes MAE/RMSE/SSIM/diff evidence, and
  records structure counts plus conversion warnings.

## 0.3.4

- Preserve legacy ShadeScale background gradients, including their angle and
  symmetric color stops.
- Preserve picture line color, width, and dash properties.
- Emit standard WMF payloads so OLE preview images retain transparent
  backgrounds across Office-compatible renderers.
- Preserve custom bend positions and endpoint direction for three-segment
  elbow connectors.

## 0.3.3

- Preserve legacy line widths, dash styles, arrowhead types and sizes, including
  scaling inside nested groups.
- Improve text fidelity with paragraph margins, indents, spacing, tab rulers,
  text insets, vertical alignment, baseline shifts, and Symbol-font character
  normalization.
- Preserve transparent-color image cutouts with DrawingML color-change effects.
- Restore additional legacy autoshapes, connector geometry, freeform arcs,
  pattern fills, and approximate two-color gradients.
- Improve slide layering so backgrounds, plots, equations, annotations, and
  small overlay images retain their intended visibility.
- Revalidate the converter against 14 supported real-world presentations and a
  complex 30-slide presentation through rendering and overflow checks.

## 0.3.2

- Preserve paragraph nesting, level-specific master text styles, bullet
  characters, vertical anchoring, wrapping, and both legacy text autofit modes.
- Apply nested group rotation, flipping, scaling, and large child-anchor
  coordinates consistently to text, pictures, and editable shapes.
- Preserve the geometry of text-bearing autoshapes and emit valid OOXML preset
  names for legacy arrows.
- Respect explicit no-line flags and underline recovered hyperlinks.
- Harden version 3 CFB stream-size handling and truncated optional paragraph
  records.
- Fix packed freeform vertex parsing so open paths retain their final vertices
  without gaining an unintended closing segment.
- Expand inspection diagnostics with paragraph, layout, geometry, fill, and line
  details.
- Validate the converter against 15 PowerPoint 97–2003 samples and a complex
  30-slide presentation through conversion, rendering, strict parsing, and
  overflow checks.

## 0.3.1

- Preserve hidden-slide state in generated presentations.
- Verify release tags against the package version and validate distributions
  before publishing them to PyPI through trusted publishing.

## 0.3.0

- Add RC4 CryptoAPI password verification and record-level decryption, plus
  `--password` and safer `--password-file` CLI options.
- Recover legacy external slide text and inherit master text formatting and
  color schemes.
- Preserve character styles, paragraph alignment and bullets, hyperlinks,
  comments, speaker notes, headers, footers, dates, and slide-number fields.
- Preserve PNG, JPEG, GIF, TIFF, EMF, WMF, and PICT media with cropping,
  rotation, and flips.
- Recreate common editable shapes and solid or two-color gradient backgrounds.
- Distinguish normal slides from master and notes persist lists.
- Add a real-file regression runner and validate 15 PowerPoint 97–2003 files
  through conversion and LibreOffice rendering.
- Report encrypted metadata and remaining advanced feature approximations.

## 0.2.0

- Preserve legacy slide dimensions and editable text-box positions.
- Preserve embedded PNG, JPEG, GIF, and TIFF pictures with their slide positions.
- Preserve common OLE core properties such as title, author, and timestamps.
- Group text by its OfficeArt shape instead of flattening all slide text.
- Write slide master, blank layout, theme, and complete package relationships.
- Validate every generated XML part before atomically replacing the output.
- Add recursive directory conversion with per-file JSON results.

## 0.1.0

- Read CFB containers, PowerPoint persist directories, slide records, and text atoms.
- Write a basic editable PPTX package and structured conversion diagnostics.
