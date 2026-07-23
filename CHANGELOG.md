# Changelog

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
