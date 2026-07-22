# Changelog

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
