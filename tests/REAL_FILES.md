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
