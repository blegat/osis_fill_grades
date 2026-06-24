# OSIS fill grades

Fill the **Score** column of an official UCLouvain grade file (`session_*.xlsx`)
from your own spreadsheet of points, matching students **by email**.

You keep your grades wherever you compute them; this tool copies them into the
exact column the university platform expects, in a clean single-sheet file that
uploads without error.

## Why use it

- **Real warnings instead of silent failure.** The UCLouvain upload page gives
  *no* feedback when the file is wrong: you click *Submit*, nothing happens, and
  the only trace is a `500 (Internal Server Error)` in the browser's JS console.
  This tool tells you, *before* uploading, exactly what is off:
  - students in the official file with **no point** in your file (Score left empty),
  - emails in your file that are **absent** from the official file (ignored),
  - **duplicate** emails in your file,
  - **decimal** scores (often rejected for a given learning unit).
- **No more `XLOOKUP` inside the official file.** Doing the lookup directly in
  the downloaded file tempts you to add an extra worksheet — and that extra sheet
  is precisely what makes the platform crash with the 500 above. Here the lookup
  lives in *your* file; the official file is only ever written to, never restructured.
- **The official file is never modified.** Output goes to a new
  `*_filled.xlsx`, with a **single sheet** preserved, so the upload just works.
- **Robust matching.** Emails are compared case-insensitively and trimmed, and
  the email/points columns in your file are detected automatically.

## Install

```bash
pip install -e .          # add '.[test]' to also get pytest
```

## Usage

```bash
fill-grade TARGET.xlsx CUSTOM.xlsx [-o OUTPUT.xlsx]
# or without installing:
python3 fill_grades.py TARGET.xlsx CUSTOM.xlsx
```

- `TARGET.xlsx` — the official file (header `Email` in column **G**, `Score` in **H**).
- `CUSTOM.xlsx` — your file with emails and points.

If your custom file has several sheets, you are asked which one to use. The email
and points columns are detected automatically; if the points column is ambiguous,
you are shown a sample of each candidate and asked to pick.

### Options

| Option | Description |
| --- | --- |
| `-o, --output` | Output file (default `<target>_filled.xlsx`). |
| `--sheet` | Custom sheet name (asked if several and omitted). |
| `--email-col` | Force the custom email column, e.g. `B`. |
| `--score-col` | Force the custom points column, e.g. `C`. |
| `--target-email-col` | Target email column (default `G`). |
| `--target-score-col` | Target score column (default `H`). |

## Recommended workflow

1. Compute your points in a **separate** workbook (do your `XLOOKUP` there).
2. Run `fill-grade` to produce `*_filled.xlsx`.
3. Read the warnings, fix what needs fixing, re-run.
4. Upload `*_filled.xlsx` — a single-sheet file the platform accepts.

## Development

```bash
pip install -e '.[test]'
pytest
```

Tests build throwaway `.xlsx` files in a temp directory, so no real student data
is involved. CI runs them on Python 3.9 / 3.11 / 3.13 (see
`.github/workflows/ci.yml`).
