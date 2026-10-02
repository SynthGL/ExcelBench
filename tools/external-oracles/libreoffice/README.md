# LibreOffice helpers

Two independent helpers live here:

- `libreoffice_uno_adapter.py` + `uno/excelbench_uno.py`: the `libreoffice`
  benchmark adapter (read, write, modify) used by
  `excelbench.harness.adapters.libreoffice_adapter` and
  `excelbench.modifiable.LibreofficeEngine`.
- `libreoffice_oracle.py`: the older open/save and render oracle (not a
  benchmark adapter).

Both locate LibreOffice through `LIBREOFFICE_BIN`, `soffice`, `libreoffice`, or
`/Applications/LibreOffice.app/Contents/MacOS/soffice`.

## Benchmark adapter (`libreoffice`)

### Run

```bash
cd tools/external-oracles/libreoffice
python libreoffice_uno_adapter.py < request.json
```

The request/response protocol is the ExcelBench JSON-model helper protocol
(`describe`, `read_model`, `write_model`, `mutate`) documented in
`src/excelbench/harness/adapters/json_model_adapter.py`. Any Python 3 works for
the launcher; ExcelBench runs it with its own interpreter.

### Why a macro inside soffice

The obvious designs do not work on macOS arm64 with LibreOffice 26.8:

- LibreOffice's bundled standalone interpreter
  (`LibreOffice.app/Contents/Resources/python`) is killed with SIGKILL on
  launch.
- Importing `uno` from a regular Python interpreter segfaults.

What does work is LibreOffice's embedded Python running a user macro inside
`soffice` itself. The launcher therefore:

1. Reads one JSON request from stdin and writes it to a temporary file.
2. Resolves `soffice` (order above).
3. Takes an exclusive file lock on a reusable profile directory
   (`$XDG_CACHE_HOME/excelbench/libreoffice-uno/profile`, default
   `~/.cache/...`; override with `EXCELBENCH_LIBREOFFICE_PROFILE`). Two soffice
   processes on one profile do not run side by side: the second forwards its
   arguments to the first over IPC and exits, so concurrent requests would
   silently never run. The lock serializes them.
4. Copies `uno/excelbench_uno.py` into `<profile>/user/Scripts/python/` on every
   run, so the macro always matches the checkout.
5. Runs
   `soffice -env:UserInstallation=<profile> --headless --invisible --norestore --nologo --nodefault --nolockcheck 'vnd.sun.star.script:excelbench_uno.py$main?language=Python&location=user'`
   with the request and response file paths in `EXCELBENCH_UNO_REQUEST` and
   `EXCELBENCH_UNO_RESPONSE`. The default timeout is 150 s
   (`EXCELBENCH_LIBREOFFICE_TIMEOUT` overrides it); on timeout the whole
   process group is killed.
6. Prints the macro's JSON response. If soffice exits without writing one, it
   prints `{"error": "libreoffice_failed", ...}` with the soffice return code
   and stderr tail and exits 1.

The first launch with a fresh profile takes about 5 s; a warm profile opens,
reads, and closes a fixture in about 3 to 4 s. Every operation is one soffice
process.

### Version recording

`describe` reads `ooSetupVersionAboutBox` (plus `ooSetupVersionAboutBoxSuffix`)
from the `/org.openoffice.Setup/Product` configuration node and the `buildid`
from the installation's `versionrc` via the UNO macro expander, for example
`26.8.0.3 (build bce0998afefdbc355585ca324285661a2170ba77)`.

### Read path

`read_model` opens the workbook with `loadComponentFromURL` (`Hidden=True`), so
every value is what Calc's OOXML import produced, then walks the Calc UNO API.
No workbook XML is parsed. Conversions that are not plain property reads:

- Formulas, defined names, conditional-format and validation formulas, and
  internal hyperlink targets are Calc token arrays or Calc-grammar strings,
  printed through `com.sun.star.sheet.FormulaParser` with the OOXML op-code map
  and the `XL_OOX` address convention.
- A formula cell whose result is an error reports `type: "error"` with the
  error text (contract rule). Calc recalculates on import; for example the
  fixture formula `=A3*2` over a text cell reads as `#VALUE!`.
- Calc imports OOXML boolean cells as `=TRUE()` / `=FALSE()` formula cells, and
  the API reports them as formulas, so that is what the adapter reports.
- Date cells: numeric cells whose number format type is DATE, converted with the
  document `NullDate`.
- Row height: Calc stores twips; points = round(height_hmm * 1440 / 2540) / 20.
  Only rows without `OptimalHeight` are reported.
- Column width: `Width` (1/100 mm) divided by the widest digit of the workbook
  default font (`Default` cell style, measured on the document reference device
  as `XclRoot::SetCharWidth` does), truncated to two decimals as
  `XclExpColinfo::SaveXml` writes it. A 20.83203125 OOXML width therefore reads
  as 20.83, and after ExcelBench strips the Excel padding, 19.998.
- Borders: `TopBorder2`... `DiagonalTLBR2` mapped to OOXML styles with the
  width thresholds of `lclGetBorderLine` in `sc/source/filter/excel/xestyle.cxx`.
- Rotation and indent: `XclTools::GetXclRotation` and the export's
  `(twips + 100) / 200` indent level.
- Conditional format kinds are classified by the properties each entry object
  exposes (the numeric `getType()` codes are not consistent across entry
  implementations).
- Explicit validation lists (`"Red";"Green";"Blue"` in Calc) are reported as one
  quoted comma-joined literal, which is how Calc's xlsx export writes them.
- Chart and image `from` cells come from the shape `Anchor`; the chart `to` cell
  is the cell holding the shape's bottom-right point using Calc's own cell
  positions, with a point on a boundary assigned to the preceding cell. That is
  the `xdr:to` Calc's xlsx export writes for the same shape.
- Freeze panes come from the controller (`hasFrozenPanes`, `getSplitColumn`,
  `getSplitRow`); non-frozen splits come from the document view settings.
- Sheet protection: `XProtectable.isProtected()`. Whether a password is set is
  probed with `unprotect("")` (it succeeds only without a password) after every
  other read of that sheet.

### Write path

`write_model` creates a Calc document (not `Hidden`: soffice already runs
headless, and with `Hidden=True` Calc silently drops view operations such as
`freezeAtPosition`), inserts the requested sheets, removes the default sheet,
replays each op through the UNO API, and stores with the `Calc Office Open XML`
filter. Number formats are registered under the document default locale; an
explicit locale makes the export prefix codes with `[$-409]`. `mutate` loads the
template, sets values with `setString` / `setValue`, and stores the same way.

### Unsupported declarations

`describe.unsupported_write` and `model.unsupported` are both empty: Calc's UNO
API can read and write every one of the 22 benchmark features at least in part.
The gaps below are specific attributes; they are not filled in or faked, so the
affected test cases fail.

| Area | Gap | Reason |
| --- | --- | --- |
| Hyperlinks | `tooltip` is always null (read and write) | Calc's URL text field (`com.sun.star.text.TextField.URL`) has no ScreenTip property. |
| Conditional formatting | `priority` and `stop_if_true` are null (read) and not written | Calc's conditional-format API has neither concept; entries are evaluated in insertion order. |
| Conditional formatting | color scale rules are not written | `createEntry(COLORSCALE)` yields a format with zero scale entries, and setting `ColorScaleEntries` indexes existing entries without adding them (`ScColorScaleFormatObj`), which aborts soffice with SIGABRT. No other UNO call adds scale entries. |
| Sheet protection | granular flags (`format_cells`, `insert_rows`, `select_locked_cells`, `select_unlocked_cells`, `sort`, `auto_filter`) are null on read and not written | `XProtectable` exposes only `protect(password)`, `unprotect(password)`, and `isProtected()`. |
| Freeze panes | non-frozen splits do not survive a write | `XViewSplitable.splitAtPosition` takes pixels (OOXML split offsets are twips, so small offsets round to 0 px), and Calc's own xlsx round-trip also drops split panes. Frozen panes work. |
| Comments | `author` is not written | `XSheetAnnotation.Author` is read-only; new notes carry the profile's user name. |
| Pivot tables | `target_cell` reads Calc's `OutputRange` start | That is the cell Calc reports for the pivot output; no other location is exposed. |

LibreOffice behaviors that surface as failed cases (observed, not adapter gaps):
OOXML booleans import as `TRUE()`/`FALSE()` formulas; number format codes are
reported in Calc's dialect (`YYYY-MM-DD`) and exported with Calc's escaping
(`\$#,##0.00`, `"USD "0.00`); conditional-format fills are exported as a
`patternFill` with only `bgColor`; references to undefined names are
lowercased; a header-only table range is not exported; fit-to-page values are
ignored without the `fitToPage` flag and the export always writes `scale` and
`fitToWidth`/`fitToHeight`; the drawing `to` anchor of a shape whose end lies on
a cell boundary is written as the preceding cell plus a full-size offset.

## Open/save and render oracle (`libreoffice_oracle.py`)

```bash
python libreoffice_oracle.py < request.json
```

Missing LibreOffice returns a structured skip.

- `open_save_validate`: converts an input workbook back to `.xlsx` using the
  `Calc Office Open XML` filter.
- `render_validate` / `render_pdf`: exports an input workbook to PDF using the
  `calc_pdf_Export` filter.
