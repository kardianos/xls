XLS Reader
==========

Package `xls` reads binary Excel files (`.xls`): BIFF8 (Excel 97-2003) and
BIFF5/BIFF7 (Excel 5.0/95) workbooks.

It is derived from https://github.com/extrame/xls. The parser has been
rewritten to be safe for untrusted input. Damaged files return an error
wrapping `ErrCorrupt` (files that are not compound files at all wrap
`ole2.ErrNotCompoundFile`) rather than panicking, hanging or exhausting
memory.
Parsing time is linear in the file size and memory is bounded by a small
multiple of it; see the package documentation for the exact guarantees.

```go
wb, err := xls.Open("report.xls", "")
if err != nil {
	return err
}
sheet, err := wb.GetSheet(0)
if err != nil {
	return err
}
for _, row := range sheet.Rows() {
	for c := row.FirstCol(); c < row.LastCol(); c++ {
		fmt.Print(row.Col(c), "\t")
	}
	fmt.Println()
}
```

`Row.Col` returns cell text; dates are formatted with the cell's date format.
`Row.Value` returns the typed value: kind, number, `time.Time` for dates,
boolean, error text, and the number format code. `WorkSheet.Rows` lists the
rows that exist, so the loop above does work proportional to the file;
`Row(i)` for `i` up to `MaxRow` also works.

## What is supported

- Strings, numbers, booleans, error values and cached formula results,
  including formulas whose string result follows a shared or array formula.
- Dates and times in the 1900 and 1904 date systems, from 1900 to 9999,
  including elapsed time formats such as `[h]:mm`.
- Merged cells (`WorkSheet.MergedCells`; as in Excel, `Row.Col` and
  `Row.Value` show the top-left value in every covered cell of rows that
  exist, and `Row.ColExact` shows what the cell itself stores) and
  hyperlinks (`WorkSheet.Hyperlinks`, `Row.Hyperlink`).
- Worksheets and Excel 4.0 macro sheets (`WorkSheet.Type`); chart sheets and
  modules have no rows.
- BIFF5 text in the file's code page, or one you choose with `charset`.
  BIFF5 support is tested with hand-built records only; current LibreOffice
  cannot write Excel 95 files. BIFF8 is also tested with files written by
  LibreOffice (`testdata/real`).
- Compound files with 512 or 4096 byte sectors.

Not supported: encrypted workbooks (`ErrEncrypted`), BIFF2-4 files
(`ErrUnsupportedVersion`), formula evaluation, and number formats other than
dates (numbers are shown in their shortest exact decimal form).

## Changes from v0.0.1

This version changes the API:

- `WorkBook.Close` does nothing. `Open` closes its file before returning,
  and `OpenReader` no longer closes the reader it is given.
- `CellValue` has a `Kind`, and also `Bool` and `Time` fields. `Float` always
  holds the number, and `Int` holds it when it is a whole number.
- `Row.LastCol` is one past the last column, as its loop idiom suggests.
  `WorkSheet.Rows`, `WorkSheet.Type` and `Row.Index` are new. A damaged
  sheet makes `GetSheet` fail for that sheet only.
- `WorkSheet.Selected` reports the selected-tab flag (it read the
  displayed-sheet flag before). `WorkSheet.RightToLeft` is new.
- Errors: `ErrCorrupt`, `ErrInternal`, `ErrEncrypted`, `ErrNoWorkbook`,
  `ErrUnsupportedVersion` and `ErrTooLarge` can be tested with `errors.Is`.
- `ReadAll(0)` means no row limit. Merged cells hold their value only in the
  top-left cell, as in a CSV export. A sheet that cannot be read has nil data
  and its error is joined into the returned error; the other sheets are
  still returned. `ReadAll` returns `ErrTooLarge` when its dense grid would
  be far larger than the file.
- Internal types (`XfRk`, `RK`, `MulrkCol`, `NumberCol`, `Format`, `Font`,
  `XF` and others) and the exported `WorkBook` fields other than `Author`
  were removed. `CellRange` now has plain int fields. The `ole2` and `yymmdd`
  APIs were rewritten; `yymmdd.Compile` and `Layout.FormatSerial` format
  Excel serial numbers the way Excel shows them.

## Fuzzing

The fuzz targets (`FuzzOpen` in `.` and `./ole2`, `FuzzFormat` in
`./yymmdd`) refuse to fuzz without a memory cap. A parser bug found while
fuzzing can otherwise take all of the machine's memory. Use `fuzz.sh`,
which runs `go test` in a memory-limited cgroup (or with `ulimit -v`):

```sh
./fuzz.sh . FuzzOpen                       # 60s, 4 GiB cap
./fuzz.sh -t 10m -m 2G ./ole2 FuzzOpen
./fuzz.sh -- -run 'FuzzOpen/<id>' .        # replay a crasher under the cap
```

Go's fuzzer pauses while it minimizes new inputs; passing
`-fuzzminimizetime 5s` (through `./fuzz.sh -- -run '^$' -fuzz ...`) keeps
throughput up.

Date and time rendering can be checked against the Excel-derived tables of
[SheetJS SSF](https://github.com/SheetJS/ssf) (`test/*.tsv`,
`test/is_date.json`): put them in a directory and run
`SSF_DIR=<dir> go test -run TestSSF ./yymmdd`.
