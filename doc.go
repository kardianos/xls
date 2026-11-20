// Package xls reads binary Excel files (.xls): BIFF8 (Excel 97-2003) and
// BIFF5/BIFF7 (Excel 5.0/95) workbooks stored in a compound file.
//
// The reader is meant to be safe for untrusted input. Malformed files
// return an error rather than panicking or hanging. Parsing takes time
// linear in the file size, and memory bounded by a small multiple of it
// (at most about 40 times, for adversarial merged ranges or long date
// formats; ReadAll has its own documented budget). Row.Col, Row.Value and
// Row.Hyperlink take time logarithmic in the number of cells, merged ranges
// or hyperlinks, plus time linear in the length of a number format, which
// is capped at Excel's 255 characters; the text of a merged range is
// rendered once and reused. MergedCells, Hyperlinks and Rows return copies.
// Iterating WorkSheet.Rows, and each row's columns from FirstCol to
// LastCol, visits at most 256 columns per row that has a record, so the
// work is linear in the file size: a file can declare 256 columns in one
// 20-byte ROW record, about 13 column visits per byte. A visit takes tens
// to hundreds of nanoseconds; rendering a date takes time proportional to
// its text, up to some tens of microseconds for a 255-character format
// whose parts expand (such as dddd). Row.Cols lists only the cells a row
// stores, which is cheaper on sparse rows; ReadAll iterates that way. Memory is bounded relative to the size
// the io.ReadSeeker reports, so a reader that reports far more data than it
// holds can still cost memory.
//
// Structural damage that affects cell data (the compound file, record
// framing, cell records, string records, references to missing shared
// strings) is an error that wraps ErrCorrupt. Damage inside one sheet,
// including a sheet that runs into the next one, makes GetSheet fail for
// that sheet only. Damage confined to extra
// information is skipped: hyperlinks, merged ranges, window settings, the
// author name, and number formats longer than 255 characters (cells using
// them show General numbers). A shared string table with fewer strings than
// it declares is accepted, as long as no cell refers to a missing string.
// A panic while parsing is returned as an error wrapping ErrInternal.
// GetSheet failing means the sheet's contents are unknown; Excel might
// repair some of these files.
//
// As in Excel, the cells of a merged range show the range's top-left value
// through Col and Value, even a covered cell that holds a value of its own
// (ColExact returns that). Comments, text boxes, headers and footers,
// charts and defined names are not read.
//
// The workbook stream is found the way Excel finds it, by searching the
// compound file's directory tree. If that fails (a tree out of order), a
// stream with a unique matching name is used, as LibreOffice and xlrd do.
//
// Encrypted workbooks return ErrEncrypted. Formulas are not evaluated; a
// formula cell has the value Excel cached when it last saved the file.
//
// Cell text is rendered as follows. Strings are returned as stored.
// Booleans become "TRUE" or "FALSE" and errors their error text, such as
// "#N/A". Numbers in a date or time format are formatted with that format
// (see package yymmdd). Other numbers use the shortest decimal form that
// reads back as the same float64, with an exponent below 1e-6 and from 1e21
// on. Number formats such as currency, percent and digit grouping are not
// applied. Use Row.Value to get the typed value and its format code.
package xls
