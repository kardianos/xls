package xls

import (
	"errors"
	"fmt"
	"math"
	"sort"
	"strings"
	"unicode/utf8"

	"golang.org/x/text/encoding"
)

var (
	// ErrEncrypted is returned for password-protected workbooks.
	ErrEncrypted = errors.New("xls: workbook is encrypted")
	// ErrNoWorkbook is returned when a compound file has no Workbook stream.
	ErrNoWorkbook = errors.New("xls: no Excel workbook stream found")
	// ErrUnsupportedVersion is returned for BIFF2, BIFF3 and BIFF4 files.
	ErrUnsupportedVersion = errors.New("xls: unsupported BIFF version")
	// ErrTooLarge is returned by ReadAll when the grid it would build is
	// much larger than the file. Read such files with GetSheet and Row.
	ErrTooLarge = errors.New("xls: sheet grid too large for ReadAll")
)

// WorkBook is a parsed XLS file. After Open returns, a WorkBook and its
// sheets are safe for concurrent use.
type WorkBook struct {
	// Author is the user name in the file's WRITEACCESS record, if any.
	Author string

	stream   []byte
	biff     int // 5 for BIFF5/BIFF7, 8 for BIFF8
	date1904 bool
	enc      encoding.Encoding // decodes 8-bit strings in BIFF5 files
	formats  map[uint16]string
	xfs      []xfInfo
	sst      []string
	sheets   []*WorkSheet
}

// parseWorkbook parses the workbook globals substream ([MS-XLS] 2.1.7.20.1).
// biff is the version implied by the stream name. It is used when the BOF
// record does not give one.
func parseWorkbook(stream []byte, biff int, charset encoding.Encoding) (*WorkBook, error) {
	w := &WorkBook{
		stream:  stream,
		biff:    biff,
		formats: make(map[uint16]string),
	}
	it := &recordIter{b: stream}
	r, ok, err := it.next()
	if err != nil {
		return nil, err
	}
	if !ok {
		return nil, corruptf("empty workbook stream")
	}
	switch r.id {
	case recBOF:
	case recBOF2, recBOF3, recBOF4:
		return nil, ErrUnsupportedVersion
	default:
		return nil, corruptf("workbook stream does not start with a BOF record")
	}
	rd := newReader(r)
	vers, dt := rd.u16(), rd.u16()
	if rd.bad {
		return nil, corruptf("short BOF record")
	}
	switch vers {
	case 0x0600:
		w.biff = 8
	case 0x0500:
		w.biff = 5
	}
	if dt != 0x0005 {
		return nil, corruptf("first substream has type %#x, not workbook globals", dt)
	}

	// Find the code page first. In BIFF5 it follows WRITEACCESS, which
	// already needs it.
	w.enc = charset
	if w.enc == nil {
		w.enc = codepageEncoding(1252)
		if cp, ok := findCodepage(stream[it.off:]); ok {
			w.enc = codepageEncoding(cp)
		}
	}

	var xfFormats []uint16
	for {
		r, ok, err := it.next()
		if err != nil {
			return nil, err
		}
		if !ok {
			return nil, corruptf("workbook globals have no EOF record")
		}
		if r.id == recEOF {
			break
		}
		rd := newReader(r)
		switch r.id {
		case recFILEPASS:
			return nil, ErrEncrypted
		case recBOF:
			return nil, corruptf("unexpected BOF record at offset %d", r.pos)
		case recDATEMODE:
			w.date1904 = rd.u16() != 0
		case recWRITEACCESS:
			var name string
			if w.biff == 8 {
				name = rd.unicodeString(2)
			} else {
				name = rd.byteString(1, w.enc)
			}
			if !rd.bad {
				w.Author = strings.TrimRight(name, " \x00")
			}
			continue // a malformed user name does not make the file unreadable
		case recFORMAT:
			idx := rd.u16()
			var code string
			if w.biff == 8 {
				code = rd.unicodeString(2)
			} else {
				code = rd.byteString(1, w.enc)
			}
			if !rd.bad {
				if utf8.RuneCountInString(code) > maxFormatLen {
					code = "General"
				}
				w.formats[idx] = code
			}
		case recXF:
			rd.u16() // font
			xfFormats = append(xfFormats, rd.u16())
		case recBOUNDSHEET:
			pos := rd.u32()
			state, _ := rd.u8(), rd.u8()
			var name string
			if w.biff == 8 {
				name = rd.unicodeString(1)
			} else {
				name = rd.byteString(1, w.enc)
			}
			if rd.bad {
				break
			}
			s := &WorkSheet{
				Name:       name,
				Visibility: TWorkSheetVisibility(state & 0x03),
				wb:         w,
				offset:     int(pos),
			}
			if int64(pos) >= int64(len(stream)) {
				s.scanErr = corruptf("sheet %q starts at %d, past the end of the stream", name, pos)
			}
			w.sheets = append(w.sheets, s)
		case recSST:
			if err := w.parseSST(rd); err != nil {
				return nil, err
			}
			continue
		}
		if rd.bad {
			return nil, corruptf("record %#04x at offset %d is too short", r.id, r.pos)
		}
	}
	w.resolveFormats(xfFormats)
	if err := w.findSheetExtents(it.off); err != nil {
		return nil, err
	}
	return w, nil
}

// findSheetExtents finds where each sheet substream ends. Sheets must
// follow the workbook globals and must not overlap, so that the total work
// of parsing every sheet is bounded by the stream size. A sheet with bad
// framing, or one that runs into the next sheet, gets an error that
// GetSheet reports; the other sheets stay readable.
func (w *WorkBook) findSheetExtents(globalsEnd int) error {
	var order []*WorkSheet
	for _, s := range w.sheets {
		switch {
		case s.scanErr != nil:
		case s.offset < globalsEnd:
			s.scanErr = corruptf("sheet %q starts inside the workbook globals", s.Name)
		default:
			order = append(order, s)
		}
	}
	sort.SliceStable(order, func(i, j int) bool { return order[i].offset < order[j].offset })
	// Sheets that share an offset are all damaged; none of them is read.
	kept := order[:0]
	for i := 0; i < len(order); {
		j := i + 1
		for j < len(order) && order[j].offset == order[i].offset {
			j++
		}
		if j-i > 1 {
			for _, s := range order[i:j] {
				s.scanErr = corruptf("sheet %q starts at the same offset as another sheet", s.Name)
			}
		} else {
			kept = append(kept, order[i])
		}
		i = j
	}
	for i, s := range kept {
		limit := len(w.stream)
		if i+1 < len(kept) {
			limit = kept[i+1].offset
		}
		s.scanErr = w.scanSheet(s, limit)
	}
	return nil
}

// scanSheet walks the records of one sheet, which must end at or before
// limit, and records where it ends and its type. It reads only
// stream[:limit], so the scans of all sheets together read each byte of the
// stream at most once.
func (w *WorkBook) scanSheet(s *WorkSheet, limit int) error {
	it := &recordIter{b: w.stream[:limit], off: s.offset}
	depth := 0
	for {
		r, ok, err := it.next()
		if err != nil {
			return fmt.Errorf("sheet %q: %w", s.Name, err)
		}
		if !ok {
			return corruptf("sheet %q has no EOF record before the next sheet or the end of the stream", s.Name)
		}
		if r.pos == s.offset {
			rd := newReader(r)
			rd.u16()
			s.Type = SheetType(rd.u16())
			if r.id != recBOF || rd.bad {
				return corruptf("sheet %q does not start with a BOF record", s.Name)
			}
		}
		switch r.id {
		case recBOF:
			depth++
		case recEOF:
			depth--
			if depth == 0 {
				s.end = it.off
				return nil
			}
		}
	}
}

func findCodepage(b []byte) (uint16, bool) {
	it := &recordIter{b: b}
	for {
		r, ok, err := it.next()
		if err != nil || !ok || r.id == recEOF {
			return 0, false
		}
		if r.id == recCODEPAGE {
			rd := newReader(r)
			cp := rd.u16()
			return cp, !rd.bad
		}
	}
}

// parseSST reads the shared string table ([MS-XLS] 2.4.265).
func (w *WorkBook) parseSST(rd *reader) error {
	rd.u32() // total references
	unique := rd.u32()
	if rd.bad {
		return corruptf("short SST record")
	}
	// Each string takes at least 3 bytes, which limits the allocation.
	n := int(unique)
	if max := rd.remaining() / 3; n > max {
		n = max
	}
	w.sst = make([]string, 0, n)
	for i := uint32(0); i < unique && rd.remaining() > 0; i++ {
		s := rd.richString()
		if rd.bad {
			return corruptf("SST string %d is truncated", i)
		}
		w.sst = append(w.sst, s)
	}
	return nil
}

// NumSheets returns the number of sheets in the workbook.
func (w *WorkBook) NumSheets() int {
	return len(w.sheets)
}

// GetSheet returns sheet num (zero-based), parsing it on first use.
func (w *WorkBook) GetSheet(num int) (*WorkSheet, error) {
	if num < 0 || num >= len(w.sheets) {
		return nil, fmt.Errorf("xls: sheet index %d not found (%d total sheets)", num, len(w.sheets))
	}
	s := w.sheets[num]
	if s.scanErr != nil {
		return nil, s.scanErr
	}
	s.once.Do(func() {
		defer func() {
			if r := recover(); r != nil {
				s.err = fmt.Errorf("%w: sheet %q: %v", ErrInternal, s.Name, r)
			}
		}()
		s.err = s.parse()
	})
	if s.err != nil {
		return nil, s.err
	}
	return s, nil
}

// ReadAll returns the text of every cell in every sheet, indexed by sheet,
// row and column, along with the sheet names. Merged cells hold their value
// only in the top-left cell, as in a CSV export. Rows with no cells at all,
// not even blank ones, are nil; blank cells are "".
//
// maxPerSheet limits the number of rows returned per sheet. A value of zero
// or less means no limit.
//
// ReadAll builds a dense grid: 24 bytes for every row up to the last one,
// 16 bytes for every column up to a row's last cell, and the text of every
// number and date. A sparse sheet can make that grid far larger than the
// file. ReadAll returns ErrTooLarge rather than use more than 64 MiB plus 64
// times the size of the workbook stream. Read such files with GetSheet and
// Row.
//
// sheetData always has one entry per sheet. A sheet that cannot be read
// has nil data, and its error is part of the returned error (see
// errors.Join); the other sheets are still returned. ErrTooLarge stops
// ReadAll, and the sheets it did not reach are nil.
func (w *WorkBook) ReadAll(maxPerSheet int) (sheetData [][][]string, sheetName []string, err error) {
	defer func() {
		for len(sheetData) < len(w.sheets) {
			sheetData = append(sheetData, nil)
		}
	}()
	sheetName = make([]string, len(w.sheets))
	for si, s := range w.sheets {
		sheetName[si] = s.Name
	}
	budget := readAllBudget(len(w.stream))
	var errs []error
	for si := range w.sheets {
		sheet, err := w.GetSheet(si)
		if err != nil {
			errs = append(errs, err)
			sheetData = append(sheetData, nil)
			continue
		}
		if len(sheet.rows) == 0 {
			sheetData = append(sheetData, [][]string{})
			continue
		}
		n := int(sheet.MaxRow) + 1
		if maxPerSheet > 0 && maxPerSheet < n {
			n = maxPerSheet
		}
		budget -= 24 * int64(n)
		for k, row := range sheet.rows {
			if int(k) < n {
				budget -= 16 * int64(row.width)
			}
		}
		if budget < 0 {
			return sheetData, sheetName, errors.Join(append(errs, fmt.Errorf("%w: sheet %q", ErrTooLarge, sheet.Name))...)
		}
		sd := make([][]string, n)
		for k, row := range sheet.rows {
			if int(k) >= n || row.width == 0 {
				continue
			}
			data := make([]string, row.width)
			for i := range row.cells {
				c := &row.cells[i]
				if c.kind != cellNumber {
					data[c.col] = sheet.text(c) // shared with the sheet
					continue
				}
				// Charge the most the text can take before rendering it,
				// so an oversized grid is refused without the work.
				most := int64(w.xf(c.xf).maxLen)
				if budget -= most; budget < 0 {
					return sheetData, sheetName, errors.Join(append(errs, fmt.Errorf("%w: sheet %q", ErrTooLarge, sheet.Name))...)
				}
				data[c.col] = sheet.text(c)
				budget += most - int64(len(data[c.col]))
			}
			sd[k] = data
		}
		sheetData = append(sheetData, sd)
	}
	return sheetData, sheetName, errors.Join(errs...)
}

// readAllBudget is the most memory ReadAll uses for a stream of n bytes.
var readAllBudget = func(n int) int64 { return 64<<20 + 64*int64(n) }

// isIntegral reports whether f has no fractional part and fits in an int64.
func isIntegral(f float64) bool {
	return f == math.Trunc(f) && f >= -(1<<63) && f < 1<<63
}
