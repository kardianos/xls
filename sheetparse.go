package xls

import (
	"bytes"
	"encoding/binary"
	"math"
	"sort"
	"strings"
	"sync/atomic"
	"unicode/utf16"
)

// parse reads the sheet substream ([MS-XLS] 2.1.7.20.5). Substreams nested
// in it, such as embedded charts, are skipped.
func (w *WorkSheet) parse() error {
	w.rows = make(map[uint16]*Row)
	wb := w.wb
	it := &recordIter{b: wb.stream[:w.end], off: w.offset}
	if _, _, err := it.next(); err != nil { // the BOF record, checked by scanSheet
		return err
	}
	if w.Type != SheetWorksheet && w.Type != SheetMacro {
		// Chart sheets and Visual Basic modules have no cells.
		return nil
	}

	var pending *cell // a formula whose string result is in a later STRING record
	depth := 0
	for {
		r, ok, err := it.next()
		if err != nil {
			return err
		}
		if !ok {
			return corruptf("sheet %q has no EOF record", w.Name)
		}
		switch r.id {
		case recBOF:
			depth++
			continue
		case recEOF:
			if depth == 0 {
				w.finish()
				return nil
			}
			depth--
			continue
		}
		if depth > 0 {
			continue
		}
		switch r.id {
		case recSTRING, recSHRFMLA, recARRAY, recTABLE:
		default:
			pending = nil
		}

		rd := newReader(r)
		switch r.id {
		case recROW:
			if rd.remaining() < 16 {
				return corruptf("ROW record at offset %d is %d bytes, not 16", r.pos, rd.remaining())
			}
			rw, mic, mac := rd.u16(), rd.u16(), rd.u16()
			row := w.row(rw)
			row.hasInfo, row.colMic, row.colMac = true, min(mic, maxCols-1), min(mac, maxCols)
		case recNUMBER:
			rw, col, xf := rd.u16(), rd.u16(), rd.u16()
			f := math.Float64frombits(rd.u64())
			if !rd.bad {
				w.add(rw, cell{col: col, xf: xf, kind: cellNumber, num: f})
			}
		case recRK:
			rw, col, xf := rd.u16(), rd.u16(), rd.u16()
			rk := rd.u32()
			if !rd.bad {
				w.add(rw, cell{col: col, xf: xf, kind: cellNumber, num: rkValue(rk)})
			}
		case recMULRK:
			rw, first := rd.u16(), rd.u16()
			n := (rd.remaining() - 2) / 6
			if rd.bad || n < 1 || rd.remaining() != 6*n+2 || int(first)+n-1 > math.MaxUint16 {
				return corruptf("malformed MULRK record at offset %d", r.pos)
			}
			for i := 0; i < n; i++ {
				xf, rk := rd.u16(), rd.u32()
				w.add(rw, cell{col: first + uint16(i), xf: xf, kind: cellNumber, num: rkValue(rk)})
			}
			if last := rd.u16(); int(last) != int(first)+n-1 {
				return corruptf("MULRK record at offset %d: column range does not match its %d values", r.pos, n)
			}
		case recMULBLANK:
			rw, first := rd.u16(), rd.u16()
			n := (rd.remaining() - 2) / 2
			if rd.bad || n < 1 || rd.remaining() != 2*n+2 {
				return corruptf("malformed MULBLANK record at offset %d", r.pos)
			}
			rd.skip(2 * n)
			if last := rd.u16(); int(last) != int(first)+n-1 {
				return corruptf("MULBLANK record at offset %d: column range does not match its %d cells", r.pos, n)
			}
			w.blank(rw, first, first+uint16(n-1))
		case recBLANK:
			rw, col := rd.u16(), rd.u16()
			if !rd.bad {
				w.blank(rw, col, col)
			}
		case recLABELSST:
			rw, col, xf := rd.u16(), rd.u16(), rd.u16()
			idx := rd.u32()
			if rd.bad {
				break
			}
			if int64(idx) >= int64(len(wb.sst)) {
				return corruptf("LABELSST record at offset %d refers to string %d of %d", r.pos, idx, len(wb.sst))
			}
			w.add(rw, cell{col: col, xf: xf, kind: cellString, str: wb.sst[idx]})
		case recLABEL, recRSTRING:
			rw, col, xf := rd.u16(), rd.u16(), rd.u16()
			var s string
			if wb.biff == 8 {
				s = rd.unicodeString(2)
			} else {
				s = rd.byteString(2, wb.enc)
			}
			if !rd.bad {
				w.add(rw, cell{col: col, xf: xf, kind: cellString, str: s})
			}
		case recBOOLERR:
			rw, col, xf := rd.u16(), rd.u16(), rd.u16()
			v, isErr := rd.u8(), rd.u8()
			if rd.bad {
				break
			}
			kind := uint8(cellBool)
			if isErr != 0 {
				kind = cellError
			}
			w.add(rw, cell{col: col, xf: xf, kind: kind, b: v})
		case recFORMULA:
			rw, col, xf := rd.u16(), rd.u16(), rd.u16()
			val := rd.bytes(8)
			if rd.bad {
				break
			}
			c := cell{col: col, xf: xf}
			if val[6] == 0xFF && val[7] == 0xFF {
				// [MS-XLS] 2.5.133 FormulaValue
				switch val[0] {
				case 0:
					c.kind = cellEmptyString // replaced by the STRING record
				case 1:
					c.kind, c.b = cellBool, val[2]
				case 2:
					c.kind, c.b = cellError, val[2]
				case 3:
					c.kind = cellEmptyString
				default:
					return corruptf("FORMULA record at offset %d has unknown result type %d", r.pos, val[0])
				}
			} else {
				c.kind = cellNumber
				c.num = math.Float64frombits(binary.LittleEndian.Uint64(val))
			}
			added := w.add(rw, c)
			if val[6] == 0xFF && val[7] == 0xFF && val[0] == 0 {
				// Only STRING, SHRFMLA, ARRAY and TABLE records may come
				// before the STRING record, and none of them adds a cell,
				// so the pointer stays valid.
				pending = added
			}
		case recSTRING:
			if pending == nil {
				break // no formula waits for a string; ignore it
			}
			var s string
			if wb.biff == 8 {
				s = rd.unicodeString(2)
			} else {
				s = rd.byteString(2, wb.enc)
			}
			if rd.bad {
				return corruptf("malformed STRING record at offset %d", r.pos)
			}
			pending.kind, pending.str = cellString, s
			pending = nil
		case recMERGECELLS:
			n := int(rd.u16())
			for i := 0; i < n && !rd.bad; i++ {
				m := CellRange{FirstRow: int(rd.u16()), LastRow: int(rd.u16()), FirstCol: int(rd.u16()), LastCol: int(rd.u16())}
				if !rd.bad && m.FirstRow <= m.LastRow && m.FirstCol <= m.LastCol && m.LastCol < maxCols {
					w.merges = append(w.merges, toRef(m))
				}
			}
			rd.bad = false // keep the ranges read before any damage
		case recHLINK:
			if h, ok := parseHyperlink(rd, wb); ok {
				w.links = append(w.links, h)
			}
			rd.bad = false // a malformed hyperlink is skipped
		case recWINDOW2:
			flags := rd.u16()
			if !rd.bad {
				w.RightToLeft = flags&0x0040 != 0
				w.Selected = flags&0x0200 != 0
			}
			rd.bad = false
		}
		if rd.bad {
			return corruptf("record %#04x at offset %d is too short", r.id, r.pos)
		}
		if w.badCol {
			return corruptf("record %#04x at offset %d names a column past %d", r.id, r.pos, maxCols-1)
		}
	}
}

// rkValue decodes an RK number ([MS-XLS] 2.5.217).
func rkValue(rk uint32) float64 {
	var f float64
	if rk&0x02 != 0 {
		f = float64(int32(rk) >> 2)
	} else {
		f = math.Float64frombits(uint64(rk&0xFFFFFFFC) << 32)
	}
	if rk&0x01 != 0 {
		f /= 100
	}
	return f
}

func (w *WorkSheet) row(i uint16) *Row {
	r := w.rows[i]
	if r == nil {
		r = &Row{sheet: w, index: int(i), minCol: -1}
		w.rows[i] = r
		if i > w.MaxRow {
			w.MaxRow = i
		}
	}
	return r
}

func (r *Row) extend(first, last uint16) {
	if r.minCol < 0 || int(first) < r.minCol {
		r.minCol = int(first)
	}
	if int(last)+1 > r.width {
		r.width = int(last) + 1
	}
}

// maxCols is the number of columns a BIFF5 or BIFF8 sheet can have.
const maxCols = 256

// add stores a cell and returns it, or returns nil if the column is out
// of range.
func (w *WorkSheet) add(rw uint16, c cell) *cell {
	if c.col >= maxCols {
		w.badCol = true
		return nil
	}
	r := w.row(rw)
	r.extend(c.col, c.col)
	r.cells = append(r.cells, c)
	return &r.cells[len(r.cells)-1]
}

func (w *WorkSheet) blank(rw, first, last uint16) {
	if last >= maxCols {
		w.badCol = true
		return
	}
	w.row(rw).extend(first, last)
}

// finish sorts the cells of each row by column and keeps the last record
// for a column that appears more than once.
func (w *WorkSheet) finish() {
	for _, r := range w.rows {
		cells := r.cells
		sort.SliceStable(cells, func(i, j int) bool { return cells[i].col < cells[j].col })
		out := cells[:0]
		for i := range cells {
			if len(out) > 0 && out[len(out)-1].col == cells[i].col {
				out[len(out)-1] = cells[i]
				continue
			}
			out = append(out, cells[i])
		}
		r.cells = out
	}
	w.mergeIdx = newRangeIndex(w.merges)
	w.mergeText = make([]atomic.Pointer[string], len(w.merges))
	w.rowList = make([]*Row, 0, len(w.rows))
	for _, r := range w.rows {
		w.rowList = append(w.rowList, r)
	}
	sort.Slice(w.rowList, func(i, j int) bool { return w.rowList[i].index < w.rowList[j].index })
	ranges := make([]ref, len(w.links))
	for i, h := range w.links {
		ranges[i] = toRef(h.CellRange)
	}
	w.linkIdx = newRangeIndex(ranges)
}

// HyperLink is a hyperlink attached to a range of cells.
type HyperLink struct {
	CellRange
	// Description is the display text of the link.
	Description string
	// TextMark is the location in the target, such as a cell reference.
	TextMark string
	// TargetFrame is the frame the link opens in.
	TargetFrame string
	// URL is the target of a URL link.
	URL string
	// ShortedFilePath is the 8.3 path of a file link.
	ShortedFilePath string
	// ExtendedFilePath is the full Unicode path of a file link.
	ExtendedFilePath string
	// IsURL reports whether the link targets a URL.
	IsURL bool
}

var (
	clsidURLMoniker  = []byte{0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B}
	clsidFileMoniker = []byte{0x03, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}
)

// maxParentRefs limits the "..\" references of a file hyperlink, so a
// small record cannot expand into a long path.
const maxParentRefs = 64

// parseHyperlink reads an HLINK record ([MS-XLS] 2.4.140) and the
// Hyperlink Object it holds ([MS-OSHARED] 2.3.7.1).
func parseHyperlink(rd *reader, wb *WorkBook) (HyperLink, bool) {
	var h HyperLink
	h.FirstRow, h.LastRow = int(rd.u16()), int(rd.u16())
	h.FirstCol, h.LastCol = int(rd.u16()), int(rd.u16())
	rd.skip(16) // hlinkClsid
	rd.u32()    // streamVersion
	flags := rd.u32()
	if rd.bad || h.FirstRow > h.LastRow || h.FirstCol > h.LastCol || h.LastCol >= maxCols {
		return h, false
	}
	if flags&0x10 != 0 {
		h.Description = hyperlinkString(rd)
	}
	if flags&0x80 != 0 {
		h.TargetFrame = hyperlinkString(rd)
	}
	if flags&0x01 != 0 {
		if flags&0x100 != 0 {
			h.URL = hyperlinkString(rd)
			h.IsURL = true
		} else {
			clsid := rd.bytes(16)
			switch {
			case bytes.Equal(clsid, clsidURLMoniker):
				n := int(rd.u32())
				h.URL = utf16z(rd.bytes(n))
				h.IsURL = true
			case bytes.Equal(clsid, clsidFileMoniker):
				// cAnti counts the "..\" parent references that the
				// paths leave out.
				anti := int(rd.u16())
				if anti > maxParentRefs {
					return h, false
				}
				up := strings.Repeat(`..\`, anti)
				n := int(rd.u32())
				ansi := rd.bytes(n)
				if i := bytes.IndexByte(ansi, 0); i >= 0 {
					ansi = ansi[:i]
				}
				h.ShortedFilePath = up + decode8(ansi, wb.enc)
				rd.skip(24) // endServer, versionNumber, reserved1, reserved2
				if size := rd.u32(); size > 0 {
					n := int(rd.u32())
					rd.u16() // usKeyValue
					h.ExtendedFilePath = up + utf16z(rd.bytes(n))
				}
			default:
				// Other monikers have no length field, so the fields
				// that follow cannot be found.
				return h, !rd.bad
			}
		}
	}
	if flags&0x08 != 0 {
		h.TextMark = hyperlinkString(rd)
	}
	return h, !rd.bad
}

// hyperlinkString reads a HyperlinkString: a count of UTF-16 characters
// that includes a terminating NUL, then the characters.
func hyperlinkString(rd *reader) string {
	n := int(rd.u32())
	if n > rd.remaining()/2 {
		rd.bad = true
		return ""
	}
	return utf16z(rd.bytes(2 * n))
}

// utf16z decodes little-endian UTF-16 up to the first NUL.
func utf16z(b []byte) string {
	u := make([]uint16, 0, len(b)/2)
	for i := 0; i+1 < len(b); i += 2 {
		c := uint16(b[i]) | uint16(b[i+1])<<8
		if c == 0 {
			break
		}
		u = append(u, c)
	}
	return string(utf16.Decode(u))
}
