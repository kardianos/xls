package xls

import (
	"bytes"
	"encoding/binary"
	"math"
	"testing"

	"github.com/kardianos/xls/internal/cfbtest"
)

// Helpers that build BIFF records and workbooks for tests.

func le(vals ...interface{}) []byte {
	var buf bytes.Buffer
	for _, v := range vals {
		if err := binary.Write(&buf, binary.LittleEndian, v); err != nil {
			panic(err)
		}
	}
	return buf.Bytes()
}

func cat(parts ...[]byte) []byte { return bytes.Join(parts, nil) }

func rec(id uint16, data ...[]byte) []byte {
	d := cat(data...)
	return cat(le(id, uint16(len(d))), d)
}

func bofRec(vers, dt uint16) []byte {
	return rec(recBOF, le(vers, dt, uint16(0), uint16(0), uint32(0), uint32(0)))
}

func eofRec() []byte { return rec(recEOF) }

// ustr is a BIFF8 XLUnicodeString with a 16-bit length, compressed when
// every character fits in 8 bits.
func ustr(s string) []byte {
	r := []rune(s)
	wide := false
	for _, c := range r {
		if c > 0xFF {
			wide = true
		}
	}
	if !wide {
		b := cat(le(uint16(len(r)), uint8(0)))
		for _, c := range r {
			b = append(b, byte(c))
		}
		return b
	}
	b := cat(le(uint16(len(r)), uint8(1)))
	for _, c := range r {
		b = append(b, le(uint16(c))...)
	}
	return b
}

func xfRec(fmtIdx uint16) []byte {
	return rec(recXF, le(uint16(0), fmtIdx), make([]byte, 16))
}

func formatRec(idx uint16, code string) []byte { return rec(recFORMAT, le(idx), ustr(code)) }

func numberRec(row, col, xf uint16, f float64) []byte {
	return rec(recNUMBER, le(row, col, xf, f))
}

func rkRec(row, col, xf uint16, rk uint32) []byte { return rec(recRK, le(row, col, xf, rk)) }

func rkInt(v int32) uint32 { return uint32(v)<<2 | 2 }

func labelRec(row, col uint16, s string) []byte {
	return rec(recLABEL, le(row, col, uint16(0)), ustr(s))
}

func sstRec(strs ...string) []byte {
	var body []byte
	for _, s := range strs {
		body = append(body, ustr(s)...)
	}
	return rec(recSST, le(uint32(len(strs)), uint32(len(strs))), body)
}

func labelSSTRec(row, col, xf uint16, idx uint32) []byte {
	return rec(recLABELSST, le(row, col, xf, idx))
}

func formulaRec(row, col, xf uint16, val [8]byte) []byte {
	return rec(recFORMULA, le(row, col, xf), val[:], le(uint16(0), uint32(0), uint16(0)))
}

func numVal(f float64) (v [8]byte) {
	binary.LittleEndian.PutUint64(v[:], math.Float64bits(f))
	return v
}

func specialVal(typ, b byte) [8]byte { return [8]byte{typ, 0, b, 0, 0, 0, 0xFF, 0xFF} }

// book describes a workbook stream.
type book struct {
	biff    int // 8 (default) or 5
	globals [][]byte
	sheets  []sheet
}

type sheet struct {
	name string
	dt   uint16 // BOF substream type; 0 means worksheet
	recs [][]byte
}

func (b book) stream() []byte {
	vers := uint16(0x0600)
	if b.biff == 5 {
		vers = 0x0500
	}
	var g bytes.Buffer
	g.Write(bofRec(vers, 0x0005))
	for _, r := range b.globals {
		g.Write(r)
	}
	var patch []int
	for _, s := range b.sheets {
		patch = append(patch, g.Len()+4)
		var name []byte
		if b.biff == 5 {
			name = cat([]byte{byte(len(s.name))}, []byte(s.name))
		} else {
			name = cat([]byte{byte(len(s.name)), 0}, []byte(s.name))
		}
		g.Write(rec(recBOUNDSHEET, le(uint32(0), uint8(0), uint8(0)), name))
	}
	g.Write(eofRec())
	out := g.Bytes()
	for i, s := range b.sheets {
		binary.LittleEndian.PutUint32(out[patch[i]:], uint32(len(out)))
		dt := s.dt
		if dt == 0 {
			dt = 0x0010
		}
		out = append(out, bofRec(vers, dt)...)
		for _, r := range s.recs {
			out = append(out, r...)
		}
		out = append(out, eofRec()...)
	}
	return out
}

// file wraps the stream in a compound file.
func (b book) file() []byte {
	name := "Workbook"
	if b.biff == 5 {
		name = "Book"
	}
	f, _ := cfbtest.Build([]cfbtest.Stream{{Name: name, Data: b.stream()}}, cfbtest.Options{})
	return f
}

func openBook(t *testing.T, b book) (*WorkBook, error) {
	t.Helper()
	return OpenReader(bytes.NewReader(b.file()), "")
}

// mustSheet opens a one-sheet workbook made of globals and cell records.
func mustSheet(t *testing.T, globals [][]byte, recs ...[]byte) (*WorkBook, *WorkSheet) {
	t.Helper()
	wb, err := openBook(t, book{globals: globals, sheets: []sheet{{name: "Sheet1", recs: recs}}})
	if err != nil {
		t.Fatalf("open: %v", err)
	}
	s, err := wb.GetSheet(0)
	if err != nil {
		t.Fatalf("GetSheet: %v", err)
	}
	return wb, s
}

// col returns the text of a cell, or a marker if the row is missing.
func col(s *WorkSheet, row, c int) string {
	r := s.Row(row)
	if r == nil {
		return "<no row>"
	}
	return r.Col(c)
}
