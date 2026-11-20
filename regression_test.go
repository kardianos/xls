package xls

import (
	"bytes"
	"encoding/binary"
	"errors"
	"math"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/kardianos/xls/internal/cfbtest"
	"github.com/kardianos/xls/ole2"
)

// These tests cover the problems found in the November 2026 review.

func TestOpenReaderErrorDoesNotPanic(t *testing.T) {
	// The reader is not an io.Closer; the old code called Close on nil here.
	stream := book{sheets: []sheet{{name: "S"}}}.stream()
	for _, n := range []int{2, 10, 30} {
		f, _ := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream[:n]}}, cfbtest.Options{})
		if _, err := OpenReader(bytes.NewReader(f), ""); !errors.Is(err, ErrCorrupt) {
			t.Errorf("stream cut to %d bytes: err = %v", n, err)
		}
	}
}

func TestOpenClosesFile(t *testing.T) {
	if runtime.GOOS != "linux" {
		t.Skip("counts /proc/self/fd")
	}
	dir := t.TempDir()
	bad := filepath.Join(dir, "bad.xls")
	good := filepath.Join(dir, "good.xls")
	os.WriteFile(bad, make([]byte, 1024), 0o644)
	os.WriteFile(good, book{sheets: []sheet{{name: "S"}}}.file(), 0o644)
	count := func() int { e, _ := os.ReadDir("/proc/self/fd"); return len(e) }
	before := count()
	for i := 0; i < 20; i++ {
		if _, err := Open(bad, ""); err == nil {
			t.Fatal("no error")
		}
		if _, err := Open(good, ""); err != nil {
			t.Fatal(err)
		}
	}
	if after := count(); after > before {
		t.Errorf("open files: %d before, %d after", before, after)
	}
}

func urlHyperlink(row, c uint16, url, desc string) []byte {
	us := func(s string) []byte {
		b := le(uint32(len([]rune(s)) + 1))
		for _, r := range s {
			b = append(b, le(uint16(r))...)
		}
		return append(b, 0, 0)
	}
	urlData := us(url)[4:] // the URL moniker stores a byte count, then the string
	return rec(recHLINK,
		le(row, row, c, c), make([]byte, 16), le(uint32(2), uint32(0x01|0x02|0x10|0x08)),
		us(desc),
		clsidURLMoniker, le(uint32(len(urlData))), urlData,
		us("Sheet1!A1"),
	)
}

func TestHyperlinkOnCell(t *testing.T) {
	_, s := mustSheet(t, [][]byte{xfRec(0)},
		labelRec(0, 0, "click me"),
		urlHyperlink(0, 0, "https://example.com/", "Example"),
		urlHyperlink(2, 1, "https://example.org/", "Empty cell"),
	)
	if got := col(s, 0, 0); got != "click me" {
		t.Errorf("Col = %q", got)
	}
	h, ok := s.Row(0).Hyperlink(0)
	if !ok || h.URL != "https://example.com/" || !h.IsURL || h.Description != "Example" || h.TextMark != "Sheet1!A1" {
		t.Errorf("Hyperlink = %+v, %v", h, ok)
	}
	if n := len(s.Hyperlinks()); n != 2 {
		t.Errorf("%d hyperlinks", n)
	}
}

func TestMalformedHyperlinkSkipped(t *testing.T) {
	bad := rec(recHLINK, le(uint16(0), uint16(0), uint16(0), uint16(0)), make([]byte, 16), le(uint32(2), uint32(0x10), uint32(0xFFFFFFFF)))
	_, s := mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "a"), bad)
	if len(s.Hyperlinks()) != 0 || col(s, 0, 0) != "a" {
		t.Errorf("malformed hyperlink not skipped")
	}
}

func TestDuplicateCellLastWins(t *testing.T) {
	_, s := mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "a"), labelRec(0, 0, "b"))
	if got := col(s, 0, 0); got != "b" {
		t.Errorf("Col = %q, want b", got)
	}
}

func TestFormulaResults(t *testing.T) {
	_, s := mustSheet(t, [][]byte{xfRec(0)},
		formulaRec(0, 0, 0, numVal(42)),
		formulaRec(0, 1, 0, specialVal(1, 1)),
		formulaRec(0, 2, 0, specialVal(2, 0x07)),
		formulaRec(0, 3, 0, specialVal(3, 0)),
		formulaRec(0, 4, 0, specialVal(0, 0)),
		rec(recSHRFMLA, make([]byte, 10)),
		rec(recSTRING, ustr("shared")),
		formulaRec(0, 5, 0, specialVal(0, 0)),
		rec(recARRAY, make([]byte, 14)),
		rec(recSTRING, ustr("array")),
		rec(recSTRING, ustr("orphan")),
	)
	want := []string{"42", "TRUE", "#DIV/0!", "", "shared", "array"}
	for i, w := range want {
		if got := col(s, 0, i); got != w {
			t.Errorf("col %d = %q, want %q", i, got, w)
		}
	}
	r := s.Row(0)
	if v := r.Value(0); v.Kind != KindNumber || v.Float != 42 || v.Int != 42 {
		t.Errorf("Value(0) = %+v", v)
	}
	if v := r.Value(1); v.Kind != KindBool || !v.Bool {
		t.Errorf("Value(1) = %+v", v)
	}
	if v := r.Value(2); v.Kind != KindError || v.Text != "#DIV/0!" {
		t.Errorf("Value(2) = %+v", v)
	}
	if v := r.Value(3); v.Kind != KindString || v.Text != "" {
		t.Errorf("Value(3) = %+v", v)
	}
}

func TestStringSplitAcrossContinue(t *testing.T) {
	// A STRING record for a formula, continued with 16-bit characters.
	_, s := mustSheet(t, [][]byte{xfRec(0)},
		formulaRec(0, 0, 0, specialVal(0, 0)),
		rec(recSTRING, le(uint16(4), uint8(0)), []byte("ab")),
		rec(recCONTINUE, []byte{1}, le(uint16('c'), uint16(0x00e9))),
	)
	if got := col(s, 0, 0); got != "abcé" {
		t.Errorf("got %q", got)
	}
}

func TestValueGeneralFormat(t *testing.T) {
	_, s := mustSheet(t, [][]byte{xfRec(0)}, numberRec(0, 0, 0, 1.5), rkRec(0, 1, 0, rkInt(2)), numberRec(0, 2, 99, 3.25))
	r := s.Row(0)
	if v := r.Value(0); v.Kind != KindNumber || v.Float != 1.5 || v.Int != 0 || v.Format != "General" {
		t.Errorf("Value(0) = %+v", v)
	}
	if v := r.Value(1); v.Kind != KindNumber || v.Float != 2 || v.Int != 2 {
		t.Errorf("Value(1) = %+v", v)
	}
	if got := r.Col(2); got != "3.25" { // XF index out of range: General
		t.Errorf("Col(2) = %q", got)
	}
	if v := r.Value(7); v.Kind != KindEmpty {
		t.Errorf("Value(7) = %+v", v)
	}
}

func TestNumberFormats(t *testing.T) {
	f8 := `"$"#,##0.00_);[Red]\("$"#,##0.00\)`
	_, s := mustSheet(t,
		[][]byte{formatRec(8, f8), formatRec(165, "HH:MM:SS"), formatRec(166, "[h]:mm"), formatRec(167, `#,##0 "USD"`),
			xfRec(8), xfRec(14), xfRec(165), xfRec(166), xfRec(167), xfRec(22)},
		numberRec(0, 0, 0, 1234.5),
		rkRec(0, 1, 0, rkInt(1234)),
		numberRec(1, 0, 1, 43831),
		rkRec(1, 1, 1, rkInt(43831)),
		numberRec(2, 0, 2, 0.5630208333333334),
		numberRec(2, 1, 3, 1.5208333333333333),
		numberRec(2, 2, 4, 1234),
		numberRec(2, 3, 5, 43831.75),
		numberRec(3, 0, 1, 2958465),
		numberRec(3, 1, 1, 109575),
		numberRec(3, 2, 1, -1),
	)
	cases := []struct {
		row, col int
		want     string
	}{
		{0, 0, "1234.5"},
		{0, 1, "1234"},
		{1, 0, "2020-01-01"},
		{1, 1, "2020-01-01"},
		{2, 0, "13:30:45"},
		{2, 1, "36:30"},
		{2, 2, "1234"},
		{2, 3, "2020-01-01 18:00"},
		{3, 0, "9999-12-31"},
		{3, 1, "2200-01-01"},
		{3, 2, "-1"},
	}
	for _, c := range cases {
		if got := col(s, c.row, c.col); got != c.want {
			t.Errorf("(%d,%d) = %q, want %q", c.row, c.col, got, c.want)
		}
	}
	v := s.Row(1).Value(0)
	if v.Kind != KindDate || !v.Time.Equal(time.Date(2020, 1, 1, 0, 0, 0, 0, time.UTC)) || v.Format != "yyyy-mm-dd" {
		t.Errorf("date Value = %+v", v)
	}
}

func TestDateMode1904(t *testing.T) {
	wb, s := mustSheet(t, [][]byte{rec(recDATEMODE, le(uint16(1))), xfRec(14)}, numberRec(0, 0, 0, 0))
	if got := col(s, 0, 0); got != "1904-01-01" {
		t.Errorf("got %q", got)
	}
	if !wb.ToDateTime(1).Equal(time.Date(1904, 1, 2, 0, 0, 0, 0, time.UTC)) {
		t.Errorf("ToDateTime(1) = %v", wb.ToDateTime(1))
	}
}

func TestGlobalsStopAtEOF(t *testing.T) {
	// A DATEMODE record inside a sheet must not change the workbook.
	wb, _ := mustSheet(t, [][]byte{xfRec(14)}, rec(recDATEMODE, le(uint16(1))), numberRec(0, 0, 0, 1))
	if wb.date1904 {
		t.Error("sheet record changed the workbook date system")
	}
}

func TestSSTContinue(t *testing.T) {
	cases := []struct {
		name string
		recs [][]byte
		want []string
	}{
		{
			"rich runs split",
			[][]byte{
				rec(recSST, le(uint32(2), uint32(2)), le(uint16(2), uint8(0x08), uint16(2)), []byte("AB"), le(uint32(1))),
				rec(recCONTINUE, le(uint32(2)), ustr("CD")),
			},
			[]string{"AB", "CD"},
		},
		{
			"rich runs all in continue",
			[][]byte{
				rec(recSST, le(uint32(2), uint32(2)), le(uint16(2), uint8(0x08), uint16(2)), []byte("AB")),
				rec(recCONTINUE, le(uint32(1), uint32(2)), ustr("CD")),
			},
			[]string{"AB", "CD"},
		},
		{
			"phonetic split",
			[][]byte{
				rec(recSST, le(uint32(2), uint32(2)), le(uint16(1), uint8(0x04), uint32(6)), []byte("A"), []byte{1, 2}),
				rec(recCONTINUE, []byte{3, 4, 5, 6}, ustr("B")),
			},
			[]string{"A", "B"},
		},
		{
			"characters split, width changes",
			[][]byte{
				rec(recSST, le(uint32(2), uint32(2)), le(uint16(4), uint8(0)), []byte("ab")),
				rec(recCONTINUE, []byte{1}, le(uint16(0x4E2D), uint16(0x6587)), ustr("x")),
			},
			[]string{"ab中文", "x"},
		},
		{
			"header at end of record",
			[][]byte{
				rec(recSST, le(uint32(2), uint32(2)), ustr("first"), le(uint16(3), uint8(0))),
				rec(recCONTINUE, []byte{0}, []byte("xyz")),
			},
			[]string{"first", "xyz"},
		},
		{
			"string starts in continue",
			[][]byte{
				rec(recSST, le(uint32(2), uint32(2)), ustr("first")),
				rec(recCONTINUE, ustr("second")),
			},
			[]string{"first", "second"},
		},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			globals := append([][]byte{xfRec(0)}, c.recs...)
			wb, err := openBook(t, book{globals: globals, sheets: []sheet{{name: "S"}}})
			if err != nil {
				t.Fatal(err)
			}
			if len(wb.sst) != len(c.want) {
				t.Fatalf("sst = %q, want %q", wb.sst, c.want)
			}
			for i := range c.want {
				if wb.sst[i] != c.want[i] {
					t.Errorf("sst = %q, want %q", wb.sst, c.want)
				}
			}
		})
	}
}

func TestSSTHugeCountSmallAllocation(t *testing.T) {
	f := book{globals: [][]byte{rec(recSST, le(uint32(0xFFFFFFFF), uint32(0xFFFFFFFF)), ustr("x"))}, sheets: []sheet{{name: "S"}}}.file()
	var before, after runtime.MemStats
	runtime.ReadMemStats(&before)
	wb, err := OpenReader(bytes.NewReader(f), "")
	runtime.ReadMemStats(&after)
	if err != nil {
		t.Fatal(err)
	}
	if len(wb.sst) != 1 {
		t.Errorf("sst = %q", wb.sst)
	}
	if n := after.TotalAlloc - before.TotalAlloc; n > 1<<20 {
		t.Errorf("allocated %d bytes for a %d byte file", n, len(f))
	}
}

func TestBIFF5Codepage(t *testing.T) {
	label5 := func(row, c uint16, b []byte) []byte {
		return rec(recLABEL, le(row, c, uint16(0), uint16(len(b))), b)
	}
	name := []byte("Caf\xe9")
	cases := []struct {
		codepage uint16
		charset  string
		want     string
	}{
		{0, "", "Café"},
		{1252, "", "Café"},
		{1251, "", "Cafй"},
		{1252, "windows-1251", "Cafй"},
		{1251, "utf-8", "Cafй"},
	}
	for _, c := range cases {
		var globals [][]byte
		if c.codepage != 0 {
			globals = append(globals, rec(recCODEPAGE, le(c.codepage)))
		}
		globals = append(globals, xfRec(0), rec(recWRITEACCESS, []byte{byte(len(name))}, name, bytes.Repeat([]byte(" "), 26)))
		b := book{biff: 5, globals: globals, sheets: []sheet{{name: "Ünï", recs: [][]byte{label5(0, 0, name)}}}}
		wb, err := OpenReader(bytes.NewReader(b.file()), c.charset)
		if err != nil {
			t.Fatalf("%+v: %v", c, err)
		}
		s, err := wb.GetSheet(0)
		if err != nil {
			t.Fatal(err)
		}
		if got := col(s, 0, 0); got != c.want || wb.Author != c.want {
			t.Errorf("%+v: cell %q author %q, want %q", c, got, wb.Author, c.want)
		}
	}
	if _, err := OpenReader(bytes.NewReader(book{biff: 5}.file()), "no-such-charset"); err == nil {
		t.Error("unknown charset accepted")
	}
}

func TestPreferWorkbookStream(t *testing.T) {
	biff8 := book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "Eight", recs: [][]byte{labelRec(0, 0, "8")}}}}.stream()
	biff5 := book{biff: 5, globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "Five"}}}.stream()
	other := book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "Embedded"}}}.stream()
	f, _ := cfbtest.Build([]cfbtest.Stream{
		{Name: "Book", Data: biff5},
		{Name: "Workbook", Data: other, Storage: "MBD0001"},
		{Name: "Workbook", Data: biff8},
	}, cfbtest.Options{})
	wb, err := OpenReader(bytes.NewReader(f), "")
	if err != nil {
		t.Fatal(err)
	}
	s, err := wb.GetSheet(0)
	if err != nil || s.Name != "Eight" {
		t.Errorf("opened sheet %q, %v", s.Name, err)
	}
}

func TestReadAllLimits(t *testing.T) {
	wb, _ := mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "a"), labelRec(1, 0, "b"), labelRec(2, 1, "c"))
	for _, n := range []int{-1, 0, 100} {
		data, _, err := wb.ReadAll(n)
		if err != nil || len(data[0]) != 3 {
			t.Errorf("ReadAll(%d) = %v, %v", n, data, err)
		}
	}
	data, _, _ := wb.ReadAll(2)
	if len(data[0]) != 2 {
		t.Errorf("ReadAll(2) = %v", data)
	}
	// A sheet whose only row is row 0.
	wb, _ = mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "only"))
	data, _, _ = wb.ReadAll(0)
	if len(data[0]) != 1 || data[0][0][0] != "only" {
		t.Errorf("ReadAll = %v", data)
	}
}

func TestIndexRange(t *testing.T) {
	_, s := mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "row0col0"))
	if s.Row(65536) != nil || s.Row(-1) != nil {
		t.Error("Row out of range returned a row")
	}
	r := s.Row(0)
	if r.Col(65536) != "" || r.Col(-1) != "" || r.ColExact(1<<20) != "" || r.Value(65536).Kind != KindEmpty {
		t.Error("Col out of range returned a value")
	}
}

func TestStreamSizeEnforced(t *testing.T) {
	stream := book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "S", recs: [][]byte{labelRec(0, 0, "a")}}}}.stream()
	stream = append(stream, make([]byte, 5000-len(stream))...) // keep it out of the mini stream
	f, l := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream}}, cfbtest.Options{})
	// Hide a DATEMODE record in the unused bytes after the stream's end.
	off := l.SectorOffset(l.StreamStart["Workbook"]) + len(stream)
	copy(f[off:], rec(recDATEMODE, le(uint16(1))))
	wb, err := OpenReader(bytes.NewReader(f), "")
	if err != nil {
		t.Fatal(err)
	}
	if wb.date1904 {
		t.Error("record after the end of the stream was parsed")
	}
}

func TestCorruptContainer(t *testing.T) {
	stream := book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "S"}}}.stream()
	stream = append(stream, make([]byte, 5000-len(stream))...)
	for name, sid := range map[string]uint32{"cycle": 0, "beyond table": 130, "beyond file": 1 << 20} {
		f, l := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream}}, cfbtest.Options{})
		start := l.StreamStart["Workbook"]
		if name == "cycle" {
			sid = start
		}
		l.SetFAT(f, start, sid)
		if name == "beyond table" {
			f = append(f, make([]byte, 140*512)...)
		}
		done := make(chan error, 1)
		go func() { _, err := OpenReader(bytes.NewReader(f), ""); done <- err }()
		select {
		case err := <-done:
			var ce ole2.CorruptError
			if !errors.As(err, &ce) {
				t.Errorf("%s: err = %v", name, err)
			}
		case <-time.After(10 * time.Second):
			t.Fatalf("%s: OpenReader did not return", name)
		}
	}
}

func TestEncrypted(t *testing.T) {
	_, err := openBook(t, book{globals: [][]byte{rec(recFILEPASS, make([]byte, 54))}, sheets: []sheet{{name: "S"}}})
	if !errors.Is(err, ErrEncrypted) {
		t.Errorf("err = %v", err)
	}
}

func TestStructuralErrors(t *testing.T) {
	cases := map[string]book{
		"LABELSST index":   {globals: [][]byte{sstRec("a")}, sheets: []sheet{{name: "S", recs: [][]byte{labelSSTRec(0, 0, 0, 5)}}}},
		"MULRK range":      {sheets: []sheet{{name: "S", recs: [][]byte{rec(recMULRK, le(uint16(0), uint16(0), uint16(0), rkInt(1), uint16(5)))}}}},
		"MULRK overflow":   {sheets: []sheet{{name: "S", recs: [][]byte{rec(recMULRK, le(uint16(0), uint16(65535), uint16(0), rkInt(1), uint16(0), rkInt(2), uint16(0)))}}}},
		"MULRK short":      {sheets: []sheet{{name: "S", recs: [][]byte{rec(recMULRK, le(uint16(0)))}}}},
		"MULBLANK range":   {sheets: []sheet{{name: "S", recs: [][]byte{rec(recMULBLANK, le(uint16(0), uint16(4), uint16(0), uint16(0), uint16(1)))}}}},
		"short NUMBER":     {sheets: []sheet{{name: "S", recs: [][]byte{rec(recNUMBER, le(uint16(0), uint16(0)))}}}},
		"short STRING":     {sheets: []sheet{{name: "S", recs: [][]byte{formulaRec(0, 0, 0, specialVal(0, 0)), rec(recSTRING, le(uint16(10), uint8(0)), []byte("ab"))}}}},
		"string past end":  {sheets: []sheet{{name: "S", recs: [][]byte{rec(recLABEL, le(uint16(0), uint16(0), uint16(0), uint16(500), uint8(1)))}}}},
		"duplicate offset": {sheets: []sheet{{name: "A"}, {name: "B"}}},
	}
	for name, b := range cases {
		stream := b.stream()
		if name == "duplicate offset" {
			// Point the second BOUNDSHEET at the first sheet.
			i := bytes.Index(stream, []byte{0x85, 0x00})
			j := bytes.Index(stream[i+1:], []byte{0x85, 0x00}) + i + 1
			copy(stream[j+4:j+8], stream[i+4:i+8])
		}
		f, _ := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream}}, cfbtest.Options{})
		wb, err := OpenReader(bytes.NewReader(f), "")
		if err == nil {
			_, err = wb.GetSheet(0)
		}
		if !errors.Is(err, ErrCorrupt) {
			t.Errorf("%s: err = %v, want a corrupt-workbook error", name, err)
		}
	}
}

func TestMissingSheetEOF(t *testing.T) {
	stream := book{sheets: []sheet{{name: "S", recs: [][]byte{labelRec(0, 0, "a")}}}}.stream()
	stream = stream[:len(stream)-4] // drop the sheet's EOF record
	f, _ := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream}}, cfbtest.Options{})
	wb, err := OpenReader(bytes.NewReader(f), "")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
		t.Errorf("err = %v", err)
	}
}

func TestMergedCells(t *testing.T) {
	merge := rec(recMERGECELLS, le(uint16(2), uint16(0), uint16(1), uint16(0), uint16(2), uint16(5), uint16(5), uint16(3), uint16(3)))
	_, s := mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "top"), labelRec(1, 0, "own"), rec(recBLANK, le(uint16(1), uint16(2), uint16(0))), merge)
	if got := s.Row(0).Col(2); got != "top" {
		t.Errorf("Col in merged range = %q", got)
	}
	if got := s.Row(0).ColExact(2); got != "" {
		t.Errorf("ColExact in merged range = %q", got)
	}
	// As in Excel, a covered cell shows the top-left value; ColExact
	// still returns the value stored in the cell.
	if got, own := s.Row(1).Col(0), s.Row(1).ColExact(0); got != "top" || own != "own" {
		t.Errorf("covered cell with its own value: Col %q, ColExact %q", got, own)
	}
	if got := s.Row(1).Col(1); got != "top" {
		t.Errorf("Col(1,1) = %q", got)
	}
	if n := len(s.MergedCells()); n != 2 {
		t.Errorf("%d merged ranges", n)
	}
}

func TestNestedSubstream(t *testing.T) {
	_, s := mustSheet(t, [][]byte{xfRec(0)},
		labelRec(0, 0, "before"),
		bofRec(0x0600, 0x0020), labelRec(0, 1, "chart"), eofRec(),
		labelRec(0, 2, "after"),
	)
	if col(s, 0, 0) != "before" || col(s, 0, 1) != "" || col(s, 0, 2) != "after" {
		t.Errorf("row = %q %q %q", col(s, 0, 0), col(s, 0, 1), col(s, 0, 2))
	}
}

func TestCellRecords(t *testing.T) {
	_, s := mustSheet(t, [][]byte{xfRec(0), sstRec("shared")},
		rec(recROW, le(uint16(0), uint16(0), uint16(10), uint16(0), uint16(0), uint16(0), uint32(0))),
		rec(recMULRK, le(uint16(0), uint16(1), uint16(0), rkInt(10), uint16(0), uint32(0x3FF00000)|1, uint16(2))),
		rec(recMULBLANK, le(uint16(0), uint16(3), uint16(0), uint16(0), uint16(4))),
		rec(recBOOLERR, le(uint16(0), uint16(5), uint16(0), uint8(0), uint8(0))),
		rec(recBOOLERR, le(uint16(0), uint16(6), uint16(0), uint8(0x2A), uint8(1))),
		labelSSTRec(0, 7, 0, 0),
		rec(recRSTRING, le(uint16(0), uint16(8), uint16(0)), ustr("rich"), le(uint16(0))),
		rec(recWINDOW2, le(uint16(0x0200|0x0040))),
	)
	want := []string{"", "10", "0.01", "", "", "FALSE", "#N/A", "shared", "rich"}
	for i, w := range want {
		if got := col(s, 0, i); got != w {
			t.Errorf("col %d = %q, want %q", i, got, w)
		}
	}
	r := s.Row(0)
	if r.FirstCol() != 0 || r.LastCol() != 10 {
		t.Errorf("FirstCol, LastCol = %d, %d", r.FirstCol(), r.LastCol())
	}
	if !s.Selected || !s.RightToLeft {
		t.Error("WINDOW2 flags not read")
	}
}

func TestAuthor(t *testing.T) {
	wa := rec(recWRITEACCESS, ustr("Jane Doe"), bytes.Repeat([]byte(" "), 100))
	wb, err := openBook(t, book{globals: [][]byte{wa}, sheets: []sheet{{name: "S"}}})
	if err != nil || wb.Author != "Jane Doe" {
		t.Errorf("Author = %q, %v", wb.Author, err)
	}
}

func TestChartSheet(t *testing.T) {
	wb, err := openBook(t, book{sheets: []sheet{{name: "Chart", dt: 0x0020, recs: [][]byte{labelRec(0, 0, "x")}}, {name: "Data", recs: [][]byte{labelRec(0, 0, "y")}}}})
	if err != nil {
		t.Fatal(err)
	}
	c, err := wb.GetSheet(0)
	if err != nil || c.Row(0) != nil {
		t.Errorf("chart sheet: %v", err)
	}
	d, err := wb.GetSheet(1)
	if err != nil || col(d, 0, 0) != "y" {
		t.Errorf("data sheet: %v", err)
	}
}

func TestUnsupportedVersion(t *testing.T) {
	stream := cat(rec(recBOF4, le(uint16(0), uint16(0x10))), eofRec())
	f, _ := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream}}, cfbtest.Options{})
	if _, err := OpenReader(bytes.NewReader(f), ""); !errors.Is(err, ErrUnsupportedVersion) {
		t.Errorf("err = %v", err)
	}
	f, _ = cfbtest.Build([]cfbtest.Stream{{Name: "Other", Data: stream}}, cfbtest.Options{})
	if _, err := OpenReader(bytes.NewReader(f), ""); !errors.Is(err, ErrNoWorkbook) {
		t.Errorf("err = %v", err)
	}
	if _, err := OpenReader(bytes.NewReader([]byte("PK\x03\x04 not an xls file at all, padding padding")), ""); !errors.Is(err, ole2.ErrNotCompoundFile) {
		t.Errorf("err = %v", err)
	}
}

func TestConcurrentGetSheet(t *testing.T) {
	var sheets []sheet
	for i := 0; i < 4; i++ {
		sheets = append(sheets, sheet{name: string(rune('A' + i)), recs: [][]byte{labelRec(0, 0, "v"), numberRec(1, 0, 0, float64(i))}})
	}
	wb, err := openBook(t, book{globals: [][]byte{xfRec(0)}, sheets: sheets})
	if err != nil {
		t.Fatal(err)
	}
	var wg sync.WaitGroup
	for g := 0; g < 8; g++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for i := 0; i < 4; i++ {
				s, err := wb.GetSheet(i)
				if err != nil || col(s, 0, 0) != "v" {
					t.Errorf("sheet %d: %v", i, err)
				}
			}
			wb.ReadAll(0)
		}()
	}
	wg.Wait()
}

func TestDirectoryEntryHighSizeIgnored(t *testing.T) {
	// Version 3 files must ignore the high 32 bits of a stream size.
	stream := book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "S", recs: [][]byte{labelRec(0, 0, "a")}}}}.stream()
	f, l := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream}}, cfbtest.Options{})
	binary.LittleEndian.PutUint32(f[l.EntryOffset["Workbook"]+124:], 0xDEADBEEF)
	wb, err := OpenReader(bytes.NewReader(f), "")
	if err != nil {
		t.Fatal(err)
	}
	if s, err := wb.GetSheet(0); err != nil || col(s, 0, 0) != "a" {
		t.Errorf("%v", err)
	}
}

func TestColumnLimit(t *testing.T) {
	for name, r := range map[string][]byte{
		"LABEL":    labelRec(0, 256, "x"),
		"BLANK":    rec(recBLANK, le(uint16(0), uint16(300), uint16(0))),
		"MULRK":    rec(recMULRK, le(uint16(0), uint16(255), uint16(0), rkInt(1), uint16(0), rkInt(2), uint16(256))),
		"MULBLANK": rec(recMULBLANK, le(uint16(0), uint16(250), uint16(0), uint16(0), uint16(0), uint16(0), uint16(0), uint16(0), uint16(0), uint16(256))),
	} {
		wb, err := openBook(t, book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "S", recs: [][]byte{r}}}})
		if err != nil {
			t.Fatal(err)
		}
		if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
			t.Errorf("%s: err = %v", name, err)
		}
	}
	_, s := mustSheet(t, [][]byte{xfRec(0)}, rec(recROW, le(uint16(0), uint16(0xFFFF), uint16(0xFFFF), uint16(0), uint16(0), uint16(0), uint32(0))))
	if r := s.Row(0); r.FirstCol() > 255 || r.LastCol() > 256 {
		t.Errorf("ROW record not clamped: %d, %d", r.FirstCol(), r.LastCol())
	}
}

func TestManyRangesStayFast(t *testing.T) {
	// Overlapping full-height merges and hyperlinks are invalid, but must not
	// make each lookup scan all of them.
	var merges, links [][]byte
	const n = 20000
	for i := 0; i < n; i += 1000 {
		body := le(uint16(1000))
		for k := 0; k < 1000; k++ {
			body = append(body, le(uint16(0), uint16(65535), uint16(k%256), uint16(255))...)
		}
		merges = append(merges, rec(recMERGECELLS, body))
	}
	for i := 0; i < 2000; i++ {
		links = append(links, rec(recHLINK, le(uint16(0), uint16(65535), uint16(0), uint16(255)), make([]byte, 16), le(uint32(2), uint32(0))))
	}
	recs := append([][]byte{labelRec(0, 0, "top")}, merges...)
	recs = append(recs, links...)
	for r := 1; r < 2000; r++ {
		recs = append(recs, rec(recBLANK, le(uint16(r), uint16(255), uint16(0))))
	}
	_, s := mustSheet(t, [][]byte{xfRec(0)}, recs...)
	start := time.Now()
	for r := 0; r < 2000; r++ {
		row := s.Row(r)
		for c := row.FirstCol(); c < row.LastCol(); c++ {
			row.Col(c)
			row.Hyperlink(c)
		}
	}
	if d := time.Since(start); d > slow*2*time.Second {
		t.Errorf("512000 lookups took %v", d)
	}
}

func TestRangeIndex(t *testing.T) {
	ranges := []CellRange{{0, 0, 0, 3}, {2, 5, 1, 1}, {2, 3, 2, 4}, {65535, 65535, 8, 255}, {10, 65535, 7, 7}}
	refs := make([]ref, len(ranges))
	for i, r := range ranges {
		refs[i] = toRef(r)
	}
	x := newRangeIndex(refs)
	for r := 0; r < maxRows; r += 7 {
		if r+7 >= maxRows {
			r = maxRows - 1
		}
		for c := 0; c < 256; c += 3 {
			want := -1
			for i, g := range ranges {
				if r >= g.FirstRow && r <= g.LastRow && c >= g.FirstCol && c <= g.LastCol {
					want = i
				}
			}
			if got := x.find(r, c); got != want {
				t.Fatalf("find(%d, %d) = %d, want %d", r, c, got, want)
			}
		}
	}
	if x.find(65535, 9) != 3 || x.find(-1, 0) != -1 || x.find(70000, 0) != -1 {
		t.Error("edge rows")
	}
}

func TestReadAllBudget(t *testing.T) {
	// Each sheet is a few dozen bytes but asks for a 65536-row grid.
	var sheets []sheet
	for i := 0; i < 1000; i++ {
		sheets = append(sheets, sheet{name: string(rune(0x4E00 + i)), recs: [][]byte{rec(recBLANK, le(uint16(65535), uint16(255), uint16(0)))}})
	}
	wb, err := openBook(t, book{globals: [][]byte{xfRec(0)}, sheets: sheets})
	if err != nil {
		t.Fatal(err)
	}
	var before, after runtime.MemStats
	runtime.ReadMemStats(&before)
	_, _, err = wb.ReadAll(0)
	runtime.ReadMemStats(&after)
	if !errors.Is(err, ErrTooLarge) {
		t.Errorf("err = %v, want ErrTooLarge", err)
	}
	if n := after.TotalAlloc - before.TotalAlloc; n > 200<<20 {
		t.Errorf("allocated %d MiB", n>>20)
	}
	if _, _, err := wb.ReadAll(100); err != nil {
		t.Errorf("ReadAll(100): %v", err)
	}
}

func TestFormulaPastLastColumn(t *testing.T) {
	// Found by fuzzing: a string formula in column 256 crashed the parser.
	for _, recs := range [][][]byte{
		{formulaRec(0, 300, 0, specialVal(0, 0)), rec(recSTRING, ustr("x"))},
		{labelRec(0, 0, "a"), formulaRec(0, 300, 0, specialVal(0, 0)), rec(recSTRING, ustr("x"))},
	} {
		wb, err := openBook(t, book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "S", recs: recs}}})
		if err != nil {
			t.Fatal(err)
		}
		if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
			t.Errorf("err = %v, want a corrupt-workbook error", err)
		}
	}
}

// overlapBook returns a workbook whose first sheet's BOF record carries a
// second sheet's BOF and cells inside its own data, or inside an unknown
// record after it.
func overlapBook(inUnknown bool) []byte {
	inner := cat(bofRec(0x0600, 0x0010), labelRec(0, 0, "x"), eofRec())
	var first []byte
	if inUnknown {
		first = cat(bofRec(0x0600, 0x0010), rec(0x1234, inner), eofRec())
	} else {
		first = cat(rec(recBOF, le(uint16(0x0600), uint16(0x0010)), make([]byte, 12), inner), eofRec())
	}
	var g bytes.Buffer
	g.Write(bofRec(0x0600, 0x0005))
	g.Write(xfRec(0))
	p1 := g.Len() + 4
	g.Write(rec(recBOUNDSHEET, le(uint32(0), uint8(0), uint8(0)), []byte{1, 0, 'A'}))
	p2 := g.Len() + 4
	g.Write(rec(recBOUNDSHEET, le(uint32(0), uint8(0), uint8(0)), []byte{1, 0, 'B'}))
	g.Write(eofRec())
	out := g.Bytes()
	start := len(out)
	out = append(out, first...)
	binary.LittleEndian.PutUint32(out[p1:], uint32(start))
	innerAt := start + bytes.Index(first[4:], inner) + 4
	binary.LittleEndian.PutUint32(out[p2:], uint32(innerAt))
	f, _ := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: out}}, cfbtest.Options{})
	return f
}

func TestOverlappingSheets(t *testing.T) {
	// The sheet that runs into the next one is an error; the next sheet
	// is read once, on its own.
	for _, inUnknown := range []bool{false, true} {
		wb, err := OpenReader(bytes.NewReader(overlapBook(inUnknown)), "")
		if err != nil {
			t.Fatalf("inUnknown=%v: %v", inUnknown, err)
		}
		if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
			t.Errorf("inUnknown=%v: sheet A err = %v, want ErrCorrupt", inUnknown, err)
		}
		if s, err := wb.GetSheet(1); err != nil || col(s, 0, 0) != "x" {
			t.Errorf("inUnknown=%v: sheet B: %v", inUnknown, err)
		}
	}
	// A sheet without an EOF record does not hide the sheet after it.
	stream := book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "A", recs: [][]byte{labelRec(0, 0, "a")}}, {name: "B", recs: [][]byte{labelRec(0, 0, "b")}}}}.stream()
	i := bytes.Index(stream, cat(labelRec(0, 0, "a"), eofRec()))
	stream = append(stream[:i+len(labelRec(0, 0, "a"))], stream[i+len(labelRec(0, 0, "a"))+4:]...)
	j := bytes.Index(stream, []byte{0x85, 0x00})
	k := bytes.Index(stream[j+1:], []byte{0x85, 0x00}) + j + 1
	binary.LittleEndian.PutUint32(stream[k+4:], binary.LittleEndian.Uint32(stream[k+4:])-4)
	f, _ := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream}}, cfbtest.Options{})
	wb, err := OpenReader(bytes.NewReader(f), "")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
		t.Errorf("sheet without EOF: err = %v", err)
	}
	if s, err := wb.GetSheet(1); err != nil || col(s, 0, 0) != "b" {
		t.Errorf("sheet after it: %v", err)
	}
	// A sheet that starts inside the workbook globals.
	stream = book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "S"}}}.stream()
	i = bytes.Index(stream, []byte{0x85, 0x00})
	binary.LittleEndian.PutUint32(stream[i+4:], 0)
	f, _ = cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream}}, cfbtest.Options{})
	wb, err = OpenReader(bytes.NewReader(f), "")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
		t.Errorf("sheet in globals: err = %v", err)
	}
}

func TestLongFormatIgnored(t *testing.T) {
	long := "yyyy" + strings.Repeat("e", 300)
	_, s := mustSheet(t, [][]byte{formatRec(164, long), formatRec(165, "yyyy"+strings.Repeat(" ", 251)), xfRec(164), xfRec(165)},
		numberRec(0, 0, 0, 43831), numberRec(0, 1, 1, 43831))
	if got := col(s, 0, 0); got != "43831" {
		t.Errorf("over-long format: %q", got)
	}
	if got := col(s, 0, 1); got != "2020"+strings.Repeat(" ", 251) {
		t.Errorf("255-character format: %q", got)
	}
}

func TestReadAllTextBudget(t *testing.T) {
	defer func(f func(int) int64) { readAllBudget = f }(readAllBudget)
	readAllBudget = func(n int) int64 { return 100 << 10 }
	code := strings.Repeat("ey", 127) // renders about 760 bytes
	var recs [][]byte
	for r := 0; r < 200; r++ {
		recs = append(recs, numberRec(uint16(r), 0, 0, 43831))
	}
	wb, _ := mustSheet(t, [][]byte{formatRec(164, code), xfRec(164)}, recs...)
	if _, _, err := wb.ReadAll(0); !errors.Is(err, ErrTooLarge) {
		t.Errorf("err = %v, want ErrTooLarge", err)
	}
}

func TestFormulaUnknownResult(t *testing.T) {
	wb, err := openBook(t, book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "S", recs: [][]byte{formulaRec(0, 0, 0, specialVal(9, 0))}}}})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
		t.Errorf("err = %v", err)
	}
}

func TestBareBIFF(t *testing.T) {
	for _, id := range []uint16{recBOF2, recBOF3, recBOF4, recBOF} {
		b := cat(rec(id, le(uint16(0), uint16(0x10))), eofRec(), make([]byte, 600))
		if _, err := OpenReader(bytes.NewReader(b), ""); !errors.Is(err, ErrUnsupportedVersion) {
			t.Errorf("BOF %#x: err = %v", id, err)
		}
	}
}

func TestMergedRanges(t *testing.T) {
	// A1:C4 with a value in A1. Rows 1 and 2 have blank cells, as Excel
	// writes them; row 3 has no records at all.
	merge := rec(recMERGECELLS, le(uint16(1), uint16(0), uint16(3), uint16(0), uint16(2)))
	blanks := func(r uint16) []byte {
		return rec(recMULBLANK, le(r, uint16(0), uint16(0), uint16(0), uint16(0), uint16(2)))
	}
	_, s := mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "top"), rec(recMULBLANK, le(uint16(0), uint16(1), uint16(0), uint16(0), uint16(2))), blanks(1), blanks(2), merge)
	if s.MaxRow != 2 || s.Row(3) != nil {
		t.Errorf("MaxRow = %d, Row(3) = %v", s.MaxRow, s.Row(3))
	}
	rows := s.Rows()
	if len(rows) != 3 {
		t.Fatalf("%d rows", len(rows))
	}
	for i, row := range rows {
		if row.Index() != i || row.FirstCol() != 0 || row.LastCol() != 3 {
			t.Errorf("row %d: index %d, columns %d-%d", i, row.Index(), row.FirstCol(), row.LastCol())
		}
		for c := row.FirstCol(); c < row.LastCol(); c++ {
			if got := row.Col(c); got != "top" {
				t.Errorf("Col(%d,%d) = %q", i, c, got)
			}
			if v := row.Value(c); v.Kind != KindString || v.Text != "top" {
				t.Errorf("Value(%d,%d) = %+v", i, c, v)
			}
			if got := row.ColExact(c); (got != "") != (i == 0 && c == 0) {
				t.Errorf("ColExact(%d,%d) = %q", i, c, got)
			}
		}
	}
}

func TestOneMergeDoesNotGrowTheGrid(t *testing.T) {
	// One 8-byte range covering the whole sheet must not make the row loop
	// visit 16 million cells.
	merge := rec(recMERGECELLS, le(uint16(1), uint16(0), uint16(65535), uint16(0), uint16(255)))
	_, s := mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "top"), merge)
	if s.MaxRow != 0 || len(s.Rows()) != 1 || s.Row(0).LastCol() != 1 {
		t.Errorf("MaxRow %d, %d rows, LastCol %d", s.MaxRow, len(s.Rows()), s.Row(0).LastCol())
	}
}

func TestNumberText(t *testing.T) {
	cases := map[float64]string{
		0: "0", 123.5: "123.5", -1184955: "-1184955", 0.1: "0.1", 1e-6: "0.000001",
		1e-7: "1e-07", 1e20: "100000000000000000000", 1e21: "1e+21", 1e300: "1e+300", -2.5e-300: "-2.5e-300",
	}
	for f, want := range cases {
		if got := formatNumber(f); got != want {
			t.Errorf("formatNumber(%v) = %q, want %q", f, got, want)
		}
	}
}

func TestZeroSection(t *testing.T) {
	_, s := mustSheet(t, [][]byte{formatRec(164, "yyyy-mm-dd;;"), xfRec(164)}, numberRec(0, 0, 0, 0), numberRec(0, 1, 0, 43831))
	if col(s, 0, 0) != "" || col(s, 0, 1) != "2020-01-01" {
		t.Errorf("got %q %q", col(s, 0, 0), col(s, 0, 1))
	}
}

// splitRecord writes data as a record followed by CONTINUE records of at
// most n bytes each.
func splitRecord(id uint16, data []byte, n int) []byte {
	out := rec(id, data[:min(n, len(data))])
	for i := n; i < len(data); i += n {
		out = append(out, rec(recCONTINUE, data[i:min(i+n, len(data))])...)
	}
	return out
}

func TestManyContinuePiecesStayLinear(t *testing.T) {
	timed := func(name string, b book) {
		t.Helper()
		f := b.file()
		start := time.Now()
		wb, err := OpenReader(bytes.NewReader(f), "")
		if err == nil {
			_, err = wb.GetSheet(0)
		}
		if d := time.Since(start); d > slow*3*time.Second {
			t.Errorf("%s: %d byte file took %v (%v)", name, len(f), d, err)
		}
	}
	// SST with tiny strings followed by 100,000 empty CONTINUE records.
	var sst []byte
	for i := 0; i < 20000; i++ {
		sst = append(sst, ustr("x")...)
	}
	sstRecs := cat(rec(recSST, le(uint32(20000), uint32(20000)), sst[:8000]))
	sstRecs = append(sstRecs, splitRecord(recCONTINUE, sst[8000:], 8000)[4:]...)
	for i := 0; i < 100000; i++ {
		sstRecs = append(sstRecs, rec(recCONTINUE)...)
	}
	timed("SST", book{globals: [][]byte{xfRec(0), sstRecs}, sheets: []sheet{{name: "S"}}})

	// MERGECELLS and MULRK split into one-byte pieces.
	body := le(uint16(20000))
	for i := 0; i < 20000; i++ {
		body = append(body, le(uint16(i), uint16(i), uint16(0), uint16(1))...)
	}
	timed("MERGECELLS", book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "S", recs: [][]byte{splitRecord(recMERGECELLS, body, 1)}}}})
	mulrk := le(uint16(0), uint16(0))
	for i := 0; i < 200; i++ {
		mulrk = append(mulrk, le(uint16(0), rkInt(int32(i)))...)
	}
	mulrk = append(mulrk, le(uint16(199))...)
	var recs [][]byte
	for i := 0; i < 400; i++ {
		binary.LittleEndian.PutUint16(mulrk, uint16(i))
		recs = append(recs, splitRecord(recMULRK, mulrk, 1))
	}
	timed("MULRK", book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "S", recs: recs}}})
}

func TestBadStringsAllocateLittle(t *testing.T) {
	// Each WRITEACCESS claims 65535 characters but holds none.
	var globals [][]byte
	for i := 0; i < 20000; i++ {
		globals = append(globals, rec(recWRITEACCESS, le(uint16(0xFFFF), uint8(1))))
	}
	f := book{globals: globals, sheets: []sheet{{name: "S"}}}.file()
	var before, after runtime.MemStats
	runtime.ReadMemStats(&before)
	if _, err := OpenReader(bytes.NewReader(f), ""); err != nil {
		t.Fatal(err)
	}
	runtime.ReadMemStats(&after)
	if n := after.TotalAlloc - before.TotalAlloc; n > 32<<20 {
		t.Errorf("allocated %d MiB for a %d KiB file", n>>20, len(f)>>10)
	}
}

func TestSheetTypes(t *testing.T) {
	wb, err := openBook(t, book{globals: [][]byte{xfRec(0)}, sheets: []sheet{
		{name: "Macro", dt: 0x0040, recs: [][]byte{labelRec(0, 0, "=ALERT(1)"), formulaRec(1, 0, 0, numVal(1))}},
		{name: "Chart", dt: 0x0020},
	}})
	if err != nil {
		t.Fatal(err)
	}
	m, err := wb.GetSheet(0)
	if err != nil || m.Type != SheetMacro || col(m, 0, 0) != "=ALERT(1)" || col(m, 1, 0) != "1" {
		t.Errorf("macro sheet: %v, type %#x", err, m.Type)
	}
	c, err := wb.GetSheet(1)
	if err != nil || c.Type != SheetChart || len(c.Rows()) != 0 {
		t.Errorf("chart sheet: %v", err)
	}
	// A worksheet whose BOF record is cut short.
	stream := book{sheets: []sheet{{name: "S"}}}.stream()
	i := bytes.LastIndex(stream, bofRec(0x0600, 0x0010))
	stream = cat(stream[:i], rec(recBOF, le(uint16(0x0600))), eofRec())
	f, _ := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream}}, cfbtest.Options{})
	wb, err = OpenReader(bytes.NewReader(f), "")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
		t.Errorf("short BOF: %v", err)
	}
}

func TestOverlappingRangeLookup(t *testing.T) {
	link := func(c0, c1 uint16, url string) []byte {
		h := urlHyperlink(0, 0, url, "d")
		binary.LittleEndian.PutUint16(h[8:], c0)
		binary.LittleEndian.PutUint16(h[10:], c1)
		return h
	}
	_, s := mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "a"),
		link(0, 10, "https://wide/"), link(2, 3, "https://narrow/"),
		rec(recMERGECELLS, le(uint16(2), uint16(5), uint16(6), uint16(0), uint16(10), uint16(5), uint16(6), uint16(2), uint16(3))),
		labelRec(5, 0, "m"), rec(recBLANK, le(uint16(5), uint16(9), uint16(0))))
	for _, c := range []int{0, 1, 2, 3, 4, 5, 10} {
		if _, ok := s.Row(0).Hyperlink(c); !ok {
			t.Errorf("Hyperlink(%d) not found", c)
		}
	}
	if _, ok := s.Row(0).Hyperlink(11); ok {
		t.Error("Hyperlink(11) found")
	}
	if got := s.Row(5).Col(9); got != "m" {
		t.Errorf("overlapping merge: Col(9) = %q", got)
	}
}

func TestContainerErrorsWrapErrCorrupt(t *testing.T) {
	f := book{sheets: []sheet{{name: "S"}}}.file()
	f = f[:len(f)-600]
	_, err := OpenReader(bytes.NewReader(f), "")
	var ce ole2.CorruptError
	if !errors.Is(err, ErrCorrupt) || !errors.As(err, &ce) {
		t.Errorf("err = %v", err)
	}
}

func TestFileHyperlinkParents(t *testing.T) {
	fileLink := func(anti uint16) []byte {
		path := []byte("x.xls\x00")
		return rec(recHLINK, le(uint16(0), uint16(0), uint16(0), uint16(0)), make([]byte, 16), le(uint32(2), uint32(0x01)),
			clsidFileMoniker, le(anti, uint32(len(path))), path, le(uint16(0xFFFF), uint16(0xDEAD)), make([]byte, 20), le(uint32(0)))
	}
	_, s := mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "a"), fileLink(2))
	if h, ok := s.Row(0).Hyperlink(0); !ok || h.ShortedFilePath != `..\..\x.xls` {
		t.Errorf("hyperlink = %+v, %v", h, ok)
	}
	_, s = mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "a"), fileLink(1000))
	if len(s.Hyperlinks()) != 0 {
		t.Error("hyperlink with 1000 parent references kept")
	}
}

func TestWorkbookChoiceMatchesTreeSearch(t *testing.T) {
	biff8 := book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "Eight"}}}.stream()
	biff5 := book{biff: 5, globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "Five"}}}.stream()
	// Unsorted: Book is the root's child and Workbook its right sibling,
	// which is where a tree search looks, so Workbook is found.
	f, l := cfbtest.Build([]cfbtest.Stream{{Name: "Book", Data: biff5}, {Name: "Workbook", Data: biff8}}, cfbtest.Options{Unsorted: true})
	name := func(f []byte) string {
		wb, err := OpenReader(bytes.NewReader(f), "")
		if err != nil {
			t.Fatal(err)
		}
		return wb.sheets[0].Name
	}
	if got := name(f); got != "Eight" {
		t.Errorf("sorted placement: opened %q", got)
	}
	// Move Workbook to Book's left, where a tree search does not look.
	// Excel would then open Book, and so must this reader.
	off := l.EntryOffset["Book"]
	binary.LittleEndian.PutUint32(b4(f, off+68), binary.LittleEndian.Uint32(f[off+72:]))
	binary.LittleEndian.PutUint32(b4(f, off+72), 0xFFFFFFFF)
	if got := name(f); got != "Five" {
		t.Errorf("hidden Workbook: opened %q", got)
	}
}

func b4(b []byte, off int) []byte { return b[off : off+4] }

func TestNilRow(t *testing.T) {
	var r *Row
	if r.Col(0) != "" || r.ColExact(0) != "" || r.Value(0).Kind != KindEmpty || r.FirstCol() != 0 || r.LastCol() != 0 || r.Index() != -1 {
		t.Error("nil Row is not empty")
	}
	if _, ok := r.Hyperlink(0); ok {
		t.Error("nil Row has a hyperlink")
	}
}

func TestSheetScanLinear(t *testing.T) {
	// Many sheets point into one long run of empty CONTINUE records. Each
	// sheet's scan must stop at the next sheet's start.
	const n = 20000
	var g bytes.Buffer
	g.Write(bofRec(0x0600, 0x0005))
	var patch []int
	for i := 0; i < n; i++ {
		patch = append(patch, g.Len()+4)
		g.Write(rec(recBOUNDSHEET, le(uint32(0), uint8(0), uint8(0)), []byte{1, 0, 'S'}))
	}
	g.Write(eofRec())
	g.Write(rec(0x1234)) // so the globals' EOF does not absorb the run
	out := g.Bytes()
	base := len(out)
	for i := 0; i < n; i++ {
		binary.LittleEndian.PutUint32(out[patch[i]:], uint32(base+4*i))
		out = append(out, rec(recCONTINUE)...)
	}
	f, _ := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: out}}, cfbtest.Options{})
	start := time.Now()
	wb, err := OpenReader(bytes.NewReader(f), "")
	if err != nil {
		t.Fatal(err)
	}
	if d := time.Since(start); d > slow*3*time.Second {
		t.Errorf("%d byte file took %v", len(f), d)
	}
	if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
		t.Errorf("err = %v", err)
	}
}

func TestMergedRenderCost(t *testing.T) {
	// A merged range over 4096 rows x 256 columns whose top-left cell is a
	// date with a 255-character format. The README loop must render that
	// value once, not once per covered cell.
	code := strings.Repeat("d/", 127) + "d"
	recs := [][]byte{numberRec(0, 0, 0, 43831), rec(recMERGECELLS, le(uint16(1), uint16(0), uint16(4095), uint16(0), uint16(255)))}
	for r := 0; r < 4096; r++ {
		recs = append(recs, rec(recROW, le(uint16(r), uint16(0), uint16(256), uint16(0), uint16(0), uint16(0), uint32(0))))
	}
	_, s := mustSheet(t, [][]byte{formatRec(164, code), xfRec(164)}, recs...)
	want := strings.Repeat("1/", 127) + "1"
	start := time.Now()
	n := 0
	for _, row := range s.Rows() {
		for c := row.FirstCol(); c < row.LastCol(); c++ {
			if got := row.Col(c); got != want {
				t.Fatalf("(%d,%d) = %q", row.Index(), c, got)
			}
			n++
		}
	}
	if d := time.Since(start); d > slow*2*time.Second {
		t.Errorf("%d cells took %v", n, d)
	}
}

func TestReadAllKeepsGoodSheets(t *testing.T) {
	bad := rec(recMULRK, le(uint16(0), uint16(0), uint16(0), rkInt(1), uint16(5)))
	wb, err := openBook(t, book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "Bad", recs: [][]byte{bad}}, {name: "Good", recs: [][]byte{labelRec(0, 0, "ok")}}}})
	if err != nil {
		t.Fatal(err)
	}
	data, names, err := wb.ReadAll(0)
	if !errors.Is(err, ErrCorrupt) || len(data) != 2 || data[0] != nil || data[1][0][0] != "ok" || names[1] != "Good" {
		t.Errorf("ReadAll = %q, %q, %v", data, names, err)
	}
}

func TestCoveredCellShowsTopLeft(t *testing.T) {
	_, s := mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 0, "top"), labelRec(0, 1, "hidden"),
		rec(recMERGECELLS, le(uint16(1), uint16(0), uint16(0), uint16(0), uint16(1))))
	r := s.Row(0)
	if r.Col(1) != "top" || r.Value(1).Text != "top" || r.ColExact(1) != "hidden" {
		t.Errorf("Col %q, Value %q, ColExact %q", r.Col(1), r.Value(1).Text, r.ColExact(1))
	}
}

func TestRangeIndexBuildLinear(t *testing.T) {
	// Ranges spread over every level of the tree.
	build := func(n int) time.Duration {
		refs := make([]ref, n)
		for i := range refs {
			r0 := i * 7919 % 65536
			r1 := r0 + (i%60000)%(65536-r0)
			refs[i] = ref{uint16(r0), uint16(r1), uint16(i % 256), uint16(i % 256)}
		}
		start := time.Now()
		newRangeIndex(refs)
		return time.Since(start)
	}
	build(1000) // warm up
	small, large := build(50000), build(400000)
	if large > 16*small+slow*50*time.Millisecond {
		t.Errorf("8x the ranges took %v vs %v", large, small)
	}
}

func TestRound4Fixes(t *testing.T) {
	// A mixed format code: numbers below 1, dates from 1.
	_, s := mustSheet(t, [][]byte{formatRec(164, "[<1]0.0000;yyyy"), xfRec(164)}, numberRec(0, 0, 0, 43831), numberRec(0, 1, 0, 0.5))
	if col(s, 0, 0) != "2020" || col(s, 0, 1) != "0.5" || s.Row(0).Value(0).Kind != KindDate || s.Row(0).Value(1).Kind != KindNumber {
		t.Errorf("mixed format: %q %q", col(s, 0, 0), col(s, 0, 1))
	}
	// Special numbers.
	_, s = mustSheet(t, [][]byte{xfRec(0)}, numberRec(0, 0, 0, math.Copysign(0, -1)), numberRec(0, 1, 0, math.NaN()), numberRec(0, 2, 0, math.Inf(-1)))
	if col(s, 0, 0) != "0" || col(s, 0, 1) != "#NUM!" || col(s, 0, 2) != "#NUM!" {
		t.Errorf("special numbers: %q %q %q", col(s, 0, 0), col(s, 0, 1), col(s, 0, 2))
	}
	// Cols lists only stored cells.
	_, s = mustSheet(t, [][]byte{xfRec(0)}, labelRec(0, 3, "a"), rec(recBLANK, le(uint16(0), uint16(200), uint16(0))), labelRec(0, 9, "b"))
	if got := s.Row(0).Cols(); len(got) != 2 || got[0] != 3 || got[1] != 9 || s.Row(0).LastCol() != 201 {
		t.Errorf("Cols = %v", got)
	}
	// ROW records must be complete.
	wb, err := openBook(t, book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "S", recs: [][]byte{rec(recROW, le(uint16(0), uint16(0), uint16(256)))}}}})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
		t.Errorf("short ROW: %v", err)
	}
	// Any non-zero DATEMODE is the 1904 system.
	wb, _ = mustSheet(t, [][]byte{rec(recDATEMODE, le(uint16(2))), xfRec(0)})
	if !wb.date1904 {
		t.Error("DATEMODE 2 is not 1904")
	}
	// Spellings of UTF-8 all mean the file's code page.
	for _, cs := range []string{"UTF-8", "unicode-1-1-utf-8", "utf8"} {
		if e, err := charsetEncoding(cs); e != nil || err != nil {
			t.Errorf("charset %q = %v, %v", cs, e, err)
		}
	}
	// An over-long format on a built-in date index becomes General.
	_, s = mustSheet(t, [][]byte{formatRec(14, strings.Repeat("y", 300)), xfRec(14)}, numberRec(0, 0, 0, 43831))
	if got := col(s, 0, 0); got != "43831" {
		t.Errorf("over-long format: %q", got)
	}
	// One sheet's bad offset does not hide the others.
	stream := book{globals: [][]byte{xfRec(0)}, sheets: []sheet{{name: "A"}, {name: "B", recs: [][]byte{labelRec(0, 0, "b")}}}}.stream()
	i := bytes.Index(stream, []byte{0x85, 0x00})
	binary.LittleEndian.PutUint32(stream[i+4:], 1<<30)
	f, _ := cfbtest.Build([]cfbtest.Stream{{Name: "Workbook", Data: stream}}, cfbtest.Options{})
	wb, err = OpenReader(bytes.NewReader(f), "")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := wb.GetSheet(0); !errors.Is(err, ErrCorrupt) {
		t.Errorf("bad offset: %v", err)
	}
	if s, err := wb.GetSheet(1); err != nil || col(s, 0, 0) != "b" {
		t.Errorf("good sheet: %v", err)
	}
	if data, names, _ := wb.ReadAll(0); len(data) != len(names) || data[0] != nil {
		t.Errorf("ReadAll = %v, %v", data, names)
	}
}
