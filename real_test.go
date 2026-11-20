package xls

import (
	"path/filepath"
	"testing"
	"time"
)

// The files in testdata/real were written by LibreOffice Calc (Excel 97
// export), not by this package's test helpers. The constant TRUE in E1 is
// stored by LibreOffice as the number 1, so it reads as "1".

func TestRealLibreOfficeFeatures(t *testing.T) {
	wb, err := Open(filepath.Join("testdata", "real", "libreoffice-features.xls"), "")
	if err != nil {
		t.Fatal(err)
	}
	if wb.NumSheets() != 2 || wb.Author != "Calc" {
		t.Errorf("sheets %d, author %q", wb.NumSheets(), wb.Author)
	}
	s, err := wb.GetSheet(0)
	if err != nil {
		t.Fatal(err)
	}
	want := map[[2]int]string{
		{0, 0}: "Text", {0, 1}: "123.5", {0, 2}: "2020-01-01", {0, 3}: "13:30:45", {0, 4}: "1",
		{0, 5}: "#DIV/0!", {0, 6}: "247", {0, 7}: "ab", {0, 8}: "Example", {0, 9}: "TRUE",
		{1, 0}: "Merged", {1, 1}: "Merged", {1, 2}: "Merged", {1, 3}: "after",
		{2, 0}: "Merged", {2, 2}: "Merged", {2, 3}: "row3",
		{3, 0}: "2020-01-01 18:00", {3, 1}: "36:30", {3, 2}: "9999-12-31", {3, 3}: "1e+300",
		{4, 0}: "Café ünï", {4, 1}: "Café ünï!",
		{5, 0}: "Café ünï6", {5, 1}: "12", {14, 0}: "Café ünï15", {14, 1}: "30",
	}
	for rc, w := range want {
		if got := col(s, rc[0], rc[1]); got != w {
			t.Errorf("(%d,%d) = %q, want %q", rc[0], rc[1], got, w)
		}
	}
	if got := s.Row(1).ColExact(1); got != "" {
		t.Errorf("ColExact inside merged range = %q", got)
	}
	kinds := map[[2]int]CellKind{
		{0, 0}: KindString, {0, 1}: KindNumber, {0, 2}: KindDate, {0, 3}: KindDate, {0, 5}: KindError,
		{0, 6}: KindNumber, {0, 7}: KindString, {0, 9}: KindBool, {3, 1}: KindDate, {5, 0}: KindString,
	}
	for rc, k := range kinds {
		if got := s.Row(rc[0]).Value(rc[1]).Kind; got != k {
			t.Errorf("(%d,%d) kind = %v, want %v", rc[0], rc[1], got, k)
		}
	}
	v := s.Row(3).Value(0)
	if !v.Time.Equal(time.Date(2020, 1, 1, 18, 0, 0, 0, time.UTC)) {
		t.Errorf("date-time Value = %+v", v)
	}
	h, ok := s.Row(0).Hyperlink(8)
	if !ok || h.URL != "https://example.com/" || h.Description != "Example" {
		t.Errorf("hyperlink = %+v, %v", h, ok)
	}
	if m := s.MergedCells(); len(m) != 1 || m[0] != (CellRange{FirstRow: 1, LastRow: 2, FirstCol: 0, LastCol: 2}) {
		t.Errorf("merged cells = %v", m)
	}
	s2, err := wb.GetSheet(1)
	if err != nil || s2.Name != "Second" || col(s2, 0, 0) != "x" {
		t.Errorf("second sheet: %v", err)
	}
}

func TestRealLibreOffice1904(t *testing.T) {
	wb, err := Open(filepath.Join("testdata", "real", "libreoffice-1904.xls"), "")
	if err != nil {
		t.Fatal(err)
	}
	s, err := wb.GetSheet(0)
	if err != nil {
		t.Fatal(err)
	}
	if !wb.date1904 || col(s, 0, 0) != "2020-01-01" || col(s, 0, 1) != "1904-01-02 12:00" {
		t.Errorf("1904=%v, %q %q", wb.date1904, col(s, 0, 0), col(s, 0, 1))
	}
	if v := s.Row(0).Value(1); v.Float != 1.5 || !v.Time.Equal(time.Date(1904, 1, 2, 12, 0, 0, 0, time.UTC)) {
		t.Errorf("Value = %+v", v)
	}
}
