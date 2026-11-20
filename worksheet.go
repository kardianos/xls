package xls

import (
	"math"
	"sort"
	"sync"
	"sync/atomic"
)

// TWorkSheetVisibility is the visibility of a sheet tab.
type TWorkSheetVisibility byte

// Sheet visibility values.
const (
	WorkSheetVisible    TWorkSheetVisibility = 0
	WorkSheetHidden     TWorkSheetVisibility = 1
	WorkSheetVeryHidden TWorkSheetVisibility = 2
)

// CellRange is a rectangle of cells. All bounds are zero-based and
// inclusive.
type CellRange struct {
	FirstRow, LastRow int
	FirstCol, LastCol int
}

// WorkSheet is one sheet of a WorkBook. Get one with WorkBook.GetSheet.
//
// The exported fields are set when the sheet is read and must not be
// changed; a WorkSheet is shared by every caller of GetSheet.
type WorkSheet struct {
	Name string
	// Selected reports whether the sheet tab is selected.
	Selected bool
	// RightToLeft reports whether the sheet is shown right to left.
	RightToLeft bool
	Visibility  TWorkSheetVisibility
	// Type is the kind of sheet: worksheet, chart, macro sheet or
	// Visual Basic module.
	Type SheetType
	// MaxRow is the index of the last row that has a ROW record or a cell.
	// A sheet with no rows also has MaxRow 0.
	MaxRow uint16

	wb      *WorkBook
	offset  int   // stream offset of the sheet's BOF record
	end     int   // stream offset after the sheet's EOF record
	scanErr error // framing error found when the workbook was opened

	once     sync.Once
	err      error
	rows     map[uint16]*Row
	rowList  []*Row // rows sorted by index
	merges   []ref
	mergeIdx rangeIndex
	// mergeText caches the text of each merged range's top-left cell, so
	// a loop over the covered cells renders it once.
	mergeText []atomic.Pointer[string]
	links     []HyperLink
	linkIdx   rangeIndex
	badCol    bool // a cell record named a column past the last one
}

// SheetType is the kind of a sheet, from its BOF record.
type SheetType uint16

// Sheet types. Worksheets and macro sheets have cells; the others are
// returned with no rows.
const (
	SheetWorksheet SheetType = 0x0010
	SheetChart     SheetType = 0x0020
	SheetMacro     SheetType = 0x0040 // Excel 4.0 macro sheet
	SheetVBModule  SheetType = 0x0006
)

// Row returns row i (zero-based), or nil if the row has no ROW record and
// no cells.
func (w *WorkSheet) Row(i int) *Row {
	if i < 0 || i > math.MaxUint16 {
		return nil
	}
	return w.rows[uint16(i)]
}

// Rows returns the rows that have a ROW record or a cell, in order. Unlike
// a loop from 0 to MaxRow, its length is bounded by the size of the file.
func (w *WorkSheet) Rows() []*Row {
	return append([]*Row(nil), w.rowList...)
}

// MergedCells returns the merged cell ranges of the sheet.
func (w *WorkSheet) MergedCells() []CellRange {
	out := make([]CellRange, len(w.merges))
	for i, m := range w.merges {
		out[i] = m.cellRange()
	}
	return out
}

// Hyperlinks returns the hyperlinks of the sheet.
func (w *WorkSheet) Hyperlinks() []HyperLink {
	return append([]HyperLink(nil), w.links...)
}

// mergedAt returns the merged range that contains a cell, and its index.
func (w *WorkSheet) mergedAt(row, col int) (CellRange, int, bool) {
	if i := w.mergeIdx.find(row, col); i >= 0 {
		return w.merges[i].cellRange(), i, true
	}
	return CellRange{}, -1, false
}

// Row is one row of a sheet. A nil *Row, as WorkSheet.Row returns for a
// missing row, behaves as an empty row.
type Row struct {
	sheet *WorkSheet
	index int

	hasInfo        bool   // a ROW record was seen
	colMic, colMac uint16 // from the ROW record; colMac is exclusive

	minCol int // first column with a cell or blank, or -1
	width  int // one past the last column with a cell or blank
	cells  []cell
}

// find returns the cell in column col, or nil.
func (r *Row) find(col int) *cell {
	if r == nil || col < 0 || col > math.MaxUint16 {
		return nil
	}
	i := sort.Search(len(r.cells), func(i int) bool { return int(r.cells[i].col) >= col })
	if i < len(r.cells) && int(r.cells[i].col) == col {
		return &r.cells[i]
	}
	return nil
}

// Col returns the text of column n (zero-based), or "" if there is no
// value. As in Excel, every cell of a merged range shows the range's
// top-left value, even a covered cell that holds a value of its own (use
// ColExact for that). Row returns nil for a row with no records even if a
// merged range covers it; Excel writes blank cells for the covered area,
// so in its files those rows exist.
func (r *Row) Col(n int) string {
	if r == nil {
		return ""
	}
	if i, top := r.mergedTop(n); i >= 0 {
		if top == nil {
			return ""
		}
		if p := r.sheet.mergeText[i].Load(); p != nil {
			return *p
		}
		s := r.sheet.text(top)
		r.sheet.mergeText[i].Store(&s)
		return s
	}
	if c := r.find(n); c != nil {
		return r.sheet.text(c)
	}
	return ""
}

// ColExact returns the text of column n (zero-based), or "" if there is no
// value. Merged ranges have their value only in the top-left cell.
func (r *Row) ColExact(n int) string {
	if r == nil {
		return ""
	}
	if c := r.find(n); c != nil {
		return r.sheet.text(c)
	}
	return ""
}

// Value returns the typed value of column n (zero-based). Like Col, every
// cell of a merged range returns the range's top-left value. Empty cells
// return a CellValue of kind KindEmpty.
func (r *Row) Value(n int) CellValue {
	if r == nil {
		return CellValue{}
	}
	if i, top := r.mergedTop(n); i >= 0 {
		if top == nil {
			return CellValue{}
		}
		return r.sheet.value(top)
	}
	if c := r.find(n); c != nil {
		return r.sheet.value(c)
	}
	return CellValue{}
}

// mergedTop reports whether column n of this row is covered by a merged
// range, other than as its top-left cell. If so, it returns the range's
// index and its top-left cell, which is nil if that cell is empty.
func (r *Row) mergedTop(n int) (int, *cell) {
	m, i, ok := r.sheet.mergedAt(r.index, n)
	if !ok || (m.FirstRow == r.index && m.FirstCol == n) {
		return -1, nil
	}
	return i, r.sheet.Row(m.FirstRow).find(m.FirstCol)
}

// Cols returns the columns of the row that hold a value (not blank
// cells), in order. A loop over Cols visits only stored cells, which is
// cheaper than a loop from FirstCol to LastCol on sparse rows. Covered
// cells of a merged range that hold no value of their own are not
// included.
func (r *Row) Cols() []int {
	if r == nil {
		return nil
	}
	out := make([]int, len(r.cells))
	for i := range r.cells {
		out[i] = int(r.cells[i].col)
	}
	return out
}

// Index returns the row's index (zero-based), or -1 for a nil Row.
func (r *Row) Index() int {
	if r == nil {
		return -1
	}
	return r.index
}

// Hyperlink returns the hyperlink that covers column n of this row, if any.
func (r *Row) Hyperlink(n int) (HyperLink, bool) {
	if r == nil {
		return HyperLink{}, false
	}
	if i := r.sheet.linkIdx.find(r.index, n); i >= 0 {
		return r.sheet.links[i], true
	}
	return HyperLink{}, false
}

// FirstCol returns the index of the first column of the row.
func (r *Row) FirstCol() int {
	if r == nil {
		return 0
	}
	first := r.minCol
	if r.hasInfo && (first < 0 || int(r.colMic) < first) {
		first = int(r.colMic)
	}
	if first < 0 {
		return 0
	}
	return first
}

// LastCol returns one past the index of the last column of the row, so
// columns FirstCol() <= c < LastCol() hold the row's cells. It is at most
// 256, the number of columns in a sheet.
func (r *Row) LastCol() int {
	if r == nil {
		return 0
	}
	last := r.width
	if r.hasInfo && int(r.colMac) > last {
		last = int(r.colMac)
	}
	return last
}
