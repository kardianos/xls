package xls

import (
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"

	"github.com/kardianos/xls/yymmdd"
)

type CellValue struct {
	Text   string
	Float  float64
	Int    int64
	Format string
}

//content type
type contentHandler interface {
	String(*WorkBook) []string
	Value(*WorkBook) CellValue
	FirstCol() uint16
	LastCol() uint16
}
type columner interface {
	Row() uint16
}

type contentColumner interface {
	contentHandler
	columner
}

type Col struct {
	RowB      uint16
	FirstColB uint16
}

func (c *Col) Row() uint16 {
	return c.RowB
}

func (c *Col) FirstCol() uint16 {
	return c.FirstColB
}

func (c *Col) LastCol() uint16 {
	return c.FirstColB
}

func (c *Col) String(wb *WorkBook) []string {
	return []string{"default"}
}

type XfRk struct {
	Index uint16
	Rk    RK
}

func (xf *XfRk) String(wb *WorkBook) string {
	idx := int(xf.Index)
	if len(wb.XF) <= idx {
		return xf.Rk.String()
	}
	fNo := wb.XF[idx].formatNo()
	if fNo >= 164 { // user defined format
		if formatter := wb.Formats[fNo]; formatter != nil {
			formatterLower := strings.ToLower(formatter.str)
			if formatterLower == "general" ||
				strings.Contains(formatter.str, "#") ||
				strings.Contains(formatter.str, ".00") ||
				strings.Contains(formatterLower, "m/y") ||
				strings.Contains(formatterLower, "d/y") ||
				strings.Contains(formatterLower, "m.y") ||
				strings.Contains(formatterLower, "d.y") ||
				strings.Contains(formatterLower, "h:") ||
				strings.Contains(formatterLower, "д.г") {
				//If format contains # or .00 then this is a number
				return xf.Rk.String()
			} else {
				i, f, isFloat := xf.Rk.number()
				if !isFloat {
					f = float64(i)
				}
				t := timeFromExcelTime(f, wb.dateMode == 1)
				return yymmdd.Format(t, formatter.str)
			}
		}
		// see http://www.openoffice.org/sc/excelfileformat.pdf Page #174
	} else if 14 <= fNo && fNo <= 17 || fNo == 22 || 27 <= fNo && fNo <= 36 || 50 <= fNo && fNo <= 58 { // jp. date format
		i, f, isFloat := xf.Rk.number()
		if !isFloat {
			f = float64(i)
		}
		t := timeFromExcelTime(f, wb.dateMode == 1)
		return t.Format(time.RFC3339)
	}
	return xf.Rk.String()
}
func (xf *XfRk) Value(wb *WorkBook) CellValue {
	return CellValue{
		Text: xf.String(wb),
	}
}

type RK uint32

func (rk RK) number() (intNum int64, floatNum float64, isFloat bool) {
	multiplied := rk & 1
	isInt := rk & 2
	val := int32(rk) >> 2
	if isInt == 0 {
		isFloat = true
		floatNum = math.Float64frombits(uint64(val) << 34)
		if multiplied != 0 {
			floatNum = floatNum / 100
		}
		return
	}
	if multiplied != 0 {
		isFloat = true
		floatNum = float64(val) / 100
		return
	}
	return int64(val), 0, false
}

func (rk RK) String() string {
	i, f, isFloat := rk.number()
	if isFloat {
		return strconv.FormatFloat(f, 'f', -1, 64)
	}
	return strconv.FormatInt(i, 10)
}

var ErrIsInt = fmt.Errorf("is int")

func (rk RK) Float() (float64, error) {
	_, f, isFloat := rk.number()
	if !isFloat {
		return 0, ErrIsInt
	}
	return f, nil
}
func (rk RK) Value(wb *WorkBook) CellValue {
	i, f, _ := rk.number()
	return CellValue{
		Int:   int64(i),
		Float: f,
	}
}

var _ contentHandler = &MulrkCol{}

type MulrkCol struct {
	Col
	Xfrks    []XfRk
	LastColB uint16
}

func (c *MulrkCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulrkCol) String(wb *WorkBook) []string {
	var res = make([]string, len(c.Xfrks))
	for i := 0; i < len(c.Xfrks); i++ {
		xfrk := c.Xfrks[i]
		res[i] = xfrk.String(wb)
	}
	return res
}
func (c *MulrkCol) Value(wb *WorkBook) CellValue {
	if len(c.Xfrks) == 0 {
		return CellValue{}
	}
	return c.Xfrks[0].Rk.Value(wb)
}

var _ contentHandler = &MulBlankCol{}

type MulBlankCol struct {
	Col
	Xfs      []uint16
	LastColB uint16
}

func (c *MulBlankCol) LastCol() uint16 {
	return c.LastColB
}

func (c *MulBlankCol) String(wb *WorkBook) []string {
	return make([]string, len(c.Xfs))
}
func (c *MulBlankCol) Value(wb *WorkBook) CellValue {
	return CellValue{}
}

var _ contentHandler = &NumberCol{}

type NumberCol struct {
	Col
	Index uint16
	Float float64
}

func (c *NumberCol) String(wb *WorkBook) []string {
	fNo := wb.XF[c.Index].formatNo()
	fo, ok := wb.Formats[fNo]
	if !ok {
		return []string{strconv.FormatFloat(c.Float, 'f', -1, 64)}
	}
	fs := fo.str

	switch {
	default:
		return []string{strconv.FormatFloat(c.Float, 'f', -1, 64)}
	case strings.ContainsAny(fs, "YyMmDdHhSs"):
		t := wb.ToDateTime(c.Float)
		return []string{yymmdd.Format(t, fs)}
	case strings.ContainsAny(fs, "0.#"):
		// TODO: actually format number.
		return []string{strconv.FormatFloat(c.Float, 'f', -1, 64)}
	}
}

func (c *NumberCol) Value(wb *WorkBook) CellValue {
	fNo := wb.XF[c.Index].formatNo()
	fo, _ := wb.Formats[fNo]
	return CellValue{
		Format: fo.str,
		Float:  c.Float,
	}
}

var _ contentHandler = &FormulaStringCol{}

type FormulaStringCol struct {
	Col
	RenderedValue string
}

func (c *FormulaStringCol) String(wb *WorkBook) []string {
	return []string{c.RenderedValue}
}
func (c *FormulaStringCol) Value(wb *WorkBook) CellValue {
	return CellValue{
		Text: c.RenderedValue,
	}
}

//str, err = wb.get_string(buf_item, size)
//wb.sst[offset_pre] = wb.sst[offset_pre] + str

type FormulaCol struct {
	Header struct {
		Col
		IndexXf uint16
		Result  [8]byte
		Flags   uint16
		_       uint32
	}
	Bts []byte
}

func (c *FormulaCol) String(wb *WorkBook) []string {
	return []string{"FormulaCol"}
}
func (c *FormulaCol) Value(wb *WorkBook) CellValue {
	return CellValue{}
}

var _ contentHandler = &RkCol{}

type RkCol struct {
	Col
	Xfrk XfRk
}

func (c *RkCol) String(wb *WorkBook) []string {
	return []string{c.Xfrk.String(wb)}
}
func (c *RkCol) Value(wb *WorkBook) CellValue {
	return c.Xfrk.Value(wb)
}

var _ contentHandler = &LabelsstCol{}

type LabelsstCol struct {
	Col
	Xf  uint16
	Sst uint32
}

func (c *LabelsstCol) String(wb *WorkBook) []string {
	return []string{wb.sst[int(c.Sst)]}
}
func (c *LabelsstCol) Value(wb *WorkBook) CellValue {
	return CellValue{
		Text: wb.sst[int(c.Sst)],
	}
}

var _ contentHandler = &labelCol{}

type labelCol struct {
	BlankCol
	Str string
}

func (c *labelCol) String(wb *WorkBook) []string {
	return []string{c.Str}
}
func (c *labelCol) Value(wb *WorkBook) CellValue {
	return CellValue{
		Text: c.Str,
	}
}

var _ contentHandler = &BlankCol{}

type BlankCol struct {
	Col
	Xf uint16
}

func (c *BlankCol) String(wb *WorkBook) []string {
	return []string{""}
}

func (c *BlankCol) Value(wb *WorkBook) CellValue {
	return CellValue{}
}
