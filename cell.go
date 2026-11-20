package xls

import (
	"fmt"
	"time"

	"github.com/kardianos/xls/yymmdd"
)

// CellKind is the type of a cell value.
type CellKind uint8

// Cell kinds. A formula cell has the kind of its cached result.
const (
	KindEmpty  CellKind = iota // no value
	KindString                 // Text holds the string
	KindNumber                 // Float holds the number
	KindDate                   // a number in a date or time format: Float holds the serial number, Time the date
	KindBool                   // Bool holds the value
	KindError                  // Text holds the error, such as "#N/A"
)

func (k CellKind) String() string {
	switch k {
	case KindEmpty:
		return "empty"
	case KindString:
		return "string"
	case KindNumber:
		return "number"
	case KindDate:
		return "date"
	case KindBool:
		return "bool"
	case KindError:
		return "error"
	}
	return fmt.Sprintf("CellKind(%d)", uint8(k))
}

// CellValue is the typed value of a cell.
type CellValue struct {
	Kind CellKind
	// Text is the string value of a KindString cell, or the error text of
	// a KindError cell.
	Text string
	// Float is the value of a KindNumber cell, or the date serial number of
	// a KindDate cell.
	Float float64
	// Int is Float as an integer when Float has no fractional part and fits
	// in an int64. Otherwise it is zero.
	Int int64
	// Bool is the value of a KindBool cell.
	Bool bool
	// Time is the value of a KindDate cell, in UTC.
	Time time.Time
	// Format is the number format code of a number or date cell.
	Format string
}

// cell kinds stored in a sheet.
const (
	cellString = iota + 1
	cellNumber
	cellBool
	cellError
	cellEmptyString // a formula whose cached result is the empty string
)

type cell struct {
	col  uint16
	xf   uint16
	kind uint8
	b    uint8 // bool value or error code
	num  float64
	str  string
}

var errorText = map[uint8]string{
	0x00: "#NULL!",
	0x07: "#DIV/0!",
	0x0F: "#VALUE!",
	0x17: "#REF!",
	0x1D: "#NAME?",
	0x24: "#NUM!",
	0x2A: "#N/A",
	0x2B: "#GETTING_DATA",
}

func errText(code uint8) string {
	if s, ok := errorText[code]; ok {
		return s
	}
	return fmt.Sprintf("#ERR%d", code)
}

func (w *WorkSheet) value(c *cell) CellValue {
	switch c.kind {
	case cellString:
		return CellValue{Kind: KindString, Text: c.str}
	case cellEmptyString:
		return CellValue{Kind: KindString}
	case cellBool:
		return CellValue{Kind: KindBool, Bool: c.b != 0}
	case cellError:
		return CellValue{Kind: KindError, Text: errText(c.b)}
	case cellNumber:
		info := w.wb.xf(c.xf)
		v := CellValue{Kind: KindNumber, Float: c.num, Format: info.code}
		if isIntegral(c.num) {
			v.Int = int64(c.num)
		}
		if info.layout != nil && info.layout.IsDate(c.num, w.wb.date1904) {
			if t, ok := yymmdd.SerialTime(c.num, w.wb.date1904); ok {
				v.Kind = KindDate
				v.Time = t
			}
		}
		return v
	}
	return CellValue{}
}

// text renders a cell as a string. Dates are formatted with their number
// format; other numbers use formatNumber.
func (w *WorkSheet) text(c *cell) string {
	switch c.kind {
	case cellString:
		return c.str
	case cellBool:
		if c.b != 0 {
			return "TRUE"
		}
		return "FALSE"
	case cellError:
		return errText(c.b)
	case cellNumber:
		if info := w.wb.xf(c.xf); info.layout != nil {
			if s, ok := info.layout.FormatSerial(c.num, w.wb.date1904); ok {
				return s
			}
		}
		return formatNumber(c.num)
	}
	return ""
}
