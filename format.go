package xls

import "github.com/kardianos/xls/yymmdd"

// maxFormatLen is the longest number format code Excel accepts. Longer
// codes are ignored, and cells using them show General numbers.
const maxFormatLen = 255

// builtinFormats are the number formats that have fixed indexes and are not
// always written to the file as FORMAT records ([MS-XLS] 2.4.126).
//
// Some built-in formats depend on the locale. Excel shows format 14 as the
// system short date, and formats 27-36 and 50-58 differ between Chinese,
// Japanese and Korean locales. This package uses ISO 8601 ordering for them.
var builtinFormats = map[uint16]string{
	0:  "General",
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	5:  `"$"#,##0_);\("$"#,##0\)`,
	6:  `"$"#,##0_);[Red]\("$"#,##0\)`,
	7:  `"$"#,##0.00_);\("$"#,##0.00\)`,
	8:  `"$"#,##0.00_);[Red]\("$"#,##0.00\)`,
	9:  "0%",
	10: "0.00%",
	11: "0.00E+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "yyyy-mm-dd",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "h:mm",
	21: "h:mm:ss",
	22: "yyyy-mm-dd h:mm",
	37: "#,##0 ;(#,##0)",
	38: "#,##0 ;[Red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[Red](#,##0.00)",
	41: `_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`,
	42: `_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)`,
	43: `_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)`,
	44: `_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)`,
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mm:ss.0",
	48: "##0.0E+0",
	49: "@",
}

func init() {
	for _, i := range []uint16{27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 57, 58} {
		builtinFormats[i] = "yyyy-mm-dd"
	}
	for _, i := range []uint16{32, 33, 55, 56} {
		builtinFormats[i] = "h:mm:ss"
	}
}

// xfInfo is the number format of a cell style (XF record).
type xfInfo struct {
	code   string         // format code
	layout *yymmdd.Layout // compiled code, for date formats only
	maxLen int            // upper bound on the text of a number in this format
}

// resolveFormats computes the number format of every XF record. A format
// code is compiled once, however many indexes and styles use it; codes that
// have a date or time part in any section keep the compiled layout.
func (w *WorkBook) resolveFormats(xfFormats []uint16) {
	w.xfs = make([]xfInfo, len(xfFormats))
	byIndex := make(map[uint16]xfInfo)
	byCode := make(map[string]*yymmdd.Layout)
	for i, idx := range xfFormats {
		info, ok := byIndex[idx]
		if !ok {
			code, found := w.formats[idx]
			if !found {
				code, found = builtinFormats[idx]
			}
			if !found {
				code = "General"
			}
			l, seen := byCode[code]
			if !seen {
				if l = yymmdd.Compile(code); !l.HasDate() {
					l = nil
				}
				byCode[code] = l
			}
			info = xfInfo{code: code, layout: l, maxLen: maxNumberLen}
			if l != nil {
				info.maxLen = max(maxNumberLen, l.MaxLen())
			}
			byIndex[idx] = info
		}
		w.xfs[i] = info
	}
}

// xf returns the number format for an XF index, or General if the index is
// not valid.
func (w *WorkBook) xf(i uint16) xfInfo {
	if int(i) < len(w.xfs) {
		return w.xfs[i]
	}
	return xfInfo{code: "General", maxLen: maxNumberLen}
}

// maxNumberLen bounds the text formatNumber returns.
const maxNumberLen = 32
