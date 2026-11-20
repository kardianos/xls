package xls

import (
	"math"
	"strconv"
	"time"

	"github.com/kardianos/xls/yymmdd"
)

// ToDateTime converts an Excel date serial number to a time in UTC, using
// the workbook's date system (1900 or 1904). It returns the zero time if the
// value is not a valid Excel date (negative, from 10000-01-01 on, or NaN).
// See yymmdd.SerialTime for how Excel's 1900 leap-year error is handled.
func (w *WorkBook) ToDateTime(f float64) time.Time {
	t, _ := yymmdd.SerialTime(f, w.date1904)
	return t
}

// formatNumber renders a number in the shortest decimal form that reads
// back as the same float64. Very large and very small magnitudes use an
// exponent, so the text is never longer than about 25 bytes.
func formatNumber(f float64) string {
	switch {
	case math.IsNaN(f) || math.IsInf(f, 0):
		return "#NUM!" // Excel cannot hold these; it shows #NUM!
	case f == 0:
		return "0" // including -0
	}
	if a := math.Abs(f); a >= 1e21 || a < 1e-6 {
		return strconv.FormatFloat(f, 'g', -1, 64)
	}
	return strconv.FormatFloat(f, 'f', -1, 64)
}
