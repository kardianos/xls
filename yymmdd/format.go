// Package yymmdd formats and parses dates with Excel number format codes,
// such as "yyyy-mm-dd", "d-mmm-yy" or "[h]:mm:ss".
//
// It supports these parts of a format code:
//   - year (y, yy, yyyy, e), Buddhist era year (b, bbbb), month (m to
//     mmmm, and mmmmm for the first letter), day (d to dddd)
//   - hour (h, hh), minute (m, mm), second (s, ss) and fractions (ss.000)
//   - elapsed time ([h], [mm], [ss])
//   - AM/PM (always shown in capitals, as Excel does) and A/P
//   - quoted text, backslash escapes, _ padding and * fill
//   - bracketed colors, conditions and locales ([Red], [$-409]); the
//     currency text in [$€-407] is kept
//   - the Hijri calendar selector B2 is not supported: dates show in the
//     Gregorian calendar
//   - sections separated by ';' (FormatSerial uses the third section for
//     zero, as Excel does)
//
// Like Excel, m and mm mean minutes when they come directly after an hour or
// directly before a second (ignoring literal text), and months otherwise.
// Letters are case insensitive. Month and day names are in English.
//
// All functions take time linear in the length of the format code.
package yymmdd

import (
	"math"
	"strconv"
	"strings"
	"time"
	"unicode"
)

type kind uint8

const (
	kLit kind = iota
	kYear
	kMonth
	kMinute
	kMonthOrMinute
	kDay
	kHour
	kSecond
	kFrac
	kAMPM
	kElapsedH
	kElapsedM
	kElapsedS
	kBuddhistYear
)

// token is one part of a format code. Literal text, and an AM/PM marker
// as written, is text[lo:hi] of the Layout's text.
type token struct {
	kind   kind
	n      int32 // number of letters, or fraction digits
	lo, hi int32
}

func clamp32(n int) int32 {
	if n > math.MaxInt32 {
		return math.MaxInt32
	}
	return int32(n)
}

func hasPrefixFold(rs []rune, prefix string) bool {
	if len(rs) < len(prefix) {
		return false
	}
	for i := 0; i < len(prefix); i++ {
		if unicode.ToLower(rs[i]) != rune(prefix[i]|0x20) && rs[i] != rune(prefix[i]) {
			return false
		}
	}
	return true
}

// closer finds the next ']' at or after a position. It remembers the last
// answer, so a scan with increasing positions takes linear time overall.
type closer struct {
	rs []rune
	at int // index of the next ']' after the last query, or len(rs)
}

func (c *closer) next(i int) int {
	if c.at >= i && c.at > 0 {
		return c.at
	}
	j := i
	for j < len(c.rs) && c.rs[j] != ']' {
		j++
	}
	c.at = j
	return j
}

// condition is a bracketed comparison such as [<1] or [>=100] that selects
// a section.
type condition struct {
	op  string
	val float64
}

func (c *condition) match(v float64) bool {
	switch c.op {
	case "<":
		return v < c.val
	case "<=":
		return v <= c.val
	case ">":
		return v > c.val
	case ">=":
		return v >= c.val
	case "=":
		return v == c.val
	case "<>":
		return v != c.val
	}
	return false
}

func parseCondition(inner string) *condition {
	for _, op := range []string{"<=", ">=", "<>", "<", ">", "="} {
		if rest, ok := strings.CutPrefix(inner, op); ok {
			v, err := strconv.ParseFloat(strings.TrimSpace(rest), 64)
			if err != nil {
				return nil
			}
			return &condition{op: op, val: v}
		}
	}
	return nil
}

type section struct {
	toks    []token
	cond    *condition
	numeric bool // has number placeholders (0 # ?) or General
}

// isDate reports whether the section shows a date or time part. A
// fraction of a second alone, as in "0.000", is not one.
func (sec section) isDate() bool {
	for _, tk := range sec.toks {
		if tk.kind != kLit && tk.kind != kFrac {
			return true
		}
	}
	return false
}

// tokenizeSection reads one section of a format code starting at rs[i] and
// returns it and the index after the ';' that ends it. Literal text is
// appended to lit.
func tokenizeSection(rs []rune, i int, cl *closer, lit *strings.Builder) (section, int) {
	var cond *condition
	numeric := false
	var toks []token
	start := lit.Len() // start of the pending literal text
	flush := func() {
		if lit.Len() > start {
			toks = append(toks, token{kind: kLit, lo: int32(start), hi: int32(lit.Len())})
			start = lit.Len()
		}
	}
	add := func(k kind, n int) {
		flush()
		toks = append(toks, token{kind: k, n: clamp32(n)})
	}
	marker := func(m []rune) {
		flush()
		lit.WriteString(string(m))
		toks = append(toks, token{kind: kAMPM, lo: int32(start), hi: int32(lit.Len())})
		start = lit.Len()
	}
	run := func(i int, letter rune) int {
		n := 1
		for i+n < len(rs) && unicode.ToLower(rs[i+n]) == letter {
			n++
		}
		return n
	}
	for ; i < len(rs); i++ {
		c := rs[i]
		lc := unicode.ToLower(c)
		switch {
		case c == ';':
			flush()
			return section{resolveMinutes(toks), cond, numeric}, i + 1
		case c == '"':
			j := i + 1
			for ; j < len(rs) && rs[j] != '"'; j++ {
				lit.WriteRune(rs[j])
			}
			i = j
		case c == '\\':
			if i+1 < len(rs) {
				i++
				lit.WriteRune(rs[i])
			}
		case c == '_':
			i++
			lit.WriteByte(' ')
		case c == '*':
			i++
		case c == '@':
		case c == '[':
			j := cl.next(i + 1)
			if j >= len(rs) {
				lit.WriteRune(c)
				continue
			}
			inner := strings.ToLower(string(rs[i+1 : j]))
			switch {
			case inner != "" && strings.Trim(inner, "h") == "":
				add(kElapsedH, len(inner))
			case inner != "" && strings.Trim(inner, "m") == "":
				add(kElapsedM, len(inner))
			case inner != "" && strings.Trim(inner, "s") == "":
				add(kElapsedS, len(inner))
			case cond == nil && parseCondition(inner) != nil:
				cond = parseCondition(inner)
			case strings.HasPrefix(inner, "$"):
				cur := rs[i+2 : j]
				for k, r := range cur {
					if r == '-' {
						cur = cur[:k]
						break
					}
				}
				lit.WriteString(string(cur))
			}
			i = j
		case hasPrefixFold(rs[i:], "AM/PM"):
			marker(rs[i : i+5])
			i += 4
		case hasPrefixFold(rs[i:], "A/P"):
			marker(rs[i : i+3])
			i += 2
		case hasPrefixFold(rs[i:], "General"):
			numeric = true
			i += 6
		case lc == 'y':
			n := run(i, 'y')
			add(kYear, n)
			i += n - 1
		case lc == 'e' && i+1 < len(rs) && (rs[i+1] == '+' || rs[i+1] == '-'):
			// Scientific notation, as in 0.00E+00.
			numeric = true
			lit.WriteRune(c)
			lit.WriteRune(rs[i+1])
			i++
		case lc == 'e':
			n := run(i, 'e')
			add(kYear, 4)
			i += n - 1
		case lc == 'm':
			n := run(i, 'm')
			add(kMonthOrMinute, n)
			i += n - 1
		case lc == 'd':
			n := run(i, 'd')
			add(kDay, n)
			i += n - 1
		case lc == 'h':
			n := run(i, 'h')
			add(kHour, n)
			i += n - 1
		case lc == 's':
			n := run(i, 's')
			add(kSecond, n)
			i += n - 1
		case lc == 'g':
			// Japanese era name: not supported, dropped.
			i += run(i, 'g') - 1
		case lc == 'b' && i+1 < len(rs) && (rs[i+1] == '1' || rs[i+1] == '2'):
			// Calendar selector (B2 is the Hijri calendar): not
			// supported, dropped; dates show in the Gregorian calendar.
			i++
		case lc == 'b':
			// Buddhist era year: the Gregorian year plus 543.
			n := run(i, 'b')
			add(kBuddhistYear, n)
			i += n - 1
		case c == '.' && i+1 < len(rs) && rs[i+1] == '0':
			// In a date format, a point followed by zeros is a
			// fraction of a second, wherever it appears.
			n := 0
			for i+1 < len(rs) && rs[i+1] == '0' {
				n++
				i++
			}
			add(kFrac, min(n, 9))
		case c == '0' || c == '#' || c == '?':
			numeric = true
			lit.WriteRune(c)
		default:
			lit.WriteRune(c)
		}
	}
	flush()
	return section{resolveMinutes(toks), cond, numeric}, i
}

// resolveMinutes decides whether each m or mm is a month or a minute.
// m and mm are minutes directly after an hour or directly before a second;
// mmm and longer are always months.
func resolveMinutes(toks []token) []token {
	for i := range toks {
		if toks[i].kind != kMonthOrMinute {
			continue
		}
		toks[i].kind = kMonth
		if toks[i].n > 2 {
			continue
		}
		for j := i - 1; j >= 0; j-- {
			if toks[j].kind != kLit {
				if toks[j].kind == kHour || toks[j].kind == kElapsedH {
					toks[i].kind = kMinute
				}
				break
			}
		}
		for j := i + 1; j < len(toks) && toks[i].kind == kMonth; j++ {
			if toks[j].kind != kLit {
				if toks[j].kind == kSecond || toks[j].kind == kElapsedS {
					toks[i].kind = kMinute
				}
				break
			}
		}
	}
	return toks
}

// Layout is a compiled format code. It is safe for concurrent use. The zero
// Layout formats everything as the empty string.
type Layout struct {
	sections []section
	text     string // literal text of all tokens
	date     bool   // some section shows a date or time part
}

// MaxLen returns an upper bound on the length in bytes of what
// FormatSerial returns.
func (l *Layout) MaxLen() int {
	most := 0
	for _, sec := range l.sections {
		n := 0
		for _, tk := range sec.toks {
			switch tk.kind {
			case kLit, kAMPM:
				n += int(tk.hi - tk.lo)
			case kMonth, kDay:
				n += 9 // September, Wednesday
			case kFrac:
				n += 1 + int(tk.n)
			case kElapsedH, kElapsedM, kElapsedS:
				n += 1 + max(int(tk.n), 12) // sign and up to 12 digits
			default:
				n += max(int(tk.n), 5) // years up to 10542 (Buddhist), two-digit parts
			}
		}
		most = max(most, n)
	}
	return most
}

// HasDate reports whether some section of the format code shows a date or
// time part. Unlike IsDateFormat, it looks at the code's structure, so a
// code such as "[<1]0.000;yyyy" that shows numbers for some values and
// dates for others has a date part.
func (l *Layout) HasDate() bool { return l.date }

// first returns the tokens of the first section.
func (l *Layout) first() []token {
	if len(l.sections) == 0 {
		return nil
	}
	return l.sections[0].toks
}

// pick chooses the section Excel would use for value v. With conditions,
// the first section whose condition holds, or the first section without a
// condition, is used. Otherwise the third section, if any, is for zero.
func (l *Layout) pick(v float64) (section, bool) {
	if len(l.sections) == 0 {
		return section{}, true
	}
	conditional := false
	for _, sec := range l.sections {
		if sec.cond != nil {
			conditional = true
		}
	}
	if conditional {
		for _, sec := range l.sections {
			if sec.cond == nil || sec.cond.match(v) {
				return sec, true
			}
		}
		return section{}, false
	}
	if v == 0 && len(l.sections) >= 3 {
		return l.sections[2], true
	}
	return l.sections[0], true
}

// IsDate reports whether FormatSerial shows serial as a date or time: the
// value is in range and the section chosen for it has a date or time part.
func (l *Layout) IsDate(serial float64, date1904 bool) bool {
	if !l.date || !(serial >= 0 && serial < float64(maxSerial(date1904))) {
		return false
	}
	sec, ok := l.pick(serial)
	return ok && sec.isDate()
}

// Compile parses a format code for repeated use. Every string is a valid
// format code; text that is not a date part is shown literally.
func Compile(layout string) *Layout {
	rs := []rune(layout)
	cl := &closer{rs: rs}
	l := &Layout{}
	var lit strings.Builder
	defer func() { l.text = lit.String() }()
	for i := 0; ; {
		sec, next := tokenizeSection(rs, i, cl, &lit)
		l.sections = append(l.sections, sec)
		if next >= len(rs) || len(l.sections) == 4 {
			// A trailing ';' starts an empty section.
			if next == len(rs) && next > 0 && rs[next-1] == ';' && len(l.sections) < 4 {
				l.sections = append(l.sections, section{})
			}
			break
		}
		i = next
	}
	for _, sec := range l.sections {
		if sec.isDate() {
			l.date = true
		}
	}
	return l
}

const msPerDay = 24 * 60 * 60 * 1000

var epoch1900 = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)

// Format formats t with the format code layout, using its first section.
// Elapsed-time parts such as [h] count from 1899-12-30, the start of
// Excel's 1900 date system.
func Format(t time.Time, layout string) string {
	return Compile(layout).Format(t)
}

// Format formats t with the first section of l. See the package function
// Format.
func (l *Layout) Format(t time.Time) string {
	toks := l.first()
	if hasTime(toks) {
		t = t.Round(precision(toks))
	}
	y, m, d := t.Date()
	f := fields{
		year: y, month: int(m), day: d, weekday: t.Weekday(),
		hour: t.Hour(), minute: t.Minute(), second: t.Second(), nsec: t.Nanosecond(),
	}
	secs := t.Unix() - epoch1900.Unix()
	if secs > math.MaxInt64/1000 || secs < math.MinInt64/1000 {
		secs = max(min(secs, math.MaxInt64/1000), math.MinInt64/1000)
	}
	f.elapsed = secs*1000 + int64(t.Nanosecond()/1e6)
	return render(l.text, toks, f)
}

// FormatSerial formats an Excel date serial number in the 1900 or 1904 date
// system, as Excel shows it. It reports false for values Excel cannot show
// as a date (negative values, values from 10000-01-01 on, NaN and
// infinities), for a format code with no date or time part, when no
// conditional section applies, and when the section chosen for the value
// is a number format (such as "0.00" in "[<1]0.00;yyyy"), which the caller
// should show as a number.
//
// Like Excel, it shows serial number 0 as 1900-01-00 and serial number 60 as
// 1900-02-29 in the 1900 date system. It chooses the section the way Excel
// does: by bracketed conditions such as [<1] if there are any, otherwise
// the third section, if there is one, for zero. Times are rounded to the
// precision shown, such as the nearest second; a format without a time
// part shows the date the value falls on.
func (l *Layout) FormatSerial(serial float64, date1904 bool) (string, bool) {
	limit := maxSerial(date1904)
	if !l.date || !(serial >= 0 && serial < float64(limit)) {
		return "", false
	}
	sec, ok := l.pick(serial)
	if !ok || (sec.numeric && !sec.isDate()) {
		return "", false
	}
	toks := sec.toks
	var ms int64
	if hasTime(toks) {
		// Round once, to the precision shown (a second, or up to three
		// decimal places of one).
		digits := min(fracDigits(toks), 3)
		unit := int64(1000)
		scale := 86400.0
		for i := 0; i < digits; i++ {
			unit /= 10
			scale *= 10
		}
		ms = int64(math.Round(serial*scale)) * unit
	} else {
		ms = int64(math.Round(serial * msPerDay))
	}
	if last := limit*msPerDay - 1; ms > last {
		ms = last / 1000 * 1000
	}
	f := serialFields(ms, date1904)
	return render(l.text, toks, f), true
}

func maxSerial(date1904 bool) int64 {
	if date1904 {
		return 2957004 // 10000-01-01
	}
	return 2958466 // 10000-01-01
}

// SerialTime converts an Excel date serial number to a time in UTC, rounded
// to the millisecond. It reports false for the values FormatSerial rejects.
//
// In the 1900 date system, serial numbers 1 to 59 count from 1899-12-31,
// because Excel treats 1900 as a leap year. Serial number 60, Excel's
// nonexistent 1900-02-29, becomes 1900-02-28, and times below 1 (serial
// number 0, shown by Excel as 1900-01-00) fall on 1899-12-30.
func SerialTime(serial float64, date1904 bool) (time.Time, bool) {
	limit := maxSerial(date1904)
	if !(serial >= 0 && serial < float64(limit)) {
		return time.Time{}, false
	}
	ms := int64(math.Round(serial * msPerDay))
	if ms >= limit*msPerDay {
		return time.Time{}, false
	}
	days, rest := ms/msPerDay, ms%msPerDay
	var t time.Time
	switch {
	case date1904:
		t = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC).AddDate(0, 0, int(days))
	case days >= 1 && days < 60:
		t = epoch1900.AddDate(0, 0, int(days)+1)
	default:
		t = epoch1900.AddDate(0, 0, int(days))
	}
	return t.Add(time.Duration(rest) * time.Millisecond), true
}

// fields are the parts of a date and time that a format code can show.
type fields struct {
	year, month, day     int
	weekday              time.Weekday
	hour, minute, second int
	nsec                 int
	elapsed              int64 // milliseconds, for [h], [mm] and [ss]
}

// serialFields splits a serial number, in milliseconds, into fields the
// way Excel shows them.
func serialFields(ms int64, date1904 bool) fields {
	days, rest := ms/msPerDay, ms%msPerDay
	f := fields{
		hour:    int(rest / 3600000),
		minute:  int(rest / 60000 % 60),
		second:  int(rest / 1000 % 60),
		nsec:    int(rest%1000) * 1e6,
		elapsed: ms,
	}
	var t time.Time
	switch {
	case date1904:
		t = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC).AddDate(0, 0, int(days))
	case days == 0:
		f.year, f.month, f.day, f.weekday = 1900, 1, 0, time.Saturday
		return f
	case days == 60:
		f.year, f.month, f.day, f.weekday = 1900, 2, 29, time.Wednesday
		return f
	case days < 60:
		t = epoch1900.AddDate(0, 0, int(days)+1)
	default:
		t = epoch1900.AddDate(0, 0, int(days))
	}
	y, m, d := t.Date()
	f.year, f.month, f.day = y, int(m), d
	f.weekday = t.Weekday()
	if !date1904 {
		// Excel's weekdays before 1900-03-01 follow its leap-year error.
		f.weekday = time.Weekday((days + 6) % 7)
	}
	return f
}

// hasTime reports whether a section shows a time of day or elapsed time.
func hasTime(toks []token) bool {
	for _, tk := range toks {
		switch tk.kind {
		case kHour, kMinute, kSecond, kFrac, kAMPM, kElapsedH, kElapsedM, kElapsedS:
			return true
		}
	}
	return false
}

// fracDigits is the number of decimal places of a second a section shows.
func fracDigits(toks []token) int {
	digits := 0
	for _, tk := range toks {
		if tk.kind == kFrac && int(tk.n) > digits {
			digits = int(tk.n)
		}
	}
	return digits
}

// precision is the smallest unit a section shows: a second, or a fraction
// of one.
func precision(toks []token) time.Duration {
	digits := fracDigits(toks)
	p := time.Second
	for i := 0; i < digits; i++ {
		p /= 10
	}
	return p
}

func render(text string, toks []token, f fields) string {
	ampm := false
	size := 0
	for _, tk := range toks {
		switch tk.kind {
		case kAMPM:
			ampm = true
		case kLit:
			size += int(tk.hi - tk.lo)
		default:
			size += 4
		}
	}
	b := make([]byte, 0, size)
	num := func(v int64, width int32) {
		if v < 0 {
			b = append(b, '-')
			v = -v
		}
		var tmp [20]byte
		d := strconv.AppendInt(tmp[:0], v, 10)
		for i := int32(len(d)); i < width; i++ {
			b = append(b, '0')
		}
		b = append(b, d...)
	}
	signed := false
	for _, tk := range toks {
		switch tk.kind {
		case kLit:
			b = append(b, text[tk.lo:tk.hi]...)
		case kYear, kBuddhistYear:
			y := int64(f.year)
			if tk.kind == kBuddhistYear {
				y += 543
			}
			if tk.n <= 2 {
				y %= 100
				if y < 0 {
					y = -y
				}
				num(y, 2)
			} else {
				num(y, 4)
			}
		case kMonth:
			name := time.Month(f.month).String()
			switch tk.n {
			case 1, 2:
				num(int64(f.month), tk.n)
			case 3:
				b = append(b, name[:3]...)
			case 5:
				b = append(b, name[:1]...)
			default:
				b = append(b, name...)
			}
		case kDay:
			name := f.weekday.String()
			switch tk.n {
			case 1, 2:
				num(int64(f.day), tk.n)
			case 3:
				b = append(b, name[:3]...)
			default:
				b = append(b, name...)
			}
		case kHour:
			h := f.hour
			if ampm {
				h %= 12
				if h == 0 {
					h = 12
				}
			}
			num(int64(h), min(tk.n, 2))
		case kMinute:
			num(int64(f.minute), min(tk.n, 2))
		case kSecond:
			num(int64(f.second), min(tk.n, 2))
		case kFrac:
			b = append(b, '.')
			v := int64(f.nsec)
			for i := tk.n; i < 9; i++ {
				v /= 10
			}
			num(v, tk.n)
		case kAMPM:
			// Excel shows AM/PM in capitals however the code spells it,
			// and A/P as written.
			m := text[tk.lo:tk.hi]
			am := f.hour < 12
			switch {
			case len(m) == 5 && am:
				b = append(b, "AM"...)
			case len(m) == 5:
				b = append(b, "PM"...)
			case am:
				b = append(b, m[:1]...)
			default:
				b = append(b, m[2:]...)
			}
		case kElapsedH, kElapsedM, kElapsedS:
			e := f.elapsed
			if e < 0 {
				if !signed {
					b = append(b, '-')
					signed = true
				}
				e = -e
			}
			unit := int64(3600000)
			switch tk.kind {
			case kElapsedM:
				unit = 60000
			case kElapsedS:
				unit = 1000
			}
			num(e/unit, tk.n)
		}
	}
	return string(b)
}

// IsDateFormat reports whether a number format code formats dates or times.
// It uses the same test as most spreadsheet readers. It ignores quoted text,
// escaped characters, bracketed text such as colors and locales, and the
// zeros of a fraction of a second. The code is a date format if it has more
// date characters (y m d h s e) than number placeholders (0 # ?).
func IsDateFormat(code string) bool {
	rs := []rune(code)
	cl := &closer{rs: rs}
	var dates, nums int
	afterSecond := false
	for i := 0; i < len(rs); i++ {
		c := rs[i]
		wasSecond := afterSecond
		afterSecond = false
		switch c {
		case '"':
			for i++; i < len(rs) && rs[i] != '"'; i++ {
			}
		case '\\', '_', '*':
			i++
		case '[':
			j := cl.next(i + 1)
			if j >= len(rs) {
				continue // no closing bracket: an ordinary character
			}
			inner := strings.ToLower(string(rs[i+1 : j]))
			if inner != "" && strings.Trim(inner, "hms") == "" {
				dates++
				afterSecond = strings.Trim(inner, "s") == ""
			}
			i = j
		case 'G', 'g':
			if hasPrefixFold(rs[i:], "General") {
				i += 6
			}
		case 'A', 'a':
			if hasPrefixFold(rs[i:], "AM/PM") {
				dates++
				i += 4
			} else if hasPrefixFold(rs[i:], "A/P") {
				dates++
				i += 2
			}
		case 'y', 'Y', 'm', 'M', 'd', 'D', 'h', 'H', 'e', 'E':
			dates++
		case 's', 'S':
			dates++
			afterSecond = true
		case '.':
			if wasSecond {
				for i+1 < len(rs) && rs[i+1] == '0' {
					i++
				}
			}
		case '0', '#', '?':
			nums++
		}
	}
	return dates > nums
}
