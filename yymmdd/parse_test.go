package yymmdd

import (
	"math"
	"strings"
	"testing"
	"time"

	"github.com/kardianos/xls/internal/fuzzcap"
)

var date = time.Date(2018, time.March, 15, 12, 0, 0, 0, time.UTC)

func ymd(y int, m time.Month, d int) time.Time {
	return time.Date(y, m, d, 0, 0, 0, 0, time.UTC)
}

var tests = []struct {
	format   string
	expected string
	parsed   time.Time
}{
	{"yy-", "18-", ymd(2018, 1, 1)},
	{"yyyymm", "201803", ymd(2018, 3, 1)},
	{"yyyy-mm", "2018-03", ymd(2018, 3, 1)},
	{"mmm yyyy", "Mar 2018", ymd(2018, 3, 1)},
	{"yyyy-mm-dd", "2018-03-15", ymd(2018, 3, 15)},
	{"DD/MM/YY", "15/03/18", ymd(2018, 3, 15)},
	{"mmmm d, yyyy", "March 15, 2018", ymd(2018, 3, 15)},
	{"yyyy-mm-dd hh:mm:ss", "2018-03-15 12:00:00", date},
	{"h:mm AM/PM", "12:00 PM", time.Date(0, 1, 1, 12, 0, 0, 0, time.UTC)},
}

func TestFormat(t *testing.T) {
	for i, tc := range tests {
		if got := Format(date, tc.format); got != tc.expected {
			t.Errorf("case %d %q: got %q, want %q", i, tc.format, got, tc.expected)
		}
	}
}

func TestParse(t *testing.T) {
	for i, tc := range tests {
		got, err := Parse(tc.expected, tc.format)
		if err != nil || !got.Equal(tc.parsed) {
			t.Errorf("case %d %q: got %v, %v; want %v", i, tc.format, got, err, tc.parsed)
		}
	}
}

func TestFormatCodes(t *testing.T) {
	tm := time.Date(2018, 3, 5, 14, 7, 9, 250e6, time.UTC)
	cases := []struct{ layout, want string }{
		{"hh:mm:ss", "14:07:09"},
		{"HH:MM:SS", "14:07:09"},
		{"h:mm AM/PM", "2:07 PM"},
		{"h:mm am/pm", "2:07 PM"},
		{"h mmm", "14 Mar"},
		{"h mmmm s", "14 March 9"},
		{"h:mm A/P", "2:07 P"},
		{"dddd, mmmm d, yyyy", "Monday, March 5, 2018"},
		{"ddd mmm", "Mon Mar"},
		{"mmmmm", "M"},
		{"m/d/yy h:mm", "3/5/18 14:07"},
		{`yyyy"Jan"`, "2018Jan"},
		{`yyyy\J\a\n`, "2018Jan"},
		{"[$-409]d-mmm-yy", "5-Mar-18"},
		{"[$€-407] yyyy", "€ 2018"},
		{"[Red]yyyy;[Blue]mm", "2018"},
		{"mm:ss.0", "07:09.3"},
		{"hh:mm:ss.000", "14:07:09.250"},
		{"h:mm:ss", "14:07:09"},
		{`H:"12";@`, "14:12"},
		{"yyyy年m月d日", "2018年3月5日"},
		{"_(yyyy_)", " 2018 "},
		{"* yyyy", "2018"},
		{"e", "2018"},
		{"y", "18"},
		{"yyy", "2018"},
		{"Jan 2006 15:04 yyyy", "Jan 2006 15:04 2018"}, // Go layout words are literal text
	}
	for _, c := range cases {
		if got := Format(tm, c.layout); got != c.want {
			t.Errorf("Format(%q) = %q, want %q", c.layout, got, c.want)
		}
	}
}

func TestFormatRounding(t *testing.T) {
	tm := time.Date(1899, 12, 30, 23, 59, 59, 600e6, time.UTC)
	if got, want := Format(tm, "hh:mm:ss"), "00:00:00"; got != want {
		t.Errorf("got %q, want %q", got, want)
	}
	if got, want := Format(tm, "hh:mm:ss.0"), "23:59:59.6"; got != want {
		t.Errorf("got %q, want %q", got, want)
	}
}

func TestFormatSerial(t *testing.T) {
	cases := []struct {
		serial float64
		d1904  bool
		layout string
		want   string
	}{
		{1.5208333333333333, false, "[h]:mm:ss", "36:30:00"},
		{1.5208333333333333, false, "[hh]:mm", "36:30"},
		{1.5208333333333333, false, "[mm]:ss", "2190:00"},
		{1.5208333333333333, false, "[ss]", "131400"},
		{109575, false, "[h]:mm:ss", "2629800:00:00"},
		{109575, false, "[ss]", "9467280000"},
		{2958465.5, false, "[h]:mm:ss", "71003172:00:00"},
		{2958465.5, false, "yyyy-mm-dd hh:mm", "9999-12-31 12:00"},
		{2958465.99999999, false, "yyyy-mm-dd hh:mm:ss", "9999-12-31 23:59:59"},
		{0, false, "yyyy-mm-dd dddd", "1900-01-00 Saturday"},
		{60, false, "yyyy-mm-dd", "1900-02-29"},
		{59, false, "yyyy-mm-dd", "1900-02-28"},
		{61, false, "yyyy-mm-dd dddd", "1900-03-01 Thursday"},
		{1, false, "yyyy-mm-dd dddd", "1900-01-01 Sunday"},
		{43831.75, false, "yyyy-mm-dd h:mm AM/PM ddd", "2020-01-01 6:00 PM Wed"},
		{0.5630208333333334, false, "hh:mm:ss", "13:30:45"},
		{0.9999884259259259, false, "hh:mm:ss", "23:59:59"},
		{0.99999999, false, "hh:mm:ss.000", "23:59:59.999"},
		{0.99999999, false, "yyyy-mm-dd hh:mm:ss", "1900-01-01 00:00:00"},
		{0, false, "yyyy;;", ""},
		{0, false, `yyyy;;"zero"`, "zero"},
		{5, false, `yyyy;;"zero"`, "1900"},
		{0, true, "yyyy-mm-dd", "1904-01-01"},
		{100.5, true, "yyyy-mm-dd hh:mm", "1904-04-10 12:00"},
	}
	for _, c := range cases {
		got, ok := Compile(c.layout).FormatSerial(c.serial, c.d1904)
		if !ok || got != c.want {
			t.Errorf("FormatSerial(%v, %v, %q) = %q, %v; want %q", c.serial, c.d1904, c.layout, got, ok, c.want)
		}
	}
	for _, v := range []float64{-1, 2958466, math.NaN(), math.Inf(1)} {
		if _, ok := Compile("yyyy").FormatSerial(v, false); ok {
			t.Errorf("FormatSerial(%v) accepted", v)
		}
	}
	tm := time.Date(1899, 12, 31, 12, 30, 15, 0, time.UTC)
	if got := Format(tm, "[h]:mm"); got != "36:30" {
		t.Errorf("Format([h]:mm) = %q", got)
	}
	if got := Format(time.Date(2300, 1, 1, 0, 0, 0, 0, time.UTC), "[h]"); got != "3506376" {
		t.Errorf("Format(2300, [h]) = %q", got)
	}
}

func TestSerialTime(t *testing.T) {
	d := func(y int, m time.Month, day, h, min, s, ms int) time.Time {
		return time.Date(y, m, day, h, min, s, ms*1e6, time.UTC)
	}
	cases := []struct {
		serial float64
		d1904  bool
		want   time.Time
		ok     bool
	}{
		{0, false, d(1899, 12, 30, 0, 0, 0, 0), true},
		{0.5630208333333334, false, d(1899, 12, 30, 13, 30, 45, 0), true},
		{0.00011574074074074075, false, d(1899, 12, 30, 0, 0, 10, 0), true},
		{0.99999999, false, d(1899, 12, 30, 23, 59, 59, 999), true},
		{0.999999999, false, d(1900, 1, 1, 0, 0, 0, 0), true},
		{1, false, d(1900, 1, 1, 0, 0, 0, 0), true},
		{59, false, d(1900, 2, 28, 0, 0, 0, 0), true},
		{60, false, d(1900, 2, 28, 0, 0, 0, 0), true},
		{61, false, d(1900, 3, 1, 0, 0, 0, 0), true},
		{109575, false, d(2200, 1, 1, 0, 0, 0, 0), true},
		{2958465.999988426, false, d(9999, 12, 31, 23, 59, 59, 0), true},
		{2958465.999999995, false, time.Time{}, false},
		{2958466, false, time.Time{}, false},
		{-1, false, time.Time{}, false},
		{math.NaN(), false, time.Time{}, false},
		{100.5, true, d(1904, 4, 10, 12, 0, 0, 0), true},
		{2957004, true, time.Time{}, false},
	}
	for _, c := range cases {
		got, ok := SerialTime(c.serial, c.d1904)
		if ok != c.ok || !got.Equal(c.want) {
			t.Errorf("SerialTime(%v, %v) = %v, %v; want %v, %v", c.serial, c.d1904, got, ok, c.want, c.ok)
		}
	}
}

func TestIsDateFormat(t *testing.T) {
	cases := map[string]bool{
		"General": false, "@": false, "0.00": false, "0.00E+00": false, "##0.0E+0": false,
		`"$"#,##0.00_);[Red]\("$"#,##0.00\)`: false, `#,##0 "USD"`: false, "0%": false,
		"yyyy-mm-dd": true, "DD/MM/YY": true, "HH:MM:SS": true, `H:"12";@`: true,
		"[h]:mm:ss": true, "mm:ss.0": true, "ss.000": true, "[ss].000": true, "e": true,
		"[$-409]d-mmm-yy": true, `0 \d`: false, "[Red]0": false,
		"d" + strings.Repeat("[", 1000): true,
	}
	for code, want := range cases {
		if got := IsDateFormat(code); got != want {
			t.Errorf("IsDateFormat(%q) = %v, want %v", code, got, want)
		}
	}
}

func TestLinearTime(t *testing.T) {
	// Unclosed brackets used to make tokenizing quadratic.
	for _, code := range []string{"d" + strings.Repeat("[", 200000), strings.Repeat("[h", 100000), strings.Repeat(`"`, 100000)} {
		start := time.Now()
		l := Compile(code)
		l.FormatSerial(43831.5, false)
		IsDateFormat(code)
		if d := time.Since(start); d > slow*time.Second {
			t.Errorf("format of length %d took %v", len(code), d)
		}
	}
}

func TestParseErrors(t *testing.T) {
	cases := []struct{ value, layout string }{
		{"2018-02-30", "yyyy-mm-dd"},
		{"2018-13-01", "yyyy-mm-dd"},
		{"2018-03-15x", "yyyy-mm-dd"},
		{"25:00", "hh:mm"},
		{"13:00 PM", "h:mm AM/PM"},
		{"Foo 2018", "mmm yyyy"},
		{"", "yyyy"},
	}
	for _, c := range cases {
		if got, err := Parse(c.value, c.layout); err == nil {
			t.Errorf("Parse(%q, %q) = %v, want error", c.value, c.layout, got)
		}
	}
}

func FuzzFormat(f *testing.F) {
	for _, tc := range tests {
		f.Add(tc.format, tc.expected)
	}
	f.Add(`[h]:mm:ss.000;"x"`, "1:02:03")
	f.Add(`[$€-407]"`, "")
	f.Add("d[[[[[;yyyy;;", "")
	fuzzcap.Require(f)
	f.Fuzz(func(t *testing.T, layout, value string) {
		if len(layout) > 4096 {
			return
		}
		l := Compile(layout)
		l.Format(date)
		l.Format(time.Date(-5000, 1, 1, 0, 0, 0, 0, time.UTC))
		for _, v := range []float64{0, 0.5, 60, 43831.123456, 2958465.9999} {
			l.FormatSerial(v, false)
			l.FormatSerial(v, true)
		}
		IsDateFormat(layout)
		Parse(value, layout)
	})
}

func TestSectionsAndEdges(t *testing.T) {
	cases := []struct {
		layout string
		serial float64
		want   string
		ok     bool
	}{
		{"[<1]h:mm;yyyy-mm-dd", 0.75, "18:00", true},
		{"[<1]h:mm;yyyy-mm-dd", 43831.75, "2020-01-01", true},
		{"[>=43831]yyyy;[<10]h", 43831, "2020", true},
		{"[>=43831]yyyy;[<10]h", 100, "", false}, // no section applies
		{"yyyy-mm-dd", 43831.999999, "2020-01-01", true},
		{"yyyy-mm-dd hh:mm:ss", 43831.999999, "2020-01-02 00:00:00", true},
		{"General", 43831, "", false},
		{`"text only"`, 43831, "", false},
	}
	for _, c := range cases {
		got, ok := Compile(c.layout).FormatSerial(c.serial, false)
		if got != c.want || ok != c.ok {
			t.Errorf("FormatSerial(%q, %v) = %q, %v; want %q, %v", c.layout, c.serial, got, ok, c.want, c.ok)
		}
	}
	var zero Layout
	if got := zero.Format(date); got != "" {
		t.Errorf("zero Layout: %q", got)
	}
	if _, ok := zero.FormatSerial(1, false); ok {
		t.Error("zero Layout formatted a serial")
	}
	if got := Format(time.Date(-5, 3, 1, 0, 0, 0, 0, time.UTC), "yyyy yy"); got != "-0005 05" {
		t.Errorf("negative year: %q", got)
	}
}

func TestParseElapsedAndExcelDates(t *testing.T) {
	for _, c := range []struct{ value, layout string }{
		{"36:30", "[h]:mm"}, {"5000:00", "[mm]:ss"}, {"90000", "[ss]"}, {"999999999:00", "[h]:mm"},
	} {
		tm, err := Parse(c.value, c.layout)
		if err != nil {
			t.Errorf("Parse(%q, %q): %v", c.value, c.layout, err)
			continue
		}
		if got := Format(tm, c.layout); got != c.value {
			t.Errorf("round trip %q via %q gave %q", c.value, c.layout, got)
		}
	}
	for value, want := range map[string]time.Time{
		"1900-01-00": time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC),
		"1900-02-29": time.Date(1900, 2, 28, 0, 0, 0, 0, time.UTC),
	} {
		got, err := Parse(value, "yyyy-mm-dd")
		if err != nil || !got.Equal(want) {
			t.Errorf("Parse(%q) = %v, %v", value, got, err)
		}
	}
	if _, err := Parse("2020-01-00", "yyyy-mm-dd"); err == nil {
		t.Error("day 0 accepted outside 1900-01")
	}
}

func TestRoundingAndSections(t *testing.T) {
	cases := []struct {
		layout string
		serial float64
		want   string
		ok     bool
	}{
		{"ss.00", 0.942257, "51.00", true}, // rounded once: 51.0048 s
		{"hh:mm:ss", (43200 + 0.4996) / 86400, "12:00:00", true},
		{"yyyy-mm-dd hh:mm:ss", 43831 + (86399+0.4996)/86400, "2020-01-01 23:59:59", true},
		{"hh .0", 0.5 + 0.8/86400, "12 .8", true},
		{"General;yyyy-mm-dd", 5, "", false},
		{"[>=1]0;h:mm:ss", 5, "", false},
		{"[<1]0.000;yyyy-mm-dd", 0.5, "", false},
		{"[<1]0.000;yyyy-mm-dd", 5, "1900-01-05", true},
		{"h:mm A/P", 0.75, "6:00 P", true},
		{"h:mm a/p", 0.25, "6:00 a", true},
	}
	for _, c := range cases {
		got, ok := Compile(c.layout).FormatSerial(c.serial, false)
		if got != c.want || ok != c.ok {
			t.Errorf("FormatSerial(%q, %v) = %q, %v; want %q, %v", c.layout, c.serial, got, ok, c.want, c.ok)
		}
		if isDate := Compile(c.layout).IsDate(c.serial, false); isDate != c.ok {
			t.Errorf("IsDate(%q, %v) = %v", c.layout, c.serial, isDate)
		}
	}
	for _, code := range []string{"A/P", "h AM/PM", "am/pm"} {
		if !IsDateFormat(code) {
			t.Errorf("IsDateFormat(%q) = false", code)
		}
	}
}

func TestRound4Codes(t *testing.T) {
	cases := map[string]string{
		"mmmmmm":   "March",
		"mmmmm":    "M",
		"bbbb-mm":  "2561-03",
		"bb":       "61",
		"yyyy E+0": "2018 E+0",
	}
	for code, want := range cases {
		if got := Format(date, code); got != want {
			t.Errorf("Format(%q) = %q, want %q", code, got, want)
		}
	}
	if Compile("0.00E+00").HasDate() || !Compile("[<1]0.0000;yyyy").HasDate() || Compile("General").HasDate() {
		t.Error("HasDate")
	}
	if got, err := Parse("2561-03", "bbbb-mm"); err != nil || got.Year() != 2018 {
		t.Errorf("Parse Buddhist year = %v, %v", got, err)
	}
}

func TestMaxLen(t *testing.T) {
	for _, code := range []string{"yyyy-mm-dd", "dddd, mmmm d, yyyy h:mm:ss.000 AM/PM", "[h]:mm;[mm];[ss]", strings.Repeat("dddd", 60), "bbbb e", `"text"`} {
		l := Compile(code)
		for _, v := range []float64{0, 0.5, 60, 2958465.99, 1e6 + 0.123456} {
			if s, ok := l.FormatSerial(v, false); ok && len(s) > l.MaxLen() {
				t.Errorf("%q at %v: %d bytes > MaxLen %d", code, v, len(s), l.MaxLen())
			}
		}
	}
}
