package yymmdd

import (
	"fmt"
	"strings"
	"time"
)

// Parse parses value with the first section of the Excel number format code
// layout and returns the time in UTC. Parts missing from the layout default
// to year 0, January, day 1 and midnight. Two-digit years 00-29 are
// 2000-2029 and 30-99 are 1930-1999, as in Excel. Excel's 1900-01-00 and
// 1900-02-29 parse as SerialTime maps them. A value with elapsed-time parts
// ([h], [mm], [ss]) is that span after 1899-12-30, so Format gives it back.
func Parse(value, layout string) (time.Time, error) {
	year, month, day := 0, 1, 1
	hour, minute, sec, nsec := 0, 0, 0, 0
	pm, hasAMPM, elapsedHours := false, false, false
	s := value
	fail := func(what string) (time.Time, error) {
		return time.Time{}, fmt.Errorf("yymmdd: cannot parse %q as %q: bad %s", value, layout, what)
	}
	var v int
	var ok bool
	elapsed := false
	l := Compile(layout)
	for _, tk := range l.first() {
		switch tk.kind {
		case kLit:
			lit := l.text[tk.lo:tk.hi]
			if len(s) < len(lit) || !strings.EqualFold(s[:len(lit)], lit) {
				return fail("literal text")
			}
			s = s[len(lit):]
		case kBuddhistYear:
			if tk.n <= 2 {
				return fail("Buddhist year") // two digits do not give the century
			}
			if year, s, ok = digits(s, 4, 4); !ok {
				return fail("year")
			}
			year -= 543
		case kYear:
			if tk.n <= 2 {
				if v, s, ok = digits(s, 2, 2); !ok {
					return fail("year")
				}
				year = 1900 + v
				if v < 30 {
					year = 2000 + v
				}
			} else {
				if year, s, ok = digits(s, 4, 4); !ok {
					return fail("year")
				}
			}
		case kMonth:
			switch {
			case tk.n <= 2:
				month, s, ok = digits(s, 1, 2)
			case tk.n == 5:
				ok = false // a single letter is ambiguous
			default:
				month, s, ok = name(s, tk.n == 3, func(i int) string { return time.Month(i).String() }, 1, 12)
			}
			if !ok || month < 1 || month > 12 {
				return fail("month")
			}
		case kDay:
			if tk.n <= 2 {
				if day, s, ok = digits(s, 1, 2); !ok || day > 31 {
					return fail("day")
				}
			} else if _, s, ok = name(s, tk.n == 3, func(i int) string { return time.Weekday(i).String() }, 0, 6); !ok {
				return fail("weekday")
			}
		case kHour:
			if hour, s, ok = digits(s, 1, 2); !ok {
				return fail("hour")
			}
		case kMinute:
			if minute, s, ok = digits(s, 1, 2); !ok || minute > 59 {
				return fail("minute")
			}
		case kSecond:
			if sec, s, ok = digits(s, 1, 2); !ok || sec > 59 {
				return fail("second")
			}
		case kFrac:
			if v, s, ok = digits(s, int(tk.n), int(tk.n)); !ok {
				return fail("fraction")
			}
			nsec = v
			for i := tk.n; i < 9; i++ {
				nsec *= 10
			}
		case kAMPM:
			hasAMPM = true
			long := tk.hi-tk.lo == 5
			switch {
			case long && len(s) >= 2 && strings.EqualFold(s[:2], "AM"):
				s = s[2:]
			case long && len(s) >= 2 && strings.EqualFold(s[:2], "PM"):
				pm, s = true, s[2:]
			case !long && len(s) >= 1 && (s[0] == 'A' || s[0] == 'a'):
				s = s[1:]
			case !long && len(s) >= 1 && (s[0] == 'P' || s[0] == 'p'):
				pm, s = true, s[1:]
			default:
				return fail("AM/PM")
			}
		case kElapsedH:
			if hour, s, ok = digits(s, 1, 9); !ok {
				return fail("elapsed hours")
			}
			elapsed, elapsedHours = true, true
		case kElapsedM:
			if minute, s, ok = digits(s, 1, 9); !ok {
				return fail("elapsed minutes")
			}
			elapsed = true
		case kElapsedS:
			if sec, s, ok = digits(s, 1, 9); !ok {
				return fail("elapsed seconds")
			}
			elapsed = true
		}
	}
	if s != "" {
		return fail("trailing text")
	}
	if hasAMPM {
		if hour < 1 || hour > 12 {
			return fail("hour")
		}
		hour %= 12
		if pm {
			hour += 12
		}
	} else if hour > 23 && !elapsedHours {
		return fail("hour")
	}
	if elapsed {
		// Elapsed time counts from 1899-12-30, as in Format.
		total := int64(hour)*3600 + int64(minute)*60 + int64(sec)
		return epoch1900.AddDate(0, 0, int(total/86400)).Add(time.Duration(total%86400)*time.Second + time.Duration(nsec)), nil
	}
	switch {
	case year == 1900 && month == 1 && day == 0:
		// Excel shows serial number 0 as 1900-01-00; see SerialTime.
		year, month, day = 1899, 12, 30
	case year == 1900 && month == 2 && day == 29:
		// Excel's 1900-02-29, serial number 60; see SerialTime.
		day = 28
	case day == 0:
		return fail("day")
	}
	t := time.Date(year, time.Month(month), day, hour, minute, sec, nsec, time.UTC)
	if t.Day() != day {
		return fail("day") // such as February 30
	}
	return t, nil
}

// digits reads between lo and hi ASCII digits.
func digits(s string, lo, hi int) (int, string, bool) {
	n, v := 0, 0
	for n < len(s) && n < hi && s[n] >= '0' && s[n] <= '9' {
		v = v*10 + int(s[n]-'0')
		n++
	}
	if n < lo {
		return 0, s, false
	}
	return v, s[n:], true
}

// name reads an English month or weekday name, abbreviated to three letters
// if short, and returns its index.
func name(s string, short bool, names func(int) string, lo, hi int) (int, string, bool) {
	for i := lo; i <= hi; i++ {
		n := names(i)
		if short {
			n = n[:3]
		}
		if len(s) >= len(n) && strings.EqualFold(s[:len(n)], n) {
			return i, s[len(n):], true
		}
	}
	return 0, s, false
}
