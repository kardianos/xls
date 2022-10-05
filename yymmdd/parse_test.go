package yymmdd

import (
	"testing"
	"time"
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
}

func TestFormat(t *testing.T) {
	for i, testCase := range tests {
		_, tokens := lexLayout(testCase.format)
		ds := parse(tokens)
		actual := ds.Format(date)

		if actual != testCase.expected {
			t.Errorf("Case index %d failed. Expected: %s, Got: %s", i, testCase.expected, actual)
		}
	}
}

func TestParse(t *testing.T) {
	for i, testCase := range tests {
		_, tokens := lexLayout(testCase.format)
		ds := parse(tokens)
		actual, err := ds.Parse(testCase.expected)

		if err != nil || actual.UnixNano() != testCase.parsed.UnixNano() {
			t.Errorf("Case index %d failed. Expected: %s, Got: %s", i, testCase.expected, actual)
		}
	}
}
