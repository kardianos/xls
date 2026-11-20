package yymmdd

import (
	"encoding/json"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// TestSSF compares FormatSerial and IsDateFormat with the Excel-derived test
// tables of SheetJS SSF (https://github.com/SheetJS/ssf, test/*.tsv and
// test/is_date.json). Set SSF_DIR to a directory holding any of times.tsv,
// dates.tsv and is_date.json to run it. Values Excel shows as "####" (out of
// range) are skipped: FormatSerial reports them as not dates.
func TestSSF(t *testing.T) {
	dir := os.Getenv("SSF_DIR")
	if dir == "" {
		t.Skip("set SSF_DIR to run")
	}
	tsv, _ := filepath.Glob(filepath.Join(dir, "*.tsv"))
	if _, err := os.Stat(filepath.Join(dir, "is_date.json")); len(tsv) == 0 && err != nil {
		t.Fatalf("SSF_DIR=%q has no SSF test files (a relative path is resolved from the package directory)", dir)
	}
	for _, name := range tsv {
		b, err := os.ReadFile(name)
		if err != nil {
			t.Fatal(err)
		}
		lines := strings.Split(strings.TrimRight(string(b), "\n"), "\n")
		hdr := strings.Split(lines[0], "\t")
		checks, bad := 0, 0
		for _, ln := range lines[1:] {
			f := strings.Split(ln, "\t")
			v, err := strconv.ParseFloat(f[0], 64)
			if err != nil || len(f) != len(hdr) {
				continue
			}
			for i := 1; i < len(hdr); i++ {
				if strings.HasPrefix(f[i], "#") {
					continue
				}
				checks++
				if got, _ := Compile(hdr[i]).FormatSerial(v, false); got != f[i] {
					if bad++; bad <= 5 {
						t.Errorf("%s: FormatSerial(%v, %q) = %q, Excel %q", filepath.Base(name), v, hdr[i], got, f[i])
					}
				}
			}
		}
		t.Logf("%s: %d checks, %d differ", filepath.Base(name), checks, bad)
	}
	if b, err := os.ReadFile(filepath.Join(dir, "is_date.json")); err == nil {
		var cases [][2]interface{}
		if err := json.Unmarshal(b, &cases); err != nil {
			t.Fatal(err)
		}
		for _, c := range cases {
			code, _ := c[0].(string)
			want, _ := c[1].(bool)
			if got := IsDateFormat(code); got != want {
				t.Errorf("IsDateFormat(%q) = %v, SSF %v", code, got, want)
			}
		}
	}
}
