package xls

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// TestCompareCSV compares each testdata/compare/*.xls file with the CSV
// export next to it. Files named skip_* are skipped.
func TestCompareCSV(t *testing.T) {
	root := filepath.Join("testdata", "compare")
	files, err := os.ReadDir(root)
	if err != nil {
		t.Fatalf("Cant read testdata directory contents: %s", err)
	}
	for _, f := range files {
		fn := f.Name()
		if filepath.Ext(fn) != ".xls" {
			continue
		}
		t.Run(fn, func(t *testing.T) {
			if strings.HasPrefix(fn, "skip_") {
				t.Skip("skipping compare")
			}
			csvName := strings.TrimSuffix(fn, filepath.Ext(fn)) + ".csv"
			if err := compareXLX(filepath.Join(root, fn), filepath.Join(root, csvName)); err != nil {
				t.Fatalf("XLS file %s and CSV file are not equal: %v", fn, err)
			}
		})
	}
}
