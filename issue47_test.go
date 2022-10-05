package xls

import (
	"io/ioutil"
	"path"
	"path/filepath"
	"strings"
	"testing"
)

func TestIssue47(t *testing.T) {
	root := path.Join("testdata", "compare")
	files, err := ioutil.ReadDir(root)
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
			xlsName := fn
			csvName := strings.TrimSuffix(xlsName, filepath.Ext(xlsName)) + ".csv"
			err := compareXLX(
				path.Join(root, xlsName),
				path.Join(root, csvName),
			)
			if err != nil {
				t.Fatalf("XLS file %s an CSV file are not equal: %v", xlsName, err)
			}
		})
	}

}
