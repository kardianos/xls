package xls

import (
	"bytes"
	"errors"
	"os"
	"path/filepath"
	"runtime"
	"testing"
	"time"

	"github.com/kardianos/xls/internal/fuzzcap"
)

// FuzzOpen feeds arbitrary files to the reader. Run it with ./fuzz.sh, which
// caps memory. The target refuses to fuzz without a cap.
func FuzzOpen(f *testing.F) {
	files, _ := filepath.Glob(filepath.Join("testdata", "*.xls"))
	more, _ := filepath.Glob(filepath.Join("testdata", "compare", "*.xls"))
	real, _ := filepath.Glob(filepath.Join("testdata", "real", "*.xls"))
	for _, name := range append(append(files, more...), real...) {
		if b, err := os.ReadFile(name); err == nil && len(b) < 64<<10 {
			f.Add(b)
		}
	}
	f.Add(book{globals: [][]byte{xfRec(0), formatRec(164, "yyyy-mm-dd hh:mm"), xfRec(164), sstRec("a", "b")},
		sheets: []sheet{{name: "S", recs: [][]byte{
			labelSSTRec(0, 0, 0, 1), numberRec(0, 1, 1, 43831.5), formulaRec(0, 2, 0, specialVal(0, 0)), rec(recSTRING, ustr("s")),
			urlHyperlink(0, 0, "https://example.com/", "x"),
			rec(recMERGECELLS, le(uint16(1), uint16(1), uint16(2), uint16(0), uint16(1))),
		}}}}.file())
	f.Add(book{biff: 5, globals: [][]byte{rec(recCODEPAGE, le(uint16(1251)))}, sheets: []sheet{{name: "S"}}}.file())
	fuzzcap.Require(f)
	// A small ReadAll budget, so the allocation bound below can be tight.
	defer func(old func(int) int64) { readAllBudget = old }(readAllBudget)
	readAllBudget = func(n int) int64 { return 1<<20 + 64*int64(n) }
	f.Fuzz(func(t *testing.T, b []byte) {
		// Memory must stay within a fixed multiple of the input size. The
		// allowance covers ReadAll's grid budget (64 MiB + 64x) and parsing.
		var before runtime.MemStats
		runtime.ReadMemStats(&before)
		defer func() {
			var after runtime.MemStats
			runtime.ReadMemStats(&after)
			if n, max := after.TotalAlloc-before.TotalAlloc, uint64(8<<20+300*len(b)); n > max {
				t.Fatalf("allocated %d bytes for a %d byte input", n, len(b))
			}
		}()
		start := time.Now()
		defer func() {
			// Parsing and reading are linear in the input; even a slow,
			// busy machine handles a fuzz input in well under this.
			if d := time.Since(start); d > time.Second+time.Duration(len(b))*20*time.Microsecond {
				t.Fatalf("took %v for a %d byte input", d, len(b))
			}
		}()
		wb, err := OpenReader(bytes.NewReader(b), "")
		if errors.Is(err, ErrInternal) {
			t.Fatal(err)
		}
		if err != nil {
			return
		}
		for i := 0; i < wb.NumSheets(); i++ {
			s, err := wb.GetSheet(i)
			if errors.Is(err, ErrInternal) {
				t.Fatal(err)
			}
			if err != nil {
				continue
			}
			for _, row := range s.Rows() {
				for c := row.FirstCol(); c < row.LastCol(); c++ {
					row.Col(c)
					row.Value(c)
					row.Hyperlink(c)
				}
				for _, c := range row.Cols() {
					row.ColExact(c)
				}
			}
		}
		wb.ReadAll(0)
	})
}
