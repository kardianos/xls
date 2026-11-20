// Package fuzzcap stops fuzzing that would run without a memory cap.
//
// A parser bug found by fuzzing can allocate memory without limit. Without a
// cap that can exhaust the machine and take other programs down with it. Run
// fuzz targets through fuzz.sh in the repository root, which applies a cgroup
// or address-space limit.
package fuzzcap

import (
	"flag"
	"os"
	"testing"
)

// maxCap is the largest limit accepted as a cap.
const maxCap = 16 << 30

// Require fails f when it is fuzzing (go test -fuzz) and the process has no
// memory limit of at most min(16 GiB, half of physical memory). Running the
// seed corpus with plain go test is not affected. Set
// XLS_FUZZ_UNCAPPED=1 to override.
func Require(f *testing.F) {
	f.Helper()
	fl := flag.Lookup("test.fuzz")
	if fl == nil || fl.Value.String() == "" || os.Getenv("XLS_FUZZ_UNCAPPED") == "1" {
		return
	}
	want := uint64(maxCap)
	if total, ok := physicalMemory(); ok && total/2 < want {
		want = total / 2
	}
	if limit, ok := memoryLimit(); ok && limit <= want {
		return
	}
	f.Fatalf("refusing to fuzz without a memory cap of at most %d MiB; run ./fuzz.sh (see README) or set XLS_FUZZ_UNCAPPED=1", want>>20)
}
