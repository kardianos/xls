package ole2

import (
	"bytes"
	"encoding/binary"
	"errors"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"testing"
	"time"

	"github.com/kardianos/xls/internal/cfbtest"
	"github.com/kardianos/xls/internal/fuzzcap"
)

func pattern(n int, seed byte) []byte {
	b := make([]byte, n)
	for i := range b {
		b[i] = byte(i*7) + seed
	}
	return b
}

func openBytes(t *testing.T, b []byte) (*File, error) {
	t.Helper()
	return Open(bytes.NewReader(b))
}

func readNamed(t *testing.T, f *File, storage *Entry, name string) []byte {
	t.Helper()
	e := f.Lookup(storage, name)
	if e == nil {
		t.Fatalf("%q not found", name)
	}
	data, err := f.ReadStream(e)
	if err != nil {
		t.Fatalf("ReadStream(%q): %v", name, err)
	}
	return data
}

// onlyReadSeeker hides the io.ReaderAt method of its reader.
type onlyReadSeeker struct{ io.ReadSeeker }

func TestRoundTrip(t *testing.T) {
	streams := []cfbtest.Stream{
		{Name: "Big", Data: pattern(10000, 1)},
		{Name: "Small", Data: pattern(100, 2)},
		{Name: "Small2", Data: pattern(4095, 3)},
		{Name: "Exact", Data: pattern(4096, 4)},
		{Name: "Empty"},
		{Name: "Inner", Data: pattern(50, 5), Storage: "Sub"},
	}
	for _, opt := range []cfbtest.Options{{}, {SectorShift: 12}} {
		b, _ := cfbtest.Build(streams, opt)
		for _, rs := range []io.ReadSeeker{bytes.NewReader(b), onlyReadSeeker{bytes.NewReader(b)}} {
			f, err := Open(rs)
			if err != nil {
				t.Fatalf("%+v: Open: %v", opt, err)
			}
			root := f.Root()
			for _, s := range streams {
				if s.Storage != "" {
					continue
				}
				if got := readNamed(t, f, root, s.Name); !bytes.Equal(got, s.Data) {
					t.Errorf("%+v: stream %q differs", opt, s.Name)
				}
			}
			if f.Lookup(root, "inner") != nil {
				t.Errorf("Lookup found a grandchild of the root")
			}
			sub := f.Lookup(root, "SUB")
			if sub == nil || sub.Type != TypeStorage {
				t.Fatalf("storage not found by case-insensitive name")
			}
			if got := readNamed(t, f, sub, "Inner"); !bytes.Equal(got, streams[5].Data) {
				t.Errorf("nested stream differs")
			}
		}
	}
}

func TestManyFATSectors(t *testing.T) {
	// More than 109 FAT sectors needs a DIFAT chain of two sectors.
	big := pattern(110*128*512, 9)
	b, l := cfbtest.Build([]cfbtest.Stream{{Name: "Big", Data: big}}, cfbtest.Options{})
	if len(l.DIFSectors) == 0 {
		t.Fatal("test file has no DIFAT sectors")
	}
	f, err := openBytes(t, b)
	if err != nil {
		t.Fatal(err)
	}
	if got := readNamed(t, f, f.Root(), "Big"); !bytes.Equal(got, big) {
		t.Fatal("stream differs")
	}
}

func TestLookup(t *testing.T) {
	data := func(i byte) []byte { return pattern(100, i) }
	// U+212A KELVIN SIGN folds to k under Unicode case folding, but
	// upper-cases to itself, so [MS-CFB] treats the names as different.
	streams := []cfbtest.Stream{
		{Name: "Wor\u212Abook", Data: data(1)},
		{Name: "Workbook", Data: data(2)},
		{Name: "a", Data: data(3)},
		{Name: "LongerName", Data: data(4)},
		{Name: "Bb", Data: data(5)},
	}
	for _, unsorted := range []bool{false, true} {
		b, _ := cfbtest.Build(streams, cfbtest.Options{Unsorted: unsorted})
		f, err := openBytes(t, b)
		if err != nil {
			t.Fatal(err)
		}
		find := f.Lookup
		if unsorted {
			// Lookup follows the tree, which is out of order; Search
			// does not depend on it.
			find = f.Search
			if f.Lookup(f.Root(), "Workbook") != nil && f.Lookup(f.Root(), "a") != nil && f.Lookup(f.Root(), "Bb") != nil {
				t.Error("Lookup found every name in an unsorted tree")
			}
		}
		for _, c := range []struct {
			name string
			want byte
		}{{"Workbook", 2}, {"WORKBOOK", 2}, {"workBook", 2}, {"A", 3}, {"longername", 4}, {"bB", 5}} {
			e := find(f.Root(), c.name)
			if e == nil {
				t.Errorf("unsorted=%v: %q not found", unsorted, c.name)
				continue
			}
			got, err := f.ReadStream(e)
			if err != nil || !bytes.Equal(got, data(c.want)) {
				t.Errorf("unsorted=%v: %q found the wrong entry (%q)", unsorted, c.name, e.Name)
			}
		}
		if find(f.Root(), "Workbooks") != nil || find(f.Root(), "WorKbook") == nil {
			t.Error("bad match")
		}
	}

	// Two children with the same name in a tree that is out of order: the
	// tree search fails, and the fallback refuses to choose.
	// The first child is longer than "Workbook", so the search goes left and
	// finds nothing; the two other children are both named Workbook.
	b, l := cfbtest.Build([]cfbtest.Stream{{Name: "Zzzzzzzzzz", Data: data(1)}, {Name: "Workbook", Data: data(2)}, {Name: "Other", Data: data(3)}}, cfbtest.Options{Unsorted: true})
	nameOff := l.EntryOffset["Other"]
	for i, c := range "Workbook" {
		binary.LittleEndian.PutUint16(b[nameOff+2*i:], uint16(c))
	}
	binary.LittleEndian.PutUint16(b[nameOff+64:], 18)
	f, err := openBytes(t, b)
	if err != nil {
		t.Fatal(err)
	}
	if e := f.Lookup(f.Root(), "Workbook"); e != nil {
		t.Errorf("Lookup resolved %q in an out-of-order tree", e.Name)
	}
	if e := f.Search(f.Root(), "Workbook"); e != nil {
		t.Errorf("ambiguous name resolved to %q", e.Name)
	}
}

func TestLookupNil(t *testing.T) {
	b, _ := cfbtest.Build([]cfbtest.Stream{{Name: "A", Data: pattern(10, 1)}}, cfbtest.Options{})
	f, err := openBytes(t, b)
	if err != nil {
		t.Fatal(err)
	}
	if f.Lookup(nil, "A") != nil || f.Search(nil, "A") != nil || f.Children(nil) != nil {
		t.Error("nil entry found something")
	}
}

func TestTreeIsCanonical(t *testing.T) {
	streams := []cfbtest.Stream{{Name: "A", Data: pattern(10, 1), Storage: "S"}, {Name: "B", Data: pattern(10, 2), Storage: "T"}, {Name: "Top", Data: pattern(10, 3)}}
	b, l := cfbtest.Build(streams, cfbtest.Options{})
	// Make storage S's child point back at the root's child, and storage
	// T's child point at S: cycles across storages.
	rootChild := binary.LittleEndian.Uint32(b[l.EntryOffset["Root Entry"]+76:])
	sID := binary.LittleEndian.Uint32(b[l.EntryOffset["Root Entry"]+76:])
	binary.LittleEndian.PutUint32(b[l.EntryOffset["S"]+76:], rootChild)
	binary.LittleEndian.PutUint32(b[l.EntryOffset["T"]+76:], sID)
	f, err := openBytes(t, b)
	if err != nil {
		t.Fatal(err)
	}
	visits := map[*Entry]int{}
	var walk func(e *Entry, depth int)
	walk = func(e *Entry, depth int) {
		if depth > 100 {
			t.Fatal("walk did not terminate")
		}
		for _, c := range f.Children(e) {
			visits[c]++
			walk(c, depth+1)
		}
	}
	walk(f.Root(), 0)
	for e, n := range visits {
		if n != 1 {
			t.Errorf("%q visited %d times", e.Name, n)
		}
	}
}

func TestHeaderChecks(t *testing.T) {
	streams := []cfbtest.Stream{{Name: "Big", Data: pattern(5000, 1)}}
	for name, mutate := range map[string]func(b []byte, l cfbtest.Layout){
		"version 3 with 4096-byte sectors": func(b []byte, l cfbtest.Layout) { binary.LittleEndian.PutUint16(b[30:], 12) },
		"version 5":                        func(b []byte, l cfbtest.Layout) { binary.LittleEndian.PutUint16(b[26:], 5) },
		"mini cutoff":                      func(b []byte, l cfbtest.Layout) { binary.LittleEndian.PutUint32(b[56:], 100) },
		"FAT sector is also the directory": func(b []byte, l cfbtest.Layout) { binary.LittleEndian.PutUint32(b[48:], l.FATSectors[0]) },
	} {
		b, l := cfbtest.Build(streams, cfbtest.Options{})
		mutate(b, l)
		_, err := openBytes(t, b)
		var ce CorruptError
		if !errors.As(err, &ce) {
			t.Errorf("%s: err = %v", name, err)
		}
	}
	b, _ := cfbtest.Build(streams, cfbtest.Options{})
	_, err := openBytes(t, b[:300])
	var ce CorruptError
	if !errors.As(err, &ce) {
		t.Errorf("truncated header: err = %v", err)
	}
}

func TestNotCompound(t *testing.T) {
	for _, b := range [][]byte{nil, make([]byte, 100), make([]byte, 1024)} {
		if _, err := openBytes(t, b); !errors.Is(err, ErrNotCompoundFile) {
			t.Errorf("len %d: err = %v, want ErrNotCompoundFile", len(b), err)
		}
	}
}

func wantCorrupt(t *testing.T, err error) {
	t.Helper()
	var ce CorruptError
	if !errors.As(err, &ce) {
		t.Fatalf("err = %v, want CorruptError", err)
	}
}

func TestCorruption(t *testing.T) {
	streams := []cfbtest.Stream{{Name: "Big", Data: pattern(5000, 1)}, {Name: "Small", Data: pattern(300, 2)}}
	build := func() ([]byte, cfbtest.Layout) { return cfbtest.Build(streams, cfbtest.Options{}) }
	readAll := func(b []byte) error {
		f, err := Open(bytes.NewReader(b))
		if err != nil {
			return err
		}
		for _, e := range f.Children(f.Root()) {
			if _, err := f.ReadStream(e); err != nil {
				return err
			}
		}
		return nil
	}

	t.Run("self-loop", func(t *testing.T) {
		b, l := build()
		s := l.StreamStart["Big"]
		l.SetFAT(b, s, s)
		wantCorrupt(t, readAll(b))
	})
	t.Run("loop", func(t *testing.T) {
		b, l := build()
		s := l.StreamStart["Big"]
		l.SetFAT(b, s+1, s)
		wantCorrupt(t, readAll(b))
	})
	t.Run("out of range", func(t *testing.T) {
		b, l := build()
		l.SetFAT(b, l.StreamStart["Big"], 5000)
		wantCorrupt(t, readAll(b))
	})
	t.Run("short chain", func(t *testing.T) {
		b, l := build()
		l.SetFAT(b, l.StreamStart["Big"], 0xFFFFFFFE)
		wantCorrupt(t, readAll(b))
	})
	t.Run("free sector in chain", func(t *testing.T) {
		b, l := build()
		l.SetFAT(b, l.StreamStart["Big"], 0xFFFFFFFF)
		wantCorrupt(t, readAll(b))
	})
	t.Run("huge stream size", func(t *testing.T) {
		b, l := build()
		binary.LittleEndian.PutUint32(b[l.EntryOffset["Big"]+120:], 0xFFFFFFF0)
		wantCorrupt(t, readAll(b))
	})
	t.Run("directory loop", func(t *testing.T) {
		b, l := build()
		l.SetFAT(b, l.DirStart, l.DirStart)
		wantCorrupt(t, readAll(b))
	})
	t.Run("FAT count", func(t *testing.T) {
		b, _ := build()
		binary.LittleEndian.PutUint32(b[44:], 0xFFFFFFFF)
		wantCorrupt(t, readAll(b))
	})
	t.Run("FAT sector out of range", func(t *testing.T) {
		b, _ := build()
		binary.LittleEndian.PutUint32(b[76:], 0x7FFFFFFF)
		wantCorrupt(t, readAll(b))
	})
	t.Run("DIFAT loop", func(t *testing.T) {
		b, l := cfbtest.Build([]cfbtest.Stream{{Name: "Big", Data: pattern(110*128*512, 1)}}, cfbtest.Options{})
		d := l.DIFSectors[0]
		binary.LittleEndian.PutUint32(b[l.SectorOffset(d)+508:], d) // next DIFAT sector is itself
		binary.LittleEndian.PutUint32(b[44:], 300)                  // and more FAT sectors are needed
		wantCorrupt(t, readAll(b))
	})
	t.Run("DIFAT ends early", func(t *testing.T) {
		b, _ := build()
		binary.LittleEndian.PutUint32(b[44:], 120) // claims FAT sectors beyond the header, no DIFAT
		wantCorrupt(t, readAll(b))
	})
	t.Run("mini loop", func(t *testing.T) {
		b, l := build()
		s := l.StreamStart["Small"]
		// The mini FAT is the last sector before the directory.
		off := l.SectorOffset(l.DirStart-1) + int(s+1)*4
		binary.LittleEndian.PutUint32(b[off:], s)
		wantCorrupt(t, readAll(b))
	})
	t.Run("root not first", func(t *testing.T) {
		b, l := build()
		b[l.EntryOffset["Root Entry"]+66] = 2
		wantCorrupt(t, readAll(b))
	})
	t.Run("bad sector shift", func(t *testing.T) {
		b, _ := build()
		binary.LittleEndian.PutUint16(b[30:], 31)
		wantCorrupt(t, readAll(b))
	})
	t.Run("name length", func(t *testing.T) {
		for _, n := range []uint16{0, 1, 0xFFFF} {
			b, l := build()
			binary.LittleEndian.PutUint16(b[l.EntryOffset["Big"]+64:], n)
			if err := readAll(b); err != nil {
				t.Errorf("name length %d: %v", n, err)
			}
		}
	})
	t.Run("tree cycle", func(t *testing.T) {
		b, l := build()
		// Point Big's links back at itself and at the root; Children and
		// Lookup must still terminate.
		off := l.EntryOffset["Big"]
		binary.LittleEndian.PutUint32(b[off+72:], 1)
		binary.LittleEndian.PutUint32(b[off+68:], 0)
		f, err := openBytes(t, b)
		if err != nil {
			t.Fatal(err)
		}
		if n := len(f.Children(f.Root())); n != 2 {
			t.Errorf("got %d children", n)
		}
		if f.Lookup(f.Root(), "Missing") != nil {
			t.Error("found a missing name")
		}
	})
	t.Run("shared chain", func(t *testing.T) {
		b, l := build()
		// Point Small's directory entry at Big's sectors, with a regular size.
		off := l.EntryOffset["Small"]
		binary.LittleEndian.PutUint32(b[off+116:], l.StreamStart["Big"])
		binary.LittleEndian.PutUint32(b[off+120:], 5000)
		wantCorrupt(t, readAll(b))
	})
	t.Run("stream over directory", func(t *testing.T) {
		b, l := build()
		binary.LittleEndian.PutUint32(b[l.EntryOffset["Big"]+116:], l.DirStart)
		wantCorrupt(t, readAll(b))
	})
	t.Run("shared mini chain", func(t *testing.T) {
		b, l := cfbtest.Build([]cfbtest.Stream{{Name: "A", Data: pattern(300, 1)}, {Name: "B", Data: pattern(300, 2)}}, cfbtest.Options{})
		binary.LittleEndian.PutUint32(b[l.EntryOffset["B"]+116:], l.StreamStart["A"])
		wantCorrupt(t, readAll(b))
	})
	t.Run("bad mini stream only affects small streams", func(t *testing.T) {
		b, l := build()
		s := l.StreamStart["Small"]
		off := l.SectorOffset(l.DirStart-1) + int(s+1)*4
		binary.LittleEndian.PutUint32(b[off:], s)
		f, err := openBytes(t, b)
		if err != nil {
			t.Fatal(err)
		}
		if got := readNamed(t, f, f.Root(), "Big"); !bytes.Equal(got, streams[0].Data) {
			t.Error("Big differs")
		}
		_, err = f.ReadStream(f.Lookup(f.Root(), "Small"))
		wantCorrupt(t, err)
	})
	t.Run("foreign entry", func(t *testing.T) {
		b, _ := build()
		f, err := openBytes(t, b)
		if err != nil {
			t.Fatal(err)
		}
		for _, e := range []*Entry{nil, {Type: TypeStream, Size: -1}, {Type: TypeStream, Size: 100}} {
			if _, err := f.ReadStream(e); err == nil {
				t.Errorf("ReadStream(%+v) accepted", e)
			}
		}
		e := *f.Lookup(f.Root(), "Big")
		if _, err := f.ReadStream(&e); err == nil {
			t.Error("ReadStream accepted a copied entry")
		}
	})
	t.Run("version 4 size", func(t *testing.T) {
		b, l := cfbtest.Build(streams, cfbtest.Options{SectorShift: 12})
		binary.LittleEndian.PutUint32(b[l.EntryOffset["Big"]+124:], 1)
		wantCorrupt(t, readAll(b))
	})
	t.Run("truncated", func(t *testing.T) {
		b, _ := build()
		for _, n := range []int{512, 600, 1024, len(b) - 1000} {
			if err := readAll(b[:n]); err == nil {
				t.Errorf("truncated to %d: no error", n)
			}
		}
	})
}

func TestTestdata(t *testing.T) {
	files, _ := filepath.Glob(filepath.Join("..", "testdata", "*.xls"))
	more, _ := filepath.Glob(filepath.Join("..", "testdata", "compare", "*.xls"))
	for _, name := range append(files, more...) {
		b, err := os.ReadFile(name)
		if err != nil {
			t.Fatal(err)
		}
		f, err := openBytes(t, b)
		if err != nil {
			t.Errorf("%s: %v", name, err)
			continue
		}
		if e := f.Lookup(f.Root(), "Workbook"); e == nil {
			t.Errorf("%s: no Workbook stream", name)
		} else if _, err := f.ReadStream(e); err != nil {
			t.Errorf("%s: %v", name, err)
		}
	}
}

func FuzzOpen(f *testing.F) {
	b, _ := cfbtest.Build([]cfbtest.Stream{{Name: "Big", Data: pattern(5000, 1)}, {Name: "Small", Data: pattern(300, 2)}}, cfbtest.Options{})
	f.Add(b)
	b, _ = cfbtest.Build([]cfbtest.Stream{{Name: "A", Data: pattern(20, 1), Storage: "S"}}, cfbtest.Options{SectorShift: 12})
	f.Add(b)
	fuzzcap.Require(f)
	f.Fuzz(func(t *testing.T, b []byte) {
		var before runtime.MemStats
		runtime.ReadMemStats(&before)
		start := time.Now()
		cf, err := Open(bytes.NewReader(b))
		if err == nil {
			var walk func(e *Entry)
			walk = func(e *Entry) {
				for _, c := range cf.Children(e) {
					cf.ReadStream(c)
					cf.Lookup(e, c.Name)
					cf.Search(e, c.Name)
					walk(c)
				}
			}
			walk(cf.Root())
		}
		var after runtime.MemStats
		runtime.ReadMemStats(&after)
		if n, max := after.TotalAlloc-before.TotalAlloc, uint64(256<<10+40*len(b)); n > max {
			t.Fatalf("allocated %d bytes for a %d byte input", n, len(b))
		}
		if d := time.Since(start); d > time.Second+time.Duration(len(b))*10*time.Microsecond {
			t.Fatalf("took %v", d)
		}
	})
}
