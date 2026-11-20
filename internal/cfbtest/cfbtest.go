// Package cfbtest builds compound files for tests.
package cfbtest

import (
	"encoding/binary"
	"sort"
	"unicode"
	"unicode/utf16"
)

// Less orders directory entry names as [MS-CFB] 2.6.4 does: shorter names
// first, then by upper-cased UTF-16 code units.
func Less(a, b string) bool {
	x, y := utf16.Encode([]rune(a)), utf16.Encode([]rune(b))
	if len(x) != len(y) {
		return len(x) < len(y)
	}
	for i := range x {
		p, q := unicode.ToUpper(rune(x[i])), unicode.ToUpper(rune(y[i]))
		if p != q {
			return p < q
		}
	}
	return false
}

const (
	endOfChain = 0xFFFFFFFE
	freeSect   = 0xFFFFFFFF
	fatSect    = 0xFFFFFFFD
	difSect    = 0xFFFFFFFC
	noStream   = 0xFFFFFFFF
)

// Stream is a stream to store. A non-empty Storage puts it inside a storage
// of that name under the root.
type Stream struct {
	Name    string
	Data    []byte
	Storage string
}

// Options control the layout.
type Options struct {
	SectorShift int  // 9 (default) or 12
	Unsorted    bool // link siblings in input order instead of a sorted tree
}

// Layout describes where Build put things, so tests can damage them.
type Layout struct {
	SectorSize  int
	FATSectors  []uint32
	DIFSectors  []uint32
	DirStart    uint32
	StreamStart map[string]uint32 // first sector (regular) or mini sector
	EntryOffset map[string]int    // file offset of each directory entry
}

// SectorOffset returns the file offset of sector sid.
func (l Layout) SectorOffset(sid uint32) int { return (int(sid) + 1) * l.SectorSize }

// SetFAT changes FAT entry sid in file b.
func (l Layout) SetFAT(b []byte, sid, val uint32) {
	per := uint32(l.SectorSize / 4)
	off := l.SectorOffset(l.FATSectors[sid/per]) + int(sid%per)*4
	binary.LittleEndian.PutUint32(b[off:], val)
}

// Build returns a compound file holding streams.
func Build(streams []Stream, o Options) ([]byte, Layout) {
	shift := o.SectorShift
	if shift == 0 {
		shift = 9
	}
	ss := 1 << shift
	per := ss / 4
	l := Layout{SectorSize: ss, StreamStart: map[string]uint32{}, EntryOffset: map[string]int{}}

	var sectors [][]byte // regular sector contents
	var fat []uint32
	alloc := func(data []byte) uint32 {
		if len(data) == 0 {
			return endOfChain
		}
		start := uint32(len(sectors))
		for off := 0; off < len(data); off += ss {
			sec := make([]byte, ss)
			copy(sec, data[off:])
			sectors = append(sectors, sec)
			fat = append(fat, uint32(len(sectors)))
		}
		fat[len(fat)-1] = endOfChain
		return start
	}

	// Mini stream.
	var mini []byte
	var miniFAT []uint32
	starts := make([]uint32, len(streams))
	for i, s := range streams {
		if len(s.Data) >= 4096 || len(s.Data) == 0 {
			continue
		}
		starts[i] = uint32(len(mini) / 64)
		n := (len(s.Data) + 63) / 64
		for k := 0; k < n; k++ {
			miniFAT = append(miniFAT, uint32(len(miniFAT)+1))
		}
		miniFAT[len(miniFAT)-1] = endOfChain
		buf := make([]byte, n*64)
		copy(buf, s.Data)
		mini = append(mini, buf...)
	}
	for i, s := range streams {
		switch {
		case len(s.Data) == 0:
			starts[i] = endOfChain
		case len(s.Data) >= 4096:
			starts[i] = alloc(s.Data)
		}
		l.StreamStart[s.Name] = starts[i]
	}
	miniStart := alloc(mini)
	var miniFATBytes []byte
	for _, v := range miniFAT {
		miniFATBytes = binary.LittleEndian.AppendUint32(miniFATBytes, v)
	}
	for len(miniFATBytes)%ss != 0 {
		miniFATBytes = binary.LittleEndian.AppendUint32(miniFATBytes, freeSect)
	}
	miniFATStart := alloc(miniFATBytes)

	// Directory: root, then storages, then streams. Siblings form a
	// balanced binary search tree in [MS-CFB] order (or, with Unsorted, a
	// right-linked list in input order).
	type ent struct {
		name        string
		typ         byte
		start, size uint32
		left, right uint32
		child       uint32
	}
	ents := []ent{{name: "Root Entry", typ: 5, start: miniStart, size: uint32(len(mini)), left: noStream, right: noStream, child: noStream}}
	storageID := map[string]int{}
	var storageOrder []string
	for _, s := range streams {
		if s.Storage != "" && storageID[s.Storage] == 0 {
			storageID[s.Storage] = len(ents)
			storageOrder = append(storageOrder, s.Storage)
			ents = append(ents, ent{name: s.Storage, typ: 1, left: noStream, right: noStream, child: noStream})
		}
	}
	children := map[int][]int{}
	for _, name := range storageOrder {
		children[0] = append(children[0], storageID[name])
	}
	for i, s := range streams {
		id := len(ents)
		ents = append(ents, ent{name: s.Name, typ: 2, start: starts[i], size: uint32(len(s.Data)), left: noStream, right: noStream, child: noStream})
		parent := 0
		if s.Storage != "" {
			parent = storageID[s.Storage]
		}
		children[parent] = append(children[parent], id)
	}
	var build func(ids []int) uint32
	build = func(ids []int) uint32 {
		if len(ids) == 0 {
			return noStream
		}
		if o.Unsorted {
			for k := 0; k+1 < len(ids); k++ {
				ents[ids[k]].right = uint32(ids[k+1])
			}
			return uint32(ids[0])
		}
		m := len(ids) / 2
		ents[ids[m]].left = build(ids[:m])
		ents[ids[m]].right = build(ids[m+1:])
		return uint32(ids[m])
	}
	for parent, ids := range children {
		if !o.Unsorted {
			sort.SliceStable(ids, func(a, b int) bool { return Less(ents[ids[a]].name, ents[ids[b]].name) })
		}
		ents[parent].child = build(ids)
	}
	dir := make([]byte, ((len(ents)*128+ss-1)/ss)*ss)
	for i, e := range ents {
		b := dir[i*128:]
		u := utf16.Encode([]rune(e.name))
		for k, c := range u {
			binary.LittleEndian.PutUint16(b[2*k:], c)
		}
		binary.LittleEndian.PutUint16(b[64:], uint16(2*(len(u)+1)))
		b[66] = e.typ
		b[67] = 1 // black
		binary.LittleEndian.PutUint32(b[68:], e.left)
		binary.LittleEndian.PutUint32(b[72:], e.right)
		binary.LittleEndian.PutUint32(b[76:], e.child)
		binary.LittleEndian.PutUint32(b[116:], e.start)
		binary.LittleEndian.PutUint32(b[120:], e.size)
	}
	l.DirStart = alloc(dir)
	for i, e := range ents {
		l.EntryOffset[e.name] = l.SectorOffset(l.DirStart) + i*128 // dir sectors are contiguous
	}

	// FAT and DIFAT sectors, sized so the FAT covers itself.
	nOther := len(sectors)
	nFAT, nDIF := 1, 0
	for {
		inHeader := 109
		dif := 0
		if nFAT > inHeader {
			dif = (nFAT - inHeader + per - 2) / (per - 1)
		}
		need := (nOther + nFAT + dif + per - 1) / per
		if need <= nFAT && dif == nDIF {
			nDIF = dif
			break
		}
		if need > nFAT {
			nFAT = need
		}
		nDIF = dif
	}
	for i := 0; i < nFAT; i++ {
		l.FATSectors = append(l.FATSectors, uint32(nOther+i))
		fat = append(fat, fatSect)
	}
	for i := 0; i < nDIF; i++ {
		l.DIFSectors = append(l.DIFSectors, uint32(nOther+nFAT+i))
		fat = append(fat, difSect)
	}
	for len(fat) < nFAT*per {
		fat = append(fat, freeSect)
	}
	for i := 0; i < nFAT; i++ {
		sec := make([]byte, ss)
		for k := 0; k < per; k++ {
			binary.LittleEndian.PutUint32(sec[4*k:], fat[i*per+k])
		}
		sectors = append(sectors, sec)
	}
	var headerList, difList []uint32
	if nFAT > 109 {
		headerList, difList = l.FATSectors[:109], l.FATSectors[109:]
	} else {
		headerList = l.FATSectors
	}
	for i := 0; i < nDIF; i++ {
		sec := make([]byte, ss)
		for k := 0; k < per-1; k++ {
			v := uint32(freeSect)
			if j := i*(per-1) + k; j < len(difList) {
				v = difList[j]
			}
			binary.LittleEndian.PutUint32(sec[4*k:], v)
		}
		next := uint32(endOfChain)
		if i+1 < nDIF {
			next = l.DIFSectors[i+1]
		}
		binary.LittleEndian.PutUint32(sec[ss-4:], next)
		sectors = append(sectors, sec)
	}

	h := make([]byte, ss)
	copy(h, []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1})
	binary.LittleEndian.PutUint16(h[24:], 0x3E)
	major := uint16(3)
	if shift == 12 {
		major = 4
		binary.LittleEndian.PutUint32(h[40:], uint32(len(dir)/ss))
	}
	binary.LittleEndian.PutUint16(h[26:], major)
	binary.LittleEndian.PutUint16(h[28:], 0xFFFE)
	binary.LittleEndian.PutUint16(h[30:], uint16(shift))
	binary.LittleEndian.PutUint16(h[32:], 6)
	binary.LittleEndian.PutUint32(h[44:], uint32(nFAT))
	binary.LittleEndian.PutUint32(h[48:], l.DirStart)
	binary.LittleEndian.PutUint32(h[56:], 4096)
	binary.LittleEndian.PutUint32(h[60:], miniFATStart)
	binary.LittleEndian.PutUint32(h[64:], uint32(len(miniFATBytes)/ss))
	difStart := uint32(endOfChain)
	if nDIF > 0 {
		difStart = l.DIFSectors[0]
	}
	binary.LittleEndian.PutUint32(h[68:], difStart)
	binary.LittleEndian.PutUint32(h[72:], uint32(nDIF))
	for i := 0; i < 109; i++ {
		v := uint32(freeSect)
		if i < len(headerList) {
			v = headerList[i]
		}
		binary.LittleEndian.PutUint32(h[76+4*i:], v)
	}

	out := h
	for _, s := range sectors {
		out = append(out, s...)
	}
	return out, l
}
