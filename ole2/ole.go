// Package ole2 reads Microsoft Compound File Binary files (also called OLE2 or
// structured storage), as specified in [MS-CFB].
//
// The reader is meant for untrusted input. Every sector reference is bounds
// checked, sector chains are checked for cycles, and the memory used is bounded
// by the size of the file.
//
// [MS-CFB]: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
package ole2

import (
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"unicode"
	"unicode/utf16"
)

// ErrNotCompoundFile is returned by Open when the input does not start with a
// compound file header.
var ErrNotCompoundFile = errors.New("ole2: not a compound file")

// CorruptError reports a structurally invalid compound file.
type CorruptError string

func (e CorruptError) Error() string { return "ole2: corrupt compound file: " + string(e) }

func corrupt(format string, args ...interface{}) error {
	return CorruptError(fmt.Sprintf(format, args...))
}

const (
	maxRegSect = 0xFFFFFFFA
	endOfChain = 0xFFFFFFFE
	freeSect   = 0xFFFFFFFF
	noStream   = 0xFFFFFFFF

	headerSize       = 512
	dirEntrySize     = 128
	miniSectorSize   = 64
	miniStreamCutoff = 4096 // [MS-CFB] 2.2: MUST be 4096
)

var signature = [8]byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

// EntryType is the object type of a directory entry.
type EntryType uint8

// Directory entry object types.
const (
	TypeUnused  EntryType = 0
	TypeStorage EntryType = 1
	TypeStream  EntryType = 2
	TypeRoot    EntryType = 5
)

// Entry is a directory entry: a storage, a stream or the root storage.
type Entry struct {
	Name string
	Type EntryType
	Size int64

	id                 uint32
	name16             []uint16
	start              uint32
	left, right, child uint32
}

// File is an open compound file. It is not safe for concurrent use.
//
// The memory a File uses, including every stream read from it, is bounded
// by a small multiple of the size the reader reports: no sector can belong
// to two streams. (A reader that reports a size far larger than the data
// it holds, such as a sparse file, can make Open allocate in proportion to
// the reported size.)
type File struct {
	r          io.ReaderAt
	major      uint16
	sectorSize int
	numSectors uint32
	fat        []uint32
	entries    []Entry
	kids       [][]uint32 // children of each storage, from buildTree

	seen     bitset // scratch for chain, one bit per sector
	miniSeen bitset // scratch for chain, one bit per mini sector

	// owner records which stream each sector belongs to (entry id + 1),
	// or one of the owner constants for the file's own structures, so no
	// sector is read as part of two of them.
	owner     []int32
	miniOwner []int32

	miniFATStart uint32
	miniLoaded   bool
	miniErr      error
	miniFAT      []uint32
	mini         []byte
}

// Owners of the sectors that hold the file's own structures.
const (
	ownerDIFAT     = -1
	ownerFAT       = -2
	ownerDirectory = -3
	ownerMiniFAT   = -4
)

// claim marks sectors as belonging to owner. A sector that already belongs
// to another stream or to the file's own structures makes the file corrupt.
func claim(table []int32, sids []uint32, owner int32) error {
	for _, sid := range sids {
		if o := table[sid]; o != 0 && o != owner {
			return corrupt("sector %d is used by two streams", sid)
		}
	}
	for _, sid := range sids {
		table[sid] = owner
	}
	return nil
}

type seekReaderAt struct{ rs io.ReadSeeker }

func (s seekReaderAt) ReadAt(p []byte, off int64) (int, error) {
	if _, err := s.rs.Seek(off, io.SeekStart); err != nil {
		return 0, err
	}
	return io.ReadFull(s.rs, p)
}

// Open reads the header, allocation tables and directory of a compound file.
// Stream contents are read by ReadStream.
func Open(r io.ReadSeeker) (*File, error) {
	size, err := r.Seek(0, io.SeekEnd)
	if err != nil {
		return nil, err
	}
	ra, ok := r.(io.ReaderAt)
	if !ok {
		ra = seekReaderAt{r}
	}
	return open(ra, size)
}

func le16(b []byte) uint16 { return binary.LittleEndian.Uint16(b) }
func le32(b []byte) uint32 { return binary.LittleEndian.Uint32(b) }

func open(r io.ReaderAt, size int64) (*File, error) {
	if size < headerSize {
		var sig [8]byte
		if n, _ := r.ReadAt(sig[:], 0); n == len(sig) && sig == signature {
			return nil, corrupt("file of %d bytes is shorter than the header", size)
		}
		return nil, ErrNotCompoundFile
	}
	var h [headerSize]byte
	if n, err := r.ReadAt(h[:], 0); n != len(h) {
		if err == nil || err == io.EOF {
			err = io.ErrUnexpectedEOF
		}
		return nil, err
	}
	if [8]byte(h[:8]) != signature {
		return nil, ErrNotCompoundFile
	}
	if le16(h[28:]) != 0xFFFE {
		return nil, corrupt("invalid byte order mark %#x", le16(h[28:]))
	}
	f := &File{r: r, major: le16(h[26:])}
	switch shift := le16(h[30:]); {
	case f.major == 3 && shift == 9:
		f.sectorSize = 512
	case f.major == 4 && shift == 12:
		f.sectorSize = 4096
	default:
		return nil, corrupt("unsupported version %d with sector shift %d", f.major, shift)
	}
	if shift := le16(h[32:]); shift != 6 {
		return nil, corrupt("unsupported mini sector shift %d", shift)
	}
	if cutoff := le32(h[56:]); cutoff != miniStreamCutoff {
		return nil, corrupt("mini stream cutoff %d, not %d", cutoff, miniStreamCutoff)
	}
	ss := int64(f.sectorSize)
	if size > ss {
		n := (size - ss + ss - 1) / ss
		if n > maxRegSect {
			n = maxRegSect
		}
		f.numSectors = uint32(n)
	}
	f.owner = make([]int32, f.numSectors)
	f.seen = newBitset(int(f.numSectors))

	if err := f.readFAT(h[:]); err != nil {
		return nil, err
	}
	if err := f.readDirectory(le32(h[48:])); err != nil {
		return nil, err
	}
	f.buildTree()
	f.miniFATStart = le32(h[60:])
	return f, nil
}

// readSector reads sector sid into dst (len(dst) <= sector size). A final
// sector cut short by the end of the file is zero filled.
func (f *File) readSector(sid uint32, dst []byte) error {
	if sid >= f.numSectors {
		return corrupt("sector %d out of range (%d sectors)", sid, f.numSectors)
	}
	off := (int64(sid) + 1) * int64(f.sectorSize)
	n, err := f.r.ReadAt(dst, off)
	if n == len(dst) {
		return nil
	}
	if err == io.EOF || err == io.ErrUnexpectedEOF {
		for i := n; i < len(dst); i++ {
			dst[i] = 0
		}
		return nil
	}
	if err == nil {
		err = io.ErrUnexpectedEOF
	}
	return err
}

type bitset []uint64

func newBitset(n int) bitset       { return make(bitset, (n+63)/64) }
func (b bitset) has(i uint32) bool { return b[i/64]&(1<<(i%64)) != 0 }
func (b bitset) set(i uint32)      { b[i/64] |= 1 << (i % 64) }
func (b bitset) clear(i uint32)    { b[i/64] &^= 1 << (i % 64) }

func (f *File) readFAT(h []byte) error {
	numFAT := le32(h[44:])
	if numFAT == 0 || numFAT > f.numSectors {
		return corrupt("FAT sector count %d invalid for a file of %d sectors", numFAT, f.numSectors)
	}
	fatSectors := make([]uint32, 0, numFAT)
	for i := 0; i < 109 && uint32(len(fatSectors)) < numFAT; i++ {
		fatSectors = append(fatSectors, le32(h[76+4*i:]))
	}
	buf := make([]byte, f.sectorSize)
	perDIF := f.sectorSize/4 - 1
	seen := newBitset(int(f.numSectors))
	for next := le32(h[68:]); uint32(len(fatSectors)) < numFAT; {
		if next >= f.numSectors {
			return corrupt("DIFAT chain ends after %d of %d FAT sectors", len(fatSectors), numFAT)
		}
		if seen.has(next) {
			return corrupt("cycle in DIFAT chain at sector %d", next)
		}
		seen.set(next)
		if err := claim(f.owner, []uint32{next}, ownerDIFAT); err != nil {
			return err
		}
		if err := f.readSector(next, buf); err != nil {
			return err
		}
		for j := 0; j < perDIF && uint32(len(fatSectors)) < numFAT; j++ {
			fatSectors = append(fatSectors, le32(buf[4*j:]))
		}
		next = le32(buf[f.sectorSize-4:])
	}

	perFAT := f.sectorSize / 4
	f.fat = make([]uint32, 0, int(numFAT)*perFAT)
	for _, sid := range fatSectors {
		if err := f.readSector(sid, buf); err != nil {
			return err
		}
		if err := claim(f.owner, []uint32{sid}, ownerFAT); err != nil {
			return err
		}
		for j := 0; j < perFAT; j++ {
			f.fat = append(f.fat, le32(buf[4*j:]))
		}
	}
	return nil
}

// chain follows a sector chain through table starting at start. Sector ids
// must be below limit. When want >= 0, it returns the first want sectors and
// reports an error if the chain is shorter; when want < 0 it returns the
// whole chain. Cycles are reported as errors. seen must have at least limit
// bits, all clear; chain clears the bits it sets before returning, so the
// cost is linear in the chain length.
func chain(start uint32, table []uint32, limit uint32, want int, seen bitset) (out []uint32, err error) {
	if want == 0 {
		return nil, nil
	}
	if limit > uint32(len(table)) {
		limit = uint32(len(table))
	}
	if want > int(limit) {
		return nil, corrupt("stream needs %d sectors but only %d exist", want, limit)
	}
	if want > 0 {
		out = make([]uint32, 0, want)
	}
	defer func() {
		for _, sid := range out {
			seen.clear(sid)
		}
		if err != nil {
			out = nil
		}
	}()
	for sid := start; ; {
		if want > 0 && len(out) == want {
			return out, nil
		}
		if sid == endOfChain {
			if want > 0 {
				return out, corrupt("sector chain has %d sectors, stream needs %d", len(out), want)
			}
			return out, nil
		}
		if sid >= limit {
			return out, corrupt("sector id %#x out of range", sid)
		}
		if seen.has(sid) {
			return out, corrupt("cycle in sector chain at sector %d", sid)
		}
		seen.set(sid)
		out = append(out, sid)
		sid = table[sid]
	}
}

func (f *File) readDirectory(start uint32) error {
	sids, err := chain(start, f.fat, f.numSectors, -1, f.seen)
	if err != nil {
		return err
	}
	if len(sids) == 0 {
		return corrupt("empty directory")
	}
	if err := claim(f.owner, sids, ownerDirectory); err != nil {
		return err
	}
	buf := make([]byte, f.sectorSize)
	f.entries = make([]Entry, 0, len(sids)*f.sectorSize/dirEntrySize)
	for _, sid := range sids {
		if err := f.readSector(sid, buf); err != nil {
			return err
		}
		for off := 0; off+dirEntrySize <= len(buf); off += dirEntrySize {
			e := f.parseEntry(buf[off : off+dirEntrySize])
			e.id = uint32(len(f.entries))
			f.entries = append(f.entries, e)
		}
	}
	if f.entries[0].Type != TypeRoot {
		return corrupt("first directory entry is not the root storage")
	}
	return nil
}

func (f *File) parseEntry(b []byte) Entry {
	e := Entry{
		Type:  EntryType(b[66]),
		left:  le32(b[68:]),
		right: le32(b[72:]),
		child: le32(b[76:]),
		start: le32(b[116:]),
	}
	if e.Type == TypeUnused {
		return e
	}
	// The name length counts bytes including the terminating NUL.
	n := int(le16(b[64:]))/2 - 1
	if n < 0 {
		n = 0
	}
	if n > 31 {
		n = 31
	}
	name := make([]uint16, 0, n)
	for i := 0; i < n; i++ {
		c := le16(b[2*i:])
		if c == 0 {
			break
		}
		name = append(name, c)
	}
	e.name16 = name
	e.Name = string(utf16.Decode(name))
	// Version 3 files only define the low 32 bits of the size ([MS-CFB] 2.6.3).
	e.Size = int64(le32(b[120:]))
	if f.major >= 4 {
		if hi := le32(b[124:]); hi >= 0x80000000 {
			e.Size = -1
		} else {
			e.Size |= int64(hi) << 32
		}
	}
	return e
}

// loadMini reads the mini FAT and the mini stream, the first time a stream
// stored in the mini stream is read.
func (f *File) loadMini() error {
	if f.miniLoaded {
		return f.miniErr
	}
	f.miniLoaded = true
	f.miniErr = func() error {
		sids, err := chain(f.miniFATStart, f.fat, f.numSectors, -1, f.seen)
		if err != nil {
			return fmt.Errorf("mini FAT: %w", err)
		}
		if err := claim(f.owner, sids, ownerMiniFAT); err != nil {
			return err
		}
		buf := make([]byte, f.sectorSize)
		per := f.sectorSize / 4
		f.miniFAT = make([]uint32, 0, len(sids)*per)
		for _, sid := range sids {
			if err := f.readSector(sid, buf); err != nil {
				return err
			}
			for j := 0; j < per; j++ {
				f.miniFAT = append(f.miniFAT, le32(buf[4*j:]))
			}
		}
		root := &f.entries[0]
		if root.Size == 0 || len(f.miniFAT) == 0 {
			return nil
		}
		f.mini, err = f.readRegular(root)
		if err != nil {
			return fmt.Errorf("mini stream: %w", err)
		}
		f.miniOwner = make([]int32, (len(f.mini)+miniSectorSize-1)/miniSectorSize)
		f.miniSeen = newBitset(len(f.miniOwner))
		return nil
	}()
	return f.miniErr
}

func (f *File) readRegular(e *Entry) ([]byte, error) {
	ss := int64(f.sectorSize)
	size := e.Size
	if size < 0 || size > int64(f.numSectors)*ss {
		return nil, corrupt("stream size %d exceeds file size", size)
	}
	sids, err := chain(e.start, f.fat, f.numSectors, int((size+ss-1)/ss), f.seen)
	if err != nil {
		return nil, err
	}
	if err := claim(f.owner, sids, int32(e.id)+1); err != nil {
		return nil, err
	}
	buf := make([]byte, size)
	for i, sid := range sids {
		lo := int64(i) * ss
		hi := lo + ss
		if hi > size {
			hi = size
		}
		if err := f.readSector(sid, buf[lo:hi]); err != nil {
			return nil, err
		}
	}
	return buf, nil
}

// Root returns the root storage entry.
func (f *File) Root() *Entry { return &f.entries[0] }

// buildTree assigns each directory entry to one parent storage, walking
// the tree from the root. An entry reachable from more than one place, or
// through a cycle, is kept only where it is first found, so the result is a
// tree. It takes time linear in the number of entries.
func (f *File) buildTree() {
	f.kids = make([][]uint32, len(f.entries))
	placed := newBitset(len(f.entries))
	placed.set(0)
	queue := []uint32{0}
	for len(queue) > 0 {
		parent := queue[0]
		queue = queue[1:]
		stack := []uint32{f.entries[parent].child}
		for len(stack) > 0 {
			id := stack[len(stack)-1]
			stack = stack[:len(stack)-1]
			if id >= uint32(len(f.entries)) || placed.has(id) {
				continue
			}
			placed.set(id)
			c := &f.entries[id]
			if c.Type == TypeUnused {
				continue
			}
			f.kids[parent] = append(f.kids[parent], id)
			if c.Type == TypeStorage {
				queue = append(queue, id)
			}
			stack = append(stack, c.right, c.left)
		}
	}
}

// Children returns the entries directly contained in storage e. Every
// entry belongs to at most one storage, so a recursive walk from the root
// visits each entry once.
func (f *File) Children(e *Entry) []*Entry {
	if e == nil || int(e.id) >= len(f.entries) || e != &f.entries[e.id] {
		return nil
	}
	out := make([]*Entry, 0, len(f.kids[e.id]))
	for _, id := range f.kids[e.id] {
		out = append(out, &f.entries[id])
	}
	return out
}

// compareNames orders directory entry names as [MS-CFB] 2.6.4 does:
// shorter names first, then by upper-cased UTF-16 code units.
func compareNames(a, b []uint16) int {
	if len(a) != len(b) {
		return len(a) - len(b)
	}
	for i := range a {
		x, y := upper(a[i]), upper(b[i])
		if x != y {
			return int(x) - int(y)
		}
	}
	return 0
}

func upper(c uint16) rune {
	if c >= 0xD800 && c < 0xE000 {
		return rune(c) // surrogates compare as they are
	}
	return unicode.ToUpper(rune(c))
}

// Lookup returns the child of storage e named name, or nil if there is
// none. It searches the red-black tree of children as [MS-CFB] 2.6.4
// describes: names compare by length, then by upper-cased code units. This
// is the lookup that other readers, such as Excel, perform. A file whose
// tree is out of order may hide a child from it; see Search.
func (f *File) Lookup(e *Entry, name string) *Entry {
	if e == nil || (e.Type != TypeStorage && e.Type != TypeRoot) {
		return nil
	}
	key := utf16.Encode([]rune(name))
	id := e.child
	for steps := 0; steps < len(f.entries) && id != noStream && id != 0 && id < uint32(len(f.entries)); steps++ {
		c := &f.entries[id]
		if c.Type == TypeUnused {
			return nil
		}
		switch cmp := compareNames(key, c.name16); {
		case cmp == 0:
			return c
		case cmp < 0:
			id = c.left
		default:
			id = c.right
		}
	}
	return nil
}

// Search returns the only child of storage e named name, comparing names as
// Lookup does but without relying on the order of the tree. It returns nil
// if no child or more than one child has the name.
func (f *File) Search(e *Entry, name string) *Entry {
	key := utf16.Encode([]rune(name))
	var found *Entry
	for _, c := range f.Children(e) {
		if compareNames(key, c.name16) == 0 {
			if found != nil {
				return nil
			}
			found = c
		}
	}
	return found
}

// ReadStream returns the contents of stream e, which must be an entry of
// f. A stream that shares sectors with another stream, or with the file's
// allocation tables or directory, is reported as corrupt.
func (f *File) ReadStream(e *Entry) ([]byte, error) {
	if e == nil || int(e.id) >= len(f.entries) || e != &f.entries[e.id] {
		return nil, errors.New("ole2: entry is not from this file")
	}
	if e.Type != TypeStream {
		return nil, fmt.Errorf("ole2: %q is not a stream", e.Name)
	}
	if e.Size < 0 {
		return nil, corrupt("stream %q has an invalid size", e.Name)
	}
	if e.Size >= miniStreamCutoff {
		return f.readRegular(e)
	}
	if e.Size == 0 {
		return []byte{}, nil
	}
	if err := f.loadMini(); err != nil {
		return nil, err
	}
	numMini := uint32(len(f.miniOwner))
	sids, err := chain(e.start, f.miniFAT, numMini, int((e.Size+miniSectorSize-1)/miniSectorSize), f.miniSeen)
	if err != nil {
		return nil, fmt.Errorf("mini stream %q: %w", e.Name, err)
	}
	if err := claim(f.miniOwner, sids, int32(e.id)+1); err != nil {
		return nil, err
	}
	buf := make([]byte, e.Size)
	for i, sid := range sids {
		off := int(sid) * miniSectorSize
		end := off + miniSectorSize
		if end > len(f.mini) {
			end = len(f.mini)
		}
		copy(buf[i*miniSectorSize:], f.mini[off:end])
	}
	return buf, nil
}
