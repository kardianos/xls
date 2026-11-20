package xls

import (
	"encoding/binary"
	"errors"
	"fmt"
	"strings"
	"unicode/utf16"

	"golang.org/x/text/encoding"
)

// BIFF record identifiers used by this package.
const (
	recFORMULA     = 0x0006
	recEOF         = 0x000A
	recDATEMODE    = 0x0022
	recFILEPASS    = 0x002F
	recCONTINUE    = 0x003C
	recCODEPAGE    = 0x0042
	recWRITEACCESS = 0x005C
	recBOUNDSHEET  = 0x0085
	recMULRK       = 0x00BD
	recMULBLANK    = 0x00BE
	recRSTRING     = 0x00D6
	recXF          = 0x00E0
	recMERGECELLS  = 0x00E5
	recSST         = 0x00FC
	recLABELSST    = 0x00FD
	recHLINK       = 0x01B8
	recBLANK       = 0x0201
	recNUMBER      = 0x0203
	recLABEL       = 0x0204
	recBOOLERR     = 0x0205
	recSTRING      = 0x0207
	recROW         = 0x0208
	recARRAY       = 0x0221
	recTABLE       = 0x0236
	recWINDOW2     = 0x023E
	recRK          = 0x027E
	recFORMAT      = 0x041E
	recSHRFMLA     = 0x04BC
	recBOF         = 0x0809

	// BOF records of BIFF2, BIFF3 and BIFF4, which are not supported.
	recBOF2 = 0x0009
	recBOF3 = 0x0209
	recBOF4 = 0x0409
)

// ErrCorrupt is wrapped by every error that reports malformed BIFF data.
// Errors about the compound file around it wrap ole2.ErrNotCompoundFile or
// are of type ole2.CorruptError.
var ErrCorrupt = errors.New("xls: corrupt workbook")

func corruptf(format string, args ...interface{}) error {
	return fmt.Errorf("%w: %s", ErrCorrupt, fmt.Sprintf(format, args...))
}

// record is a BIFF record together with the CONTINUE records that follow it.
type record struct {
	id   uint16
	pos  int      // offset of the record header in the stream
	segs [][]byte // record data, then the data of each CONTINUE record
}

// recordIter walks the records of a BIFF stream.
type recordIter struct {
	b   []byte
	off int
}

// next returns the next record. ok is false at the end of the stream.
func (it *recordIter) next() (r record, ok bool, err error) {
	if it.off == len(it.b) {
		return r, false, nil
	}
	hdr, data, err := it.raw()
	if err != nil {
		return r, false, err
	}
	r = record{id: hdr, pos: it.off - len(data) - 4, segs: [][]byte{data}}
	for len(it.b)-it.off >= 4 && binary.LittleEndian.Uint16(it.b[it.off:]) == recCONTINUE {
		_, data, err := it.raw()
		if err != nil {
			return r, false, err
		}
		r.segs = append(r.segs, data)
	}
	return r, true, nil
}

func (it *recordIter) raw() (id uint16, data []byte, err error) {
	if len(it.b)-it.off < 4 {
		return 0, nil, corruptf("truncated record header at offset %d", it.off)
	}
	id = binary.LittleEndian.Uint16(it.b[it.off:])
	n := int(binary.LittleEndian.Uint16(it.b[it.off+2:]))
	if len(it.b)-it.off-4 < n {
		return 0, nil, corruptf("record %#04x at offset %d: length %d exceeds stream", id, it.off, n)
	}
	data = it.b[it.off+4 : it.off+4+n]
	it.off += 4 + n
	return id, data, nil
}

// reader decodes the fields of one record. Reads past the end of the data
// set bad and return zero values, so a handler checks bad once at the end.
//
// The data may be split into segments by CONTINUE records. Fixed-size
// fields are read across segment boundaries as if the data were
// contiguous. Character data is different: [MS-XLS] 2.5.293 says that when
// the characters of a string continue in the next segment, that segment
// starts with a one-byte flag that gives the new character width.
type reader struct {
	segs [][]byte
	seg  int
	off  int
	left int // unread bytes in all segments
	bad  bool
}

func newReader(r record) *reader {
	n := 0
	for _, s := range r.segs {
		n += len(s)
	}
	return &reader{segs: r.segs, left: n}
}

// advance consumes k bytes of the current segment.
func (r *reader) advance(k int) {
	r.off += k
	r.left -= k
}

// cur returns the unread part of the current segment, moving past segments
// that are used up. Over a whole record, the moves take time linear in the
// number of segments.
func (r *reader) cur() []byte {
	for r.seg < len(r.segs) && r.off >= len(r.segs[r.seg]) {
		r.seg++
		r.off = 0
	}
	if r.seg >= len(r.segs) {
		return nil
	}
	return r.segs[r.seg][r.off:]
}

// remaining returns the number of unread bytes in all segments.
func (r *reader) remaining() int { return r.left }

// bytes returns the next n bytes, joining segments if needed.
func (r *reader) bytes(n int) []byte {
	if r.bad || n < 0 || n > r.left {
		r.bad = true
		return nil
	}
	if c := r.cur(); len(c) >= n {
		r.advance(n)
		return c[:n:n]
	}
	out := make([]byte, 0, n)
	for len(out) < n {
		c := r.cur()
		k := min(n-len(out), len(c))
		out = append(out, c[:k]...)
		r.advance(k)
	}
	return out
}

func (r *reader) skip(n int) {
	if r.bad || n < 0 || n > r.left {
		r.bad = true
		return
	}
	for n > 0 {
		k := min(n, len(r.cur()))
		r.advance(k)
		n -= k
	}
}

func (r *reader) u8() uint8 {
	if b := r.bytes(1); b != nil {
		return b[0]
	}
	return 0
}

func (r *reader) u16() uint16 {
	if b := r.bytes(2); b != nil {
		return binary.LittleEndian.Uint16(b)
	}
	return 0
}

func (r *reader) u32() uint32 {
	if b := r.bytes(4); b != nil {
		return binary.LittleEndian.Uint32(b)
	}
	return 0
}

func (r *reader) u64() uint64 {
	if b := r.bytes(8); b != nil {
		return binary.LittleEndian.Uint64(b)
	}
	return 0
}

// chars reads cch characters. high selects 16-bit characters; otherwise
// characters are 8-bit and hold the low byte of a UTF-16 code unit.
func (r *reader) chars(cch int, high bool) string {
	if r.bad || cch > r.left {
		// Every character takes at least one byte, so this checks the
		// length before allocating for it.
		r.bad = true
		return ""
	}
	u := make([]uint16, 0, cch)
	for need := cch; need > 0; {
		if r.seg < len(r.segs) && r.off >= len(r.segs[r.seg]) {
			// The string continues in the next segment, which starts
			// with a flags byte.
			r.seg++
			r.off = 0
			if r.seg >= len(r.segs) || len(r.segs[r.seg]) == 0 {
				r.bad = true
				return ""
			}
			high = r.segs[r.seg][0]&0x01 != 0
			r.advance(1)
			continue
		}
		if r.seg >= len(r.segs) {
			r.bad = true
			return ""
		}
		c := r.segs[r.seg][r.off:]
		if high {
			k := min(len(c)/2, need)
			if k == 0 {
				r.bad = true // a character split across segments
				return ""
			}
			for i := 0; i < k; i++ {
				u = append(u, binary.LittleEndian.Uint16(c[2*i:]))
			}
			r.advance(2 * k)
			need -= k
		} else {
			k := min(len(c), need)
			for i := 0; i < k; i++ {
				u = append(u, uint16(c[i]))
			}
			r.advance(k)
			need -= k
		}
	}
	return string(utf16.Decode(u))
}

// unicodeString reads a BIFF8 XLUnicodeString ([MS-XLS] 2.5.294) whose
// character count is lenBytes (1 or 2) wide.
func (r *reader) unicodeString(lenBytes int) string {
	var cch int
	if lenBytes == 1 {
		cch = int(r.u8())
	} else {
		cch = int(r.u16())
	}
	flags := r.u8()
	return r.chars(cch, flags&0x01 != 0)
}

// richString reads a BIFF8 XLUnicodeRichExtendedString ([MS-XLS] 2.5.293)
// and discards its formatting runs and phonetic data.
func (r *reader) richString() string {
	cch := int(r.u16())
	flags := r.u8()
	var runs, ext int
	if flags&0x08 != 0 {
		runs = int(r.u16())
	}
	if flags&0x04 != 0 {
		ext = int(r.u32())
	}
	s := r.chars(cch, flags&0x01 != 0)
	r.skip(4 * runs)
	r.skip(ext)
	return s
}

// byteString reads a BIFF5 string of lenBytes-wide length and 8-bit
// characters in the workbook code page.
func (r *reader) byteString(lenBytes int, enc encoding.Encoding) string {
	var n int
	if lenBytes == 1 {
		n = int(r.u8())
	} else {
		n = int(r.u16())
	}
	return decode8(r.bytes(n), enc)
}

func decode8(b []byte, enc encoding.Encoding) string {
	if len(b) == 0 {
		return ""
	}
	out, err := enc.NewDecoder().Bytes(b)
	if err != nil {
		return strings.ToValidUTF8(string(b), "\uFFFD")
	}
	return string(out)
}
