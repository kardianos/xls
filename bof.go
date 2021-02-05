package xls

import (
	"encoding/binary"
	"io"
	"unicode/utf16"
)

// bof Is the unit of Information in an XLS file.
type bof struct {
	ID   uint16
	Size uint16
}

// Read as UTF-16.
func (b *bof) utf16String(buf io.ReadSeeker, count uint32) string {
	var bts = make([]uint16, count)
	binary.Read(buf, binary.LittleEndian, &bts)
	runes := utf16.Decode(bts[:len(bts)-1])
	return string(runes)
}

type biffHeader struct {
	Ver    uint16
	Type   uint16
	IDMake uint16
	Year   uint16
	Flags  uint32
	MinVer uint32
}
