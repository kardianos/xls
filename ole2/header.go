package ole2

import (
	"bytes"
	"encoding/binary"
	"fmt"
)

type Header struct {
	ID        [2]uint32
	Clid      [4]uint32
	Verminor  uint16
	Verdll    uint16
	ByteOrder uint16
	Lsectorb  uint16
	Lssectorb uint16
	_         uint16
	_         uint64

	Cfat     uint32 //Total number of sectors used for the sector allocation table
	Dirstart uint32 //SecID of first sector of the directory stream

	_ uint32

	Sectorcutoff uint32 //Minimum size of a standard stream
	Sfatstart    uint32 //SecID of first sector of the short-sector allocation table
	Csfat        uint32 //Total number of sectors used for the short-sector allocation table
	Difstart     uint32 //SecID of first sector of the master sector allocation table
	Cdif         uint32 //Total number of sectors used for the master sector allocation table
	Msat         [109]uint32
}

func parseHeader(bts []byte) (*Header, error) {
	buf := bytes.NewBuffer(bts)
	header := new(Header)
	binary.Read(buf, binary.LittleEndian, header)
	if header.ID[0] != 0xE011CFD0 || header.ID[1] != 0xE11AB1A1 || header.ByteOrder != 0xFFFE {
		return nil, fmt.Errorf("not an excel file: Header[0]=%X, Header[1]=%X, ByteOrder=%X", header.ID[0], header.ID[1], header.ByteOrder)
	}

	return header, nil
}
