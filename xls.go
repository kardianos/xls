package xls

import (
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"os"

	"github.com/kardianos/xls/ole2"
)

// ErrInternal wraps a panic recovered while parsing. Seeing it means a bug
// in this package, not just a malformed file; please report it.
var ErrInternal = errors.New("xls: internal error")

// OpenReader reads an XLS workbook from r.
//
// charset selects the character set for 8-bit text in BIFF5 (Excel 5.0/95)
// files, by any name the WHATWG Encoding Standard accepts, such as
// "windows-1251" or "shift_jis". Pass "" or "utf-8" to use the code page the
// file records, which is the usual choice. BIFF8 (Excel 97-2003) files store
// text as Unicode and do not use it.
//
// OpenReader reads all the data it needs before it returns and does not keep
// r. The caller still owns r and closes it.
func OpenReader(r io.ReadSeeker, charset string) (wb *WorkBook, err error) {
	defer func() {
		if p := recover(); p != nil {
			wb, err = nil, fmt.Errorf("%w: %v", ErrInternal, p)
		}
	}()
	enc, err := charsetEncoding(charset)
	if err != nil {
		return nil, err
	}
	cf, err := ole2.Open(r)
	if errors.Is(err, ole2.ErrNotCompoundFile) {
		if id, ok := bareBIFF(r); ok {
			what := "an Excel 2.x-4.0 file"
			if id == recBOF {
				what = "a BIFF5/BIFF8 stream outside a compound file"
			}
			return nil, fmt.Errorf("%w: %s", ErrUnsupportedVersion, what)
		}
	}
	if err != nil {
		return nil, containerError(err)
	}
	book, biff := findWorkbook(cf)
	if book == nil {
		return nil, ErrNoWorkbook
	}
	stream, err := cf.ReadStream(book)
	if err != nil {
		return nil, containerError(err)
	}
	return parseWorkbook(stream, biff, enc)
}

// containerError wraps an error from package ole2. Damage to the compound
// file also wraps ErrCorrupt, so errors.Is(err, ErrCorrupt) reports any
// damaged file.
func containerError(err error) error {
	var ce ole2.CorruptError
	if errors.As(err, &ce) {
		return fmt.Errorf("%w: %w", ErrCorrupt, err)
	}
	return fmt.Errorf("xls: %w", err)
}

// findWorkbook finds the workbook stream. Excel 5.0/95 files name it
// "Book"; files saved as "Excel 97 & 5.0/95" have both, and the BIFF8
// "Workbook" is preferred. The tree search that Excel performs is tried
// for both names first, so that this reader opens the same stream as
// Excel. Only if it finds neither, a unique name match anywhere among the
// root's children is accepted.
func findWorkbook(cf *ole2.File) (*ole2.Entry, int) {
	root := cf.Root()
	for _, find := range []func(*ole2.Entry, string) *ole2.Entry{cf.Lookup, cf.Search} {
		if e := find(root, "Workbook"); e != nil && e.Type == ole2.TypeStream {
			return e, 8
		}
		if e := find(root, "Book"); e != nil && e.Type == ole2.TypeStream {
			return e, 5
		}
	}
	return nil, 0
}

// bareBIFF reports whether r starts with a BOF record, as Excel 2.x to 4.0
// files do. Those files are not stored in a compound file.
func bareBIFF(r io.ReadSeeker) (uint16, bool) {
	var b [4]byte
	if _, err := r.Seek(0, io.SeekStart); err != nil {
		return 0, false
	}
	if _, err := io.ReadFull(r, b[:]); err != nil {
		return 0, false
	}
	switch id := binary.LittleEndian.Uint16(b[:]); id {
	case recBOF2, recBOF3, recBOF4, recBOF:
		return id, true
	}
	return 0, false
}

// Open reads the XLS file name. See OpenReader for the meaning of charset.
// The file is closed before Open returns.
func Open(name, charset string) (*WorkBook, error) {
	f, err := os.Open(name)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	return OpenReader(f, charset)
}

// Close does nothing and returns nil. The workbook keeps no reference to
// the file or reader it was read from. Close remains so existing code
// still compiles.
func (w *WorkBook) Close() error {
	return nil
}
