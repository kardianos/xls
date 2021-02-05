package xls

import (
	"fmt"
	"io"
	"os"

	"github.com/kardianos/xls/ole2"
)

// OpenReader opens an XLS file from r with charset.
// Charset may be "utf-8".
// If r is a closer, r.Close will be called when WorkBook.Close is called.
func OpenReader(r io.ReadSeeker, charset string) (*WorkBook, error) {
	ole, err := ole2.Open(r, charset)
	if err != nil {
		return nil, err
	}
	dir, err := ole.ListDir()
	if err != nil {
		return nil, err
	}
	var book, root *ole2.File
	for _, f := range dir {
		switch f.Name() {
		case "Workbook":
			book = f
		case "Book":
			book = f
		case "Root Entry":
			root = f
		}
	}
	if book == nil {
		return nil, fmt.Errorf("No OLE2 Excel Workbook found")
	}
	of := ole.OpenFile(book, root)
	wb := newWorkBookFromOle2(of)
	if c, ok := r.(io.Closer); ok {
		wb.closer = c
	}
	return wb, nil
}

// Open a XLS file from disk with the given charset.
func Open(name, charset string) (*WorkBook, error) {
	f, err := os.Open(name)
	if err != nil {
		return nil, err
	}
	wb, err := OpenReader(f, charset)
	if err != nil {
		return nil, err
	}

	return wb, nil
}

// Close WorkBook if it was opened with a Closer.
func (w *WorkBook) Close() error {
	if w.closer != nil {
		return w.closer.Close()
	}
	return nil
}
