package xls

import (
	"bytes"
	"encoding/binary"
	"io"
	"os"
	"unicode/utf16"

	"golang.org/x/text/encoding/charmap"
)

// WorkBook is the parsed XLS file.
type WorkBook struct {
	Is5ver   bool
	Type     uint16
	Codepage uint16
	XF       []XF
	Fonts    []Font
	Formats  map[uint16]*Format
	//All the sheets from the workbook
	sheets         []*WorkSheet
	Author         string
	rs             io.ReadSeeker
	sst            []string
	continue_utf16 uint16
	continue_rich  uint16
	continue_apsb  uint32
	dateMode       uint16
	closer         io.Closer
}

type sstInfo struct {
	Total uint32
	Count uint32
}

//read workbook from ole2 file
func newWorkBookFromOle2(rs io.ReadSeeker) *WorkBook {
	wb := new(WorkBook)
	wb.Formats = make(map[uint16]*Format)
	wb.rs = rs
	wb.sheets = make([]*WorkSheet, 0)
	wb.parse(rs)
	return wb
}

func (w *WorkBook) parse(buf io.ReadSeeker) {
	b := new(bof)
	bofPre := new(bof)

	offset := 0
	for {
		if err := binary.Read(buf, binary.LittleEndian, b); err == nil {
			bofPre, b, offset = w.parseBof(buf, b, bofPre, offset)
		} else {
			break
		}
	}
}

func (w *WorkBook) addXf(xf XF) {
	w.XF = append(w.XF, xf)
}

func (w *WorkBook) addFont(font *FontInfo, buf io.ReadSeeker) {
	name, _ := w.getString(buf, uint16(font.NameB))
	w.Fonts = append(w.Fonts, Font{Info: font, Name: name})
}

func (w *WorkBook) addFormat(format *Format) {
	if w.Formats == nil {
		os.Exit(1)
	}
	w.Formats[format.Head.Index] = format
}

func (w *WorkBook) parseBof(buf io.ReadSeeker, b *bof, pre *bof, offsetPre int) (after *bof, afterUsing *bof, offset int) {
	after = b
	afterUsing = pre
	var bts = make([]byte, b.Size)
	binary.Read(buf, binary.LittleEndian, bts)
	bufItem := bytes.NewReader(bts)
	switch b.ID {
	case 0x809:
		bif := new(biffHeader)
		binary.Read(bufItem, binary.LittleEndian, bif)
		if bif.Ver != 0x600 {
			w.Is5ver = true
		}
		w.Type = bif.Type
	case 0x042: // CODEPAGE
		binary.Read(bufItem, binary.LittleEndian, &w.Codepage)
	case 0x3c: // CONTINUE
		if pre.ID == 0xfc {
			var size uint16
			var err error
			if w.continue_utf16 >= 1 {
				size = w.continue_utf16
				w.continue_utf16 = 0
			} else {
				err = binary.Read(bufItem, binary.LittleEndian, &size)
			}
			for err == nil && offsetPre < len(w.sst) {
				var str string
				str, err = w.getString(bufItem, size)
				w.sst[offsetPre] = w.sst[offsetPre] + str

				if err == io.EOF {
					break
				}

				offsetPre++
				err = binary.Read(bufItem, binary.LittleEndian, &size)
			}
		}
		offset = offsetPre
		after = pre
		afterUsing = b
	case 0xfc: // SST
		info := new(sstInfo)
		binary.Read(bufItem, binary.LittleEndian, info)
		w.sst = make([]string, info.Count)
		var size uint16
		var i = 0
		// dont forget to initialize offset
		offset = 0
		for ; i < int(info.Count); i++ {
			var err error
			err = binary.Read(bufItem, binary.LittleEndian, &size)
			if err == nil {
				var str string
				str, err = w.getString(bufItem, size)
				w.sst[i] = w.sst[i] + str
			}

			if err == io.EOF {
				break
			}
		}
		offset = i
	case 0x85: // boundsheet
		var bs = new(boundsheet)
		binary.Read(bufItem, binary.LittleEndian, bs)
		// different for BIFF5 and BIFF8
		w.addSheet(bs, bufItem)
	case 0x0e0: // XF
		if w.Is5ver {
			xf := new(xf5)
			binary.Read(bufItem, binary.LittleEndian, xf)
			w.addXf(xf)
		} else {
			xf := new(xf8)
			binary.Read(bufItem, binary.LittleEndian, xf)
			w.addXf(xf)
		}
	case 0x031: // FONT
		f := new(FontInfo)
		binary.Read(bufItem, binary.LittleEndian, f)
		w.addFont(f, bufItem)
	case 0x41E: //FORMAT
		font := new(Format)
		binary.Read(bufItem, binary.LittleEndian, &font.Head)
		font.str, _ = w.getString(bufItem, font.Head.Size)
		w.addFormat(font)
	case 0x22: //DATEMODE
		binary.Read(bufItem, binary.LittleEndian, &w.dateMode)
	}
	return
}
func decodeWindows1251(enc []byte) string {
	dec := charmap.Windows1251.NewDecoder()
	out, _ := dec.Bytes(enc)
	return string(out)
}
func (w *WorkBook) getString(buf io.ReadSeeker, size uint16) (res string, err error) {
	if w.Is5ver {
		var bts = make([]byte, size)
		_, err = buf.Read(bts)
		if err != nil {
			return
		}
		res = decodeWindows1251(bts)
	} else {
		richtextNum := uint16(0)
		phoneticSize := uint32(0)
		var flag byte
		err = binary.Read(buf, binary.LittleEndian, &flag)
		if flag&0x8 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &richtextNum)
		} else if w.continue_rich > 0 {
			richtextNum = w.continue_rich
			w.continue_rich = 0
		}
		if flag&0x4 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &phoneticSize)
		} else if w.continue_apsb > 0 {
			phoneticSize = w.continue_apsb
			w.continue_apsb = 0
		}
		if flag&0x1 != 0 {
			var bts = make([]uint16, size)
			var i = uint16(0)
			for ; i < size && err == nil; i++ {
				err = binary.Read(buf, binary.LittleEndian, &bts[i])
			}

			// When eof found, we dont want to append last element.
			var runes []rune
			if err == io.EOF {
				i = i - 1
			}
			runes = utf16.Decode(bts[:i])

			res = string(runes)
			if i < size {
				w.continue_utf16 = size - i
			}

		} else {
			var bts = make([]byte, size)
			var n int
			n, err = buf.Read(bts)
			if uint16(n) < size {
				w.continue_utf16 = size - uint16(n)
				err = io.EOF
			}

			var bts1 = make([]uint16, n)
			for k, v := range bts[:n] {
				bts1[k] = uint16(v)
			}
			runes := utf16.Decode(bts1)
			res = string(runes)
		}
		if richtextNum > 0 {
			var bts []byte
			var seekSize int64
			if w.Is5ver {
				seekSize = int64(2 * richtextNum)
			} else {
				seekSize = int64(4 * richtextNum)
			}
			bts = make([]byte, seekSize)
			err = binary.Read(buf, binary.LittleEndian, bts)
			if err == io.EOF {
				w.continue_rich = richtextNum
			}
		}
		if phoneticSize > 0 {
			var bts []byte
			bts = make([]byte, phoneticSize)
			err = binary.Read(buf, binary.LittleEndian, bts)
			if err == io.EOF {
				w.continue_apsb = phoneticSize
			}
		}
	}
	return
}

func (w *WorkBook) addSheet(sheet *boundsheet, buf io.ReadSeeker) {
	name, _ := w.getString(buf, uint16(sheet.Name))
	w.sheets = append(w.sheets, &WorkSheet{bs: sheet, Name: name, wb: w, Visibility: TWorkSheetVisibility(sheet.Visible)})
}

// Reading a sheet from the compress file to memory, you should call this before you try to get anything from sheet.
func (w *WorkBook) prepareSheet(sheet *WorkSheet) {
	w.rs.Seek(int64(sheet.bs.Filepos), 0)
	sheet.parse(w.rs)
}

// GetSheet gets one sheet by its number.
func (w *WorkBook) GetSheet(num int) *WorkSheet {
	if num >= len(w.sheets) {
		return nil
	}
	s := w.sheets[num]
	if !s.parsed {
		w.prepareSheet(s)
	}
	return s
}

// NumSheets gets the number of all sheets, look into example.
func (w *WorkBook) NumSheets() int {
	return len(w.sheets)
}

// ReadAll is a helper function to read all cells from file
// Notice: the max value is the limit of the max capacity of lines.
// Warning: the helper function will need big memeory if file is large.
func (w *WorkBook) ReadAll(max int) (res [][]string) {
	res = make([][]string, 0)
	for _, sheet := range w.sheets {
		if len(res) < max {
			max = max - len(res)
			w.prepareSheet(sheet)
			if sheet.MaxRow != 0 {
				leng := int(sheet.MaxRow) + 1
				if max < leng {
					leng = max
				}
				temp := make([][]string, leng)
				for k, row := range sheet.rows {
					data := make([]string, 0)
					if len(row.cols) > 0 {
						for _, col := range row.cols {
							if uint16(len(data)) <= col.LastCol() {
								data = append(data, make([]string, col.LastCol()-uint16(len(data))+1)...)
							}
							str := col.String(w)

							for i := uint16(0); i < col.LastCol()-col.FirstCol()+1; i++ {
								data[col.FirstCol()+i] = str[i]
							}
						}
						if leng > int(k) {
							temp[k] = data
						}
					}
				}
				res = append(res, temp...)
			}
		}
	}
	return
}
