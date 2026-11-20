package xls

import (
	"fmt"
	"strings"

	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/charmap"
	"golang.org/x/text/encoding/htmlindex"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/encoding/korean"
	"golang.org/x/text/encoding/simplifiedchinese"
	"golang.org/x/text/encoding/traditionalchinese"
)

// codepageEncoding maps a CODEPAGE record value ([MS-XLS] 2.4.52) to a
// decoder for 8-bit strings. Unknown code pages fall back to Windows-1252.
func codepageEncoding(cp uint16) encoding.Encoding {
	switch cp {
	case 437:
		return charmap.CodePage437
	case 850:
		return charmap.CodePage850
	case 852:
		return charmap.CodePage852
	case 855:
		return charmap.CodePage855
	case 858:
		return charmap.CodePage858
	case 860:
		return charmap.CodePage860
	case 862:
		return charmap.CodePage862
	case 863:
		return charmap.CodePage863
	case 865:
		return charmap.CodePage865
	case 866:
		return charmap.CodePage866
	case 874:
		return charmap.Windows874
	case 932:
		return japanese.ShiftJIS
	case 936:
		return simplifiedchinese.GBK
	case 949:
		return korean.EUCKR
	case 950:
		return traditionalchinese.Big5
	case 1250:
		return charmap.Windows1250
	case 1251:
		return charmap.Windows1251
	case 1253:
		return charmap.Windows1253
	case 1254:
		return charmap.Windows1254
	case 1255:
		return charmap.Windows1255
	case 1256:
		return charmap.Windows1256
	case 1257:
		return charmap.Windows1257
	case 1258:
		return charmap.Windows1258
	case 10000, 32768:
		return charmap.Macintosh
	case 10001:
		return japanese.ShiftJIS // Mac Japanese is a Shift_JIS variant
	case 10002:
		return traditionalchinese.Big5
	case 10003:
		return korean.EUCKR
	case 10007:
		return charmap.MacintoshCyrillic
	case 10008:
		return simplifiedchinese.GBK
	}
	// 367 (ASCII), 1252, 32769, and code pages with no decoder here (such as
	// 737, 775, 857, 861, 864, 869, 1361 and most Mac code pages): text in
	// those decodes as Windows-1252. Pass a charset to OpenReader to override.
	return charmap.Windows1252
}

// charsetEncoding resolves the charset argument of Open and OpenReader.
// "" and any name of UTF-8 mean "use the code page recorded in the file".
func charsetEncoding(charset string) (encoding.Encoding, error) {
	switch strings.ToLower(strings.TrimSpace(charset)) {
	case "", "utf-8", "utf8":
		return nil, nil
	}
	e, err := htmlindex.Get(charset)
	if err != nil {
		return nil, fmt.Errorf("xls: unknown charset %q", charset)
	}
	if name, _ := htmlindex.Name(e); name == "utf-8" {
		// Any spelling of UTF-8 means the file's code page; BIFF5 text
		// is never UTF-8.
		return nil, nil
	}
	return e, nil
}
