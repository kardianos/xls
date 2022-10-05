package xls

import (
	"fmt"
	"os"
	"path/filepath"
	"testing"
)

func TestReadAll(t *testing.T) {
	f, err := os.Open(filepath.Join("testdata", "multitable.xls"))
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	wb, err := OpenReader(f, "utf-8")
	if err != nil {
		t.Fatal(err)
	}
	defer wb.Close()

	sheetData, sheetName, err := wb.ReadAll(10)
	if err != nil {
		t.Fatal(err)
	}
	const wantName = `[Table1 Table2]`
	const wantData = `[[[Code Name Description] [code1 name1 description1] [code2 name2 description2] [code3 name3 description3] [code4 name4 description4] [code5 name5 description5] [code6 name6 description6] [code7 name7 description7] [code8 name8 description8] [code9 name9 description9]] [[Key Value] [Key1 Value1] [Key2 Value2] [Key3 Value3] [Key4 Value4] [Key5 Value5] [Key6 Value6] [Key7 Value7] [Key8 Value8] [Key9 Value9]]]`
	gotName := fmt.Sprintf("%v", sheetName)
	gotData := fmt.Sprintf("%v", sheetData)

	if wantName != gotName {
		t.Fatalf("incorrect name, got: %s", gotName)
	}
	if wantData != gotData {
		t.Fatalf("incorrect data, got: %s", gotData)
	}
}
