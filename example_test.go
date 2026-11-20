package xls

import (
	"fmt"
	"log"
	"path/filepath"
)

func ExampleOpen() {
	wb, err := Open(filepath.Join("testdata", "multitable.xls"), "")
	if err != nil {
		log.Fatal(err)
	}
	for i := 0; i < wb.NumSheets(); i++ {
		sheet, err := wb.GetSheet(i)
		if err != nil {
			log.Fatal(err)
		}
		row := sheet.Row(1)
		fmt.Printf("%s: row 1 = %q %q\n", sheet.Name, row.Col(0), row.Col(1))
	}
	// Output:
	// Table1: row 1 = "code1" "name1"
	// Table2: row 1 = "Key1" "Value1"
}

func ExampleRow_Value() {
	wb, err := Open(filepath.Join("testdata", "compare", "times.xls"), "")
	if err != nil {
		log.Fatal(err)
	}
	sheet, err := wb.GetSheet(0)
	if err != nil {
		log.Fatal(err)
	}
	v := sheet.Row(0).Value(1)
	fmt.Println(v.Kind, v.Format, v.Time.Format("15:04:05"), sheet.Row(0).Col(1))
	// Output:
	// date HH:MM:SS 18:00:00 18:00:00
}
