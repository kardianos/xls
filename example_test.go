package xls

import (
	"fmt"
	"log"
	"path/filepath"
)

func ExampleOpen() {
	xlFile, err := Open(filepath.Join("testdata", "table.xls"), "utf-8")
	if err != nil {
		log.Fatal(err)
	}
	defer xlFile.Close()

	fmt.Println(xlFile.Author)
}
