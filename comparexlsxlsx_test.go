package xls

import (
	"encoding/csv"
	"fmt"
	"math"
	"os"
	"strconv"
	"strings"
)

// compareXLX compares file types.
func compareXLX(xlsName string, csvName string) error {
	xlsFile, err := Open(xlsName, "utf-8")
	if err != nil {
		return fmt.Errorf("Cant open xls file: %s", err)
	}
	defer xlsFile.Close()

	cf, err := os.Open(csvName)
	if err != nil {
		return err
	}
	cr := csv.NewReader(cf)
	all, err := cr.ReadAll()
	cf.Close()
	if err != nil {
		return err
	}

	xlsSheet := xlsFile.GetSheet(0)
	if xlsSheet == nil {
		return fmt.Errorf("Cant get xls sheet")
	}
	for rowi, row := range all {
		xlsRow := xlsSheet.Row(rowi)
		if xlsRow == nil {
			continue
		}
		for coli, cell := range row {
			csvText := strings.TrimSpace(cell)
			xlsText := strings.TrimSpace(xlsRow.ColExact(coli))
			if xlsText == csvText {
				continue
			}
			xlsFloat, xlsErr := strconv.ParseFloat(xlsText, 64)
			csvFloat, csvErr := strconv.ParseFloat(csvText, 64)

			if xlsErr == nil && csvErr == nil {
				diff := math.Abs(xlsFloat - csvFloat)
				if diff <= 0.0000001 {
					continue
				}
				return fmt.Errorf("sheet:%d, row/col: %d/%d, csv: (%s)[%d], xls: (%s)[%d], numbers difference: %f",
					0, rowi, coli, csvText, len(csvText),
					xlsText, len(xlsText), diff)
			}
			return fmt.Errorf("sheet:%d, row/col: %d/%d, csv: (%s)[%d], xls: (%s)[%d]",
				0, rowi, coli, csvText, len(csvText),
				xlsText, len(xlsText))
		}
	}

	return nil
}
