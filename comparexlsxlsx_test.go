package xls

import (
	"encoding/csv"
	"fmt"
	"math"
	"os"
	"strconv"
	"strings"
)

// compareXLX to a CSV file.
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

	xlsSheet, err := xlsFile.GetSheet(0)
	if err != nil {
		return fmt.Errorf("Cant get xls sheet: %w", err)
	}
	for rowi, row := range all {
		xlsRow := xlsSheet.Row(rowi)
		if xlsRow == nil {
			if strings.TrimSpace(strings.Join(row, "")) != "" {
				return fmt.Errorf("row %d: missing, but the CSV has %q", rowi, row)
			}
			continue
		}
		for coli, cell := range row {
			csvText := strings.TrimSpace(cell)
			xlsText := strings.TrimSpace(xlsRow.ColExact(coli))
			v := xlsRow.Value(coli)
			tm := xlsFile.ToDateTime(float64(v.Int))
			if xlsText == csvText {
				continue
			}
			xlsFloat, xlsErr := strconv.ParseFloat(xlsText, 64)
			csvFloat, csvErr := strconv.ParseFloat(csvText, 64)

			if xlsErr == nil && csvErr == nil {
				// CSV exports keep 15 significant digits.
				diff := math.Abs(xlsFloat - csvFloat)
				if diff <= 1e-14*math.Max(1, math.Max(math.Abs(xlsFloat), math.Abs(csvFloat))) {
					continue
				}
				return fmt.Errorf("sheet:%d, row/col: %d/%d, csv: (%s)[%d], xls: (%s)[%d], numbers difference: %f, v=%v, t=%s",
					0, rowi, coli, csvText, len(csvText),
					xlsText, len(xlsText), diff, v, tm)
			}
			return fmt.Errorf("sheet:%d, row/col: %d/%d, csv: (%s)[%d], xls: (%s)[%d], V=%v, t=%s",
				0, rowi, coli, csvText, len(csvText),
				xlsText, len(xlsText), v, tm)
		}
	}

	return nil
}
