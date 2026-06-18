package main

import (
	"archive/zip"
	"fmt"
	"io"
	"log"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const fileNamePattern = "HVAC leads_Dated %s - %s"

type Element struct {
	JobType  string
	Campaign string
	Category string
}

type Stats map[time.Month]int

func main() {
	result, err := processZip("files/f.zip")
	if err != nil {
		log.Fatal(err)
	}

	if err := writeResult("Result.xlsx", result); err != nil {
		log.Fatal(err)
	}
}

func processZip(zipPath string) (map[Element]Stats, error) {
	result := make(map[Element]Stats)

	reader, err := zip.OpenReader(zipPath)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	for _, zf := range reader.File {
		month, err := parseMonthFromFilename(zf.Name)
		if err != nil {
			return nil, fmt.Errorf("%s: %w", zf.Name, err)
		}

		rc, err := zf.Open()
		if err != nil {
			return nil, err
		}

		if err := processExcel(rc, month, result); err != nil {
			rc.Close()
			return nil, err
		}

		rc.Close()
	}

	return result, nil
}

func parseMonthFromFilename(filename string) (time.Month, error) {
	var start, end string

	name := strings.TrimSuffix(filepath.Base(filename), ".xlsx")

	if _, err := fmt.Sscanf(name, fileNamePattern, &start, &end); err != nil {
		return 0, err
	}

	t, err := time.Parse("01_02_06", start)
	if err != nil {
		return 0, err
	}

	return t.Month(), nil
}

func processExcel(
	r io.Reader,
	month time.Month,
	result map[Element]Stats,
) error {

	f, err := excelize.OpenReader(r)
	if err != nil {
		return err
	}
	defer f.Close()

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		return err
	}

	for _, row := range rows {
		if len(row) < 5 {
			continue
		}

		el := Element{
			JobType:  strings.TrimSpace(row[1]),
			Campaign: strings.TrimSpace(row[3]),
			Category: strings.TrimSpace(row[4]),
		}

		if result[el] == nil {
			result[el] = make(Stats)
		}

		result[el][month]++
	}

	return nil
}

func writeResult(output string, result map[Element]Stats) error {
	f := excelize.NewFile()
	defer f.Close()

	const sheet = "Sheet1"

	headers := []string{
		"Job Type",
		"Job Campaign",
		"Campaign Category",
	}

	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheet, cell, h)
	}

	months := collectMonths(result)

	for i, month := range months {
		cell, _ := excelize.CoordinatesToCellName(i+4, 1)
		f.SetCellValue(sheet, cell, month.String())
	}

	rowNum := 2

	for el, stats := range result {
		f.SetCellValue(sheet, fmt.Sprintf("A%d", rowNum), el.JobType)
		f.SetCellValue(sheet, fmt.Sprintf("B%d", rowNum), el.Campaign)
		f.SetCellValue(sheet, fmt.Sprintf("C%d", rowNum), el.Category)

		for i, month := range months {
			cell, _ := excelize.CoordinatesToCellName(i+4, rowNum)
			f.SetCellValue(sheet, cell, stats[month])
		}

		rowNum++
	}

	return f.SaveAs(output)
}

func collectMonths(result map[Element]Stats) []time.Month {
	found := make(map[time.Month]struct{})

	for _, stats := range result {
		for month := range stats {
			found[month] = struct{}{}
		}
	}

	months := make([]time.Month, 0, len(found))

	for month := range found {
		months = append(months, month)
	}

	return months
}
