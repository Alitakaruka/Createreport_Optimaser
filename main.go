package main

import (
	"archive/zip"
	"fmt"
	"io"
	"log"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type Element struct {
	Jobtype   string
	Compation string
	Catigoty  string
	Month     time.Month
}

type Value struct {
	Data map[string]int
}

func NewValue() *Value {
	d := Value{}
	d.Data = make(map[string]int)
	return &d
}

var Result map[Element]int

const formatStr = "HVAC leads_Dated %s - %s"

func main() {
	Result = make(map[Element]int)
	reader, err := zip.OpenReader("files/f.zip")

	if err != nil {
		log.Println(err)
		return
	}

	// fmt.Printf("file: %v\n", file)
	for _, f := range reader.File {
		fmt.Printf("f.Name: %v\n", f.Name)
		fileReader, err := f.Open()
		if err != nil {
			panic(err)
		}
		var start, end string

		_, err = fmt.Sscanf(strings.TrimSuffix(f.Name, ".xlsx"), formatStr, &start, &end)

		if err != nil {
			panic(err.Error())
		} else {
			fmt.Printf("start: %v\n", start)
			fmt.Printf("end: %v\n", end)
		}

		StartTime, err := time.Parse("01_02_06", start)
		// EndTime, err := time.Parse("02_01_06", end)
		parceFile(fileReader, StartTime.Month())
	}

	WriteToFile()
}

func parceFile(fileReader io.Reader, Month time.Month) {

	f, err := excelize.OpenReader(fileReader)

	if err != nil {
		log.Fatal(err)
	}

	rows, err := f.GetRows("Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	// fmt.Printf("rows: %v\n", rows)
	for _, row := range rows {
		El := Element{Jobtype: row[1], Compation: row[3], Catigoty: row[4], Month: Month}

		if _, ok := Result[El]; !ok {
			Result[El] = 1
		} else {
			Result[El] = Result[El] + 1
		}

	}

}

func WriteToFile() {
	f := excelize.NewFile()

	f.SetCellValue("Sheet1", "A1", "Job Type")
	f.SetCellValue("Sheet1", "B1", "Job Campaign")
	f.SetCellValue("Sheet1", "C1", "Campaign Category")

	find := make(map[string]bool)
	for key, _ := range Result {
		if _, ok := find[key.Month.String()]; !ok {
			find[key.Month.String()] = true
			f.SetCellValue("Sheet1", strings.ToUpper(string(intToW(4+int(key.Month))))+strconv.Itoa(1), key.Month.String())
		}
	}

	counter := 2
	for key, val := range Result {
		f.SetCellValue("Sheet1", "A"+strconv.Itoa(counter), key.Jobtype)
		f.SetCellValue("Sheet1", "B"+strconv.Itoa(counter), key.Compation)
		f.SetCellValue("Sheet1", "C"+strconv.Itoa(counter), key.Catigoty)
		f.SetCellValue("Sheet1", strings.ToUpper(string(intToW(4+int(key.Month))))+strconv.Itoa(counter), val)
		counter++

	}

	if err := f.SaveAs("Result.xlsx"); err != nil {
		log.Fatal(err)
	}
}

func intToW(a int) rune {
	return rune('a' + a - 1)
}
