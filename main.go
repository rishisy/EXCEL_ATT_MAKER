package main

import (
	"fmt"
	"log"
	"os"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	argsWithProg := os.Args
	if len(argsWithProg) < 2 {
		log.Fatal("Please provide name of month, usage: ./main.go July")
	}

	monthName := argsWithProg[1]
	fmt.Println("Processing month:", monthName)

	// Parse the month name to time.Month
	month, err := time.Parse("January", monthName)
	if err != nil {
		log.Fatalf("Invalid month name '%s': %v", monthName, err)
	}

	// Current year
	year := time.Now().Year()

	// Start date: 25th of given month
	startDate := time.Date(year, month.Month(), 25, 0, 0, 0, 0, time.Local)

	// End date: 25th of next month
	nextMonth := month.Month() + 1
	nextYear := year
	if nextMonth > 12 {
		nextMonth = 1
		nextYear++
	}
	endDate := time.Date(nextYear, nextMonth, 25, 0, 0, 0, 0, time.Local)

	// Open the Excel file
	f, err := excelize.OpenFile("base_att_template.xlsx")
	if err != nil {
		log.Fatal("Error in opening file:", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	// Use first sheet or create if not exists
	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		sheetName = "Sheet1"
		f.NewSheet(sheetName)
	}

	// Row counter starting from 1
	row := 10

	// Iterate from startDate to endDate inclusive
	for d := startDate; !d.After(endDate); d = d.AddDate(0, 0, 1) {
		dateStr := d.Format("2006-01-02") // ISO format date
		dayName := d.Weekday().String()

		// Write date to column A
		cellA := fmt.Sprintf("A%d", row)
		if err := f.SetCellValue(sheetName, cellA, dateStr); err != nil {
			log.Fatalf("Failed to set cell %s: %v", cellA, err)
		}

		// Write day name to column B
		cellB := fmt.Sprintf("B%d", row)
		if err := f.SetCellValue(sheetName, cellB, dayName); err != nil {
			log.Fatalf("Failed to set cell %s: %v", cellB, err)
		}

		// If not Saturday or Sunday, write "Present" in column C
		if d.Weekday() != time.Saturday && d.Weekday() != time.Sunday {
			cellC := fmt.Sprintf("C%d", row)
			if err := f.SetCellValue(sheetName, cellC, "Present"); err != nil {
				log.Fatalf("Failed to set cell %s: %v", cellC, err)
			}
		}

		row++
	}

	// Save the file with a new name to avoid overwriting original
    outputFile := "att_result_final.xlsx"
	if err := f.SaveAs(outputFile); err != nil {
		log.Fatal("Failed to save updated file:", err)
	}

	fmt.Printf("Updated Excel file saved as %s\n", outputFile)
}

