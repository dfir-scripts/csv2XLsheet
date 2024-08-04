package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"
)

func main() {
	// Define command-line flags
	sourceFile := flag.String("i", "", "Path to the source CSV/TSV file (required)")
	templateFile := flag.String("t", "", "Path to the Excel template file (required)")
	sheetName := flag.String("s", "", "Sheet name to write data to (required)")
	delimiter := flag.String("d", "csv", "Delimiter for the input file (options: 'csv', 'tab', or any single character) (default: 'csv')")
	outputFile := flag.String("o", "", "Output file name (required)")
	startRow := flag.Int("r", 1, "Start importing data from this line number (default: 1)")

    // Customize the help message
    flag.Usage = func() {
        fmt.Println("Appends data from CSV/TSV files onto an existing Excel (XLSX,XLTX) sheet.\nWorks with tables, pivot tables and slicers.\nLine input errors are ignored and logged.\nQuotation marks are removed during processing.")
        fmt.Printf("\nUsage: %s [-i,-t,-s,-o,-d,-r,-h]\n\n", os.Args[0])
        fmt.Println("\nOptions:")
        fmt.Println("  -i  Input Path to the source CSV/TSV file (required)")
        fmt.Println("  -t  Path to the Excel XLSX/XLTX file (required)")
        fmt.Println("  -s  Existing sheet name to append lines (required)")
        fmt.Println("  -o  Output file name (required)")
        fmt.Println("  -d  Delimiter of input file (options: 'csv', 'tab', or character(s)) (default: 'csv')")
        fmt.Println("  -r  Start appending sheet from this line number (default: 1)")
        fmt.Println("  -h  Show this help message")
        fmt.Println("\n Example: Appends CSV file prc.csv to a sheet named Pf-Table\n in an excel template named PfSlicer.xltx starting at line 2\n and outputs a file named pfoutput.xlsx \n\n\tcsv2XLsheet -i prc.csv -t PfSlicer.xltx -s Pf-Table -r 2 -o pfoutput.xlsx\n")
    }

	// Parse command-line flags
	flag.Parse()

	// Check if no parameters are passed
	if len(os.Args) == 1 {
		flag.Usage()
		os.Exit(0)
	}

	// Check required flags are provided
	if *sourceFile == "" || *templateFile == "" || *outputFile == "" || *sheetName == "" {
		flag.Usage()
		log.Fatal("\nFlags -i (input file), -t (Excel template), -s (Sheet name), and -o (Output file) must be specified")
	}

	csvFile := *sourceFile
	excelTemplate := *templateFile
	targetSheetName := *sheetName
	outputFileName := *outputFile

	// Convert delimiter based on the given input
	var delim rune
	switch *delimiter {
	case "csv":
		delim = ','
	case "tab":
		delim = '\t'
	default:
		if utf8.RuneCountInString(*delimiter) == 1 {
			delim, _ = utf8.DecodeRuneInString(*delimiter)
		} else {
			log.Fatalf("Invalid delimiter: %s", *delimiter)
		}
	}

	// Create consolidated log file name
	logFileName := strings.TrimSuffix(outputFileName, filepath.Ext(outputFileName)) + "-errors.log"
	var logFile *os.File
	var hasErrors bool
	var errorCount, notAppendedCount int

	// Open the input file
	file, err := os.Open(csvFile)
	if err != nil {
		log.Fatalf("Failed to open input file: %v", err)
	}
	defer file.Close()

	// Read the input data with the specified delimiter
	reader := csv.NewReader(file)
	reader.Comma = delim
	reader.LazyQuotes = true
	var csvData [][]string

	// Process each line and handle errors
	lineNumber := 0
	for {
		record, err := reader.Read()
		if err != nil {
			if err.Error() == "EOF" {
				break
			}
			// Open the error log file if it's not already open
			if !hasErrors {
				hasErrors = true
				logFile, err = os.Create(logFileName)
				if err != nil {
					log.Fatalf("Failed to create error log file: %v", err)
				}
				defer logFile.Close()
			}
			// Write the erroneous line to the error log
			rawLine := strings.Join(record, string(reader.Comma))
			_, _ = logFile.WriteString(fmt.Sprintf("Error reading line: %s\n", rawLine))
			errorCount++
			continue
		}
		if lineNumber >= *startRow-1 {
			// Sanitize each field by removing quotation marks
			for i := range record {
				record[i] = strings.ReplaceAll(record[i], "\"", "")
			}
			csvData = append(csvData, record)
		}
		lineNumber++
	}

	// Open the existing Excel template
	f, err := excelize.OpenFile(excelTemplate)
	if err != nil {
		log.Fatalf("Failed to open Excel template: %v", err)
	}

	// Check if the specified sheet exists
	sheetExists := false
	for _, name := range f.GetSheetList() {
		if name == targetSheetName {
			sheetExists = true
			break
		}
	}

	if !sheetExists {
		log.Fatalf("Sheet '%s' does not exist in the template file!", targetSheetName)
	}

	// Set the active sheet
	sheetIndex, err := f.GetSheetIndex(targetSheetName)
	if err != nil {
		log.Fatalf("Failed to get sheet index: %v", err)
	}
	f.SetActiveSheet(sheetIndex)

	// Get the number of columns in the template sheet
	rows, err := f.GetRows(targetSheetName)
	if err != nil {
		log.Fatalf("Failed to get rows from sheet: %v", err)
	}
	var maxCols int
	if len(rows) > 0 {
		maxCols = len(rows[0]) // Assume first row gives the number of columns
	} else {
		// If there are no rows, assume a large number of columns
		maxCols = 16384 // Excel's maximum number of columns
	}

	// Get the next empty row in the target sheet
	nextRow := len(rows) + 1

	// Append the input data to the Excel sheet
	for i, row := range csvData {
		numFields := len(row)

		// Log lines with more fields than available columns
		if numFields > maxCols {
			// Open the error log file if it's not already open
			if !hasErrors {
				hasErrors = true
				logFile, err = os.Create(logFileName)
				if err != nil {
					log.Fatalf("Failed to create error log file: %v", err)
				}
				defer logFile.Close()
			}
			// Write the line to the error log
			rawLine := strings.Join(row, string(reader.Comma))
			_, _ = logFile.WriteString(fmt.Sprintf("Not appended (too many fields): %s\n", rawLine))
			notAppendedCount++
			continue
		}

		for j, value := range row {
			cell, _ := excelize.CoordinatesToCellName(j+1, nextRow+i)
			f.SetCellValue(targetSheetName, cell, value)
		}
	}

	// Save the updated Excel file
	if err := f.SaveAs(outputFileName); err != nil {
		log.Fatalf("Failed to save updated Excel file: %v", err)
	}

	fmt.Printf("Data successfully written to file %s, sheet %s\n", outputFileName, targetSheetName)

	// Print summary messages if there were errors
	if hasErrors {
		fmt.Printf("%d lines encountered errors. See the log at %s\n", errorCount + notAppendedCount, logFileName)
	}
}

