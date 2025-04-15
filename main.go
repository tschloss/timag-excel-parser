package main

import (
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

func main() {
	if len(os.Args) < 2 {
		log.Fatalf("Usage: %s <filename.xlsx>\n", os.Args[0])
	}

	var sellFactor = flag.Float64("factor", 1.0, "Multiplication factor for Sell price")
	var markFlag = flag.Bool("po", false, "add --po if quote was purchased")
	var verbose = flag.Bool("v", false, "add --v to get beater readability instead of csv")
	var toFile = flag.Bool("tofile", false, "Output to csv file in source directory (defaults to stdout)")
	flag.Parse()

	if flag.NArg() < 1 {
		log.Fatalf("Usage: %s [--factor 1.3] <filename.xlsx>\n", os.Args[0])
	}
	purchased := "no"
	if *markFlag {
		purchased = "yes"
	}
	filename := flag.Arg(0)
	output := os.Stdout // or os.Create("output.csv")
	if *toFile {
		if *verbose {
			output = prepareOutFile(filename, ".txt")
		} else {
			output = prepareOutFile(filename, ".csv")
		}
		defer output.Close()
	}

	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatal(err)
	}
	if err != nil {
		log.Fatal(err)
	}

	headerFields, _ := findValues(f, "TIM Angebot") // adjust as needed

	rows, err := f.GetRows("TIM Angebotszeilen")
	if err != nil {
		log.Fatal(err)
	}

	if *verbose {
		fmt.Fprintf(output, "%-10s: %s\n%-10s: %s\n%-10s: %s\n%-10s: %s\n%-10s: %s\n%-10s: %.3f\n",
			"quote", headerFields[0], "date", headerFields[1], "po", purchased, "customer", headerFields[2], "xchg-rate", headerFields[3], "uplift", *sellFactor)
	} else {
		fmt.Fprintf(output, "Quote,Pos,Date,PO,Customer,Rate,SKU,Qty,List,Disc,Net,Total,Fact,Sell,Description\n")
	}

	for i := 0; i < len(rows); i++ {
		row := rows[i]

		// Check if column B (index 1) has a number → this is a line item
		if len(row) > 2 && isInteger(row[1]) {
			linepos := safeGet(row, 1)
			sku := safeGet(row, 3)                    // column A
			desc := safeGet(row, 4)                   // column C (index 2)
			qty := cleanNumber(safeGet(row, 5))       // column D
			list := cleanNumber(safeGet(row, 6))      // column E
			disc := cleanNumber(safeGet(row, 7))      // column F
			net, _ := parseUSDecimal(safeGet(row, 8)) // column G
			total := cleanNumber(safeGet(row, 9))     // column H

			// Check for an extra info row just after this
			if i+1 < len(rows) {
				next := rows[i+1]
				if len(next) > 4 && ((next[1] == "" && next[4] != "" && next[4] != "Gesamtsumme:") || next[3] == "TERM-MONTHS") {
					desc += " [" + next[4] + "]"
					i++ // skip the extra info row
				}
			}
			desc = csvEscape(desc)

			if *verbose {
				fmt.Fprintf(output, "# %s\n%-10s: %s\n%-10s: %s\n%-10s: %s\n%-10s: %s\n%-10s: %.2f\n%-10s: %s\n%-10s: %.2f\n%-10s: %s\n",
					linepos, "   SKU", sku, "   qty", qty, "   list", list, "   disc", disc, "   net", net, "   total", total, "   sell", net*(*sellFactor), "   desc", desc)
			} else {
				fmt.Fprintf(output, "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%.2f,%s,%.3f,%.2f,%s\n",
					headerFields[0], linepos, headerFields[1], purchased, headerFields[2], headerFields[3],
					sku, qty, list, disc, net, total, *sellFactor, net*(*sellFactor), desc)
			}
		}

	}
}

func isInteger(s string) bool {
	s = strings.TrimSpace(s)
	_, err := strconv.Atoi(s)
	return err == nil
}

func safeGet(row []string, index int) string {
	if index < len(row) {
		return strings.TrimSpace(row[index])
	}
	return ""
}

func cleanNumber(s string) string {
	s = strings.ReplaceAll(s, `,`, ``)
	return s
}

func csvEscape(s string) string {
	// Replace CR/LF with literal separator
	s = strings.ReplaceAll(s, "\r\n", " | ")
	s = strings.ReplaceAll(s, "\n", " | ")
	s = strings.ReplaceAll(s, "\r", " | ")

	// Escape quotes
	s = strings.ReplaceAll(s, `"`, `""`)

	// Wrap in quotes if needed
	if strings.ContainsAny(s, `",`) {
		s = `"` + s + `"`
	}

	return s
}

func parseUSDecimal(s string) (float64, error) {
	cleaned := strings.ReplaceAll(s, ",", "") // remove thousands
	return strconv.ParseFloat(cleaned, 64)
}

func findValues(f *excelize.File, sheet string) ([]string, error) {
	values := []string{"?", "?", "?", "1"}
	rows, err := f.GetRows(sheet)
	if err != nil {
		return values, err
	}
	re := regexp.MustCompile(`\d[.,]\d{2,4}\sUSD`)
	for _, row := range rows {
		for i, cell := range row {
			if (cell == "Angebots-Nr." || cell == "Angebots Nr.") && len(row) > i+1 {
				values[0] = row[i+1]
			}
			if (cell == "Angebotsdatum" || cell == "Angebots Datum" || cell == "Datum") && len(row) > i+1 {
				values = append(values, row[i+1])
				values[1] = row[i+1]
			}
			if strings.Contains(cell, "Endkunde") && len(row) > i+1 {
				values[2] = row[i+1]
			}
			if strings.Contains(cell, "USD Referenzkurs der EZB von 1 EUR") {
				cell = strings.ReplaceAll(cell, "\n", " ")
				//cell = strings.ReplaceAll(cell, "\r", " ")
				//fmt.Printf("%q\n", cell)
				//fmt.Printf("%#v\n", []byte(cell))
				rate := ""
				match := re.FindString(cell)
				if match != "" {
					rate = strings.ReplaceAll(match[:5], `,`, `.`)
				} else {
					rate = "?"
				}
				values[3] = rate
			}
		}
	}

	return values, nil
}

func prepareOutFile(inputFile string, ext string) *os.File {
	outputFile := filepath.Base(inputFile)
	outputFile = outputFile[:len(outputFile)-len(filepath.Ext(outputFile))] + ext
	fmt.Fprintln(os.Stderr, "Writing to:", outputFile)

	// Decide where to write
	var out *os.File
	var err error

	out, err = os.Create(outputFile)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Could not create output file: %v\n", err)
		os.Exit(1)
	}

	return out

}
