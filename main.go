package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"regexp"
	"strconv"
	"strings"
)

func main() {
	if len(os.Args) < 2 {
		log.Fatalf("Usage: %s <filename.xlsx>\n", os.Args[0])
	}
	filename := os.Args[1]
	factor := 1.175

	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatal(err)
	}
	if err != nil {
		log.Fatal(err)
	}

	date, _ := f.GetCellValue("TIM Angebot", "G9")      // adjust as needed
	offer, _ := f.GetCellValue("TIM Angebot", "G8")     // adjust
	customer, _ := f.GetCellValue("TIM Angebot", "G24") // adjust
	ratetext, _ := f.GetCellValue("TIM Angebot", "B4")  // adjust

	rate := ""
	re := regexp.MustCompile(`\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?\s?[A-Z]{3}`)
	match := re.FindString(ratetext)
	if match != "" {
		rate = match
	}

	rows, err := f.GetRows("TIM Angebotszeilen")
	if err != nil {
		log.Fatal(err)
	}

	output := os.Stdout // or os.Create("output.csv")
	fmt.Fprintf(output, "Pos,Quote,Date,Customer,Rate,SKU,Description,Qty,List,Disc,Net,Total,Fact,Sell\n")

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
				if len(next) > 4 && next[1] == "" && next[4] != "" {
					desc += " [" + next[4] + "]"
					i++ // skip the extra info row
				}
			}
			desc = csvEscape(desc)

			fmt.Fprintf(output, "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%.2f,%s,%.3f,%.2f\n",
				linepos, offer, date, customer, rate,
				sku, desc, qty, list, disc, net, total, factor, net*factor,
			)
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
