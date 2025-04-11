package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"regexp"
)

func main() {
	if len(os.Args) < 2 {
		log.Fatalf("Usage: %s <filename.xlsx>\n", os.Args[0])
	}
	filename := os.Args[1]

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
	fmt.Fprintf(output, "Quote,Date,Customer,Rate,Pos,SKU,Description,Qty,List,Disc,Net,Total\n")

	for i := 0; i < len(rows); i++ {
		row := rows[i]
		if len(row) < 7 {
			continue
		}

		desc := row[4]
		// look ahead for extra info
		if i+1 < len(rows) {
			next := rows[i+1]
			if len(next) > 1 && next[2] == "" && next[3] == "" && next[1] != "" {
				desc += " [" + next[4] + "]"
				i++ // skip the extra info row
			}
		}

		fmt.Fprintf(output, "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n",
			offer, date, customer, rate,
			row[1], row[3], desc, row[5], row[6], row[7], row[8], row[9],
		)
	}
}
