package main

import (
	"log"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	//basic excel write
	//basicExcel()

	//excel with chart example
	excelWithChart()

}

// func basicExcel() {
// 	f := excelize.NewFile()

// 	f.SetCellValue("Sheet1", "B2", 100)
// 	f.SetCellValue("Sheet1", "A1", 50)

// 	now := time.Now()

// 	f.SetCellValue("Sheet1", "A4", now.Format(time.ANSIC))

// 	if err := f.SaveAs("simple.xlsx"); err != nil {
// 		log.Fatal(err)
// 	}
// }

func excelWithChart() {

	categories := map[string]string{"A1": "USA", "A2": "China", "A3": "UK",
		"A4": "Russia", "A5": "South Korea", "A6": "Germany"}

	values := map[string]int{"B1": 46, "B2": 38, "B3": 29, "B4": 22, "B5": 13, "B6": 11}

	f := excelize.NewFile()

	for k, v := range categories {

		f.SetCellValue("Sheet1", k, v)
	}

	for k, v := range values {

		f.SetCellValue("Sheet1", k, v)
	}

	if err := f.AddChart("Sheet1", "E1", `{
	"type":"col", 
	"series":[
		{"name":"Sheet1!$A$2","categories":"Sheet1!$A$1:$A$6",
			"values":"Sheet1!$B$1:$B$6"}
		],
		"title":{"name":"Olympic Gold medals in London 2012"}}`); err != nil {

		log.Fatal(err)
	}

	if err := f.SaveAs("gold_medals.xlsx"); err != nil {
		log.Fatal(err)
	}
}
