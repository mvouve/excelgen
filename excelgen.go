//Package excelgen is a helper for tealeg/xlsx which generates single page
//reqorts based upon a list of structures. Headers for the report are generated
//using reflection from the field names of the structure.
//TODO: create strategy for buffering to a file.
package excelgen

import (
	"container/list"
	"log"
	"reflect"

	"github.com/tealeg/xlsx"
)

// NAXIMUM ALLOWED ROWS BY EXCEL. https://support.office.com/en-us/article/Excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
const ExcelMaxRows = 1048576

//GenerateReport takes in the name of a file and a list of structures and attempts to create a excel document using them.
func GenerateReport(fname string, elements *list.List) {
	doc := xlsx.NewFile()
	if elements.Len() <= 0 {
		return
	}
	report, _ := doc.AddSheet("Sheet 1") // TODO: make this more generalised?
	generateHeaders(elements.Front().Value, report.AddRow())
	for e := elements.Front(); e != nil; e = e.Next() {
		generateRow(e.Value, report.AddRow())
		if report.MaxRow >= ExcelMaxRows {
			log.Println("Too many entries for report, stopping at row: ", report.MaxRow)
			break
		}
	}
	doc.Save(fname + ".xlsx")
}

func generateHeaders(i interface{}, row *xlsx.Row) {
	fields := reflect.ValueOf(i)
	for i := 0; i < fields.NumField(); i++ {
		cell := row.AddCell()
		cell.SetString(fields.Type().Field(i).Name)
	}
}

func generateRow(i interface{}, row *xlsx.Row) {
	fields := reflect.ValueOf(i)
	for i := 0; i < fields.NumField(); i++ {
		cell := row.AddCell()
		cell.SetValue(fields.Field(i).Interface())
	}
}
