package excelgen

import (
	"container/list"
	"os"
	"testing"
)

type PassStruct struct {
	Ed  int
	Red int
}

//
func TestGenerateReport(t *testing.T) {
	l := list.New()
	l.PushBack(PassStruct{Ed: 5, Red: 1})
	GenerateReport("test1", l)
	if _, err := os.Stat("test1.xlsx"); os.IsNotExist(err) { // if file didn't create test failed.
		t.Fail()
	} else {
		os.Remove("test1.xlsx")
	}
}
