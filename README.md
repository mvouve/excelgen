# excelgen

Excelgen is a quick package for generating an excel file based upon a list of structures. This package was originally intended for tracking connection information for scalable server tests. By using Excel documents graphs and supporting information can be much easier to output than using an alterative.

## Basic Usage
```Go
import "github.com/mvouve/excelgen"

type ExampleStruct struct {
	FieldOne   int
	FieldTwo   string
	FieldThree double
}

func main() {
	l := list.New()
	l.PushBack(PassStruct{FieldOne: 1, FieldTwo: "Foo", FieldThree: 2.5})
	GenerateReport("examplereport", l)
}```
