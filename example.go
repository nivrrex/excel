package main

import (
	"excel"
	"fmt"
	"strconv"
)

func main() {
	e := new(excel.Excel)
	e.visible =  false
	e.readonly= false
	e.save= true

	filePath := "T:\\test.xls"
	//e.New()
	e.Open(filePath)
	fmt.Println (e.Cells(1,1))

	e,_ = (e.SheetsCount())
	fmt.Println (e.count)

	e.Sheet(1)

	v ,_:= e.Cells(1,1)
	fmt.Println (v)

	for i:=1 ; i< 9 ; i++ {
		e.CellsWrite(strconv.Itoa(i),i,1)
	}

	e.Save()
	e.SaveAs("T:\\test.result.xlsx","xlsx")

	e.Close()
}
