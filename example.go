package main

import (
	"excel"
	"strconv"
	"fmt"
)

func main() {
	
	e := &excel.Excel{Visible: false, Readonly: false, Saved: true}
	
	filePath := "T:\\test.xls"
	//e.New()
	e.Open(filePath)
	fmt.Println (e.Cells(1,1))
	
	e,_ = (e.SheetsCount())
	fmt.Println (e.Count)
	
	e.Sheet(1)
	for i:=1 ; i< 99 ; i++ {
		e.CellsWrite(strconv.Itoa(i),i,1)
	}

	v ,_:= e.Cells(10,1)
	fmt.Println (v)

	e.Save()
	e.SaveAs("T:\\test.result.xlsx","xlsx")
	
	e.Close()
}

