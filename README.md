# About excel ole

win32 excel ole implementation for golang

在Windows系统下调用github.com/mattn/go-ole库操作Excel文件，需要go-ole库的支持

## 需求

``` bash
go get github.com/mattn/go-ole
go install github.com/mattn/go-ole
go get github.com/mattn/go-ole/oleutil
go install github.com/mattn/go-ole/oleutil
```

##安装
``` bash
go get github.com/nivrrex/excel
go install github.com/nivrrex/excel
```

## 例子
``` go
package main

import (
	"github.com/nivrrex/excel"
	"strconv"
	"fmt"
)

func main() {
	
	e := &excel.Excel{Visible: false, Readonly: false, Saved: true}
	
	e.New()
	//filePath := "T:\\test.xlsx"
	//e.Open(filePath)
	
	//test error
	fmt.Println (e.Cells(1,1))
	
	e,_ = (e.SheetsCount())
	fmt.Println (e.Count)
	
	e.Sheet(1)
	for i:=1 ; i< 99 ; i++ {
		e.CellsWrite(strconv.Itoa(i),i,1)
	}

	v ,_:= e.Cells(10,1)
	fmt.Println (v)

	//e.Save()
	e.SaveAs("T:\\test.result.xls","xls")
	
	e.Close()
}

``` 

## 更新
2014.6.18 可以直接使用go get/go install进行安装使用，不再人工编译了
2014.6.15 将原有函数调用模式，更新为struct + func 的调用模式，感觉面向对象一点，看起来稍显舒服。
2012.9.07 first commit.