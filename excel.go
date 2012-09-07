package excel

import (
	"github.com/mattn/go-ole"
	"github.com/mattn/go-ole/oleutil"
)

import (
	"errors"
	"os"
	"regexp"
	"strconv"
)

func fileIsExist(filepath string) (check bool) {
	_, err := os.OpenFile(filepath, os.O_RDWR|os.O_CREATE|os.O_EXCL, 0600)
	if os.IsExist(err) {
		return true
	}
	return false
}

func ExcelBookNew(visible int) (oleObject *ole.IUnknown, err error) {
	ole.CoInitialize(0)
	excel, _ := oleutil.CreateObject("Excel.Application")
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()
	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
	}
	if visible > 1 {
		visible = 1
	} else if visible < 0 {
		visible = 0
	}
	application := oleutil.MustGetProperty(excelIDispatch, "Application").ToIDispatch()
	defer application.Release()
	oleutil.PutProperty(application, "Visible", visible)

	workBooks := oleutil.MustGetProperty(excelIDispatch, "WorkBooks").ToIDispatch()
	defer workBooks.Release()
	oleutil.MustCallMethod(workBooks, "Add").ToIDispatch()

	activeWorkbook := oleutil.MustGetProperty(excelIDispatch, "ActiveWorkbook").ToIDispatch()
	sheets := oleutil.MustGetProperty(activeWorkbook, "Sheets", 1).ToIDispatch()
	oleutil.MustCallMethod(sheets, "Select").ToIDispatch()
	defer sheets.Release()
	defer activeWorkbook.Release()

	return excel, err
}

func ExcelBookOpen(filePath string, visible int, readOnly int, password string, writePassword string) (oleObject *ole.IUnknown, err error) {
	ole.CoInitialize(0)
	excel, _ := oleutil.CreateObject("Excel.Application")
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()

	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
	}
	if visible > 1 {
		visible = 1
	} else if visible < 1 {
		visible = 0
	}
	if visible == 1 {
		oleutil.PutProperty(excelIDispatch, "Visible", true)
	} else {
		oleutil.PutProperty(excelIDispatch, "Visible", false)
	}
	if readOnly > 1 {
		readOnly = 1
	} else if readOnly < 1 {
		readOnly = 0
	}

	workbooks := oleutil.MustGetProperty(excelIDispatch, "Workbooks").ToIDispatch()
	defer workbooks.Release()

	if password != "" && writePassword != "" {
		oleutil.MustCallMethod(workbooks, "open", filePath, nil, readOnly, nil, password, writePassword).ToIDispatch()
	} else if password == "" && writePassword != "" {
		oleutil.MustCallMethod(workbooks, "open", filePath, nil, readOnly, nil, nil, writePassword).ToIDispatch()
	} else if password != "" && writePassword == "" {
		oleutil.MustCallMethod(workbooks, "open", filePath, nil, readOnly, nil, password, nil).ToIDispatch()
	} else if password == "" && writePassword == "" {
		oleutil.MustCallMethod(workbooks, "open", filePath, nil, readOnly).ToIDispatch()
	}

	activeWorkbook := oleutil.MustGetProperty(excelIDispatch, "ActiveWorkbook").ToIDispatch()
	activeWorkbookSheets := oleutil.MustGetProperty(activeWorkbook, "Sheets", 1).ToIDispatch()
	defer activeWorkbookSheets.Release()
	defer activeWorkbook.Release()
	//activeWorkbookSheets := oleutil.MustGetProperty(activeWorkbook, "Sheets","Sheet1").ToIDispatch()
	oleutil.MustCallMethod(activeWorkbookSheets, "Select").ToIDispatch()

	return excel, err
}

func ExcelBookClose(excel *ole.IUnknown, save int, alerts int) (err error) {
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()

	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
	}
	if save > 1 {
		save = 1
	} else if save < 0 {
		save = 0
	}
	if alerts > 1 {
		alerts = 1
	} else if alerts < 0 {
		alerts = 0
	}
	workbooks := oleutil.MustGetProperty(excelIDispatch, "Workbooks").ToIDispatch()
	application := oleutil.MustGetProperty(excelIDispatch, "application").ToIDispatch()
	defer workbooks.Release()
	defer application.Release()

	oleutil.PutProperty(application, "DisplayAlerts", alerts)
	oleutil.PutProperty(application, "ScreenUpdating", alerts)
	if save == 1 {
		oleutil.MustCallMethod(excelIDispatch, "Save").ToIDispatch()
	}
	//displayAlerts := oleutil.MustGetProperty(application, "DisplayAlerts").ToIDispatch()
	//screenUpdating := oleutil.MustGetProperty(application, "ScreenUpdating").ToIDispatch()

	oleutil.MustCallMethod(workbooks, "Close").ToIDispatch()
	oleutil.MustCallMethod(application, "Quit").ToIDispatch()

	defer excel.Release()
	return
}

func ExcelReadCell(excel *ole.IUnknown, rangeOrRow string, column int) (value string, err error) {
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()
	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
	}

	re, _ := regexp.Compile("[A-Z,a-z][0-9]:[A-Z,a-z][0-9]")
	if re.FindStringIndex(rangeOrRow) == nil {
		if rangeOrRowInt, _ := strconv.ParseInt(rangeOrRow, 0, 32); rangeOrRowInt < 1 {
			errors.New("error: Cant't Open excel.")
		}
		if column < 1 {
			errors.New("error: Cant't Open excel.")
		}
		activesheet := oleutil.MustGetProperty(excelIDispatch, "Activesheet").ToIDispatch()
		rangeOrRowInt, _ := strconv.ParseInt(rangeOrRow, 0, 32)
		cells := oleutil.MustGetProperty(activesheet, "Cells", rangeOrRowInt, column).ToIDispatch()
		cellsValue := oleutil.MustGetProperty(cells, "Text").ToString()
		defer cells.Release()
		defer activesheet.Release()
		return cellsValue, err
	} else {
		activesheet := oleutil.MustGetProperty(excelIDispatch, "Activesheet").ToIDispatch()
		rangeOle := oleutil.MustGetProperty(activesheet, "Range", rangeOrRow).ToIDispatch()
		rangeValue := oleutil.MustGetProperty(rangeOle, "Text").ToString()
		defer rangeOle.Release()
		defer activesheet.Release()
		return rangeValue, err
	}
	return "", err
}

func ExcelWriteCell(excel *ole.IUnknown, Value string, rangeOrRow string, column int) (err error) {
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()
	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
	}

	re, _ := regexp.Compile("[A-Z,a-z][0-9]:[A-Z,a-z][0-9]")
	if re.FindStringIndex(rangeOrRow) == nil {
		if rangeOrRowInt, _ := strconv.ParseInt(rangeOrRow, 0, 32); rangeOrRowInt < 1 {
			errors.New("error: Cant't Open excel.")
		}
		if column < 1 {
			errors.New("error: Cant't Open excel.")
		}
		activesheet := oleutil.MustGetProperty(excelIDispatch, "Activesheet").ToIDispatch()
		cells := oleutil.MustGetProperty(activesheet, "Cells", rangeOrRow, column).ToIDispatch()
		oleutil.PutProperty(cells, "Value", Value)
		defer cells.Release()
		defer activesheet.Release()
	} else {
		activesheet := oleutil.MustGetProperty(excelIDispatch, "Activesheet").ToIDispatch()
		rangeOrRow := oleutil.MustGetProperty(activesheet, "Range", rangeOrRow).ToIDispatch()
		oleutil.PutProperty(rangeOrRow, "Value", Value)
		defer rangeOrRow.Release()
		defer activesheet.Release()
	}
	return err
}

func ExcelBookSave(excel *ole.IUnknown, alerts int) (err error) {
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()
	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
	}
	if alerts > 1 {
		alerts = 1
	} else if alerts < 0 {
		alerts = 0
	}

	application := oleutil.MustGetProperty(excelIDispatch, "application").ToIDispatch()
	defer application.Release()

	oleutil.PutProperty(application, "DisplayAlerts", alerts)
	oleutil.PutProperty(application, "ScreenUpdating", alerts)

	activeWorkbook := oleutil.MustGetProperty(excelIDispatch, "ActiveWorkbook").ToIDispatch()
	defer activeWorkbook.Release()

	oleutil.MustCallMethod(activeWorkbook, "Save").ToIDispatch()
	return
}

func ExcelBookSaveAs(excel *ole.IUnknown, filePath string, typeOfString string, alerts int, password string, writePassword string) (err error) {
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()
	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
	}

	/*检查版本*/
	var typeOf, xlXLS, xlXLSX int
	application := oleutil.MustGetProperty(excelIDispatch, "Application").ToIDispatch()
	defer application.Release()
	version := oleutil.MustGetProperty(application, "Version").ToString()
	if version == "12.0" {
		xlXLS = 56
		xlXLSX = 51
	} else {
		xlXLS = -4143
	}
	xlCSV := 6
	xlTXT := -4158
	xlTemplate := 17
	xlHtml := 44

	if typeOfString == "xls" || typeOfString == "xlsx" || typeOfString == "csv" || typeOfString == "txt" || typeOfString == "template" || typeOfString == "html" {
		switch typeOfString {
		case "xls":
			typeOf = xlXLS
		case "xlsx":
			typeOf = xlXLSX
		case "csv":
			typeOf = xlCSV
		case "txt":
			typeOf = xlTXT
		case "template":
			typeOf = xlTemplate
		case "html":
			typeOf = xlHtml
		default:
			typeOf = 0
		}
	} else {
		errors.New("error: Type is error.")
		return
	}
	if alerts > 1 {
		alerts = 1
	} else if alerts < 0 {
		alerts = 0
	}

	oleutil.PutProperty(application, "DisplayAlerts", alerts)
	oleutil.PutProperty(application, "ScreenUpdating", alerts)

	activeWorkBook := oleutil.MustGetProperty(excelIDispatch, "ActiveWorkBook").ToIDispatch()
	defer activeWorkBook.Release()

	if password == "" && writePassword == "" {
		oleutil.MustCallMethod(activeWorkBook, "SaveAs", filePath, typeOf, nil, nil).ToIDispatch()
	}
	if password != "" && writePassword == "" {
		oleutil.MustCallMethod(activeWorkBook, "SaveAs", filePath, typeOf, password, nil).ToIDispatch()
	}
	if password != "" && writePassword != "" {
		oleutil.MustCallMethod(activeWorkBook, "SaveAs", filePath, typeOf, password, writePassword).ToIDispatch()
	}
	if password == "" && writePassword != "" {
		oleutil.MustCallMethod(activeWorkBook, "SaveAs", filePath, typeOf, nil, writePassword).ToIDispatch()
	}

	if alerts == 0 {
		oleutil.PutProperty(application, "DisplayAlerts", 1)
		oleutil.PutProperty(application, "ScreenUpdating", 1)
	}
	return
}
