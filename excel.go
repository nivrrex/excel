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
	"fmt"
)

/**************************struct and objcet**************************/
type Excel struct {
	excel_obj      *ole.IUnknown
	excel          *ole.IDispatch
	workbooks      *ole.IDispatch
	sheets         *ole.IDispatch
	count          int
	visible        bool
	readonly       bool
	save           bool
	displayAlerts  bool
	screenUpdating bool
}

func (this *Excel) New() (e *Excel, err error) {
	ole.CoInitialize(0)
	this.excel_obj, _ = oleutil.CreateObject("Excel.Application")
	this.excel, _ = this.excel_obj.QueryInterface(ole.IID_IDispatch)
	if this.excel == nil {
		errors.New("error: Cant't Open excel.")
	}

	oleutil.PutProperty(this.excel, "Visible", this.visible)
	oleutil.PutProperty(this.excel, "DisplayAlerts", this.displayAlerts)
	oleutil.PutProperty(this.excel, "ScreenUpdating", this.screenUpdating)

	this.workbooks = oleutil.MustGetProperty(this.excel, "WorkBooks").ToIDispatch()
	oleutil.MustCallMethod(this.workbooks, "Add").ToIDispatch()

	return this, err
}

func (this *Excel) Open(filePath string) (e *Excel, err error) {
	ole.CoInitialize(0)
	this.excel_obj, _ = oleutil.CreateObject("Excel.Application")
	this.excel, _ = this.excel_obj.QueryInterface(ole.IID_IDispatch)
	if this.excel == nil {
		errors.New("error: Cant't Open excel.")
	}

	oleutil.PutProperty(this.excel, "Visible", this.visible)
	oleutil.PutProperty(this.excel, "DisplayAlerts", this.displayAlerts)
	oleutil.PutProperty(this.excel, "ScreenUpdating", this.screenUpdating)

	this.workbooks = oleutil.MustGetProperty(this.excel, "WorkBooks").ToIDispatch()
	//no password,to do  ...
	oleutil.MustCallMethod(this.workbooks, "open", filePath, nil, this.readonly).ToIDispatch()

	return this, err
}

func (this *Excel) Close() (err error) {
	oleutil.PutProperty(this.excel, "DisplayAlerts", this.displayAlerts)
	oleutil.PutProperty(this.excel, "ScreenUpdating", this.screenUpdating)
	if this.save {
		oleutil.MustCallMethod(this.excel, "Save").ToIDispatch()
	}

	oleutil.MustCallMethod(this.workbooks, "Close").ToIDispatch()
	oleutil.MustCallMethod(this.excel, "Quit").ToIDispatch()

	if this.sheets != nil {
		defer this.sheets.Release()
	}
	defer this.workbooks.Release()
	defer this.excel.Release()
	defer this.excel_obj.Release()

	return err
}

func (this *Excel) Save() (err error) {
	oleutil.MustCallMethod(this.excel, "Save").ToIDispatch()
	return err
}

func (this *Excel) SaveAs(filepath string, filetype string) (err error) {
	//Check version
	var typeOf, xlXLS, xlXLSX int
	fmt.Println(1111)
	version := oleutil.MustGetProperty(this.excel, "Version").ToString()
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
	fmt.Println(2222)

	if filetype == "xls" || filetype == "xlsx" || filetype == "csv" || filetype == "txt" || filetype == "template" || filetype == "html" {
		switch filetype {
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
		err = errors.New("error: Type is error.")
		return err
	}

	//no password
	activeWorkBook := oleutil.MustGetProperty(this.excel, "ActiveWorkBook").ToIDispatch()
	oleutil.MustCallMethod(activeWorkBook, "SaveAs", filepath, typeOf, nil, nil).ToIDispatch()

	defer activeWorkBook.Release()
	return err
}

func (this *Excel) SheetsCount() (e *Excel, err error) {
	sheets := oleutil.MustGetProperty(this.excel, "Sheets").ToIDispatch()
	sheet_number := (int)(oleutil.MustGetProperty(sheets, "Count").Val)
	this.count = sheet_number

	defer sheets.Release()
	return this, err
}

func (this *Excel) Sheet(i int) (e *Excel, err error) {
	if this.count == 0 {
		this.SheetsCount()
	}

	this.sheets = oleutil.MustGetProperty(this.excel, "Worksheets", i).ToIDispatch()
	oleutil.MustCallMethod(this.sheets, "Select").ToIDispatch()

	return this, err
}

func (this *Excel) Cells(row int, column int) (value string, err error) {
	if this.sheets == nil {
		err = errors.New("error: please use Excel.Sheet(i) to appoint the sheet.")
		return "", err
	}
	cells := oleutil.MustGetProperty(this.sheets, "Cells", row, column).ToIDispatch()
	value = oleutil.MustGetProperty(cells, "Text").ToString()

	defer cells.Release()
	return value, err
}
func (this *Excel) CellsWrite(value string, row int, column int) (err error) {
	if this.sheets == nil {
		err = errors.New("error: please use Excel.Sheet(i) to appoint the sheet.")
		return err
	}
	cells := oleutil.MustGetProperty(this.sheets, "Cells", row, column).ToIDispatch()
	oleutil.PutProperty(cells, "Value", value)

	defer cells.Release()
	return err
}

/**************************function**************************/
func fileIsExist(filepath string) (check bool) {
	_, err := os.OpenFile(filepath, os.O_RDWR|os.O_CREATE|os.O_EXCL, 0600)
	if os.IsExist(err) {
		return true
	}
	return false
}

func ExcelBookNew(visible bool) (oleObject *ole.IUnknown, err error) {
	ole.CoInitialize(0)
	excel, _ := oleutil.CreateObject("Excel.Application")
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()
	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
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

func ExcelBookOpen(filePath string, visible bool, readOnly int, password string, writePassword string) (oleObject *ole.IUnknown, err error) {
	ole.CoInitialize(0)
	excel, _ := oleutil.CreateObject("Excel.Application")
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()

	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
	}

	oleutil.PutProperty(excelIDispatch, "Visible", visible)

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

func ExcelBookClose(excel *ole.IUnknown, save int, alerts bool) (err error) {
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

func ExcelBookSave(excel *ole.IUnknown, alerts bool) (err error) {
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()
	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
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

func ExcelBookSaveAs(excel *ole.IUnknown, filePath string, typeOfString string, alerts bool, password string, writePassword string) (err error) {
	excelIDispatch, _ := excel.QueryInterface(ole.IID_IDispatch)
	defer excelIDispatch.Release()
	if excelIDispatch == nil {
		errors.New("error: Cant't Open excel.")
	}

	//Check version
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

	if !alerts {
		oleutil.PutProperty(application, "DisplayAlerts", true)
		oleutil.PutProperty(application, "ScreenUpdating", true)
	}
	return
}
