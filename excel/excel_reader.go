package excel

import (
	"errors"
	"fmt"

	"github.com/go-gota/gota/dataframe"
	"github.com/go-gota/gota/series"
	"github.com/xuri/excelize/v2"
)

// ExcelReader Excel读取器
type ExcelReader struct {
	data        map[uint64][]string
	excel       *excelize.File
	sheetName   string
	originalRow int
	finishedRow int
	titleRow    int
	titles      []string
	content     [][]string
}

// NewExcelReader 构造函数
func NewExcelReader() *ExcelReader {
	return &ExcelReader{}
}

// AutoRead 自动读取（默认第一行是表头，从第二行开始，默认Sheet名称为：Sheet1）
func (r *ExcelReader) AutoRead(filename string, values ...any) *ExcelReader {
	return r.
		OpenFile(filename, values...).
		SetOriginalRow(2).
		SetTitleRow(1).
		SetSheetName("Sheet1").
		ReadTitle().
		Read()
}

// AutoReadBySheetName 自动读取（默认第一行是表头，从第二行开始）
func (r *ExcelReader) AutoReadBySheetName(filename, sheetName string, values ...any) *ExcelReader {
	return r.
		OpenFile(filename, values...).
		SetOriginalRow(2).
		SetTitleRow(1).
		SetSheetName(sheetName).
		ReadTitle().
		Read()
}

// ToList 获取数据（数组类型）
func (r *ExcelReader) ToList() map[uint64][]string {
	return r.data
}

// ToMap 获取数据（map类型）
func (r *ExcelReader) ToMap() map[uint64]map[string]string {
	if len(r.GetTitle()) == 0 {
		panic(errors.New("未设置表头"))
	}

	_data := make(map[uint64]map[string]string)

	for rowNumber, row := range r.ToList() {
		// if len(r.GetTitle()) != len(row) {
		// 	panic(fmt.Errorf("表头数量与实际数据列不匹配（第%d行）", rowNumber))
		// }

		_row := make(map[string]string)
		for _, title := range r.GetTitle() {
			_row[title] = "nil"
		}
		for k, v := range row {
			_row[r.GetTitle()[k]] = v
		}
		_data[rowNumber] = make(map[string]string)
		_data[rowNumber] = _row
	}

	return _data
}

// SetDataByRow 设置单行数据
func (r *ExcelReader) SetDataByRow(rowNumber uint64, data []string) *ExcelReader {
	r.data[rowNumber+1] = data
	return r
}

// GetSheetName 获取工作表名称
func (r *ExcelReader) GetSheetName() string {
	return r.sheetName
}

// SetSheetName 设置工作表名称
func (r *ExcelReader) SetSheetName(sheetName string) *ExcelReader {
	r.sheetName = sheetName
	return r
}

// GetOriginalRow 获取读取起始行
func (r *ExcelReader) GetOriginalRow() int {
	return r.originalRow
}

// SetOriginalRow 设置读取起始行
func (r *ExcelReader) SetOriginalRow(originalRow int) *ExcelReader {
	r.originalRow = originalRow - 1
	return r
}

// GetFinishedRow 获取读取终止行
func (r *ExcelReader) GetFinishedRow() int {
	return r.finishedRow
}

// SetFinishedRow 设置读取终止行
func (r *ExcelReader) SetFinishedRow(finishedRow int) *ExcelReader {
	r.finishedRow = finishedRow - 1
	return r
}

// GetTitleRow 获取表头行
func (r *ExcelReader) GetTitleRow() int {
	return r.titleRow
}

// SetTitleRow 设置表头行
func (r *ExcelReader) SetTitleRow(titleRow int) *ExcelReader {
	r.titleRow = titleRow - 1
	return r
}

// GetTitle 获取表头
func (r *ExcelReader) GetTitle() []string {
	return r.titles
}

// SetTitle 设置表头
func (r *ExcelReader) SetTitle(titles []string) *ExcelReader {
	if len(titles) == 0 {
		panic(errors.New("表头不能为空"))
	}
	r.titles = titles
	return r
}

// OpenFile 打开文件
func (r *ExcelReader) OpenFile(filename string, more ...any) *ExcelReader {
	if filename == "" {
		panic(errors.New("文件名不能为空"))
	}
	f, err := excelize.OpenFile(fmt.Sprintf(filename, more...))
	if err != nil {
		panic(fmt.Errorf("打开文件错误：%s", err.Error()))
	}
	r.excel = f

	defer func() {
		if err := r.excel.Close(); err != nil {
			panic(errors.New("文件关闭错误"))
		}
	}()

	r.SetTitleRow(1)
	r.SetOriginalRow(2)
	r.data = make(map[uint64][]string)

	return r
}

// ReadTitle 读取表头
func (r *ExcelReader) ReadTitle() *ExcelReader {
	if r.GetSheetName() == "" {
		panic(errors.New("未设置工作表名称"))
	}

	if rows, err := r.excel.GetRows(r.GetSheetName()); err != nil {
		panic(fmt.Errorf("读取表头错误：%s", err.Error()))
	} else {
		r.SetTitle(rows[r.GetTitleRow()])
	}

	return r
}

// Read 读取Excel
func (r *ExcelReader) Read() *ExcelReader {
	if r.GetSheetName() == "" {
		panic(errors.New("未设置工作表名称"))
	}

	if rows, err := r.excel.GetRows(r.GetSheetName()); err != nil {
		panic(errors.New("读取数据错误：%s"))
	} else {
		if r.finishedRow == 0 {
			r.content = rows[r.GetOriginalRow():]
		} else {
			r.content = rows[r.GetOriginalRow():r.GetFinishedRow()]
		}

		for rowNumber, row := range r.content {
			r.SetDataByRow(uint64(rowNumber), row)
		}
	}

	return r
}

// ToDataFrameDefaultType 获取DataFrame类型数据 通过Excel表头自定义数据类型
func (r *ExcelReader) ToDataFrameDefaultType() dataframe.DataFrame {
	titleWithType := make(map[string]series.Type)
	for _, title := range r.GetTitle() {
		titleWithType[title] = series.String
	}

	return r.ToDataFrame(titleWithType)
}

// ToDataFrame 获取DataFrame类型数据
func (r *ExcelReader) ToDataFrame(titleWithType map[string]series.Type) dataframe.DataFrame {
	if r.GetSheetName() == "" {
		panic(errors.New("未设置工作表名称"))
	}

	var _content [][]string

	if rows, err := r.excel.GetRows(r.GetSheetName()); err != nil {
		panic(errors.New("读取数据错误"))
	} else {
		if r.finishedRow == 0 {
			_content = rows[r.GetTitleRow():]
		} else {
			_content = rows[r.GetTitleRow():r.GetFinishedRow()]
		}
	}

	return dataframe.LoadRecords(
		_content,
		dataframe.DetectTypes(false),
		dataframe.DefaultType(series.String),
		dataframe.WithTypes(titleWithType),
	)
}

// ToDataFrameDetectType 获取DataFrame类型数据 通过自动探寻数据类型
func (r *ExcelReader) ToDataFrameDetectType() dataframe.DataFrame {
	if r.GetSheetName() == "" {
		panic(errors.New("未设置工作表名称"))
	}

	var _content [][]string

	if rows, err := r.excel.GetRows(r.GetSheetName()); err != nil {
		panic(errors.New("读取数据错误"))
	} else {
		if r.finishedRow == 0 {
			_content = rows[r.GetTitleRow():]
		} else {
			_content = rows[r.GetTitleRow():r.GetFinishedRow()]
		}
	}

	return dataframe.LoadRecords(
		_content,
		dataframe.DetectTypes(true),
		dataframe.DefaultType(series.String),
	)
}
