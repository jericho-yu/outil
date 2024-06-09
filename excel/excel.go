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
func (receiver *ExcelReader) AutoRead(filename string, values ...any) *ExcelReader {
	return receiver.
		OpenFile(filename, values...).
		SetOriginalRow(2).
		SetTitleRow(1).
		SetSheetName("Sheet1").
		ReadTitle().
		Read()
}

// AutoReadBySheetName 自动读取（默认第一行是表头，从第二行开始）
func (receiver *ExcelReader) AutoReadBySheetName(filename, sheetName string, values ...any) *ExcelReader {
	return receiver.
		OpenFile(filename, values...).
		SetOriginalRow(2).
		SetTitleRow(1).
		SetSheetName(sheetName).
		ReadTitle().
		Read()
}

// ToList 获取数据（数组类型）
func (receiver *ExcelReader) ToList() map[uint64][]string {
	return receiver.data
}

// ToMap 获取数据（map类型）
func (receiver *ExcelReader) ToMap() map[uint64]map[string]string {
	if len(receiver.GetTitle()) == 0 {
		panic(errors.New("未设置表头"))
	}

	_data := make(map[uint64]map[string]string)

	for rowNumber, row := range receiver.ToList() {
		if len(receiver.GetTitle()) != len(row) {
			panic(fmt.Errorf("表头数量与实际数据列不匹配（第%d行）", rowNumber))
		}

		_row := make(map[string]string)
		for k, v := range row {
			_row[receiver.GetTitle()[k]] = v
		}
		_data[rowNumber] = make(map[string]string)
		_data[rowNumber] = _row
	}

	return _data
}

// SetDataByRow 设置单行数据
func (receiver *ExcelReader) SetDataByRow(rowNumber uint64, data []string) *ExcelReader {
	receiver.data[rowNumber+1] = data
	return receiver
}

// GetSheetName 获取工作表名称
func (receiver *ExcelReader) GetSheetName() string {
	return receiver.sheetName
}

// SetSheetName 设置工作表名称
func (receiver *ExcelReader) SetSheetName(sheetName string) *ExcelReader {
	receiver.sheetName = sheetName
	return receiver
}

// GetOriginalRow 获取读取起始行
func (receiver *ExcelReader) GetOriginalRow() int {
	return receiver.originalRow
}

// SetOriginalRow 设置读取起始行
func (receiver *ExcelReader) SetOriginalRow(originalRow int) *ExcelReader {
	receiver.originalRow = originalRow - 1
	return receiver
}

// GetFinishedRow 获取读取终止行
func (receiver *ExcelReader) GetFinishedRow() int {
	return receiver.finishedRow
}

// SetFinishedRow 设置读取终止行
func (receiver *ExcelReader) SetFinishedRow(finishedRow int) *ExcelReader {
	receiver.finishedRow = finishedRow - 1
	return receiver
}

// GetTitleRow 获取表头行
func (receiver *ExcelReader) GetTitleRow() int {
	return receiver.titleRow
}

// SetTitleRow 设置表头行
func (receiver *ExcelReader) SetTitleRow(titleRow int) *ExcelReader {
	receiver.titleRow = titleRow - 1
	return receiver
}

// GetTitle 获取表头
func (receiver *ExcelReader) GetTitle() []string {
	return receiver.titles
}

// SetTitle 设置表头
func (receiver *ExcelReader) SetTitle(titles []string) *ExcelReader {
	if len(titles) == 0 {
		panic(errors.New("表头不能为空"))
	}
	receiver.titles = titles
	return receiver
}

// OpenFile 打开文件
func (receiver *ExcelReader) OpenFile(filename string, more ...any) *ExcelReader {
	if filename == "" {
		panic(errors.New("文件名不能为空"))
	}
	f, err := excelize.OpenFile(fmt.Sprintf(filename, more...))
	if err != nil {
		panic(fmt.Errorf("打开文件错误：%s", err.Error()))
	}
	receiver.excel = f

	defer func() {
		if err := receiver.excel.Close(); err != nil {
			panic(errors.New("文件关闭错误"))
		}
	}()

	receiver.SetTitleRow(1)
	receiver.SetOriginalRow(2)
	receiver.data = make(map[uint64][]string)

	return receiver
}

// ReadTitle 读取表头
func (receiver *ExcelReader) ReadTitle() *ExcelReader {
	if receiver.GetSheetName() == "" {
		panic(errors.New("未设置工作表名称"))
	}

	if rows, err := receiver.excel.GetRows(receiver.GetSheetName()); err != nil {
		panic(fmt.Errorf("读取表头错误：%s", err.Error()))
	} else {
		receiver.SetTitle(rows[receiver.GetTitleRow()])
	}

	return receiver
}

// Read 读取Excel
func (receiver *ExcelReader) Read() *ExcelReader {
	if receiver.GetSheetName() == "" {
		panic(errors.New("未设置工作表名称"))
	}

	if rows, err := receiver.excel.GetRows(receiver.GetSheetName()); err != nil {
		panic(errors.New("读取数据错误：%s"))
	} else {
		if receiver.finishedRow == 0 {
			receiver.content = rows[receiver.GetOriginalRow():]
		} else {
			receiver.content = rows[receiver.GetOriginalRow():receiver.GetFinishedRow()]
		}

		for rowNumber, row := range receiver.content {
			receiver.SetDataByRow(uint64(rowNumber), row)
		}
	}

	return receiver
}

// ToDataFrameDefaultType 获取DataFrame类型数据 通过Excel表头自定义数据类型
func (receiver *ExcelReader) ToDataFrameDefaultType() dataframe.DataFrame {
	titleWithType := make(map[string]series.Type)
	for _, title := range receiver.GetTitle() {
		titleWithType[title] = series.String
	}

	return receiver.ToDataFrame(titleWithType)
}

// ToDataFrame 获取DataFrame类型数据
func (receiver *ExcelReader) ToDataFrame(titleWithType map[string]series.Type) dataframe.DataFrame {
	if receiver.GetSheetName() == "" {
		panic(errors.New("未设置工作表名称"))
	}

	var _content [][]string

	if rows, err := receiver.excel.GetRows(receiver.GetSheetName()); err != nil {
		panic(errors.New("读取数据错误"))
	} else {
		if receiver.finishedRow == 0 {
			_content = rows[receiver.GetTitleRow():]
		} else {
			_content = rows[receiver.GetTitleRow():receiver.GetFinishedRow()]
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
func (receiver *ExcelReader) ToDataFrameDetectType() dataframe.DataFrame {
	if receiver.GetSheetName() == "" {
		panic(errors.New("未设置工作表名称"))
	}

	var _content [][]string

	if rows, err := receiver.excel.GetRows(receiver.GetSheetName()); err != nil {
		panic(errors.New("读取数据错误"))
	} else {
		if receiver.finishedRow == 0 {
			_content = rows[receiver.GetTitleRow():]
		} else {
			_content = rows[receiver.GetTitleRow():receiver.GetFinishedRow()]
		}
	}

	return dataframe.LoadRecords(
		_content,
		dataframe.DetectTypes(true),
		dataframe.DefaultType(series.String),
	)
}

// ExcelWriter Excel写入器
type ExcelWriter struct {
	filename  string
	excel     *excelize.File
	sheetName string
}

// NewExcelWriter 初始化
func NewExcelWriter(filename string, a ...any) *ExcelWriter {
	return (&ExcelWriter{}).Init(fmt.Sprintf(filename, a...))
}

// GetFilename 获取文件名
func (receiver *ExcelWriter) GetFilename() string {
	return receiver.filename
}

// SetFilename 设置文件名
func (receiver *ExcelWriter) SetFilename(filename string) *ExcelWriter {
	receiver.filename = filename
	return receiver
}

// Init 初始化
func (receiver *ExcelWriter) Init(filename string) *ExcelWriter {
	if filename == "" {
		panic(errors.New("文件名不能为空"))
	}
	receiver.filename = filename
	receiver.excel = excelize.NewFile()

	return receiver
}

// CreateSheet 创建工作表
func (receiver *ExcelWriter) CreateSheet(sheetName string) *ExcelWriter {
	if sheetName == "" {
		panic(errors.New("工作表名称不能为空"))
	}
	sheetIndex := receiver.excel.NewSheet(sheetName)
	receiver.excel.SetActiveSheet(sheetIndex)
	receiver.sheetName = receiver.excel.GetSheetName(sheetIndex)

	return receiver
}

// ActiveSheetByName 选择工作表（根据名称）
func (receiver *ExcelWriter) ActiveSheetByName(sheetName string) *ExcelWriter {
	if sheetName == "" {
		panic(errors.New("工作表名称不能为空"))
	}
	sheetIndex := receiver.excel.GetSheetIndex(sheetName)
	receiver.excel.SetActiveSheet(sheetIndex)
	receiver.sheetName = sheetName

	return receiver
}

// ActiveSheetByIndex 选择工作表（根据编号）
func (receiver *ExcelWriter) ActiveSheetByIndex(sheetIndex int) *ExcelWriter {
	if sheetIndex < 0 {
		panic(errors.New("工作表索引不能小于0"))
	}
	receiver.excel.SetActiveSheet(sheetIndex)
	receiver.sheetName = receiver.excel.GetSheetName(sheetIndex)
	return receiver
}

// setStyleFont 设置字体
func (receiver *ExcelWriter) setStyleFont(cell *ExcelCell) {
	if style, err := receiver.excel.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold:   cell.GetFontBold(),
			Italic: cell.GetFontItalic(),
			Family: cell.GetFontFamily(),
			Size:   cell.GetFontSize(),
			Color:  cell.GetFontColor(),
		},
	}); err != nil {
		panic(fmt.Errorf("设置字体错误：%s", cell.GetCoordinate()))
	} else {
		if err = receiver.excel.SetCellStyle(receiver.sheetName, cell.GetCoordinate(), cell.GetCoordinate(), style); err != nil {
			panic(fmt.Errorf("设置字体错误：%s", cell.GetCoordinate()))
		}
	}
}

// SetRows 设置行数据
func (receiver *ExcelWriter) SetRows(excelRows []*ExcelRow) *ExcelWriter {
	for _, row := range excelRows {
		receiver.AddRow(row)
	}
	return receiver
}

// AddRow 增加一行行数据
func (receiver *ExcelWriter) AddRow(excelRow *ExcelRow) *ExcelWriter {
	for _, cell := range excelRow.GetCells() {
		switch cell.GetContentType() {
		case ExcelCellContentTypeFormula:
			if err := receiver.excel.SetCellFormula(receiver.sheetName, cell.GetCoordinate(), cell.GetContent().(string)); err != nil {
				panic(fmt.Errorf("写入数据错误（公式）%s %s：%v", cell.GetCoordinate(), cell.GetContent(), err.Error()))
			}
		case ExcelCellContentTypeInt:
			if err := receiver.excel.SetCellInt(receiver.sheetName, cell.GetCoordinate(), cell.GetContent().(int)); err != nil {
				panic(fmt.Errorf("写入数据错误（数字）%s %s：%v", cell.GetCoordinate(), cell.GetContent(), err.Error()))
			}
		case ExcelCellContentTypeFloat64:
			if err := receiver.excel.SetCellFloat(receiver.sheetName, cell.GetCoordinate(), cell.GetContent().(float64), 4, 64); err != nil {
				panic(fmt.Errorf("写入数据错误（小数）%s %s：%v", cell.GetCoordinate(), cell.GetContent(), err.Error()))
			}
		case ExcelCellContentTypeBool:
			if err := receiver.excel.SetCellBool(receiver.sheetName, cell.GetCoordinate(), cell.GetContent().(bool)); err != nil {
				panic(fmt.Errorf("写入数据错误（布尔）%s %s：%v", cell.GetCoordinate(), cell.GetContent(), err.Error()))
			}
		default:
			if err := receiver.excel.SetCellValue(receiver.sheetName, cell.GetCoordinate(), cell.GetContent()); err != nil {
				panic(fmt.Errorf("写入数据错误（默认）%s %s：%v", cell.GetCoordinate(), cell.GetContent(), err.Error()))
			}
		}
		receiver.setStyleFont(cell)
	}

	return receiver
}

// SetTitleRow 设置标题行
func (receiver *ExcelWriter) SetTitleRow(titles []string, rowNumber uint64) *ExcelWriter {
	var (
		titleRow   *ExcelRow
		titleCells = make([]*ExcelCell, len(titles))
	)

	if len(titles) > 0 {
		for idx, title := range titles {
			titleCells[idx] = NewExcelCellAny(title)
		}

		titleRow = NewExcelRow().SetRowNumber(rowNumber).SetCells(titleCells)

		receiver.AddRow(titleRow)
	}

	return receiver
}

// Save 保存文件
func (receiver *ExcelWriter) Save() error {
	if receiver.filename == "" {
		panic(errors.New("未设置文件名"))
	}
	return receiver.excel.SaveAs(receiver.filename)
}

// Download 下载Excel
// func (receiver *ExcelWriter) Download(ctx *gin.Context) error {
// 	ctx.Writer.Header().Set("Content-Type", "application/octet-stream")
// 	ctx.Writer.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", url.QueryEscape(receiver.filename)))
// 	ctx.Writer.Header().Set("Content-Transfer-Encoding", "binary")
// 	ctx.Writer.Header().Set("Access-Control-Expose-Headers", "Content-Disposition")
// 	return receiver.excel.Write(ctx.Writer)
// }

// ColumnIndexToText 列索引转文字
func (receiver *ExcelWriter) ColumnIndexToText(columnIndex int) (string, error) {
	return excelize.ColumnNumberToName(columnIndex)
}

// ExcelRow Excel行
type ExcelRow struct {
	cells     []*ExcelCell
	rowNumber uint64
}

// NewExcelRow 构造函数
func NewExcelRow() *ExcelRow {
	return &ExcelRow{}
}

// GetCells 获取单元格组
func (receiver *ExcelRow) GetCells() []*ExcelCell {
	return receiver.cells
}

// SetCells 设置单元格组
func (receiver *ExcelRow) SetCells(cells []*ExcelCell) *ExcelRow {
	if receiver.GetRowNumber() == 0 {
		panic(errors.New("行标必须大于0"))
	}

	for colNumber, cell := range cells {
		if colText, err := excelize.ColumnNumberToName(colNumber + 1); err != nil {
			panic(fmt.Errorf("列索引转列文字失败：%d，%d", receiver.GetRowNumber(), colNumber+1))
		} else {
			cell.SetCoordinate(fmt.Sprintf("%s%d", colText, receiver.GetRowNumber()))
		}
	}
	receiver.cells = cells

	return receiver
}

// GetRowNumber 获取行标
func (receiver *ExcelRow) GetRowNumber() uint64 {
	return receiver.rowNumber
}

// SetRowNumber 设置行标
func (receiver *ExcelRow) SetRowNumber(rowNumber uint64) *ExcelRow {
	receiver.rowNumber = rowNumber
	return receiver
}

type ExcelCellContentType string

const (
	ExcelCellContentTypeAny     ExcelCellContentType = "any"
	ExcelCellContentTypeFormula ExcelCellContentType = "formula"
	ExcelCellContentTypeInt     ExcelCellContentType = "int"
	ExcelCellContentTypeFloat64 ExcelCellContentType = "float64"
	ExcelCellContentTypeBool    ExcelCellContentType = "bool"
)

// ExcelCell Excel单元格
type ExcelCell struct {
	content     any
	contentType ExcelCellContentType
	coordinate  string
	fontColor   string
	fontBold    bool
	fontItalic  bool
	fontFamily  string
	fontSize    float64
}

// NewExcelCellAny 构造函数（字符串格式）
func NewExcelCellAny(content any) *ExcelCell {
	return &ExcelCell{content: content, contentType: ExcelCellContentTypeAny}
}

// NewExcelCellFormula 构造函数（公式格式）
func NewExcelCellFormula(content string) *ExcelCell {
	return &ExcelCell{content: content, contentType: ExcelCellContentTypeFormula}
}

// NewExcelCellInt 构造函数（整数格式）
func NewExcelCellInt(content int) *ExcelCell {
	return &ExcelCell{content: content, contentType: ExcelCellContentTypeInt}
}

// NewExcelCellFloat64 构造函数（小数格式）
func NewExcelCellFloat64(content float64) *ExcelCell {
	return &ExcelCell{content: content, contentType: ExcelCellContentTypeFloat64}
}

// NewExcelCellBool 构造函数（布尔格式）
func NewExcelCellBool(content bool) *ExcelCell {
	return &ExcelCell{content: content, contentType: ExcelCellContentTypeBool}
}

// GetFontColor 获取字体颜色
func (receiver *ExcelCell) GetFontColor() string {
	return receiver.fontColor
}

// SetFontColor 设置字体颜色
func (receiver *ExcelCell) SetFontColor(fontColor string, condition bool) *ExcelCell {
	if condition {
		receiver.fontColor = fontColor
	}
	return receiver
}

// GetFontBold 获取字体粗体
func (receiver *ExcelCell) GetFontBold() bool {
	return receiver.fontBold
}

// SetFontBold 设置字体粗体
func (receiver *ExcelCell) SetFontBold(fontBold bool, condition bool) *ExcelCell {
	if condition {
		receiver.fontBold = fontBold
	}
	return receiver
}

// GetFontItalic 获取字体斜体
func (receiver *ExcelCell) GetFontItalic() bool {
	return receiver.fontItalic
}

// SetFontItalic 设置字体斜体
func (receiver *ExcelCell) SetFontItalic(fontItalic bool, condition bool) *ExcelCell {
	if condition {
		receiver.fontItalic = fontItalic
	}
	return receiver
}

// GetFontFamily 获取字体
func (receiver *ExcelCell) GetFontFamily() string {
	return receiver.fontFamily
}

// SetFontFamily 设置字体
func (receiver *ExcelCell) SetFontFamily(fontFamily string, condition bool) *ExcelCell {
	if condition {
		receiver.fontFamily = fontFamily
	}
	return receiver
}

// GetFontSize 获取字体字号
func (receiver *ExcelCell) GetFontSize() float64 {
	return receiver.fontSize
}

// SetFontSize 设置字体字号
func (receiver *ExcelCell) SetFontSize(fontSize float64) *ExcelCell {
	receiver.fontSize = fontSize
	return receiver
}

// Init 初始化
func (receiver *ExcelCell) Init(content any) *ExcelCell {
	receiver.content = content
	return receiver
}

// GetContent 获取内容
func (receiver *ExcelCell) GetContent() any {
	return receiver.content
}

// SetContent 设置内容
func (receiver *ExcelCell) SetContent(content any) *ExcelCell {
	receiver.content = content
	return receiver
}

// GetCoordinate 获取单元格坐标
func (receiver *ExcelCell) GetCoordinate() string {
	return receiver.coordinate
}

// SetCoordinate 设置单元格坐标
func (receiver *ExcelCell) SetCoordinate(coordinate string) *ExcelCell {
	receiver.coordinate = coordinate
	return receiver
}

// GetContentType 获取单元格类型
func (receiver *ExcelCell) GetContentType() ExcelCellContentType {
	return receiver.contentType
}

// SetContentType 设置单元格类型
func (receiver *ExcelCell) SetContentType(contentType ExcelCellContentType) *ExcelCell {
	receiver.contentType = contentType
	return receiver
}
