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
		if len(r.GetTitle()) != len(row) {
			panic(fmt.Errorf("表头数量与实际数据列不匹配（第%d行）", rowNumber))
		}

		_row := make(map[string]string)
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
func (r *ExcelWriter) GetFilename() string {
	return r.filename
}

// SetFilename 设置文件名
func (r *ExcelWriter) SetFilename(filename string) *ExcelWriter {
	r.filename = filename
	return r
}

// Init 初始化
func (r *ExcelWriter) Init(filename string) *ExcelWriter {
	if filename == "" {
		panic(errors.New("文件名不能为空"))
	}
	r.filename = filename
	r.excel = excelize.NewFile()

	return r
}

// CreateSheet 创建工作表
func (r *ExcelWriter) CreateSheet(sheetName string) *ExcelWriter {
	if sheetName == "" {
		panic(errors.New("工作表名称不能为空"))
	}
	sheetIndex := r.excel.NewSheet(sheetName)
	r.excel.SetActiveSheet(sheetIndex)
	r.sheetName = r.excel.GetSheetName(sheetIndex)

	return r
}

// ActiveSheetByName 选择工作表（根据名称）
func (r *ExcelWriter) ActiveSheetByName(sheetName string) *ExcelWriter {
	if sheetName == "" {
		panic(errors.New("工作表名称不能为空"))
	}
	sheetIndex := r.excel.GetSheetIndex(sheetName)
	r.excel.SetActiveSheet(sheetIndex)
	r.sheetName = sheetName

	return r
}

// ActiveSheetByIndex 选择工作表（根据编号）
func (r *ExcelWriter) ActiveSheetByIndex(sheetIndex int) *ExcelWriter {
	if sheetIndex < 0 {
		panic(errors.New("工作表索引不能小于0"))
	}
	r.excel.SetActiveSheet(sheetIndex)
	r.sheetName = r.excel.GetSheetName(sheetIndex)
	return r
}

// setStyleFont 设置字体
func (r *ExcelWriter) setStyleFont(cell *ExcelCell) {
	if style, err := r.excel.NewStyle(&excelize.Style{
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
		if err = r.excel.SetCellStyle(r.sheetName, cell.GetCoordinate(), cell.GetCoordinate(), style); err != nil {
			panic(fmt.Errorf("设置字体错误：%s", cell.GetCoordinate()))
		}
	}
}

// SetRows 设置行数据
func (r *ExcelWriter) SetRows(excelRows []*ExcelRow) *ExcelWriter {
	for _, row := range excelRows {
		r.AddRow(row)
	}
	return r
}

// AddRow 增加一行行数据
func (r *ExcelWriter) AddRow(excelRow *ExcelRow) *ExcelWriter {
	for _, cell := range excelRow.GetCells() {
		switch cell.GetContentType() {
		case ExcelCellContentTypeFormula:
			if err := r.excel.SetCellFormula(r.sheetName, cell.GetCoordinate(), cell.GetContent().(string)); err != nil {
				panic(fmt.Errorf("写入数据错误（公式）%s %s：%v", cell.GetCoordinate(), cell.GetContent(), err.Error()))
			}
		case ExcelCellContentTypeInt:
			if err := r.excel.SetCellInt(r.sheetName, cell.GetCoordinate(), cell.GetContent().(int)); err != nil {
				panic(fmt.Errorf("写入数据错误（数字）%s %s：%v", cell.GetCoordinate(), cell.GetContent(), err.Error()))
			}
		case ExcelCellContentTypeFloat64:
			if err := r.excel.SetCellFloat(r.sheetName, cell.GetCoordinate(), cell.GetContent().(float64), 4, 64); err != nil {
				panic(fmt.Errorf("写入数据错误（小数）%s %s：%v", cell.GetCoordinate(), cell.GetContent(), err.Error()))
			}
		case ExcelCellContentTypeBool:
			if err := r.excel.SetCellBool(r.sheetName, cell.GetCoordinate(), cell.GetContent().(bool)); err != nil {
				panic(fmt.Errorf("写入数据错误（布尔）%s %s：%v", cell.GetCoordinate(), cell.GetContent(), err.Error()))
			}
		default:
			if err := r.excel.SetCellValue(r.sheetName, cell.GetCoordinate(), cell.GetContent()); err != nil {
				panic(fmt.Errorf("写入数据错误（默认）%s %s：%v", cell.GetCoordinate(), cell.GetContent(), err.Error()))
			}
		}
		r.setStyleFont(cell)
	}

	return r
}

// SetTitleRow 设置标题行
func (r *ExcelWriter) SetTitleRow(titles []string, rowNumber uint64) *ExcelWriter {
	var (
		titleRow   *ExcelRow
		titleCells = make([]*ExcelCell, len(titles))
	)

	if len(titles) > 0 {
		for idx, title := range titles {
			titleCells[idx] = NewExcelCellAny(title)
		}

		titleRow = NewExcelRow().SetRowNumber(rowNumber).SetCells(titleCells)

		r.AddRow(titleRow)
	}

	return r
}

// Save 保存文件
func (r *ExcelWriter) Save() error {
	if r.filename == "" {
		panic(errors.New("未设置文件名"))
	}
	return r.excel.SaveAs(r.filename)
}

// Download 下载Excel
// func (r *ExcelWriter) Download(ctx *gin.Context) error {
// 	ctx.Writer.Header().Set("Content-Type", "application/octet-stream")
// 	ctx.Writer.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%s", url.QueryEscape(r.filename)))
// 	ctx.Writer.Header().Set("Content-Transfer-Encoding", "binary")
// 	ctx.Writer.Header().Set("Access-Control-Expose-Headers", "Content-Disposition")
// 	return r.excel.Write(ctx.Writer)
// }

// ColumnIndexToText 列索引转文字
func (r *ExcelWriter) ColumnIndexToText(columnIndex int) (string, error) {
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
func (r *ExcelRow) GetCells() []*ExcelCell {
	return r.cells
}

// SetCells 设置单元格组
func (r *ExcelRow) SetCells(cells []*ExcelCell) *ExcelRow {
	if r.GetRowNumber() == 0 {
		panic(errors.New("行标必须大于0"))
	}

	for colNumber, cell := range cells {
		if colText, err := excelize.ColumnNumberToName(colNumber + 1); err != nil {
			panic(fmt.Errorf("列索引转列文字失败：%d，%d", r.GetRowNumber(), colNumber+1))
		} else {
			cell.SetCoordinate(fmt.Sprintf("%s%d", colText, r.GetRowNumber()))
		}
	}
	r.cells = cells

	return r
}

// GetRowNumber 获取行标
func (r *ExcelRow) GetRowNumber() uint64 {
	return r.rowNumber
}

// SetRowNumber 设置行标
func (r *ExcelRow) SetRowNumber(rowNumber uint64) *ExcelRow {
	r.rowNumber = rowNumber
	return r
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
func (r *ExcelCell) GetFontColor() string {
	return r.fontColor
}

// SetFontColor 设置字体颜色
func (r *ExcelCell) SetFontColor(fontColor string, condition bool) *ExcelCell {
	if condition {
		r.fontColor = fontColor
	}
	return r
}

// GetFontBold 获取字体粗体
func (r *ExcelCell) GetFontBold() bool {
	return r.fontBold
}

// SetFontBold 设置字体粗体
func (r *ExcelCell) SetFontBold(fontBold bool, condition bool) *ExcelCell {
	if condition {
		r.fontBold = fontBold
	}
	return r
}

// GetFontItalic 获取字体斜体
func (r *ExcelCell) GetFontItalic() bool {
	return r.fontItalic
}

// SetFontItalic 设置字体斜体
func (r *ExcelCell) SetFontItalic(fontItalic bool, condition bool) *ExcelCell {
	if condition {
		r.fontItalic = fontItalic
	}
	return r
}

// GetFontFamily 获取字体
func (r *ExcelCell) GetFontFamily() string {
	return r.fontFamily
}

// SetFontFamily 设置字体
func (r *ExcelCell) SetFontFamily(fontFamily string, condition bool) *ExcelCell {
	if condition {
		r.fontFamily = fontFamily
	}
	return r
}

// GetFontSize 获取字体字号
func (r *ExcelCell) GetFontSize() float64 {
	return r.fontSize
}

// SetFontSize 设置字体字号
func (r *ExcelCell) SetFontSize(fontSize float64) *ExcelCell {
	r.fontSize = fontSize
	return r
}

// Init 初始化
func (r *ExcelCell) Init(content any) *ExcelCell {
	r.content = content
	return r
}

// GetContent 获取内容
func (r *ExcelCell) GetContent() any {
	return r.content
}

// SetContent 设置内容
func (r *ExcelCell) SetContent(content any) *ExcelCell {
	r.content = content
	return r
}

// GetCoordinate 获取单元格坐标
func (r *ExcelCell) GetCoordinate() string {
	return r.coordinate
}

// SetCoordinate 设置单元格坐标
func (r *ExcelCell) SetCoordinate(coordinate string) *ExcelCell {
	r.coordinate = coordinate
	return r
}

// GetContentType 获取单元格类型
func (r *ExcelCell) GetContentType() ExcelCellContentType {
	return r.contentType
}

// SetContentType 设置单元格类型
func (r *ExcelCell) SetContentType(contentType ExcelCellContentType) *ExcelCell {
	r.contentType = contentType
	return r
}
