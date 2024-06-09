package excel

import (
	"errors"
	"fmt"

	"github.com/xuri/excelize/v2"
)

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
