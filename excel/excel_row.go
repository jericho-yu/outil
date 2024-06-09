package excel

import (
	"errors"
	"fmt"

	"github.com/xuri/excelize/v2"
)

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
