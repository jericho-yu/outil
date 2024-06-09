package excel

import (
	"github.com/xuri/excelize/v2"
)

// ColumnIndexToText 列索引转文字
func (r ExcelWriter) ColumnIndexToText(columnIndex int) (string, error) {
	return excelize.ColumnNumberToName(columnIndex)
}
