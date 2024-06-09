package excel

type (
	// ExcelCellContentType 单元格内容类型
	ExcelCellContentType string

	// ExcelCell Excel单元格
	ExcelCell struct {
		content     any
		contentType ExcelCellContentType
		coordinate  string
		fontColor   string
		fontBold    bool
		fontItalic  bool
		fontFamily  string
		fontSize    float64
	}
)

const (
	ExcelCellContentTypeAny     ExcelCellContentType = "any"
	ExcelCellContentTypeFormula ExcelCellContentType = "formula"
	ExcelCellContentTypeInt     ExcelCellContentType = "int"
	ExcelCellContentTypeFloat64 ExcelCellContentType = "float64"
	ExcelCellContentTypeBool    ExcelCellContentType = "bool"
)

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
