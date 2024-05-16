package str

// PadLeftZeros 前置补零
func PadLeftZeros(str string, length int) string {
	for len(str) < length {
		str = "0" + str
	}
	return str
}
