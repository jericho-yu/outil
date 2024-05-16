package format

import "encoding/json"

// AnyToStruct 任意类型转结构体
func AnyToStruct(v any, r any) error {
	var (
		e        error
		jsonByte []byte
	)
	jsonByte, e = json.Marshal(v)
	if e != nil {
		return e
	}

	return json.Unmarshal(jsonByte, &r)
}
