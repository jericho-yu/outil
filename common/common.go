package common

import (
	"bytes"
	"encoding/gob"
	"reflect"
)

// DeepCopyByGob 深拷贝
func DeepCopyByGob(dst, src any) error {
	var buffer bytes.Buffer
	if err := gob.NewEncoder(&buffer).Encode(src); err != nil {
		return err
	}

	return gob.NewDecoder(&buffer).Decode(dst)
}

// IsSameType 判断两个类型是否相同
func IsSameType(a, b any) bool {
	return reflect.TypeOf(a) == reflect.TypeOf(b)
}
