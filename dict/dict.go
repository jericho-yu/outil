package dict

import (
	"errors"
	"strings"
)

// GetKeys 获取一个字典中所有的key
func GetKeys[T comparable](sources map[T]any) []T {
	keys := make([]T, 0, len(sources))
	for idx, _ := range sources {
		keys = append(keys, idx)
	}
	return keys
}

func GetVal(obj map[string]any, key string, def any) any {
	if key == "" {
		return def
	}
	keys := strings.Split(key, ".")
	if obj == nil || len(keys) == 0 {
		return def
	}

	currentKey := keys[0]
	moreKeys := keys[1:]
	if val, ok := obj[currentKey]; ok {
		if currentKey != "" {
			return GetVal(val.(map[string]any), strings.Join(moreKeys, "."), def)
		} else {
			return val
		}
	}

	return def
}

func SetVal(obj map[string]any, key string, val, def any) {
	if key == "" {
		return
	}
	keys := strings.Split(key, ".")
	currentKey := keys[0]
	moreKeys := keys[1:]

	if currentKey != "" {
		childObj, exists := obj[currentKey].(map[string]any)
		if !exists || childObj == nil {
			obj[currentKey] = make(map[string]any)
		}
		SetVal(obj[currentKey].(map[string]any), strings.Join(moreKeys, "."), val, def)
	} else {
		if val == nil {
			obj[currentKey] = def
		} else {
			obj[currentKey] = val
		}
	}
}

// Filter 过滤数组
func Filter[TVal any, TKey comparable](fn func(v TVal) (bool, TVal), values map[TKey]TVal) (ret []TVal) {
	for idx, value := range values {
		b, _ := fn(value)
		if !b {
			delete(values, idx)
		}

	}
	return
}

// Zip 压缩数据到map
func Zip[TKey ~struct{} | string | int |
	int8 | int16 | int32 | int64 | uint |
	uint8 | uint16 | uint32 | uint64,
	TVal ~struct{} | string | int |
		int8 | int16 | int32 | int64 | uint |
		uint8 | uint16 | uint32 | uint64](keys []TKey, values []TVal) (zip map[TKey]TVal, err error) {
	zip = make(map[TKey]TVal)

	if len(keys) != len(values) {
		return nil, errors.New("keys和values长度不一致")
	}
	for idx, key := range keys {
		zip[key] = values[idx]
	}
	return zip, nil
}
