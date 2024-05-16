package array

import (
	"reflect"
	"strings"
)

// JoinWithoutEmpty 去掉空值然后合并
func JoinWithoutEmpty(values []string, sep string) string {
	return strings.Join(RemoveEmpty(values), sep)
}

// RemoveEmpty 去掉数组中的空字符串
func RemoveEmpty[T comparable](slice []T) []T {
	j := 0
	for _, item := range slice {
		ref := reflect.ValueOf(item)
		if !ref.IsZero() {
			slice[j] = item
			j++
		}
	}
	return slice[:j]
}

// In 判断元素是否存在数组中
func In[T comparable](target T, elements []T) bool {
	for _, element := range elements {
		if target == element {
			return true
		}
	}
	return false
}

// Filter 过滤数组
func Filter[T any](fn func(v T) bool, values []T) (ret []T) {
	for _, value := range values {
		b := fn(value)
		if b {
			ret = append(ret, value)
		}

	}
	return
}

func FilterDemo() {
	type A struct {
		Name string
	}
	a := []*A{
		{Name: "1"},
		{Name: "2"},
		{Name: "3"},
	}
	b := Filter[*A](func(a *A) bool {
		if a.Name != "1" {
			return true
		}
		return false
	}, a)

	for _, item := range b {
		println(item.Name)
	}
}

// Max 判断数组中最大值
func Max[T int | int8 | int16 | int32 | int64 |
	uint | uint8 | uint16 | uint32 | uint64 |
	float32 | float64](values []T) (max T) {
	for _, value := range values {
		if value > max {
			max = value
		}
	}
	return
}

// Min 判断数组中最小值
func Min[T int | int8 | int16 | int32 | int64 |
	uint | uint8 | uint16 | uint32 | uint64 |
	float32 | float64](values []T) (min T) {
	for _, value := range values {
		if value < min {
			min = value
		}
	}
	return
}

// Sum 获取总和
func Sum[T int | int8 | int16 | int32 | int64 |
	uint | uint8 | uint16 | uint32 | uint64 |
	float32 | float64](numbers []T) (sum T) {
	for _, num := range numbers {
		sum += num
	}
	return sum
}

// All 判断切片中是否全部是非零值
func All[T comparable](values []T) bool {
	for _, value := range values {
		ref := reflect.ValueOf(value)
		if ref.IsZero() {
			return false
		}
	}
	return true
}

// Any 判断切片中是否包含非零值
func Any[T comparable](values []T) bool {
	for _, value := range values {
		ref := reflect.ValueOf(value)
		if !ref.IsZero() {
			return true
		}
	}
	return false
}

// NotEmptyLen 判断切片非零值的长度
func NotEmptyLen[T comparable](values []T) int {
	return len(RemoveEmpty(values))
}

// RemoveTarget 删除数组中对应的目标
func RemoveTarget[T comparable](values []T, target T) (ret []T) {
	for _, value := range values {
		if value != target {
			ret = append(ret, value)
		}
	}
	return ret
}

// RemoveTargets 删除数组中对应的多个目标
func RemoveTargets[T comparable](values []T, targets ...T) (ret []T) {
	for _, value := range values {
		if !In[T](value, targets) {
			ret = append(ret, value)
		}
	}
	return ret
}
