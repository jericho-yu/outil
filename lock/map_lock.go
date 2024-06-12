package lock

import (
	"fmt"
	"log"
	"sync"
	"time"
)

type (
	// 字典锁：一个锁的集合
	MapLock struct{ locks sync.Map }

	// 锁项：一个集合锁中的每一项，包含：锁状态、锁值、超时时间、定时器
	itemLock struct {
		inUse   bool
		val     any
		timeout time.Duration
		timer   *time.Timer
	}
)

var (
	onceMapLock   sync.Once
	mapLockIns    *MapLock
	MapLockHelper MapLock
)

// New 实例化：字典锁
func (MapLock) New() *MapLock {
	return &MapLock{locks: sync.Map{}}
}

// SingleMapLock 单例化：字典锁
func (MapLock) Single() *MapLock {
	onceMapLock.Do(func() { mapLockIns = &MapLock{locks: sync.Map{}} })
	return mapLockIns
}

// Store 创建锁
func (r *MapLock) Store(key string, val any) error {
	if _, exists := r.locks.LoadOrStore(key, &itemLock{val: val}); exists {
		return fmt.Errorf("锁[%s]已存在", key)
	}
	return nil
}

// StoreMany 批量创建锁
func (r *MapLock) StoreMany(items map[string]any) error {
	for idx, item := range items {
		err := r.Store(idx, item)
		if err != nil {
			r.DestroyAll()
			return err
		}
	}
	return nil
}

// Release 显式锁释放方法
func (r *itemLock) Release() {
	if r.timer != nil {
		r.timer.Stop()
		r.timer = nil
	}
	r.inUse = false
}

// Destroy 删除锁
func (r *MapLock) Destroy(key string) {
	if il, ok := r.locks.Load(key); ok {
		il.(*itemLock).Release()
		r.locks.Delete(key) // 删除键值对，以便垃圾回收
	}
}

// DestroyAll 删除所有锁
func (r *MapLock) DestroyAll() {
	r.locks.Range(func(key, value any) bool {
		r.Destroy(key.(string))
		return true
	})
}

// Lock 获取锁
func (r *MapLock) Lock(key string, timeout time.Duration) (*itemLock, error) {
	if item, exists := r.locks.Load(key); !exists {
		return nil, fmt.Errorf("锁[%s]不存在", key)
	} else {
		if item.(*itemLock).inUse {
			return nil, fmt.Errorf("锁[%s]被占用", key)
		}

		// 设置锁占用
		item.(*itemLock).inUse = true

		// 设置超时时间
		if timeout > 0 {
			item.(*itemLock).timeout = timeout
			item.(*itemLock).timer = time.AfterFunc(timeout, func() {
				if il, ok := r.locks.Load(key); ok {
					if il.(*itemLock).timer != nil {
						il.(*itemLock).Release()
					}
				}
			})
		}

		return item.(*itemLock), nil
	}
}

// Try 尝试获取锁
func (r *MapLock) Try(key string) error {
	if item, exist := r.locks.Load(key); !exist {
		return fmt.Errorf("锁[%s]不存在", key)
	} else {
		if item.(*itemLock).inUse {
			return fmt.Errorf("锁[%s]被占用", key)
		}
		return nil
	}
}

func (MapLock) Demo() {
	k8sLinks := map[string]any{
		"k8s-a": &struct{}{},
		"k8s-b": &struct{}{},
		"k8s-c": &struct{}{},
	}

	// 获取字典锁对象
	ml := MapLockHelper.Single()

	// 批量创建锁
	storeErr := ml.StoreMany(k8sLinks)
	if storeErr != nil {
		// 处理err
		log.Fatalln(storeErr.Error())
	}

	// 检测锁
	tryErr := ml.Try("k8s-a")
	if tryErr != nil {
		// 处理err
		log.Fatalln(tryErr.Error())
	}

	// 获取锁
	lock, lockErr := ml.Lock("k8s-a", time.Second*10) // 10秒业务处理不完也会过期 设置为：0则为永不过期
	if lockErr != nil {
		log.Fatalln(lockErr.Error())
	}
	defer lock.Release()

	// 处理业务...
}
