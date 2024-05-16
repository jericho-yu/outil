package encode

import (
	"crypto/md5"
	"encoding/hex"
)

// Md5 编码
func Md5(original []byte) (string, error) {
	hash := md5.New()

	_, err := hash.Write(original)
	if err != nil {
		return "", err
	}

	return hex.EncodeToString(hash.Sum(nil)), nil
}
