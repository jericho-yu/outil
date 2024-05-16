package string

import (
	"bytes"
	"crypto/aes"
	"crypto/cipher"
	"crypto/md5"
	"encoding/base64"
	"errors"
	"fmt"
	"otuil/outil/common"
)

type (
	Secret struct{}
)

func (Secret) PKCS7Padding(src []byte, blockSize int) []byte {
	padding := blockSize - len(src)%blockSize
	padtext := bytes.Repeat([]byte{byte(padding)}, padding)
	return append(src, padtext...)
}

func (Secret) PKCS7UnPadding(src []byte, blockSize int) ([]byte, error) {
	length := len(src)
	if blockSize <= 0 {
		panic(fmt.Errorf("invalid blockSize: %d", blockSize))
	}

	if length%blockSize != 0 || length == 0 {
		panic(errors.New("invalid data len"))
	}

	unpadding := int(src[length-1])
	if unpadding > blockSize || unpadding == 0 {
		panic(errors.New("invalid unpadding"))
	}

	padding := src[length-unpadding:]
	for i := 0; i < unpadding; i++ {
		if padding[i] != byte(unpadding) {
			panic(errors.New("invalid padding"))
		}
	}

	return src[:(length - unpadding)], nil
}

func (Secret) EncryptToken(key, secretKey string, iv []byte, randStr ...string) (encryptStr, uuid string, err error) {
	if key == "" {
		return "", "", err
	}
	// 生成随机串
	if len(randStr) > 0 {
		uuid = randStr[0]
	} else {
		uuid = Secret{}.MustEncrypt(Letters(10))
	}

	token, err := Secret{}.EncryptCBC([]byte(key+uuid), []byte(secretKey), iv)
	if err != nil {
		return "", "", err
	}
	encryptStr = string(Secret{}.Encode(token))
	return
}

func (Secret) EncryptCBC(plainText, key, iv []byte, ivs ...[]byte) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	blockSize := block.BlockSize()
	plainText = Secret{}.PKCS7Padding(plainText, blockSize)
	ivValue := ([]byte)(nil)
	if len(ivs) > 0 {
		ivValue = ivs[0]
	} else {
		ivValue = iv
	}
	blockMode := cipher.NewCBCEncrypter(block, ivValue)
	cipherText := make([]byte, len(plainText))
	blockMode.CryptBlocks(cipherText, plainText)

	return cipherText, nil
}

func (Secret) DecryptAuthorization(token, secretKey string, iv []byte) (DecryptStr, uuid string, err error) {

	if token == "" {
		panic(errors.New("decrypt Token empty"))
	}
	token64, err := Secret{}.Decode([]byte(token))
	if err != nil {
		panic(fmt.Errorf("[GFToken]decode error Token:%s %s", token, err.Error()))
	}
	decryptToken, err := Secret{}.DecryptCBC(token64, []byte(secretKey), iv)
	if err != nil {
		panic(fmt.Errorf("[GFToken]decrypt error Token:%s %s", token, err.Error()))
	}
	length := len(decryptToken)
	uuid = string(decryptToken[length-32:])
	DecryptStr = string(decryptToken[:length-32])
	return
}

func (Secret) DecryptCBC(cipherText, key, iv []byte, ivs ...[]byte) ([]byte, error) {
	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}
	blockSize := block.BlockSize()
	if len(cipherText) < blockSize {
		panic(errors.New("cipherText too short"))
	}
	ivValue := ([]byte)(nil)
	if len(ivs) > 0 {
		ivValue = ivs[0]
	} else {
		ivValue = iv
	}
	if len(cipherText)%blockSize != 0 {
		panic(errors.New("cipherText is not a multiple of the block size"))
	}
	blockModel := cipher.NewCBCDecrypter(block, ivValue)
	plainText := make([]byte, len(cipherText))
	blockModel.CryptBlocks(plainText, cipherText)
	plainText, e := Secret{}.PKCS7UnPadding(plainText, blockSize)
	if e != nil {
		return nil, e
	}
	return plainText, nil
}

func (Secret) MustEncrypt(data any) string {
	result, err := Secret{}.Encrypt(data)
	if err != nil {
		panic(err)
	}
	return result
}

func (Secret) Encrypt(data any) (encrypt string, err error) {
	return Secret{}.EncryptBytes(common.ToBytes(data))
}

func (Secret) EncryptBytes(data []byte) (encrypt string, err error) {
	h := md5.New()
	if _, err = h.Write(data); err != nil {
		return "", err
	}
	return fmt.Sprintf("%x", h.Sum(nil)), nil
}

func (Secret) Encode(src []byte) []byte {
	dst := make([]byte, base64.StdEncoding.EncodedLen(len(src)))
	base64.StdEncoding.Encode(dst, src)
	return dst
}

func (Secret) Decode(data []byte) ([]byte, error) {
	var (
		src    = make([]byte, base64.StdEncoding.DecodedLen(len(data)))
		n, err = base64.StdEncoding.Decode(src, data)
	)
	if err != nil {
		panic(errors.New("base64.StdEncoding.Decode failed"))
	}
	return src[:n], err
}
