package excelsnapshot

import (
	_ "embed"
	"fmt"
	"image/color"
	"strconv"
	"strings"

	"golang.org/x/image/font"
	"golang.org/x/image/font/opentype"
)

//go:embed fonts/SourceHanSerifSC-Regular.otf
var fontBytes []byte

// LoadDefaultFontWithSize 加载默认字体
func LoadDefaultFontWithSize(size float64) (font.Face, error) {
	f, err := opentype.Parse(fontBytes)
	if err != nil {
		return nil, err
	}

	fontFace, err := opentype.NewFace(f, &opentype.FaceOptions{
		Size:    size, // 字体大小也要跟着4倍缩放
		DPI:     72,   // 4倍DPI，匹配4倍缩放
		Hinting: font.HintingFull,
	})
	if err != nil {
		return nil, err
	}

	return fontFace, nil
}

// HexToRGBA 将十六进制颜色转换为 color.RGBA
func HexToRGBA(hex string) (color.RGBA, error) {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return color.RGBA{}, fmt.Errorf("invalid hex color: %s", hex)
	}

	r, err := strconv.ParseUint(hex[0:2], 16, 8)
	if err != nil {
		return color.RGBA{}, err
	}
	g, err := strconv.ParseUint(hex[2:4], 16, 8)
	if err != nil {
		return color.RGBA{}, err
	}
	b, err := strconv.ParseUint(hex[4:6], 16, 8)
	if err != nil {
		return color.RGBA{}, err
	}

	return color.RGBA{
		R: uint8(r),
		G: uint8(g),
		B: uint8(b),
		A: 255, // 默认不透明
	}, nil
}
