package main

import (
	"testing"

	"github.com/fogleman/gg"
)

func TestInitFontsAndSetFace(t *testing.T) {
	initFonts()
	// 创建一个小画布
	dc := gg.NewContext(100, 50)
	if err := setFontFace(dc, 12); err != nil {
		t.Fatalf("setFontFace error: %v", err)
	}
	// 指定名称：即使找不到完全匹配，也不应报错（回退到默认字体）
	if err := setFontFaceWithName(dc, "NonExistFamily", 10); err != nil {
		t.Fatalf("setFontFaceWithName fallback error: %v", err)
	}
}
