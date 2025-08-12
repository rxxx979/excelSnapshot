package main

import (
	"embed"
	"fmt"
	"log"
	"path/filepath"
	"strings"
	"sync"

	"github.com/fogleman/gg"
	"golang.org/x/image/font/opentype"
)

//go:embed static/*
var staticFS embed.FS

var (
	fontCache      = make(map[string]*opentype.Font)
	fontCacheMutex sync.RWMutex
	defaultFont    *opentype.Font
	fontInitOnce   sync.Once
)

const (
	DefaultFontSize = 11.0 // Excel默认字体大小
	DefaultDPI      = 96
)

// 初始化字体系统，使用embed的static目录字体文件
func initFonts() {
	fontInitOnce.Do(func() {
		// 扫描embed的static目录下的所有字体文件
		entries, err := staticFS.ReadDir("static")
		if err != nil {
			log.Fatal("无法读取嵌入的static目录")
		}

		// 加载所有找到的字体，优先选择中文字体作为默认字体
		loadedCount := 0
		var chineseFont *opentype.Font
		var chineseFontName string

		for _, entry := range entries {
			if !entry.IsDir() {
				name := entry.Name()
				ext := strings.ToLower(filepath.Ext(name))
				if ext == ".ttf" || ext == ".ttc" || ext == ".otf" || ext == ".otc" {
					fontPath := "static/" + name
					log.Printf("发现嵌入字体文件: %s", name)

					if font := loadEmbedFont(fontPath); font != nil {
						fontCache[name] = font
						log.Printf("已加载嵌入字体: %s", name)
						loadedCount++

						// 检查是否是中文字体
						fileName := strings.ToLower(name)
						if chineseFont == nil && (strings.Contains(fileName, "cjk") ||
							strings.Contains(fileName, "sc") ||
							strings.Contains(fileName, "chinese") ||
							strings.Contains(fileName, "simsun") ||
							strings.Contains(fileName, "yahei") ||
							strings.Contains(fileName, "pingfang")) {
							chineseFont = font
							chineseFontName = name
						}

						// 如果还没有默认字体，先设置一个
						if defaultFont == nil {
							defaultFont = font
						}
					}
				}
			}
		}

		// 如果找到了中文字体，优先使用作为默认字体
		if chineseFont != nil {
			defaultFont = chineseFont
			log.Printf("使用中文字体作为默认字体: %s", chineseFontName)
		} else if defaultFont != nil {
			log.Printf("使用第一个可用字体作为默认字体")
		}

		if defaultFont == nil {
			log.Fatal("嵌入的static目录中未找到可用的字体文件")
		}

		log.Printf("字体系统初始化完成，共加载 %d 个嵌入字体", loadedCount)
	})
}

// 加载嵌入的字体文件
func loadEmbedFont(fontPath string) *opentype.Font {
	data, err := staticFS.ReadFile(fontPath)
	if err != nil {
		log.Printf("[DEBUG] 检查嵌入字体文件: %s - %v", fontPath, err)
		return nil
	}

	// 尝试解析为单字体文件
	if font, err := opentype.Parse(data); err == nil {
		return font
	}

	// 尝试解析为字体集合文件
	if collection, err := opentype.ParseCollection(data); err == nil {
		if collection.NumFonts() > 0 {
			if font, err := collection.Font(0); err == nil {
				return font
			}
		}
	}

	return nil
}

// 设置字体面，使用默认字体
func setFontFace(dc *gg.Context, size float64) error {
	return setFontFaceWithName(dc, "", size)
}

// 设置指定名称的字体面
func setFontFaceWithName(dc *gg.Context, fontName string, size float64) error {
	if defaultFont == nil {
		return fmt.Errorf("字体系统未初始化")
	}

	var font *opentype.Font = defaultFont

	// 如果指定了字体名称，尝试从缓存中获取
	if fontName != "" {
		fontCacheMutex.RLock()
		for path, cachedFont := range fontCache {
			fileName := filepath.Base(path)
			if strings.Contains(strings.ToLower(fileName), strings.ToLower(fontName)) {
				font = cachedFont
				break
			}
		}
		fontCacheMutex.RUnlock()
	}

	// 创建字体面
	face, err := opentype.NewFace(font, &opentype.FaceOptions{
		Size: size,
		DPI:  DefaultDPI,
	})
	if err != nil {
		return fmt.Errorf("创建字体面失败: %w", err)
	}

	dc.SetFontFace(face)
	return nil
}
