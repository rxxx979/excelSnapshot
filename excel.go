package main

import (
	"fmt"
	"image/color"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func getCellDisplayValue(f *excelize.File, sheetName, axis string, forceRaw bool) string {
	if forceRaw {
		cell, err := f.GetCellValue(sheetName, axis)
		if err != nil {
			return ""
		}
		return cell
	}

	cell, err := f.GetCellValue(sheetName, axis)
	if err != nil {
		return ""
	}

	if cell == "" {
		formula, err := f.GetCellFormula(sheetName, axis)
		if err == nil && formula != "" {
			calcVal, err := f.CalcCellValue(sheetName, axis)
			if err == nil && calcVal != "" {
				return calcVal
			}
		}
	}
	return cell
}

func listNonEmptySheets(xlsxPath string) ([]string, error) {
	f, err := excelize.OpenFile(xlsxPath)
	if err != nil {
		return nil, fmt.Errorf("无法打开 Excel 文件: %w", err)
	}
	defer f.Close()

	var nonEmptySheets []string
	for _, sheet := range f.GetSheetList() {
		rows, err := f.GetRows(sheet)
		if err != nil {
			continue
		}

		hasContent := false
		for _, row := range rows {
			for _, cell := range row {
				if strings.TrimSpace(cell) != "" {
					hasContent = true
					break
				}
			}
			if hasContent {
				break
			}
		}

		if hasContent {
			nonEmptySheets = append(nonEmptySheets, sheet)
		}
	}

	return nonEmptySheets, nil
}

func getMergeCells(f *excelize.File, sheetName string) (map[string][4]int, error) {
	mergeCells, err := f.GetMergeCells(sheetName)
	if err != nil {
		return nil, fmt.Errorf("无法获取合并单元格信息: %w", err)
	}

	mergedRects := make(map[string][4]int)
	for _, mergeCell := range mergeCells {
		startCell := mergeCell.GetStartAxis()
		endCell := mergeCell.GetEndAxis()
		startCol, startRow, _ := excelize.CellNameToCoordinates(startCell)
		endCol, endRow, _ := excelize.CellNameToCoordinates(endCell)

		for r := startRow - 1; r < endRow; r++ {
			for c := startCol - 1; c < endCol; c++ {
				key := fmt.Sprintf("%d,%d", r, c)
				mergedRects[key] = [4]int{startRow - 1, startCol - 1, endRow - 1, endCol - 1}
			}
		}
	}

	return mergedRects, nil
}

func getColumnName(col int) string {
	name, _ := excelize.ColumnNumberToName(col)
	return name
}

// CellStyle 表示单元格样式信息
type CellStyle struct {
	FontName     string
	FontSize     float64
	FontColor    color.RGBA
	Bold         bool
	Italic       bool
	Underline    bool
	BgColor      color.RGBA
	BorderTop    BorderStyle
	BorderBottom BorderStyle
	BorderLeft   BorderStyle
	BorderRight  BorderStyle
	HAlign       string // 水平对齐: left, center, right
	VAlign       string // 垂直对齐: top, middle, bottom
}

// BorderStyle 表示边框样式
type BorderStyle struct {
	Style string     // 边框样式: thin, medium, thick等
	Color color.RGBA // 边框颜色
}

// 获取单元格样式信息（完整Excel样式读取）
func getCellStyle(f *excelize.File, sheetName, axis string) CellStyle {
	style := CellStyle{
		FontName:  "Calibri",
		FontSize:  11.0,                           // Excel默认字体大小
		FontColor: color.RGBA{0, 0, 0, 255},       // 黑色
		BgColor:   color.RGBA{255, 255, 255, 255}, // 白色背景
		HAlign:    "general",                      // Excel默认对齐方式
		VAlign:    "bottom",                       // Excel默认垂直对齐
	}

	// 获取单元格样式ID
	styleID, err := f.GetCellStyle(sheetName, axis)
	if err != nil {
		return style
	}

	// 获取样式详细信息
	styleInfo, err := f.GetStyle(styleID)
	if err != nil {
		return style
	}

	// 解析字体信息 - 严格按照Excel样式
	if styleInfo.Font != nil {
		if styleInfo.Font.Family != "" {
			style.FontName = styleInfo.Font.Family
		}
		if styleInfo.Font.Size > 0 {
			style.FontSize = styleInfo.Font.Size
		}
		style.Bold = styleInfo.Font.Bold
		style.Italic = styleInfo.Font.Italic
		if styleInfo.Font.Underline != "" {
			style.Underline = true
		}
		// 解析字体颜色
		if styleInfo.Font.Color != "" {
			style.FontColor = parseExcelColor(styleInfo.Font.Color)
		}
	}

	// 解析填充信息 - 严格按照Excel背景色
	if styleInfo.Fill.Type != "" && len(styleInfo.Fill.Color) > 0 {
		// 处理实体填充
		if styleInfo.Fill.Type == "pattern" && len(styleInfo.Fill.Color) > 0 {
			style.BgColor = parseExcelColor(styleInfo.Fill.Color[0])
		}
	}

	// 解析对齐信息 - 严格按照Excel对齐
	if styleInfo.Alignment != nil {
		if styleInfo.Alignment.Horizontal != "" {
			style.HAlign = styleInfo.Alignment.Horizontal
		}
		if styleInfo.Alignment.Vertical != "" {
			style.VAlign = styleInfo.Alignment.Vertical
		}
	}

	// 解析边框信息
	if styleInfo.Border != nil {
		style.BorderTop = parseBorderStyle(styleInfo.Border[0])
		style.BorderLeft = parseBorderStyle(styleInfo.Border[1])
		style.BorderBottom = parseBorderStyle(styleInfo.Border[2])
		style.BorderRight = parseBorderStyle(styleInfo.Border[3])
	}

	return style
}

// 解析Excel颜色字符串为RGBA（增强版本）
func parseExcelColor(colorStr string) color.RGBA {
	if colorStr == "" {
		return color.RGBA{0, 0, 0, 255}
	}

	// 移除#前缀
	colorStr = strings.TrimPrefix(colorStr, "#")

	// 处理Excel的特殊颜色格式
	colorStr = strings.ToUpper(colorStr)

	// Excel常见颜色映射
	excelColors := map[string]color.RGBA{
		"FF000000": {0, 0, 0, 255},       // 黑色
		"FFFFFFFF": {255, 255, 255, 255}, // 白色
		"FFFF0000": {255, 0, 0, 255},     // 红色
		"FF00FF00": {0, 255, 0, 255},     // 绿色
		"FF0000FF": {0, 0, 255, 255},     // 蓝色
		"FFFFFF00": {255, 255, 0, 255},   // 黄色
		"FFFF00FF": {255, 0, 255, 255},   // 品红
		"FF00FFFF": {0, 255, 255, 255},   // 青色
	}

	// 检查是否为预定义颜色
	if rgba, exists := excelColors[colorStr]; exists {
		return rgba
	}

	// 解析8位ARGB格式 (FFRRGGBB)
	if len(colorStr) == 8 {
		if r, err := strconv.ParseUint(colorStr[2:4], 16, 8); err == nil {
			if g, err := strconv.ParseUint(colorStr[4:6], 16, 8); err == nil {
				if b, err := strconv.ParseUint(colorStr[6:8], 16, 8); err == nil {
					return color.RGBA{uint8(r), uint8(g), uint8(b), 255}
				}
			}
		}
	}

	// 解析6位RGB格式 (RRGGBB)
	if len(colorStr) == 6 {
		if r, err := strconv.ParseUint(colorStr[0:2], 16, 8); err == nil {
			if g, err := strconv.ParseUint(colorStr[2:4], 16, 8); err == nil {
				if b, err := strconv.ParseUint(colorStr[4:6], 16, 8); err == nil {
					return color.RGBA{uint8(r), uint8(g), uint8(b), 255}
				}
			}
		}
	}

	return color.RGBA{0, 0, 0, 255}
}

// 解析颜色字符串为RGBA（保留原函数作为备用）
func parseColor(colorStr string) color.RGBA {
	return parseExcelColor(colorStr)
}

// 解析边框样式（简化版本）
func parseBorderStyle(border excelize.Border) BorderStyle {
	style := BorderStyle{
		Style: "none",
		Color: color.RGBA{0, 0, 0, 255},
	}

	// 简化边框样式处理
	if border.Style > 0 {
		style.Style = "thin"
	}
	if border.Color != "" {
		style.Color = parseColor(border.Color)
	}

	return style
}
