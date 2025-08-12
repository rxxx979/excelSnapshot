package main

import (
	"fmt"
	"log"
	"strconv"
	"strings"

	"github.com/fogleman/gg"
	"github.com/xuri/excelize/v2"
)

func renderPNGFromExcel(xlsxPath, sheet string, sheetIdx int, outPath string, forceRaw bool) error {
	f, err := excelize.OpenFile(xlsxPath)
	if err != nil {
		return fmt.Errorf("无法打开 Excel 文件: %w", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Printf("关闭 Excel 文件失败: %v", err)
		}
	}()

	var sheetName string
	if sheet != "" {
		sheetName = sheet
	} else {
		sheets := f.GetSheetList()
		if sheetIdx < 0 || sheetIdx >= len(sheets) {
			return fmt.Errorf("工作表索引 %d 越界", sheetIdx)
		}
		sheetName = sheets[sheetIdx]
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("无法读取工作表 %s: %w", sheetName, err)
	}

	if len(rows) == 0 {
		return fmt.Errorf("工作表 %s 为空", sheetName)
	}

	mergedRects, err := getMergeCells(f, sheetName)
	if err != nil {
		return fmt.Errorf("无法获取合并单元格信息: %w", err)
	}

	cols := len(rows[0])
	colWidths := make([]float64, cols)
	for c := 0; c < cols; c++ {
		// 获取Excel中的实际列宽
		colName := getColumnName(c + 1)
		width, err := f.GetColWidth(sheetName, colName)
		if err != nil || width <= 0 {
			colWidths[c] = 64 // Excel默认列宽约64像素
		} else {
			// Excel列宽单位转换为像素 (大约1个字符宽度 = 8像素)
			colWidths[c] = width * 8
		}
	}

	rowHeights := make([]float64, len(rows)) // 初始化行高（使用更精确的Excel行高）
	defaultRowHeight := 18.0                 // Excel默认行高约18像素

	// 为每行设置实际行高
	for r := 0; r < len(rows); r++ {
		rowHeight := defaultRowHeight
		if height, err := f.GetRowHeight(sheetName, r+1); err == nil && height > 0 {
			rowHeight = height * 1.33 // Excel点转像素转换
		}
		rowHeights[r] = rowHeight
	}

	// 优化列宽计算，基于实际文本内容
	for r, row := range rows {
		for c := 0; c < cols; c++ {
			if c < len(row) {
				axis := getColumnName(c+1) + strconv.Itoa(r+1)
				cellValue := getCellDisplayValue(f, sheetName, axis, forceRaw)

				if strings.TrimSpace(cellValue) != "" {
					// 更精确的文本宽度计算
					textWidth := float64(len([]rune(cellValue))) * 9 // 考虑中文字符
					padding := 12.0
					if textWidth+padding > colWidths[c] {
						colWidths[c] = textWidth + padding
					}
				}
			}
		}
	}

	imgWidth := 0.0
	for _, w := range colWidths {
		imgWidth += w
	}

	imgHeight := 0.0
	for _, h := range rowHeights {
		imgHeight += h
	}

	dc := gg.NewContext(int(imgWidth), int(imgHeight))
	dc.SetRGB(1, 1, 1)
	dc.Clear()

	currentY := 0.0

	for r := range rows {
		currentX := 0.0
		rowHeight := rowHeights[r]

		for c := 0; c < cols; c++ {
			if c >= len(colWidths) {
				break
			}
			colWidth := colWidths[c]

			key := fmt.Sprintf("%d,%d", r, c)
			if rect, ok := mergedRects[key]; ok {
				if rect[0] == r && rect[1] == c {
					// 计算合并单元格的总宽度和高度
					mergedWidth := 0.0
					for cc := rect[1]; cc <= rect[3]; cc++ {
						if cc < len(colWidths) {
							mergedWidth += colWidths[cc]
						}
					}

					mergedHeight := 0.0
					for rr := rect[0]; rr <= rect[2]; rr++ {
						if rr < len(rowHeights) {
							mergedHeight += rowHeights[rr]
						}
					}

					axis := getColumnName(c+1) + strconv.Itoa(r+1)
					cellValue := getCellDisplayValue(f, sheetName, axis, forceRaw)
					cellStyle := getCellStyle(f, sheetName, axis)

					// 渲染合并单元格
					renderCellBackground(dc, cellStyle, currentX, currentY, mergedWidth, mergedHeight)
					renderCellBorders(dc, cellStyle, currentX, currentY, mergedWidth, mergedHeight)

					if strings.TrimSpace(cellValue) != "" {
						renderCellTextWithStyle(dc, cellValue, cellStyle, currentX, currentY, mergedWidth, mergedHeight)
					}
				}
			} else if _, ok := mergedRects[key]; !ok {

				axis := getColumnName(c+1) + strconv.Itoa(r+1)
				cellValue := getCellDisplayValue(f, sheetName, axis, forceRaw)
				cellStyle := getCellStyle(f, sheetName, axis)

				// 渲染单元格背景
				renderCellBackground(dc, cellStyle, currentX, currentY, colWidth, rowHeight)

				// 渲染单元格边框
				renderCellBorders(dc, cellStyle, currentX, currentY, colWidth, rowHeight)

				// 渲染单元格文本
				if strings.TrimSpace(cellValue) != "" {
					renderCellTextWithStyle(dc, cellValue, cellStyle, currentX, currentY, colWidth, rowHeight)
				}
			}

			currentX += colWidth
		}
		currentY += rowHeight
	}

	return dc.SavePNG(outPath)
}

// 渲染单元格背景 - 严格按照Excel样式
func renderCellBackground(dc *gg.Context, style CellStyle, x, y, width, height float64) {
	// 直接使用Excel样式中的背景颜色，不做任何修改
	r := float64(style.BgColor.R) / 255.0
	g := float64(style.BgColor.G) / 255.0
	b := float64(style.BgColor.B) / 255.0
	dc.SetRGB(r, g, b)

	// 填充背景
	dc.DrawRectangle(x, y, width, height)
	dc.Fill()
}

// 渲染单元格边框 - 严格按照Excel样式
func renderCellBorders(dc *gg.Context, style CellStyle, x, y, width, height float64) {
	// 渲染上边框
	if style.BorderTop.Style != "none" {
		r := float64(style.BorderTop.Color.R) / 255.0
		g := float64(style.BorderTop.Color.G) / 255.0
		b := float64(style.BorderTop.Color.B) / 255.0
		dc.SetRGB(r, g, b)
		dc.SetLineWidth(getBorderWidth(style.BorderTop.Style))
		dc.DrawLine(x, y, x+width, y)
		dc.Stroke()
	}

	// 渲染下边框
	if style.BorderBottom.Style != "none" {
		r := float64(style.BorderBottom.Color.R) / 255.0
		g := float64(style.BorderBottom.Color.G) / 255.0
		b := float64(style.BorderBottom.Color.B) / 255.0
		dc.SetRGB(r, g, b)
		dc.SetLineWidth(getBorderWidth(style.BorderBottom.Style))
		dc.DrawLine(x, y+height, x+width, y+height)
		dc.Stroke()
	}

	// 渲染左边框
	if style.BorderLeft.Style != "none" {
		r := float64(style.BorderLeft.Color.R) / 255.0
		g := float64(style.BorderLeft.Color.G) / 255.0
		b := float64(style.BorderLeft.Color.B) / 255.0
		dc.SetRGB(r, g, b)
		dc.SetLineWidth(getBorderWidth(style.BorderLeft.Style))
		dc.DrawLine(x, y, x, y+height)
		dc.Stroke()
	}

	// 渲染右边框
	if style.BorderRight.Style != "none" {
		r := float64(style.BorderRight.Color.R) / 255.0
		g := float64(style.BorderRight.Color.G) / 255.0
		b := float64(style.BorderRight.Color.B) / 255.0
		dc.SetRGB(r, g, b)
		dc.SetLineWidth(getBorderWidth(style.BorderRight.Style))
		dc.DrawLine(x+width, y, x+width, y+height)
		dc.Stroke()
	}
}

// 根据Excel边框样式获取线宽
func getBorderWidth(borderStyle string) float64 {
	switch borderStyle {
	case "thin":
		return 0.5
	case "medium":
		return 1.0
	case "thick":
		return 2.0
	default:
		return 0.5
	}
}

// 使用样式渲染单元格文本 - 严格按照Excel样式
func renderCellTextWithStyle(dc *gg.Context, text string, style CellStyle, x, y, width, height float64) {
	if text == "" {
		return
	}

	// 设置字体 - 优化字体大小以更接近Excel显示
	fontSize := style.FontSize

	if err := setFontFaceWithName(dc, style.FontName, fontSize); err != nil {
		// 如果指定字体失败，使用默认字体
		if err := setFontFace(dc, fontSize); err != nil {
			log.Printf("设置字体失败: %v", err)
		}
	}

	// 设置文本颜色 - 严格使用Excel样式
	r := float64(style.FontColor.R) / 255.0
	g := float64(style.FontColor.G) / 255.0
	b := float64(style.FontColor.B) / 255.0
	dc.SetRGB(r, g, b)

	// 计算文本位置 - 优化Excel对齐方式
	padding := 4.0 // 调整内边距以更好匹配Excel

	// 水平对齐计算
	var textX float64
	var anchorX float64

	switch style.HAlign {
	case "center", "centerContinuous":
		textX = x + width/2
		anchorX = 0.5
	case "right":
		textX = x + width - padding
		anchorX = 1.0
	case "general", "left", "":
		// Excel的general对齐：数字右对齐，文本左对齐
		if isNumeric(text) {
			textX = x + width - padding
			anchorX = 1.0
		} else {
			textX = x + padding
			anchorX = 0.0
		}
	default:
		textX = x + padding
		anchorX = 0.0
	}

	// 垂直对齐计算 - 精细调整垂直居中
	var textY float64
	var anchorY float64

	// 获取字体度量信息以更精确对齐
	_, fontHeight := dc.MeasureString("Ag")

	switch style.VAlign {
	case "top":
		textY = y + padding + fontHeight/2
		anchorY = 0.5
	case "center", "middle":
		// 真正的垂直居中 - 向上偏移更多以避免贴下边框
		textY = y + height/2 - fontHeight*0.2
		anchorY = 0.5
	case "bottom", "":
		textY = y + height - padding - fontHeight/2
		anchorY = 0.5
	default:
		// 默认也使用向上偏移的居中
		textY = y + height/2 - fontHeight*0.15
		anchorY = 0.5
	}

	// 绘制文本
	dc.DrawStringAnchored(text, textX, textY, anchorX, anchorY)
}

// 判断文本是否为数字
func isNumeric(text string) bool {
	text = strings.TrimSpace(text)
	if text == "" {
		return false
	}

	// 简单的数字判断
	_, err := strconv.ParseFloat(text, 64)
	return err == nil
}
