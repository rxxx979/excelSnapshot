package excelsnapshot

import (
	"fmt"
	"image"
	"image/color"
	"math"

	"github.com/fogleman/gg"
	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"
	"golang.org/x/image/font"
)

// 图片缩放比例
const scale = 2.0

type SheetRenderer struct {
	logger  *zap.Logger
	fontMap map[string]font.Face
}

// NewSheetRenderer 创建 SheetRenderer
func NewSheetRenderer(logger *zap.Logger) *SheetRenderer {
	return &SheetRenderer{
		logger:  logger,
		fontMap: make(map[string]font.Face),
	}
}

// RenderSheet 渲染工作表为图片
func (sr *SheetRenderer) RenderSheet(sheet *Sheet) (image.Image, error) {
	if sheet == nil {
		return nil, fmt.Errorf("工作表为空")
	}

	w, h := sr.getSheetWidthAndHeight(sheet)
	canvas := gg.NewContext(int(w*scale), int(h*scale))
	canvas.Scale(scale, scale) // 重要：缩放坐标系，这样绘制时就是按原始尺寸计算

	canvas.SetColor(color.White)
	canvas.Clear()

	// 设置线条宽度，根据缩放调整
	canvas.SetLineWidth(scale)
	canvas.SetColor(color.Black)

	// 计算所有单元格矩形信息
	cellRects := sr.calculateCellRects(sheet)

	// 先绘制所有单元格边框（包括空单元格）
	for addr, rect := range cellRects {
		cell := sheet.cells[addr]
		// 检查是否是合并单元格的非主单元格
		if cell != nil && cell.IsMerged && cell.MergedRange[0] != addr {
			continue
		}
		// 绘制边框，使用样式颜色或默认颜色
		borderColor := sr.getBorderColor(cell)
		canvas.SetColor(borderColor)
		canvas.DrawRectangle(rect.x, rect.y, rect.w, rect.h)
		canvas.Stroke()
	}

	// 然后绘制有内容的单元格
	for addr, rect := range cellRects {
		cell := sheet.cells[addr]
		if cell == nil {
			continue
		}
		// 仅绘制主单元格
		if cell.IsMerged && cell.MergedRange[0] != addr {
			continue
		}
		sr.drawCell(canvas, rect, cell)
	}

	// 最后绘制嵌入的图片（在单元格内容之上）
	if len(sheet.images) > 0 {
		sr.logger.Info("开始渲染图片", zap.Int("数量", len(sheet.images)))
	}
	sr.drawImages(canvas, sheet, cellRects)
	sr.logger.Debug("图片渲染完成")

	// 直接返回高分辨率图片，不缩放
	return canvas.Image(), nil
}

// calculateCellRects 计算每个单元格在画布上的位置和大小
func (sr *SheetRenderer) calculateCellRects(sheet *Sheet) map[string]struct{ x, y, w, h float64 } {
	cellRects := make(map[string]struct{ x, y, w, h float64 })

	// 记录每行和每列的偏移量
	colOffsets := make([]float64, sheet.Cols+1)
	rowOffsets := make([]float64, sheet.Rows+1)

	// 计算列偏移
	for c := 1; c <= sheet.Cols; c++ {
		colName, _ := excelize.ColumnNumberToName(c)
		colWidth := sheet.GetColWidth(colName) * 7
		colOffsets[c] = colOffsets[c-1] + colWidth
	}

	// 计算行偏移
	for r := 1; r <= sheet.Rows; r++ {
		rowHeight := sheet.GetRowHeight(r) * 1.33
		rowOffsets[r] = rowOffsets[r-1] + rowHeight
	}

	// 遍历每个单元格
	for r := 1; r <= sheet.Rows; r++ {
		for c := 1; c <= sheet.Cols; c++ {
			cellAddr, _ := excelize.CoordinatesToCellName(c, r)
			cell := sheet.cells[cellAddr]

			var rectW, rectH float64
			if cell != nil && cell.IsMerged && cell.MergedRange[0] == cellAddr {
				rectW, rectH = sr.calcMergedRectOffsets(cell, colOffsets, rowOffsets)
			} else {
				rectW = colOffsets[c] - colOffsets[c-1]
				rectH = rowOffsets[r] - rowOffsets[r-1]
			}

			cellRects[cellAddr] = struct{ x, y, w, h float64 }{
				x: colOffsets[c-1],
				y: rowOffsets[r-1],
				w: rectW,
				h: rectH,
			}
		}
	}

	return cellRects
}

// calcMergedRectOffsets 计算合并单元格的宽高及位置
func (sr *SheetRenderer) calcMergedRectOffsets(cell *Cell, colOffsets, rowOffsets []float64) (float64, float64) {
	endAddr := cell.MergedRange[len(cell.MergedRange)-1]
	endCol, endRow, _ := excelize.CellNameToCoordinates(endAddr)

	width := colOffsets[endCol] - colOffsets[cell.Col-1]
	height := rowOffsets[endRow] - rowOffsets[cell.Row-1]
	return width, height
}

// drawCell 绘制单元格，包括背景色、边框和文本
func (sr *SheetRenderer) drawCell(canvas *gg.Context, rect struct{ x, y, w, h float64 }, cell *Cell) {
	// 绘制背景色
	style, err := cell.Style()
	if err != nil {
		sr.logger.Error("获取单元格样式失败", zap.Error(err))
	}
	if len(style.Fill.Color) > 0 {
		if bgColor, err := HexToRGBA(style.Fill.Color[0]); err == nil {
			canvas.SetColor(bgColor)
			canvas.DrawRectangle(rect.x, rect.y, rect.w, rect.h)
			canvas.Fill()
		}
	}

	// 绘制边框，使用样式颜色
	borderColor := sr.getBorderColor(cell)
	canvas.SetColor(borderColor)
	canvas.DrawRectangle(rect.x, rect.y, rect.w, rect.h)
	canvas.Stroke()

	// 绘制文本（避免缩放后的位图再缩放导致的模糊）
	if cell.Value != "" {
		fontSize := style.Font.Size
		bold := style.Font.Bold

		// 使用未缩放坐标系绘制文字：放大字体尺寸，使用设备像素坐标
		fontFace, err := sr.GetFont(fontSize*scale, bold)
		if err != nil {
			sr.logger.Error("获取字体失败", zap.Error(err))
			return
		}

		// 计算设备像素坐标并进行像素对齐
		dx := (rect.x + rect.w/2) * scale
		dy := (rect.y + rect.h/2) * scale
		dx = math.Round(dx)
		dy = math.Round(dy)

		var fontColor color.Color = color.Black
		if style.Font.Color != "" {
			if fc, err := HexToRGBA(style.Font.Color); err == nil {
				fontColor = fc
			}
		}

		canvas.Push()
		canvas.Identity()
		canvas.SetFontFace(fontFace)
		canvas.SetColor(fontColor)
		canvas.DrawStringAnchored(cell.Value, dx, dy, 0.5, 0.3)
		canvas.Pop()
	}
}

// getSheetWidthAndHeight 获取工作表宽高
func (sr *SheetRenderer) getSheetWidthAndHeight(sheet *Sheet) (float64, float64) {
	if sheet == nil {
		return 100, 100 // 返回默认尺寸
	}
	
	totalWidth, totalHeight := 0.0, 0.0
	for _, colWidth := range sheet.colWidthMap {
		totalWidth += colWidth * 7
	}
	for _, rowHeight := range sheet.rowHeightMap {
		totalHeight += rowHeight * 1.33
	}
	return totalWidth, totalHeight
}

// getBorderColor 获取单元格边框颜色
func (sr *SheetRenderer) getBorderColor(cell *Cell) color.Color {
	// 默认边框颜色
	defaultBorderColor := color.RGBA{R: 200, G: 200, B: 200, A: 255} // 浅灰色

	if cell == nil {
		return defaultBorderColor
	}

	// 获取单元格样式
	style, err := cell.Style()
	if err != nil {
		return defaultBorderColor
	}

	// 只使用样式中明确定义的边框颜色
	if len(style.Border) > 0 {
		// 检查各边的边框颜色（上、下、左、右）
		borders := []string{style.Border[0].Color, style.Border[1].Color, style.Border[2].Color, style.Border[3].Color}
		for _, borderColor := range borders {
			if borderColor != "" {
				if c, err := HexToRGBA(borderColor); err == nil {
					return c
				}
			}
		}
	}

	// 如果样式中没有明确的边框颜色，都使用默认颜色
	return defaultBorderColor
}

// GetFont 获取字体
func (sr *SheetRenderer) GetFont(size float64, bold bool) (font.Face, error) {
	f, err := LoadDefaultFontWithSize(size, bold)
	if err != nil {
		return nil, err
	}
	return f, nil
}

// drawImages 绘制工作表中的嵌入图片
func (sr *SheetRenderer) drawImages(canvas *gg.Context, sheet *Sheet, cellRects map[string]struct{ x, y, w, h float64 }) {
	if len(sheet.images) == 0 {
		return
	}

	for _, img := range sheet.images {
		if img.Image == nil {
			sr.logger.Warn("跳过未解码的图片", zap.String("name", img.Name))
			continue
		}

		// 计算图片在canvas上的位置
		x, y := sr.calculateImagePosition(img, cellRects)

		// 获取图片的实际尺寸
		bounds := img.Image.Bounds()
		imgWidth := float64(bounds.Dx())
		imgHeight := float64(bounds.Dy())

		// 考虑Excel中可能的缩放因子
		var finalWidth, finalHeight float64
		if img.Width > 0 && img.Height > 0 {
			// 如果有Excel指定的尺寸，使用Excel的尺寸
			finalWidth = float64(img.Width)
			finalHeight = float64(img.Height)
		} else {
			// 否则使用图片原始尺寸
			finalWidth = imgWidth
			finalHeight = imgHeight
		}

		// 绘制图片需要保持原始尺寸比例
		canvas.Push()
		canvas.Identity() // 重置变换，使用设备像素坐标

		// 计算设备像素坐标和目标尺寸
		deviceX := x * scale
		deviceY := y * scale
		targetWidth := finalWidth * scale
		targetHeight := finalHeight * scale

		// 计算缩放比例并应用变换
		scaleX := targetWidth / imgWidth
		scaleY := targetHeight / imgHeight
		canvas.Translate(deviceX, deviceY)
		canvas.Scale(scaleX, scaleY)
		canvas.DrawImage(img.Image, 0, 0)
		canvas.Pop()

	}
}

// calculateImagePosition 计算图片在canvas上的像素位置
func (sr *SheetRenderer) calculateImagePosition(img *ExcelImage, cellRects map[string]struct{ x, y, w, h float64 }) (float64, float64) {
	// 获取起始单元格的位置
	cellRect, exists := cellRects[img.Cell]
	if !exists {
		sr.logger.Warn("找不到图片起始单元格", zap.String("cell", img.Cell), zap.String("image", img.Name))
		return 0, 0
	}

	// 计算图片的最终位置：单元格位置 + 偏移量
	// Excel的偏移量通常以像素为单位，但可能需要转换
	// 根据Excel的坐标系统，偏移量可能需要不同的转换因子
	var offsetX, offsetY float64

	if img.X != 0 || img.Y != 0 {
		// 有偏移量，需要转换
		// Excel偏移量可能是EMU单位，转换为像素
		// 粗略估算：1个EMU单位 ≈ 0.0008像素（需要根据实际情况调整）
		offsetX = float64(img.X) * 0.0008
		offsetY = float64(img.Y) * 0.0008

		sr.logger.Debug("图片偏移量转换",
			zap.String("image", img.Name),
			zap.Int("originalX", img.X),
			zap.Int("originalY", img.Y),
			zap.Float64("convertedX", offsetX),
			zap.Float64("convertedY", offsetY))
	}

	x := cellRect.x + offsetX
	y := cellRect.y + offsetY

	return x, y
}
