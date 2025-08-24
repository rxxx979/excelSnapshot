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

	// 先绘制整张默认网格（浅灰色）
	sr.drawBaseGrid(canvas, sheet)

	// 再绘制单元格：背景+文本，并仅对非默认边框颜色进行覆盖
	for addr, rect := range cellRects {
		cell := sheet.cells[addr]
		if cell == nil {
			continue
		}
		// 仅绘制主单元格
		if cell.IsMerged && cell.MergedRange[0] != addr {
			continue
		}
		// 背景与文本
		sr.drawCell(canvas, rect, cell)
		// 边框覆盖（非默认颜色）
		sr.drawCellBordersOverride(canvas, rect, cell)
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

// drawCell 绘制单元格，包括背景色和文本
func (sr *SheetRenderer) drawCell(canvas *gg.Context, rect struct{ x, y, w, h float64 }, cell *Cell) {
	// 绘制背景色（样式容错）
	style, err := cell.Style()
	if err != nil || style == nil {
		if err != nil {
			sr.logger.Debug("获取单元格样式失败，采用默认样式", zap.Error(err))
		}
	} else if len(style.Fill.Color) > 0 {
		if bgColor, err := HexToRGBA(style.Fill.Color[0]); err == nil {
			canvas.SetColor(bgColor)
			canvas.DrawRectangle(rect.x, rect.y, rect.w, rect.h)
			canvas.Fill()
		}
	}

	// 绘制文本（避免缩放后的位图再缩放导致的模糊）
	if cell.Value != "" {
		// 字体参数（样式容错）
		fontSize := 11.0
		bold := false
		if style != nil && style.Font != nil {
			if style.Font.Size > 0 {
				fontSize = style.Font.Size
			}
			bold = style.Font.Bold
		}

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
		if style != nil && style.Font != nil && style.Font.Color != "" {
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

// defaultBorderColor 返回默认边框颜色（浅灰色）
func defaultBorderColor() color.Color {
	return color.RGBA{R: 200, G: 200, B: 200, A: 255}
}

// equalColor 比较两个颜色是否等价（按 RGBA 值）
func equalColor(a, b color.Color) bool {
	ar, ag, ab, aa := a.RGBA()
	br, bg, bb, ba := b.RGBA()
	return ar == br && ag == bg && ab == bb && aa == ba
}

// drawBaseGrid 使用行/列端点绘制整张默认网格
func (sr *SheetRenderer) drawBaseGrid(canvas *gg.Context, sheet *Sheet) {
	def := defaultBorderColor()
	canvas.SetColor(def)

	// 列偏移与总宽
	colOffsets := make([]float64, sheet.Cols+1)
	for c := 1; c <= sheet.Cols; c++ {
		colName, _ := excelize.ColumnNumberToName(c)
		colWidth := sheet.GetColWidth(colName) * 7
		colOffsets[c] = colOffsets[c-1] + colWidth
	}
	totalWidth := colOffsets[sheet.Cols]

	// 行偏移与总高
	rowOffsets := make([]float64, sheet.Rows+1)
	for r := 1; r <= sheet.Rows; r++ {
		rowHeight := sheet.GetRowHeight(r) * 1.33
		rowOffsets[r] = rowOffsets[r-1] + rowHeight
	}
	totalHeight := rowOffsets[sheet.Rows]

	// 竖线
	for i := 0; i <= sheet.Cols; i++ {
		x := colOffsets[i]
		canvas.DrawLine(x, 0, x, totalHeight)
		canvas.Stroke()
	}
	// 横线
	for i := 0; i <= sheet.Rows; i++ {
		y := rowOffsets[i]
		canvas.DrawLine(0, y, totalWidth, y)
		canvas.Stroke()
	}
}

// drawCellBordersOverride 仅覆盖非默认颜色的边框（按边）
func (sr *SheetRenderer) drawCellBordersOverride(canvas *gg.Context, rect struct{ x, y, w, h float64 }, cell *Cell) {
	if cell == nil {
		return
	}
	style, err := cell.Style()
	if err != nil || style == nil {
		return
	}
	def := defaultBorderColor()
	// 遍历样式边框定义，按边绘制
	for _, b := range style.Border {
		if b.Color == "" {
			continue
		}
		col, err := HexToRGBA(b.Color)
		if err != nil {
			continue
		}
		if equalColor(col, def) {
			// 与默认色一致则无需覆盖
			continue
		}
		canvas.SetColor(col)
		switch b.Type {
		case "left":
			canvas.DrawLine(rect.x, rect.y, rect.x, rect.y+rect.h)
			canvas.Stroke()
		case "right":
			canvas.DrawLine(rect.x+rect.w, rect.y, rect.x+rect.w, rect.y+rect.h)
			canvas.Stroke()
		case "top":
			canvas.DrawLine(rect.x, rect.y, rect.x+rect.w, rect.y)
			canvas.Stroke()
		case "bottom":
			canvas.DrawLine(rect.x, rect.y+rect.h, rect.x+rect.w, rect.y+rect.h)
			canvas.Stroke()
		default:
			// 其他类型（如 diagonal*），此处忽略
		}
	}
}

// GetFont 获取字体
func (sr *SheetRenderer) GetFont(size float64, bold bool) (font.Face, error) {
	mapKey := fmt.Sprintf("%f|%t", size, bold)
	if f, ok := sr.fontMap[mapKey]; ok {
		return f, nil
	}
	f, err := LoadDefaultFontWithSize(size, bold)
	if err != nil {
		return nil, err
	}
	sr.fontMap[mapKey] = f
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
