package excelsnapshot

import (
	"image"
	"image/color"
	"math"

	"github.com/fogleman/gg"
	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"
	"golang.org/x/image/font"
)

const scale = 3.5

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

// RenderSheet 渲染工作表
func (sr *SheetRenderer) RenderSheet(sheet *Sheet) (image.Image, error) {
	w, h := sr.getSheetWidthAndHeight(sheet)
	canvas := gg.NewContext(int(w*scale), int(h*scale))
	canvas.Scale(scale, scale) // 重要：缩放坐标系，这样绘制时就是按原始尺寸计算

	canvas.SetColor(color.White)
	canvas.Clear()

	// 设置线条宽度，根据缩放调整
	canvas.SetLineWidth(1)
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
		// 绘制边框
		canvas.SetColor(color.RGBA{R: 200, G: 200, B: 200, A: 255}) // 浅灰色边框
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

	// 绘制边框（深色边框用于有内容的单元格）
	canvas.SetColor(color.Black)
	canvas.DrawRectangle(rect.x, rect.y, rect.w, rect.h)
	canvas.Stroke()

	// 绘制文本（避免缩放后的位图再缩放导致的模糊）
	if cell.Value != "" {
		fontSize := style.Font.Size

		// 使用未缩放坐标系绘制文字：放大字体尺寸，使用设备像素坐标
		fontFace, err := sr.GetFont(fontSize * scale)
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
		canvas.DrawStringAnchored(cell.Value, dx, dy, 0.5, 0.5)
		canvas.Pop()
	}
}

// getSheetWidthAndHeight 获取工作表宽高
func (sr *SheetRenderer) getSheetWidthAndHeight(sheet *Sheet) (float64, float64) {
	totalWidth, totalHeight := 0.0, 0.0
	for _, colWidth := range sheet.colWidthMap {
		totalWidth += colWidth * 7
	}
	for _, rowHeight := range sheet.rowHeightMap {
		totalHeight += rowHeight * 1.33
	}
	return totalWidth, totalHeight
}

// GetFont 获取字体
func (sr *SheetRenderer) GetFont(size float64) (font.Face, error) {
	f, err := LoadDefaultFontWithSize(size)
	if err != nil {
		return nil, err
	}
	return f, nil
}
