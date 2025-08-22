package excelsnapshot

import (
	"bytes"
	"fmt"
	"image"
	"image/gif"
	"image/jpeg"
	"image/png"
	"strings"

	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"
)

// ExcelImage Excel中的图片信息
type ExcelImage struct {
	Name   string      // 图片名称
	Data   []byte      // 图片数据
	Format string      // 图片格式 (png, jpg, gif等)
	Image  image.Image // 解码后的图片
	Cell   string      // 起始单元格
	X      int         // X偏移
	Y      int         // Y偏移
	Width  int         // 宽度
	Height int         // 高度
}

type Sheet struct {
	excel      *Excel
	Name       string
	Index      int
	Rows       int
	Cols       int
	MaxColName string

	// 不同行的行高
	rowHeightMap map[int]float64
	// 不同列的列宽
	colWidthMap map[string]float64
	// 单元格
	cells map[string]*Cell
	// 工作表中的图片
	images []*ExcelImage
}

// NewSheet 构造函数，仅建立与 Excel 的关联与名称，实际数据通过 Load 加载
func NewSheet(e *Excel, name string) *Sheet {
	sheet := &Sheet{
		excel:        e,
		Name:         name,
		rowHeightMap: make(map[int]float64),
		colWidthMap:  make(map[string]float64),
		cells:        make(map[string]*Cell),
	}
	return sheet
}

// Load 加载工作表数据
func (s *Sheet) Load() error {
	// 获取所有行数据
	rows, err := s.excel.file.GetRows(s.Name)
	if err != nil {
		return err
	}
	maxRow, maxCol := 0, 0

	// 第一遍遍历：确定内容区域大小并创建单元格
	for rowIndex, row := range rows {
		if rowIndex+1 > maxRow {
			maxRow = rowIndex + 1
		}
		if len(row) > maxCol {
			maxCol = len(row)
		}

		for colIndex, value := range row {
			cellAddr, _ := excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
			s.cells[cellAddr] = &Cell{
				Sheet:   s,
				Row:     rowIndex + 1,
				Col:     colIndex + 1,
				Address: cellAddr,
				Value:   value,
			}
		}
	}

	// 使用更大的搜索范围来确保捕获所有可能的单元格
	searchRows := max(maxRow, 10)
	searchCols := max(maxCol, 10)

	for rowNum := 1; rowNum <= searchRows; rowNum++ {
		for colNum := 1; colNum <= searchCols; colNum++ {
			cellAddr, _ := excelize.CoordinatesToCellName(colNum, rowNum)
			if _, exists := s.cells[cellAddr]; !exists {
				// 检查该单元格是否被显式设置过（包括空值）
				cellValue, err := s.excel.file.GetCellValue(s.Name, cellAddr)
				if err == nil {
					// 通过获取单元格类型来判断是否存在
					cellType, err := s.excel.file.GetCellType(s.Name, cellAddr)
					if err == nil && cellType != excelize.CellTypeUnset {
						s.cells[cellAddr] = &Cell{
							Sheet:   s,
							Row:     rowNum,
							Col:     colNum,
							Address: cellAddr,
							Value:   cellValue,
						}
					}
				}
			}
		}
	}
	s.excel.logger.Info("加载工作表", zap.String("sheet", s.Name), zap.Int("rows", maxRow), zap.Int("cols", maxCol))

	// 优化：批量处理行高（利用Excel行内高度统一特性）
	for rowNum := 1; rowNum <= maxRow; rowNum++ {
		height, _ := s.excel.file.GetRowHeight(s.Name, rowNum)

		// 15 是 Excel 的默认行高，需要通过估算进行调整
		if height == 15 {
			// 只需要该行的数据来估算
			var rowData []string
			if rowNum-1 < len(rows) {
				rowData = rows[rowNum-1]
			}
			height = s.estimateRowHeight(rowNum, rowData)
		}

		s.rowHeightMap[rowNum] = height
		s.excel.logger.Debug("行：", zap.Int("row", rowNum), zap.Float64("height", height))
	}

	// 优化：批量处理列宽（利用Excel列内宽度统一特性）
	for col := 1; col <= maxCol; col++ {
		colLetter, _ := excelize.ColumnNumberToName(col)
		width, _ := s.excel.file.GetColWidth(s.Name, colLetter)
		s.colWidthMap[colLetter] = width
		s.excel.logger.Debug("列：", zap.String("col", colLetter), zap.Float64("width", width))
	}

	// 智能列宽调整：确保所有数据完整可见
	s.optimizeColumnWidths()

	// 合并单元格处理
	mergedCells, err := s.excel.file.GetMergeCells(s.Name)
	if err != nil {
		return err
	}
	for _, mergedCell := range mergedCells {
		startCol, startRow, _ := excelize.CellNameToCoordinates(mergedCell.GetStartAxis())
		endCol, endRow, _ := excelize.CellNameToCoordinates(mergedCell.GetEndAxis())
		s.excel.logger.Debug("合并单元格：", zap.String("start", mergedCell.GetStartAxis()), zap.String("end", mergedCell.GetEndAxis()))
		// 构造整个合并区域的地址列表
		var mergedRange []string
		for r := startRow; r <= endRow; r++ {
			for c := startCol; c <= endCol; c++ {
				addr, _ := excelize.CoordinatesToCellName(c, r)
				mergedRange = append(mergedRange, addr)
			}
		}

		// 给区域内的所有 Cell 打标记
		for _, addr := range mergedRange {
			cell, ok := s.cells[addr]
			if !ok {
				// 如果之前没加载到值，也新建一个 Cell
				col, row, _ := excelize.CellNameToCoordinates(addr)
				cell = &Cell{
					Sheet:   s,
					Row:     row,
					Col:     col,
					Address: addr,
					Value:   "",
				}
				s.cells[addr] = cell
			}
			cell.IsMerged = true
			cell.MergedRange = mergedRange
		}
	}

	// 加载工作表中的图片
	if err := s.loadImages(); err != nil {
		s.excel.logger.Warn("加载图片失败", zap.Error(err))
	}

	// 更新 sheet 信息
	s.Rows = maxRow
	s.Cols = maxCol
	s.MaxColName, _ = excelize.ColumnNumberToName(maxCol)
	return nil
}

// GetColWidth 获取指定列的列宽
func (s *Sheet) GetColWidth(col string) float64 {
	return s.colWidthMap[col]
}

// GetRowHeight 获取指定行的行高
func (s *Sheet) GetRowHeight(row int) float64 {
	return s.rowHeightMap[row]
}

// optimizeColumnWidths 智能调整列宽以确保数据完整可见
func (s *Sheet) optimizeColumnWidths() {
	for colLetter := range s.colWidthMap {
		maxContentWidth := s.calculateMaxContentWidth(colLetter)
		currentWidth := s.colWidthMap[colLetter]

		// 如果内容宽度超过当前列宽，则调整
		if maxContentWidth > currentWidth {
			// 增加20%的缓冲区
			s.colWidthMap[colLetter] = maxContentWidth * 1.4
			s.excel.logger.Debug("调整列宽",
				zap.String("col", colLetter),
				zap.Float64("原宽度", currentWidth),
				zap.Float64("新宽度", maxContentWidth*1.4))
		}
	}
}

// calculateMaxContentWidth 计算列中最大内容宽度
func (s *Sheet) calculateMaxContentWidth(colLetter string) float64 {
	maxWidth := 0.0

	// 遍历该列的所有单元格
	for _, cell := range s.cells {
		if cell.Address == "" {
			continue
		}

		// 检查是否属于当前列
		cellColLetter, _ := excelize.ColumnNumberToName(cell.Col)
		if cellColLetter != colLetter {
			continue
		}

		// 估算文本宽度（简单估算：字符数 * 平均字符宽度）
		contentWidth := s.estimateTextWidth(cell.Value)
		if contentWidth > maxWidth {
			maxWidth = contentWidth
		}
	}

	return maxWidth
}

// estimateTextWidth 估算文本宽度
func (s *Sheet) estimateTextWidth(text string) float64 {
	if text == "" {
		return 0
	}

	// 简单估算：中文字符约2个单位，英文字符约1个单位
	width := 0.0
	for _, char := range text {
		if char > 127 { // 非ASCII字符（中文等）
			width += 2.0
		} else {
			width += 1.0
		}
	}

	// 转换为Excel列宽单位（大约字符宽度的0.8倍）
	return width * 0.8
}

// estimateRowHeight 根据行内容和字体大小估算行高
func (s *Sheet) estimateRowHeight(rowNum int, rowData []string) float64 {
	maxHeight := 15.0 // 默认最小行高

	// 遍历这一行的所有单元格
	for colIndex, cellValue := range rowData {
		if cellValue == "" {
			continue
		}

		// 获取单元格对象以获取样式信息
		cellAddr, _ := excelize.CoordinatesToCellName(colIndex+1, rowNum)
		cell := s.cells[cellAddr]

		// 获取字体大小
		fontSize := 11.0 // 默认字体大小
		if cell != nil {
			if style, err := cell.Style(); err == nil && style.Font.Size > 0 {
				fontSize = style.Font.Size
			}
		}

		// 计算基于字体大小的基础行高
		// 经验公式：行高 ≈ 字体大小 * 1.2 到 1.4 之间
		fontBasedHeight := fontSize * 1.33

		// 检查是否有换行符
		lineCount := 1
		for _, char := range cellValue {
			if char == '\n' || char == '\r' {
				lineCount++
			}
		}

		// 根据换行数量调整高度
		if lineCount > 1 {
			estimatedHeight := fontBasedHeight * float64(lineCount)
			if estimatedHeight > maxHeight {
				maxHeight = estimatedHeight
			}
		} else {
			// 根据内容长度估算（如果内容很长，可能需要自动换行）
			colName, _ := excelize.ColumnNumberToName(colIndex + 1)
			colWidth := s.GetColWidth(colName)
			if colWidth == 0 {
				colWidth = 64 // 默认列宽，单位字符
			}

			// 基于字体大小估算每行能容纳的字符数
			// 字体越大，每行容纳的字符越少
			avgCharWidth := fontSize * 0.6                   // 估算平均字符宽度
			charsPerLine := int(colWidth * 7 / avgCharWidth) // colWidth*7 转换为像素
			if charsPerLine < 1 {
				charsPerLine = 1
			}

			// 计算内容可能占用的行数
			contentLength := len([]rune(cellValue)) // 使用 rune 处理中文字符
			estimatedLines := (contentLength + charsPerLine - 1) / charsPerLine

			if estimatedLines > 1 {
				estimatedHeight := fontBasedHeight * float64(estimatedLines)
				if estimatedHeight > maxHeight {
					maxHeight = estimatedHeight
				}
			} else {
				// 单行内容，使用基于字体的高度
				if fontBasedHeight > maxHeight {
					maxHeight = fontBasedHeight
				}
			}
		}
	}

	// 确保最小行高
	if maxHeight < 15.0 {
		maxHeight = 15.0
	}

	// 限制最大行高，避免过度拉伸
	if maxHeight > 150.0 {
		maxHeight = 150.0
	}

	return maxHeight
}

// loadImages 加载工作表中的嵌入图片
func (s *Sheet) loadImages() error {
	s.images = nil // 重置图片列表
	s.excel.logger.Info("开始加载图片", zap.String("sheet", s.Name))

	// 遍历所有单元格，查找包含图片的单元格
	cellCount := 0
	for addr := range s.cells {
		cellCount++
		pictures, err := s.excel.file.GetPictures(s.Name, addr)
		if err != nil {
			s.excel.logger.Debug("获取单元格图片失败", zap.String("cell", addr), zap.Error(err))
			continue
		}

		// 处理该单元格的所有图片
		for _, pic := range pictures {
			// 从GraphicOptions中获取位置信息
			var x, y int
			if pic.Format != nil {
				x = pic.Format.OffsetX
				y = pic.Format.OffsetY
			}

			excelImage := &ExcelImage{
				Name:   string(pic.File), // 将[]byte转换为string作为文件名
				Data:   pic.File,
				Format: pic.Extension,
				Cell:   addr,
				X:      x,
				Y:      y,
				Width:  0, // 将从解码后的图片获取实际尺寸
				Height: 0,
			}

			// 解码图片数据
			if err := s.decodeImage(excelImage); err != nil {
				s.excel.logger.Warn("解码图片失败", zap.String("cell", addr), zap.Error(err))
				continue
			}

			s.images = append(s.images, excelImage)
		}
	}

	s.excel.logger.Info("图片加载完成",
		zap.String("sheet", s.Name),
		zap.Int("检查单元格数", cellCount),
		zap.Int("加载图片数量", len(s.images)))
	return nil
}

// decodeImage 解码图片数据
func (s *Sheet) decodeImage(excelImage *ExcelImage) error {
	if len(excelImage.Data) == 0 {
		return fmt.Errorf("图片数据为空")
	}

	reader := bytes.NewReader(excelImage.Data)

	// 根据格式解码图片
	var img image.Image
	var err error

	switch strings.ToLower(excelImage.Format) {
	case "png":
		img, err = png.Decode(reader)
	case "jpg", "jpeg":
		img, err = jpeg.Decode(reader)
	case "gif":
		img, err = gif.Decode(reader)
	default:
		// 尝试自动检测格式
		reader.Seek(0, 0)
		img, excelImage.Format, err = image.Decode(reader)
	}

	if err != nil {
		return fmt.Errorf("解码图片失败: %w", err)
	}

	excelImage.Image = img

	// 设置图片的实际尺寸
	bounds := img.Bounds()
	excelImage.Width = bounds.Dx()
	excelImage.Height = bounds.Dy()

	return nil
}
