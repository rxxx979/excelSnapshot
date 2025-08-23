package excelsnapshot

import (
	"bytes"
	"fmt"
	"image"
	"image/gif"
	"image/jpeg"
	"image/png"
	"math"
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
	// 样式
	styles map[int]*excelize.Style
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
		styles:       make(map[int]*excelize.Style),
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

	// 使用列视图补齐：GetCols 能保留列内的显式空单元格（如设置为 "" 的单元格）
	if cols, err := s.excel.file.GetCols(s.Name); err == nil {
		for colIndex, colData := range cols {
			// 更新最大列数
			if colIndex+1 > maxCol {
				maxCol = colIndex + 1
			}
			// 列长度可能超过按行视图统计的 maxRow
			if len(colData) > maxRow {
				maxRow = len(colData)
			}
			// 补齐该列在 [1..len(colData)] 范围内缺失的单元格
			for r := 1; r <= len(colData); r++ {
				addr, _ := excelize.CoordinatesToCellName(colIndex+1, r)
				if _, ok := s.cells[addr]; !ok {
					val := colData[r-1]
					s.cells[addr] = &Cell{Sheet: s, Row: r, Col: colIndex + 1, Address: addr, Value: val}
				}
			}
		}
	}

	// 根据工作表维度补齐维度内缺失的空单元格（例如显式设置为空字符串的单元格）
	if dim, err := s.excel.file.GetSheetDimension(s.Name); err == nil && dim != "" {
		parts := strings.Split(dim, ":")
		if len(parts) == 2 {
			// 维度形如 A1:D10
			sc, sr, _ := excelize.CellNameToCoordinates(parts[0])
			ec, er, _ := excelize.CellNameToCoordinates(parts[1])
			if er > maxRow {
				maxRow = er
			}
			if ec > maxCol {
				maxCol = ec
			}
			for r := sr; r <= er; r++ {
				for c := sc; c <= ec; c++ {
					addr, _ := excelize.CoordinatesToCellName(c, r)
					if _, ok := s.cells[addr]; !ok {
						s.cells[addr] = &Cell{Sheet: s, Row: r, Col: c, Address: addr, Value: ""}
					}
				}
			}
		} else if len(parts) == 1 {
			// 单点维度，如 A1
			ec, er, _ := excelize.CellNameToCoordinates(parts[0])
			if er > maxRow {
				maxRow = er
			}
			if ec > maxCol {
				maxCol = ec
			}
			for r := 1; r <= er; r++ {
				for c := 1; c <= ec; c++ {
					addr, _ := excelize.CoordinatesToCellName(c, r)
					if _, ok := s.cells[addr]; !ok {
						s.cells[addr] = &Cell{Sheet: s, Row: r, Col: c, Address: addr, Value: ""}
					}
				}
			}
		}
	}

	// 仅对已存在的单元格绑定样式并缓存（避免对空区域重复扫描）
	styleBindCount := 0
	styleCacheMiss := 0
	for addr, cell := range s.cells {
		styleIndex, err := s.excel.file.GetCellStyle(s.Name, addr)
		if err != nil {
			return err
		}
		if cell.StyleIndex != styleIndex {
			cell.StyleIndex = styleIndex
		}
		if _, ok := s.styles[styleIndex]; !ok {
			st, err := s.excel.file.GetStyle(styleIndex)
			if err != nil {
				return err
			}
			s.styles[styleIndex] = st
			styleCacheMiss++
		}
		styleBindCount++
	}
	s.excel.logger.Info("加载工作表", zap.String("sheet", s.Name), zap.Int("rows", maxRow), zap.Int("cols", maxCol), zap.Int("cells", len(s.cells)), zap.Int("style_bind", styleBindCount), zap.Int("style_miss", styleCacheMiss))

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
	}

	// 优化：批量处理列宽（利用Excel列内宽度统一特性）
	for col := 1; col <= maxCol; col++ {
		colLetter, _ := excelize.ColumnNumberToName(col)
		width, _ := s.excel.file.GetColWidth(s.Name, colLetter)
		s.colWidthMap[colLetter] = width
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

		// 确保主单元格（左上角）样式已绑定并加入缓存
		mainAddr := mergedCell.GetStartAxis()
		if mainCell, ok := s.cells[mainAddr]; ok {
			// 若此前未绑定样式，则立即绑定并缓存
			if mainCell.StyleIndex == 0 {
				idx, err := s.excel.file.GetCellStyle(s.Name, mainAddr)
				if err != nil {
					return err
				}
				if mainCell.StyleIndex != idx {
					mainCell.StyleIndex = idx
				}
				if _, ok := s.styles[idx]; !ok {
					st, err := s.excel.file.GetStyle(idx)
					if err != nil {
						return err
					}
					s.styles[idx] = st
					styleCacheMiss++
				}
				styleBindCount++
			}
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
	// Excel 默认列宽（与 excelize 返回值对齐），使用容差判断
	const defaultColWidth = 9.140625
	const eps = 1e-6

	for colLetter := range s.colWidthMap {
		currentWidth := s.colWidthMap[colLetter]

		// 仅当列宽为默认值时，按内容进行估算调整；否则尊重文件中的列宽设置
		if math.Abs(currentWidth-defaultColWidth) <= eps {
			maxContentWidth := s.calculateMaxContentWidth(colLetter)
			if maxContentWidth > currentWidth {
				// 增加20%的缓冲区
				s.colWidthMap[colLetter] = maxContentWidth * 1.4
			}
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
		wrapText := true // 默认为可换行（多数场景渲染更友好）
		if cell != nil {
			if style, err := cell.Style(); err == nil && style != nil {
				if style.Font != nil && style.Font.Size > 0 {
					fontSize = style.Font.Size
				}
				// 若样式提供换行信息，则以样式为准（注意 Alignment 可能为 nil）
				if style.Alignment != nil {
					wrapText = style.Alignment.WrapText
				}
			}
		}

		// 计算基于字体大小的基础行高
		// 经验公式：行高 ≈ 字体大小 * 1.2 到 1.4 之间
		fontBasedHeight := fontSize * 1.33

		// 显式换行符统计
		explicitLines := 1
		for _, char := range cellValue {
			if char == '\n' || char == '\r' {
				explicitLines++
			}
		}

		// 列宽（Excel 列宽单位）。若未能获取，使用默认列宽
		colName, _ := excelize.ColumnNumberToName(colIndex + 1)
		colWidth := s.GetColWidth(colName)
		if colWidth <= 0 {
			colWidth = 9.140625 // Excel 默认列宽
		}

		// 基于内容宽度估算需要的行数（仅当允许换行或存在显式换行时生效）
		lines := 1
		if wrapText || explicitLines > 1 {
			if explicitLines > 1 {
				// 按显式换行切分，每段分别估算并累加
				segs := strings.FieldsFunc(cellValue, func(r rune) bool { return r == '\n' || r == '\r' })
				total := 0
				for _, seg := range segs {
					cw := s.estimateTextWidth(seg)
					if colWidth > 0 {
						need := int(cw/colWidth + 0.9999)
						if need < 1 {
							need = 1
						}
						total += need
					} else {
						total += 1
					}
				}
				if total < 1 {
					total = 1
				}
				lines = total
			} else {
				// 无显式换行，整段估算
				contentWidthUnits := s.estimateTextWidth(cellValue)
				if colWidth > 0 {
					needed := int(contentWidthUnits/colWidth + 0.9999) // ceil
					if needed < 1 {
						needed = 1
					}
					lines = needed
				} else {
					lines = 1
				}
			}
		} else {
			// 不允许自动换行且没有显式换行，认为单行
			lines = 1
		}

		estimatedHeight := fontBasedHeight * float64(lines)
		if estimatedHeight > maxHeight {
			maxHeight = estimatedHeight
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

// GetStyle 获取样式
func (s *Sheet) GetStyle(styleIndex int) (*excelize.Style, error) {
	if styleIndex < 0 {
		return nil, fmt.Errorf("非法样式索引: %d", styleIndex)
	}
	if st, ok := s.styles[styleIndex]; ok && st != nil {
		return st, nil
	}
	return nil, fmt.Errorf("样式未缓存: %d", styleIndex)
}
