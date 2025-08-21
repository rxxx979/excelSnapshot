package excelsnapshot

import (
	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"
)

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

	// 遍历行，收集行高、单元格内容
	for rowIndex, row := range rows {
		height, _ := s.excel.file.GetRowHeight(s.Name, rowIndex+1)
		s.rowHeightMap[rowIndex+1] = height
		s.excel.logger.Debug("行：", zap.Int("row", rowIndex+1), zap.Float64("height", height))

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

	// 列宽
	for col := 1; col <= maxCol; col++ {
		colLetter, _ := excelize.ColumnNumberToName(col)
		width, _ := s.excel.file.GetColWidth(s.Name, colLetter)
		s.colWidthMap[colLetter] = width
		s.excel.logger.Debug("列：", zap.String("col", colLetter), zap.Float64("width", width))
	}

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
