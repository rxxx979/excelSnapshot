package excelsnapshot

import (
	"fmt"
	"image"

	"github.com/xuri/excelize/v2"
)

type Sheet struct {
	excel *Excel
	Name  string
	Index int
	Rows  int
	Cols  int
	cells map[string]*Cell // 仅存储非空单元格
}

// NewSheet 构造函数，仅建立与 Excel 的关联与名称，实际数据通过 Load 加载
func NewSheet(e *Excel, name string) *Sheet {
	return &Sheet{
		excel: e,
		Name:  name,
		cells: make(map[string]*Cell),
	}
}

// Load 读取工作表的行列及非空单元格数据
func (s *Sheet) Load() error {
	if s.excel == nil || s.excel.file == nil {
		return fmt.Errorf("Excel 文件未打开")
	}
	rows, err := s.excel.file.GetRows(s.Name)
	if err != nil {
		return err
	}
	s.Rows = len(rows)
	maxCols := 0
	// 清空并重建 cells
	s.cells = make(map[string]*Cell)
	for rIdx, row := range rows {
		if len(row) > maxCols {
			maxCols = len(row)
		}
		for cIdx, val := range row {
			if val == "" {
				continue
			}
			addr, err := excelize.CoordinatesToCellName(cIdx+1, rIdx+1)
			if err != nil {
				return err
			}
			s.cells[addr] = &Cell{
				Sheet:   s,
				Row:     rIdx + 1,
				Col:     cIdx + 1,
				Address: addr,
				Value:   val,
			}
		}
	}
	s.Cols = maxCols
	return nil
}

// Cell 按地址获取单元格（如未缓存或为空则从文件中按需读取）
func (s *Sheet) Cell(address string) (*Cell, error) {
	if c, ok := s.cells[address]; ok {
		return c, nil
	}
	if s.excel == nil || s.excel.file == nil {
		return nil, fmt.Errorf("Excel 文件未打开")
	}
	val, err := s.excel.file.GetCellValue(s.Name, address)
	if err != nil {
		return nil, err
	}
	col, row, err := excelize.CellNameToCoordinates(address)
	if err != nil {
		return &Cell{Sheet: s, Address: address, Value: val}, nil
	}
	c := &Cell{Sheet: s, Row: row, Col: col, Address: address, Value: val}
	if val != "" {
		s.cells[address] = c
	}
	if row > s.Rows {
		s.Rows = row
	}
	if col > s.Cols {
		s.Cols = col
	}
	return c, nil
}

// CellRC 通过行列索引（1-based）获取单元格
func (s *Sheet) CellRC(row, col int) (*Cell, error) {
	addr, err := excelize.CoordinatesToCellName(col, row)
	if err != nil {
		return nil, err
	}
	return s.Cell(addr)
}

// Cells 返回当前已加载的非空单元格集合（注意：为避免外部修改，这里返回一个浅拷贝）
func (s *Sheet) Cells() map[string]*Cell {
	out := make(map[string]*Cell, len(s.cells))
	for k, v := range s.cells {
		out[k] = v
	}
	return out
}

// Size 返回工作表的最大行列（1-based，基于已加载数据推断）
func (s *Sheet) Size() (rows, cols int) {
	return s.Rows, s.Cols
}

// Render 渲染当前工作表为 image.Image
func (s *Sheet) Render() (image.Image, error) {
	return RenderSheet(s)
}
