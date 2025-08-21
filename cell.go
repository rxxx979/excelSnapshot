package excelsnapshot

import (
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Cell 表示单元格信息（最常用的信息用于后续 gg 渲染）
type Cell struct {
	Sheet       *Sheet
	Row         int    // 1-based
	Col         int    // 1-based
	Address     string // 如 "A1"
	Value       string // 解析后的显示值（已由 excelize 处理）
	IsMerged    bool
	MergedRange []string
}

// IsEmpty 判断单元格是否为空
func (c *Cell) IsEmpty() bool { return c == nil || strings.TrimSpace(c.Value) == "" }

// String 返回单元格的字符串值（空则返回空串）
func (c *Cell) String() string {
	if c == nil {
		return ""
	}
	return c.Value
}

// Float64 将单元格值转为 float64
func (c *Cell) Float64() (float64, error) {
	if c == nil {
		return 0, strconv.ErrSyntax
	}
	return strconv.ParseFloat(strings.TrimSpace(c.Value), 64)
}

// Int 将单元格值转为 int
func (c *Cell) Int() (int, error) {
	if c == nil {
		return 0, strconv.ErrSyntax
	}
	return strconv.Atoi(strings.TrimSpace(c.Value))
}

func (c *Cell) Style() (*excelize.Style, error) {
	styleIndex, err := c.Sheet.excel.file.GetCellStyle(c.Sheet.Name, c.Address)
	if err != nil {
		return nil, err
	}
	return c.Sheet.excel.file.GetStyle(styleIndex)
}
