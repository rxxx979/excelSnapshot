package excelsnapshot

import (
	"fmt"
	"image"

	"github.com/xuri/excelize/v2"
)

// Excel 结构体
type Excel struct {
	path   string
	file   *excelize.File
	sheets map[string]*Sheet
}

// NewExcel 创建 Excel struct
func NewExcel(path string) (*Excel, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	return &Excel{
		path:   path,
		file:   f,
		sheets: make(map[string]*Sheet),
	}, nil
}

// GetSheetList 获取工作表列表
func (e *Excel) GetSheetList() ([]string, error) {
	if e.file == nil {
		return nil, fmt.Errorf("Excel 文件未打开")
	}
	return e.file.GetSheetList(), nil
}

// LoadSheets 预加载所有工作表信息（名称、行列数、单元格值等）
func (e *Excel) LoadSheets() error {
	names, err := e.GetSheetList()
	if err != nil {
		return err
	}
	for idx, name := range names {
		if _, ok := e.sheets[name]; ok {
			continue
		}
		sh := NewSheet(e, name)
		sh.Index = idx
		if err := sh.Load(); err != nil {
			return fmt.Errorf("加载工作表 %s 失败: %w", name, err)
		}
		e.sheets[name] = sh
	}
	return nil
}

// GetSheet 获取指定名称的工作表（如未缓存则加载）
func (e *Excel) GetSheet(name string) (*Sheet, error) {
	if e.file == nil {
		return nil, fmt.Errorf("Excel 文件未打开")
	}
	if sh, ok := e.sheets[name]; ok {
		return sh, nil
	}
	sh := NewSheet(e, name)
	if err := sh.Load(); err != nil {
		return nil, err
	}
	e.sheets[name] = sh
	return sh, nil
}

// Sheets 返回已加载的工作表（名称到结构的映射）
func (e *Excel) Sheets() map[string]*Sheet {
	return e.sheets
}

// Path 返回 Excel 文件路径
func (e *Excel) Path() string { return e.path }

// Close 关闭 Excel 文件
func (e *Excel) Close() error {
	return e.file.Close()
}

// Render 渲染指定工作表或全部工作表。
// 约定：
// - 若 all=true，则忽略 sheetName 与 sheetIndex，渲染所有工作表；
// - 否则优先使用 sheetName，其次使用 0-based 的 sheetIndex；
// 返回：map[工作表名称]image.Image
func (e *Excel) Render(sheetName string, sheetIndex int, all bool) (map[string]image.Image, error) {
	if e.file == nil {
		return nil, fmt.Errorf("Excel 文件未打开")
	}
	out := make(map[string]image.Image)

	if all {
		names, err := e.GetSheetList()
		if err != nil {
			return nil, err
		}
		for i, name := range names {
			sh, err := e.GetSheet(name)
			if err != nil {
				return nil, err
			}
			if sh.Index == 0 { // 尝试补充索引信息
				sh.Index = i
			}
			img, err := sh.Render()
			if err != nil {
				return nil, fmt.Errorf("渲染工作表 %s 失败: %w", name, err)
			}
			out[name] = img
		}
		return out, nil
	}

	// 单个 sheet
	targetName := sheetName
	if targetName == "" {
		if sheetIndex < 0 {
			return nil, fmt.Errorf("请提供 sheet_name 或 sheet_index，或设置 all=true")
		}
		names, err := e.GetSheetList()
		if err != nil {
			return nil, err
		}
		if sheetIndex >= len(names) {
			return nil, fmt.Errorf("sheet_index 超出范围: %d >= %d", sheetIndex, len(names))
		}
		targetName = names[sheetIndex]
	}

	sh, err := e.GetSheet(targetName)
	if err != nil {
		return nil, err
	}
	img, err := sh.Render()
	if err != nil {
		return nil, err
	}
	out[targetName] = img
	return out, nil
}
