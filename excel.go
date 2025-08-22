package excelsnapshot

import (
	"fmt"

	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"
)

// Excel 结构体
type Excel struct {
	path       string
	file       *excelize.File
	sheets     map[string]*Sheet
	indexSheet map[int]string
	logger     *zap.Logger
}

// NewExcel 创建 Excel struct
func NewExcel(path string, logger *zap.Logger) (*Excel, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	excel := &Excel{
		path:       path,
		file:       f,
		sheets:     make(map[string]*Sheet),
		indexSheet: make(map[int]string),
		logger:     logger,
	}
	if err := excel.parseSheetListToMap(); err != nil {
		return nil, err
	}
	return excel, nil
}

// GetSheetList 获取工作表列表
func (e *Excel) parseSheetListToMap() error {
	if e.file == nil {
		return fmt.Errorf("Excel 文件未打开")
	}
	names := e.file.GetSheetList()
	e.sheets = make(map[string]*Sheet)
	for idx, name := range names {
		e.indexSheet[idx] = name
	}
	return nil
}

// LoadSheets 预加载所有工作表信息（名称、行列数、单元格值等）
func (e *Excel) LoadAllSheets() error {
	names := e.file.GetSheetList()
	for idx, name := range names {
		if _, ok := e.sheets[name]; ok {
			continue
		}
		e.logger.Info("加载工作表", zap.String("name", name))
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

// ParseAll 解析 Excel（加载所有工作表到内存）
func (e *Excel) ParseAll() error {
	return e.LoadAllSheets()
}

func (e *Excel) Parse(sheetName string) error {
	_, err := e.GetSheet(sheetName)
	return err
}

func (e *Excel) GetSheetNameByIndex(index int) string {
	return e.indexSheet[index]
}
