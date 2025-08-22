package excelsnapshot

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
	"go.uber.org/zap/zaptest"
)

// TestCell_Creation 测试Cell结构体的创建
func TestCell_Creation(t *testing.T) {
	tests := []struct {
		name     string
		addr     string
		value    string
		expected *Cell
	}{
		{
			name:  "字符串值",
			addr:  "A1",
			value: "测试文本",
			expected: &Cell{
				Address: "A1",
				Value:   "测试文本",
				Row:     1,
				Col:     1,
			},
		},
		{
			name:  "数字字符串",
			addr:  "B1",
			value: "123",
			expected: &Cell{
				Address: "B1",
				Value:   "123",
				Row:     1,
				Col:     2,
			},
		},
		{
			name:  "空值",
			addr:  "C1",
			value: "",
			expected: &Cell{
				Address: "C1",
				Value:   "",
				Row:     1,
				Col:     3,
			},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cell := &Cell{
				Address: tt.addr,
				Value:   tt.value,
				Row:     tt.expected.Row,
				Col:     tt.expected.Col,
			}
			
			if cell.Address != tt.expected.Address {
				t.Errorf("Cell.Address = %v, want %v", cell.Address, tt.expected.Address)
			}
			
			if cell.Value != tt.expected.Value {
				t.Errorf("Cell.Value = %v, want %v", cell.Value, tt.expected.Value)
			}
		})
	}
}

// TestCell_String 测试单元格字符串表示
func TestCell_String(t *testing.T) {
	tests := []struct {
		name     string
		cell     *Cell
		expected string
	}{
		{
			name: "字符串值",
			cell: &Cell{
				Address: "A1",
				Value:   "测试",
			},
			expected: "测试",
		},
		{
			name: "数字值",
			cell: &Cell{
				Address: "B1",
				Value:   "123",
			},
			expected: "123",
		},
		{
			name: "浮点数值",
			cell: &Cell{
				Address: "C1",
				Value:   "456.78",
			},
			expected: "456.78",
		},
		{
			name: "空值",
			cell: &Cell{
				Address: "D1",
				Value:   "",
			},
			expected: "",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.cell.String()
			if result != tt.expected {
				t.Errorf("Cell.String() = %v, want %v", result, tt.expected)
			}
		})
	}
}

// TestCell_IsEmpty 测试空单元格判断
func TestCell_IsEmpty(t *testing.T) {
	tests := []struct {
		name     string
		cell     *Cell
		expected bool
	}{
		{
			name: "非空字符串",
			cell: &Cell{
				Address: "A1",
				Value:   "测试",
			},
			expected: false,
		},
		{
			name: "空字符串",
			cell: &Cell{
				Address: "B1",
				Value:   "",
			},
			expected: true,
		},
		{
			name: "数字零",
			cell: &Cell{
				Address: "C1",
				Value:   "0",
			},
			expected: false,
		},
		{
			name: "nil值",
			cell: (*Cell)(nil),
			expected: true,
		},
		{
			name: "只有空格的字符串",
			cell: &Cell{
				Address: "E1",
				Value:   "   ",
			},
			expected: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := tt.cell.IsEmpty()
			if result != tt.expected {
				t.Errorf("Cell.IsEmpty() = %v, want %v", result, tt.expected)
			}
		})
	}
}

// TestCell_GetNumericValue 测试数值获取
func TestCell_GetNumericValue(t *testing.T) {
	tests := []struct {
		name        string
		cell        *Cell
		expectedVal float64
		expectedOk  bool
	}{
		{
			name: "整数",
			cell: &Cell{
				Address: "A1",
				Value:   "123",
			},
			expectedVal: 123.0,
			expectedOk:  true,
		},
		{
			name: "浮点数",
			cell: &Cell{
				Address: "B1",
				Value:   "456.78",
			},
			expectedVal: 456.78,
			expectedOk:  true,
		},
		{
			name: "数字字符串",
			cell: &Cell{
				Address: "C1",
				Value:   "789.12",
			},
			expectedVal: 789.12,
			expectedOk:  true,
		},
		{
			name: "非数字字符串",
			cell: &Cell{
				Address: "D1",
				Value:   "不是数字",
			},
			expectedVal: 0,
			expectedOk:  false,
		},
		{
			name: "nil值",
			cell: (*Cell)(nil),
			expectedVal: 0,
			expectedOk:  false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			val, err := tt.cell.Float64()
			ok := err == nil
			if ok != tt.expectedOk {
				t.Errorf("Cell.Float64() ok = %v, want %v", ok, tt.expectedOk)
			}
			if ok && val != tt.expectedVal {
				t.Errorf("Cell.Float64() val = %v, want %v", val, tt.expectedVal)
			}
		})
	}
}

// TestCell_WithRealExcelData 使用真实Excel文件测试单元格功能
func TestCell_WithRealExcelData(t *testing.T) {
	// 创建包含各种数据类型的Excel文件
	tempDir := t.TempDir()
	testFile := filepath.Join(tempDir, "cell_test.xlsx")
	
	f := excelize.NewFile()
	defer f.Close()
	
	// 添加各种数据类型
	f.SetCellValue("Sheet1", "A1", "文本数据")
	f.SetCellValue("Sheet1", "A2", 123)
	f.SetCellValue("Sheet1", "A3", 456.78)
	f.SetCellValue("Sheet1", "A4", true)
	f.SetCellValue("Sheet1", "A5", "")
	
	if err := f.SaveAs(testFile); err != nil {
		t.Fatalf("保存测试文件失败: %v", err)
	}

	// 加载并测试
	logger := zaptest.NewLogger(t)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		t.Fatalf("加载Excel文件失败: %v", err)
	}
	defer excel.Close()

	sheet, err := excel.GetSheet("Sheet1")
	if err != nil {
		t.Fatalf("获取工作表失败: %v", err)
	}

	// 测试各种单元格
	tests := []struct {
		addr        string
		wantEmpty   bool
		wantNumeric bool
	}{
		{"A1", false, false}, // 文本
		{"A2", false, true},  // 整数
		{"A3", false, true},  // 浮点数
		{"A4", false, false}, // 布尔值
		{"A5", true, false},  // 空值
		{"A6", true, false},  // 不存在的单元格
	}

	for _, tt := range tests {
		t.Run("Cell_"+tt.addr, func(t *testing.T) {
			cell := sheet.cells[tt.addr]
			
			if tt.addr == "A6" && cell != nil {
				t.Error("不存在的单元格应该返回nil")
				return
			}
			
			if tt.addr != "A6" && cell == nil {
				t.Error("存在的单元格不应该返回nil")
				return
			}
			
			if cell != nil {
				isEmpty := cell.IsEmpty()
				if isEmpty != tt.wantEmpty {
					t.Errorf("Cell.IsEmpty() = %v, want %v for %s", isEmpty, tt.wantEmpty, tt.addr)
				}
				
				_, err := cell.Float64()
				isNumeric := err == nil
				if isNumeric != tt.wantNumeric {
					t.Errorf("Cell.GetNumericValue() numeric = %v, want %v for %s", isNumeric, tt.wantNumeric, tt.addr)
				}
			}
		})
	}
}

// TestCell_MergedCells 测试合并单元格处理
func TestCell_MergedCells(t *testing.T) {
	tempDir := t.TempDir()
	testFile := filepath.Join(tempDir, "merged_test.xlsx")
	
	f := excelize.NewFile()
	defer f.Close()
	
	// 创建合并单元格
	f.SetCellValue("Sheet1", "A1", "合并单元格数据")
	f.MergeCell("Sheet1", "A1", "C1")
	
	if err := f.SaveAs(testFile); err != nil {
		t.Fatalf("保存测试文件失败: %v", err)
	}

	logger := zaptest.NewLogger(t)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		t.Fatalf("加载Excel文件失败: %v", err)
	}
	defer excel.Close()

	sheet, err := excel.GetSheet("Sheet1")
	if err != nil {
		t.Fatalf("获取工作表失败: %v", err)
	}

	// 测试主单元格
	if err != nil {
		t.Fatalf("获取工作表失败: %v", err)
	}

	// 测试主单元格
	cellA1 := sheet.cells["A1"]
	if cellA1 == nil {
		t.Fatal("合并单元格主单元格不应该为nil")
	}
	
	if cellA1.String() == "" {
		t.Error("合并单元格主单元格应该有值")
	}
	
	if !cellA1.IsMerged {
		t.Error("A1应该标记为合并单元格")
	}
}

// BenchmarkCell_String 基准测试单元格字符串转换
func BenchmarkCell_String(b *testing.B) {
	cell := &Cell{
		Address: "A1",
		Value:   "测试字符串值",
	}
	
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_ = cell.String()
	}
}

// BenchmarkCell_GetNumericValue 基准测试数值获取
func BenchmarkCell_GetNumericValue(b *testing.B) {
	cell := &Cell{
		Address: "A1",
		Value:   "123.45",
	}
	
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, _ = cell.Float64()
	}
}
