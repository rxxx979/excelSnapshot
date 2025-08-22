package excelsnapshot

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"
	"go.uber.org/zap/zaptest"
)

// TestNewExcel 测试Excel结构体的创建
func TestNewExcel(t *testing.T) {
	logger := zaptest.NewLogger(t)
	
	tests := []struct {
		name     string
		filename string
		wantErr  bool
	}{
		{
			name:     "不存在的文件",
			filename: "nonexistent.xlsx",
			wantErr:  true,
		},
		{
			name:     "空文件名",
			filename: "",
			wantErr:  true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, err := NewExcel(tt.filename, logger)
			if (err != nil) != tt.wantErr {
				t.Errorf("NewExcel() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

// TestExcel_GetSheetList 测试获取工作表名称
func TestExcel_GetSheetList(t *testing.T) {
	// 创建临时Excel文件进行测试
	tempDir := t.TempDir()
	testFile := filepath.Join(tempDir, "test.xlsx")
	
	// 创建一个简单的Excel文件用于测试
	if err := createTestExcelFile(testFile); err != nil {
		t.Fatalf("创建测试文件失败: %v", err)
	}

	logger := zaptest.NewLogger(t)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		t.Fatalf("加载Excel文件失败: %v", err)
	}
	defer excel.Close()

	names := excel.file.GetSheetList()
	if len(names) == 0 {
		t.Error("GetSheetList() 返回空列表")
	}
}

// TestExcel_GetSheet 测试加载工作表
func TestExcel_GetSheet(t *testing.T) {
	tempDir := t.TempDir()
	testFile := filepath.Join(tempDir, "test.xlsx")
	
	if err := createTestExcelFile(testFile); err != nil {
		t.Fatalf("创建测试文件失败: %v", err)
	}

	logger := zaptest.NewLogger(t)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		t.Fatalf("加载Excel文件失败: %v", err)
	}
	defer excel.Close()

	names := excel.file.GetSheetList()
	if len(names) == 0 {
		t.Fatal("没有可用的工作表")
	}

	tests := []struct {
		name      string
		sheetName string
		wantErr   bool
	}{
		{
			name:      "有效的工作表名",
			sheetName: names[0],
			wantErr:   false,
		},
		{
			name:      "无效的工作表名",
			sheetName: "不存在的工作表",
			wantErr:   true,
		},
		{
			name:      "空工作表名",
			sheetName: "",
			wantErr:   true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			_, err := excel.GetSheet(tt.sheetName)
			if (err != nil) != tt.wantErr {
				t.Errorf("GetSheet() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}

// TestExcel_Close 测试关闭Excel文件
func TestExcel_Close(t *testing.T) {
	tempDir := t.TempDir()
	testFile := filepath.Join(tempDir, "test.xlsx")
	
	if err := createTestExcelFile(testFile); err != nil {
		t.Fatalf("创建测试文件失败: %v", err)
	}

	logger := zaptest.NewLogger(t)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		t.Fatalf("加载Excel文件失败: %v", err)
	}

	// 测试Close方法不会panic
	excel.Close()
	
	// 多次调用Close应该是安全的
	excel.Close()
}

// createTestExcelFile 创建一个简单的测试Excel文件
func createTestExcelFile(filename string) error {
	// 使用excelize库创建一个简单的Excel文件
	f := excelize.NewFile()
	defer f.Close()

	// 添加一些测试数据
	f.SetCellValue("Sheet1", "A1", "测试数据")
	f.SetCellValue("Sheet1", "B1", "Hello")
	f.SetCellValue("Sheet1", "A2", 123)
	f.SetCellValue("Sheet1", "B2", 456.78)

	return f.SaveAs(filename)
}

// BenchmarkNewExcel 基准测试Excel文件加载性能
func BenchmarkNewExcel(b *testing.B) {
	tempDir := b.TempDir()
	testFile := filepath.Join(tempDir, "bench.xlsx")
	
	if err := createTestExcelFile(testFile); err != nil {
		b.Fatalf("创建测试文件失败: %v", err)
	}

	logger := zap.NewNop() // 使用空日志器避免影响基准测试
	
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		excel, err := NewExcel(testFile, logger)
		if err != nil {
			b.Fatalf("加载Excel文件失败: %v", err)
		}
		excel.Close()
	}
}

// BenchmarkLoadSheet 基准测试工作表加载性能
func BenchmarkLoadSheet(b *testing.B) {
	tempDir := b.TempDir()
	testFile := filepath.Join(tempDir, "bench.xlsx")
	
	if err := createTestExcelFile(testFile); err != nil {
		b.Fatalf("创建测试文件失败: %v", err)
	}

	logger := zap.NewNop()
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		b.Fatalf("加载Excel文件失败: %v", err)
	}
	defer excel.Close()

	sheetName := excel.file.GetSheetList()[0]
	
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := excel.GetSheet(sheetName)
		if err != nil {
			b.Fatalf("加载工作表失败: %v", err)
		}
	}
}
