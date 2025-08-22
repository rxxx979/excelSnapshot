package excelsnapshot

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
	"go.uber.org/zap/zaptest"
)

// TestNewSheet 测试Sheet结构体的创建
func TestNewSheet(t *testing.T) {
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

	sheetName := "Sheet1"
	sheet := NewSheet(excel, sheetName)
	
	if sheet == nil {
		t.Fatal("NewSheet() 返回 nil")
	}
	
	if sheet.Name != sheetName {
		t.Errorf("Sheet.Name = %v, want %v", sheet.Name, sheetName)
	}
	
	if sheet.excel != excel {
		t.Error("Sheet.excel 引用不正确")
	}
}

// TestSheet_Load 测试工作表加载
func TestSheet_Load(t *testing.T) {
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

	sheet := NewSheet(excel, "Sheet1")
	
	err = sheet.Load()
	if err != nil {
		t.Fatalf("Sheet.Load() 失败: %v", err)
	}
	
	// 验证基本属性
	if sheet.Rows <= 0 {
		t.Errorf("Sheet.Rows = %v, want > 0", sheet.Rows)
	}
	
	if sheet.Cols <= 0 {
		t.Errorf("Sheet.Cols = %v, want > 0", sheet.Cols)
	}
	
	if sheet.cells == nil {
		t.Error("Sheet.cells 应该被初始化")
	}
}

// TestSheet_LoadInvalidSheet 测试加载无效工作表
func TestSheet_LoadInvalidSheet(t *testing.T) {
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

	sheet := NewSheet(excel, "不存在的工作表")
	
	err = sheet.Load()
	if err == nil {
		t.Error("加载不存在的工作表应该返回错误")
	}
}

// TestSheet_GetCell 测试获取单元格
func TestSheet_GetCell(t *testing.T) {
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

	sheet := NewSheet(excel, "Sheet1")
	if err := sheet.Load(); err != nil {
		t.Fatalf("Sheet.Load() 失败: %v", err)
	}

	tests := []struct {
		name string
		addr string
		want bool // 是否应该存在单元格
	}{
		{
			name: "存在的单元格A1",
			addr: "A1",
			want: true,
		},
		{
			name: "存在的单元格B1",
			addr: "B1",
			want: true,
		},
		{
			name: "空单元格",
			addr: "Z100",
			want: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cell := sheet.cells[tt.addr]
			exists := (cell != nil)
			if exists != tt.want {
				t.Errorf("GetCell(%v) 存在性 = %v, want %v", tt.addr, exists, tt.want)
			}
		})
	}
}

// TestSheet_loadImages 测试图片加载功能
func TestSheet_loadImages(t *testing.T) {
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

	sheet := NewSheet(excel, "Sheet1")
	if err := sheet.Load(); err != nil {
		t.Fatalf("Sheet.Load() 失败: %v", err)
	}

	// 测试loadImages方法
	err = sheet.loadImages()
	if err != nil {
		t.Errorf("loadImages() 失败: %v", err)
	}
	
	// 由于测试文件没有图片，images应该为空
	if len(sheet.images) != 0 {
		t.Errorf("images长度 = %v, want 0", len(sheet.images))
	}
}

// TestSheet_calculateBounds 测试边界计算
func TestSheet_calculateBounds(t *testing.T) {
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

	sheet := NewSheet(excel, "Sheet1")
	if err := sheet.Load(); err != nil {
		t.Fatalf("Sheet.Load() 失败: %v", err)
	}

	// calculateBounds 应该已在Load中被调用
	if sheet.Rows <= 0 {
		t.Errorf("Rows = %v, want > 0", sheet.Rows)
	}
	
	if sheet.Cols <= 0 {
		t.Errorf("Cols = %v, want > 0", sheet.Cols)
	}
}

// createTestExcelWithImages 创建包含图片的测试Excel文件
func createTestExcelWithImages(filename string) error {
	f := excelize.NewFile()
	defer f.Close()

	// 添加测试数据
	f.SetCellValue("Sheet1", "A1", "测试数据")
	f.SetCellValue("Sheet1", "B1", "Hello")
	f.SetCellValue("Sheet1", "A2", 123)
	f.SetCellValue("Sheet1", "B2", 456.78)

	// 注意：在实际测试中添加图片需要真实的图片文件
	// 这里只创建基本的Excel文件，图片测试需要额外的设置

	return f.SaveAs(filename)
}

// BenchmarkSheet_Load 基准测试工作表加载性能
func BenchmarkSheet_Load(b *testing.B) {
	tempDir := b.TempDir()
	testFile := filepath.Join(tempDir, "bench.xlsx")
	
	if err := createTestExcelFile(testFile); err != nil {
		b.Fatalf("创建测试文件失败: %v", err)
	}

	logger := zaptest.NewLogger(b)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		b.Fatalf("加载Excel文件失败: %v", err)
	}
	defer excel.Close()

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		sheet := NewSheet(excel, "Sheet1")
		if err := sheet.Load(); err != nil {
			b.Fatalf("Sheet.Load() 失败: %v", err)
		}
	}
}

// BenchmarkSheet_GetCell 基准测试单元格获取性能
func BenchmarkSheet_GetCell(b *testing.B) {
	tempDir := b.TempDir()
	testFile := filepath.Join(tempDir, "bench.xlsx")
	
	if err := createTestExcelFile(testFile); err != nil {
		b.Fatalf("创建测试文件失败: %v", err)
	}

	logger := zaptest.NewLogger(b)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		b.Fatalf("加载Excel文件失败: %v", err)
	}
	defer excel.Close()

	sheet := NewSheet(excel, "Sheet1")
	if err := sheet.Load(); err != nil {
		b.Fatalf("Sheet.Load() 失败: %v", err)
	}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_ = sheet.cells["A1"]
	}
}
