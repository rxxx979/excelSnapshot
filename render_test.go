package excelsnapshot

import (
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
	"go.uber.org/zap/zaptest"
)

// TestNewSheetRenderer 测试SheetRenderer的创建
func TestNewSheetRenderer(t *testing.T) {
	logger := zaptest.NewLogger(t)
	renderer := NewSheetRenderer(logger)

	if renderer == nil {
		t.Fatal("NewSheetRenderer() 返回 nil")
	}

	if renderer.logger != logger {
		t.Error("SheetRenderer.logger 设置不正确")
	}
}

// TestSheetRenderer_RenderSheet 测试工作表渲染
func TestSheetRenderer_RenderSheet(t *testing.T) {
	// 创建测试Excel文件
	tempDir := t.TempDir()
	testFile := filepath.Join(tempDir, "render_test.xlsx")

	if err := createTestExcelFile(testFile); err != nil {
		t.Fatalf("创建测试文件失败: %v", err)
	}

	// 加载Excel和工作表
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

	// 创建渲染器并渲染
	renderer := NewSheetRenderer(logger)
	img, err := renderer.RenderSheet(sheet)

	if err != nil {
		t.Fatalf("RenderSheet() 失败: %v", err)
	}

	if img == nil {
		t.Fatal("RenderSheet() 返回 nil 图片")
	}

	// 验证图片基本属性
	bounds := img.Bounds()
	if bounds.Dx() <= 0 || bounds.Dy() <= 0 {
		t.Errorf("图片尺寸无效: %dx%d", bounds.Dx(), bounds.Dy())
	}
}

// TestSheetRenderer_calculateCellRects 测试单元格矩形计算
func TestSheetRenderer_calculateCellRects(t *testing.T) {
	tempDir := t.TempDir()
	testFile := filepath.Join(tempDir, "calc_test.xlsx")

	if err := createTestExcelFile(testFile); err != nil {
		t.Fatalf("创建测试文件失败: %v", err)
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

	renderer := NewSheetRenderer(logger)
	cellRects := renderer.calculateCellRects(sheet)

	if len(cellRects) == 0 {
		t.Error("calculateCellRects() 返回空结果")
	}

	// 验证基本单元格存在
	if _, exists := cellRects["A1"]; !exists {
		t.Error("A1单元格矩形不存在")
	}
}

// TestSheetRenderer_EmptySheet 测试空工作表渲染
func TestSheetRenderer_EmptySheet(t *testing.T) {
	tempDir := t.TempDir()
	testFile := filepath.Join(tempDir, "empty_test.xlsx")

	// 创建空的Excel文件
	f := excelize.NewFile()
	defer f.Close()
	if err := f.SaveAs(testFile); err != nil {
		t.Fatalf("保存空文件失败: %v", err)
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

	renderer := NewSheetRenderer(logger)
	img, err := renderer.RenderSheet(sheet)

	if err != nil {
		t.Fatalf("渲染空工作表失败: %v", err)
	}

	if img == nil {
		t.Fatal("空工作表渲染返回 nil")
	}
}

// TestSheetRenderer_LargeSheet 测试大工作表渲染性能
func TestSheetRenderer_LargeSheet(t *testing.T) {
	if testing.Short() {
		t.Skip("跳过大工作表测试（短测试模式）")
	}

	tempDir := t.TempDir()
	testFile := filepath.Join(tempDir, "large_test.xlsx")

	// 创建包含大量数据的Excel文件
	f := excelize.NewFile()
	defer f.Close()

	// 添加100行20列的数据
	for row := 1; row <= 100; row++ {
		for col := 1; col <= 20; col++ {
			cellAddr, _ := excelize.CoordinatesToCellName(col, row)
			f.SetCellValue("Sheet1", cellAddr, "数据")
		}
	}

	if err := f.SaveAs(testFile); err != nil {
		t.Fatalf("保存大文件失败: %v", err)
	}

	logger := zaptest.NewLogger(t)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		t.Fatalf("加载大Excel文件失败: %v", err)
	}
	defer excel.Close()

	sheet, err := excel.GetSheet("Sheet1")
	if err != nil {
		t.Fatalf("获取大工作表失败: %v", err)
	}

	renderer := NewSheetRenderer(logger)
	img, err := renderer.RenderSheet(sheet)

	if err != nil {
		t.Fatalf("渲染大工作表失败: %v", err)
	}

	if img == nil {
		t.Fatal("大工作表渲染返回 nil")
	}

	// 验证图片尺寸合理
	bounds := img.Bounds()
	if bounds.Dx() < 100 || bounds.Dy() < 100 {
		t.Errorf("大工作表图片尺寸过小: %dx%d", bounds.Dx(), bounds.Dy())
	}
}

// TestSheetRenderer_InvalidInput 测试无效输入处理
func TestSheetRenderer_InvalidInput(t *testing.T) {
	logger := zaptest.NewLogger(t)
	renderer := NewSheetRenderer(logger)

	// 测试nil工作表
	_, err := renderer.RenderSheet(nil)
	if err == nil {
		t.Error("渲染nil工作表应该返回错误")
	}
}

// BenchmarkSheetRenderer_RenderSheet 基准测试渲染性能
func BenchmarkSheetRenderer_RenderSheet(b *testing.B) {
	tempDir := b.TempDir()
	testFile := filepath.Join(tempDir, "bench_render.xlsx")

	if err := createTestExcelFile(testFile); err != nil {
		b.Fatalf("创建测试文件失败: %v", err)
	}

	logger := zaptest.NewLogger(b)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		b.Fatalf("加载Excel文件失败: %v", err)
	}
	defer excel.Close()

	sheet, err := excel.GetSheet("Sheet1")
	if err != nil {
		b.Fatalf("获取工作表失败: %v", err)
	}

	renderer := NewSheetRenderer(logger)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, err := renderer.RenderSheet(sheet)
		if err != nil {
			b.Fatalf("RenderSheet() 失败: %v", err)
		}
	}
}

// BenchmarkSheetRenderer_calculateCellRects 基准测试单元格计算性能
func BenchmarkSheetRenderer_calculateCellRects(b *testing.B) {
	tempDir := b.TempDir()
	testFile := filepath.Join(tempDir, "bench_calc.xlsx")

	if err := createTestExcelFile(testFile); err != nil {
		b.Fatalf("创建测试文件失败: %v", err)
	}

	logger := zaptest.NewLogger(b)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		b.Fatalf("加载Excel文件失败: %v", err)
	}
	defer excel.Close()

	sheet, err := excel.GetSheet("Sheet1")
	if err != nil {
		b.Fatalf("获取工作表失败: %v", err)
	}

	renderer := NewSheetRenderer(logger)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		renderer.calculateCellRects(sheet)
	}
}
