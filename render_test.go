package excelsnapshot

import (
	"fmt"
	"math/rand"
	"path/filepath"
	"testing"
	"time"

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

	// 创建测试文件
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", "Test1")
	f.SetCellValue("Sheet1", "B2", "Test2")

	if err := f.SaveAs(testFile); err != nil {
		b.Fatalf("保存测试文件失败: %v", err)
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

// createComplexWorksheet 创建包含随机数据、合并单元格和颜色的复杂工作表
func createComplexWorksheet(f *excelize.File, sheetName string, rows, cols int) error {
	rand.Seed(time.Now().UnixNano())
	
	// 定义颜色样式
	colors := []string{"FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF"}
	
	for row := 1; row <= rows; row++ {
		for col := 1; col <= cols; col++ {
			cellAddr, _ := excelize.CoordinatesToCellName(col, row)
			
			// 随机填充不同类型的数据
			switch rand.Intn(4) {
			case 0:
				// 随机数字
				f.SetCellValue(sheetName, cellAddr, rand.Float64()*1000)
			case 1:
				// 随机文本
				f.SetCellValue(sheetName, cellAddr, fmt.Sprintf("文本%d", rand.Intn(100)))
			case 2:
				// 随机整数
				f.SetCellValue(sheetName, cellAddr, rand.Intn(1000))
			case 3:
				// 随机布尔值
				f.SetCellValue(sheetName, cellAddr, rand.Intn(2) == 1)
			}
			
			// 随机添加背景颜色
			if rand.Float32() < 0.3 { // 30% 概率添加颜色
				color := colors[rand.Intn(len(colors))]
				styleID, _ := f.NewStyle(&excelize.Style{
					Fill: excelize.Fill{
						Type:    "pattern",
						Color:   []string{color},
						Pattern: 1,
					},
				})
				f.SetCellStyle(sheetName, cellAddr, cellAddr, styleID)
			}
		}
	}
	
	// 添加一些合并单元格
	mergeCount := (rows * cols) / 20 // 5% 的单元格参与合并
	for i := 0; i < mergeCount; i++ {
		startRow := rand.Intn(rows-1) + 1
		startCol := rand.Intn(cols-1) + 1
		endRow := min(startRow+rand.Intn(3), rows)
		endCol := min(startCol+rand.Intn(3), cols)
		
		startCell, _ := excelize.CoordinatesToCellName(startCol, startRow)
		endCell, _ := excelize.CoordinatesToCellName(endCol, endRow)
		f.MergeCell(sheetName, startCell, endCell)
	}
	
	return nil
}

// min 辅助函数
func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}

// BenchmarkSheetRenderer_SmallSheet 基准测试小工作表渲染性能 (8x10)
func BenchmarkSheetRenderer_SmallSheet(b *testing.B) {
	tempDir := b.TempDir()
	testFile := filepath.Join(tempDir, "bench_small.xlsx")
	
	f := excelize.NewFile()
	defer f.Close()
	
	createComplexWorksheet(f, "Sheet1", 10, 8)
	
	if err := f.SaveAs(testFile); err != nil {
		b.Fatalf("保存小工作表失败: %v", err)
	}

	logger := zaptest.NewLogger(b)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		b.Fatalf("加载小Excel文件失败: %v", err)
	}
	defer excel.Close()

	sheet, err := excel.GetSheet("Sheet1")
	if err != nil {
		b.Fatalf("获取小工作表失败: %v", err)
	}

	renderer := NewSheetRenderer(logger)
	b.ResetTimer()
	
	for i := 0; i < b.N; i++ {
		_, err := renderer.RenderSheet(sheet)
		if err != nil {
			b.Fatalf("渲染小工作表失败: %v", err)
		}
	}
}

// BenchmarkSheetRenderer_MediumSheet 基准测试中等工作表渲染性能 (16x30)
func BenchmarkSheetRenderer_MediumSheet(b *testing.B) {
	tempDir := b.TempDir()
	testFile := filepath.Join(tempDir, "bench_medium.xlsx")
	
	f := excelize.NewFile()
	defer f.Close()
	
	createComplexWorksheet(f, "Sheet1", 30, 16)
	
	if err := f.SaveAs(testFile); err != nil {
		b.Fatalf("保存中等工作表失败: %v", err)
	}

	logger := zaptest.NewLogger(b)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		b.Fatalf("加载中等Excel文件失败: %v", err)
	}
	defer excel.Close()

	sheet, err := excel.GetSheet("Sheet1")
	if err != nil {
		b.Fatalf("获取中等工作表失败: %v", err)
	}

	renderer := NewSheetRenderer(logger)
	b.ResetTimer()
	
	for i := 0; i < b.N; i++ {
		_, err := renderer.RenderSheet(sheet)
		if err != nil {
			b.Fatalf("渲染中等工作表失败: %v", err)
		}
	}
}

// BenchmarkSheetRenderer_LargeSheet 基准测试大工作表渲染性能 (25x100)
func BenchmarkSheetRenderer_LargeSheet(b *testing.B) {
	tempDir := b.TempDir()
	testFile := filepath.Join(tempDir, "bench_large.xlsx")
	
	f := excelize.NewFile()
	defer f.Close()
	
	createComplexWorksheet(f, "Sheet1", 100, 25)
	
	if err := f.SaveAs(testFile); err != nil {
		b.Fatalf("保存大工作表失败: %v", err)
	}

	logger := zaptest.NewLogger(b)
	excel, err := NewExcel(testFile, logger)
	if err != nil {
		b.Fatalf("加载大Excel文件失败: %v", err)
	}
	defer excel.Close()

	sheet, err := excel.GetSheet("Sheet1")
	if err != nil {
		b.Fatalf("获取大工作表失败: %v", err)
	}

	renderer := NewSheetRenderer(logger)
	b.ResetTimer()
	
	for i := 0; i < b.N; i++ {
		_, err := renderer.RenderSheet(sheet)
		if err != nil {
			b.Fatalf("渲染大工作表失败: %v", err)
		}
	}
}
