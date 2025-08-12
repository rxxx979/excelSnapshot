package main

import (
	"fmt"
	"log"
	"time"

	"github.com/xuri/excelize/v2"
)

// generateDemoExcel creates a demo.xlsx with formatted data, merged cells, borders and styles.
func generateDemoExcel(path string) error {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			log.Printf("关闭 Excel 文件失败: %v", err)
		}
	}()

	sheet := f.GetSheetName(0)

	// Title merged A1:E1
	_ = f.SetCellValue(sheet, "A1", "Excel Snapshot 示例表")
	_ = f.MergeCell(sheet, "A1", "E1")
	// Title style
	titleStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true, Size: 16, Family: "SourceHanSerifSC"},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFFFF2CC"}},
	})
	_ = f.SetCellStyle(sheet, "A1", "E1", titleStyle)
	_ = f.SetRowHeight(sheet, 1, 28)

	// Header row
	headers := []string{"项目", "数量", "单价", "金额", "日期"}
	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 2)
		_ = f.SetCellValue(sheet, cell, h)
	}
	headStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border:    []excelize.Border{{Type: "left", Style: 1, Color: "FF000000"}, {Type: "right", Style: 1, Color: "FF000000"}, {Type: "top", Style: 1, Color: "FF000000"}, {Type: "bottom", Style: 1, Color: "FF000000"}},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFBDD7EE"}},
	})
	_ = f.SetCellStyle(sheet, "A2", "E2", headStyle)
	_ = f.SetRowHeight(sheet, 2, 22)

	// Data rows
	type row struct {
		item  string
		qty   int
		price float64
		date  time.Time
	}
	rows := []row{
		{"苹果", 12, 3.5, time.Date(2025, 8, 1, 0, 0, 0, 0, time.Local)},
		{"香蕉", 6, 4.2, time.Date(2025, 8, 2, 0, 0, 0, 0, time.Local)},
		{"牛奶", 2, 15.8, time.Date(2025, 8, 3, 0, 0, 0, 0, time.Local)},
	}
	for i, r := range rows {
		rowIdx := 3 + i
		_ = f.SetCellValue(sheet, fmt.Sprintf("A%d", rowIdx), r.item)
		_ = f.SetCellValue(sheet, fmt.Sprintf("B%d", rowIdx), r.qty)
		_ = f.SetCellValue(sheet, fmt.Sprintf("C%d", rowIdx), r.price)
		_ = f.SetCellFormula(sheet, fmt.Sprintf("D%d", rowIdx), fmt.Sprintf("B%d*C%d", rowIdx, rowIdx))
		_ = f.SetCellValue(sheet, fmt.Sprintf("E%d", rowIdx), r.date)
	}

	// Number/date formats
	priceStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 2})  // 0.00
	amountStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 4}) // #,##0.00
	dateStyle, _ := f.NewStyle(&excelize.Style{NumFmt: 14})  // m/d/yy
	textLeft, _ := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: "left", Vertical: "center"}})
	center, _ := f.NewStyle(&excelize.Style{Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"}})
	withBorder, _ := f.NewStyle(&excelize.Style{Border: []excelize.Border{{Type: "left", Style: 1, Color: "FF000000"}, {Type: "right", Style: 1, Color: "FF000000"}, {Type: "top", Style: 1, Color: "FF000000"}, {Type: "bottom", Style: 1, Color: "FF000000"}}})

	_ = f.SetCellStyle(sheet, "A3", "A5", textLeft)
	_ = f.SetCellStyle(sheet, "B3", "B5", center)
	_ = f.SetCellStyle(sheet, "C3", "C5", priceStyle)
	_ = f.SetCellStyle(sheet, "D3", "D5", amountStyle)
	_ = f.SetCellStyle(sheet, "E3", "E5", dateStyle)
	_ = f.SetCellStyle(sheet, "A3", "E5", withBorder)

	// Summary row
	sumRow := 6
	_ = f.SetCellValue(sheet, fmt.Sprintf("A%d", sumRow), "合计")
	_ = f.SetCellFormula(sheet, fmt.Sprintf("D%d", sumRow), "SUM(D3:D5)")
	_ = f.MergeCell(sheet, fmt.Sprintf("A%d", sumRow), fmt.Sprintf("C%d", sumRow))
	sumStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "right", Vertical: "center"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"FFE2EFDA"}},
		Border:    []excelize.Border{{Type: "left", Style: 1, Color: "FF000000"}, {Type: "right", Style: 1, Color: "FF000000"}, {Type: "top", Style: 1, Color: "FF000000"}, {Type: "bottom", Style: 1, Color: "FF000000"}},
	})
	_ = f.SetCellStyle(sheet, fmt.Sprintf("A%d", sumRow), fmt.Sprintf("E%d", sumRow), sumStyle)
	_ = f.SetRowHeight(sheet, sumRow, 22)

	// Column widths
	_ = f.SetColWidth(sheet, "A", "A", 16)
	_ = f.SetColWidth(sheet, "B", "B", 8)
	_ = f.SetColWidth(sheet, "C", "C", 10)
	_ = f.SetColWidth(sheet, "D", "D", 12)
	_ = f.SetColWidth(sheet, "E", "E", 12)

	return f.SaveAs(path)
}
