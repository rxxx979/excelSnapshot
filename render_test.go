package main

import (
	"os"
	"testing"

	"github.com/xuri/excelize/v2"
)

// Integration test for renderPNGFromExcel with a tiny sheet
func TestRenderPNGFromExcel(t *testing.T) {
	initFonts()

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	sheet := f.GetSheetName(0)
	_ = f.SetCellValue(sheet, "A1", "Hello 渲染")
	_ = f.SetCellValue(sheet, "B1", 123)
	_ = f.SetCellValue(sheet, "A2", "合并")
	_ = f.SetCellValue(sheet, "B2", "单元格")
	if err := f.MergeCell(sheet, "A2", "B2"); err != nil {
		t.Fatalf("merge failed: %v", err)
	}

	dir := t.TempDir()
	xlsx := dir + "/r.xlsx"
	png := dir + "/r.png"
	if err := f.SaveAs(xlsx); err != nil {
		t.Fatalf("save: %v", err)
	}

	if err := renderPNGFromExcel(xlsx, "", 0, png, false); err != nil {
		t.Fatalf("render error: %v", err)
	}
	st, err := os.Stat(png)
	if err != nil {
		t.Fatalf("png not found: %v", err)
	}
	if st.Size() == 0 {
		t.Fatalf("png is empty")
	}
}
