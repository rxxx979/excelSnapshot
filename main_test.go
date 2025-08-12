package main

import (
	"os"
	"os/exec"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestCLI_RunSingleSheetByIndex(t *testing.T) {
	// 准备一个最小 excel
	f := excelize.NewFile()
	sheet := f.GetSheetName(0)
	_ = f.SetCellValue(sheet, "A1", "CLI TEST")
	dir := t.TempDir()
	xlsx := filepath.Join(dir, "c.xlsx")
	png := filepath.Join(dir, "c.png")
	if err := f.SaveAs(xlsx); err != nil {
		t.Fatalf("save excel: %v", err)
	}
	_ = f.Close()

	// 调用 `go run .` 执行 main
	cmd := exec.Command("go", "run", ".", "-in", xlsx, "-sheet-index", "0", "-out", png, "-force-raw")
	cmd.Env = os.Environ()
	cmd.Dir = "."
	out, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("run main failed: %v\noutput: %s", err, string(out))
	}

	st, err := os.Stat(png)
	if err != nil {
		t.Fatalf("png not generated: %v\noutput: %s", err, string(out))
	}
	if st.Size() == 0 {
		t.Fatalf("png empty\noutput: %s", string(out))
	}
}
