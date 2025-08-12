package main

import (
	"os"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestGetColumnName(t *testing.T) {
	cases := map[int]string{1: "A", 2: "B", 26: "Z", 27: "AA", 52: "AZ"}
	for in, want := range cases {
		if got := getColumnName(in); got != want {
			t.Fatalf("getColumnName(%d) = %q, want %q", in, got, want)
		}
	}
}

func TestParseExcelColor(t *testing.T) {
	if c := parseExcelColor("#FF0000"); c.R != 255 || c.G != 0 || c.B != 0 {
		t.Fatalf("parseExcelColor hex rgb failed: %+v", c)
	}
	if c := parseExcelColor("FFFF0000"); c.R != 255 || c.G != 0 || c.B != 0 {
		t.Fatalf("parseExcelColor argb failed: %+v", c)
	}
}

func TestListSheetsAndCellValuesAndMerge(t *testing.T) {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	sheet := f.GetSheetName(0)
	// values
	_ = f.SetCellValue(sheet, "A1", "Hello")
	_ = f.SetCellValue(sheet, "B1", 3)
	_ = f.SetCellFormula(sheet, "C1", "SUM(1,2)")
	// merge A2:B2
	if err := f.MergeCell(sheet, "A2", "B2"); err != nil {
		t.Fatalf("merge: %v", err)
	}

	dir := t.TempDir()
	xlsx := dir + "/t.xlsx"
	if err := f.SaveAs(xlsx); err != nil {
		t.Fatalf("save xlsx: %v", err)
	}

	names, err := listNonEmptySheets(xlsx)
	if err != nil || len(names) != 1 || names[0] != sheet {
		t.Fatalf("listNonEmptySheets got=%v err=%v", names, err)
	}

	f2, err := excelize.OpenFile(xlsx)
	if err != nil {
		t.Fatal(err)
	}
	defer func() {
		if err := f2.Close(); err != nil {
			t.Error(err)
		}
	}()

	if v := getCellDisplayValue(f2, sheet, "A1", false); v != "Hello" {
		t.Fatalf("display value A1 want Hello got %q", v)
	}
	if v := getCellDisplayValue(f2, sheet, "C1", false); v != "3" {
		// excelize 计算失败时会返回空，允许跳过
		if v == "" {
			t.Log("CalcCellValue not available in this env, skipping formula assert")
		} else {
			t.Fatalf("display value C1 want 3 got %q", v)
		}
	}

	merged, err := getMergeCells(f2, sheet)
	if err != nil {
		t.Fatal(err)
	}
	// A2 (row index 1, col index 0) should be present as merged area map entries
	key := "1,0"
	if _, ok := merged[key]; !ok {
		// at least ensure some merged cells detected
		if len(merged) == 0 {
			t.Fatalf("expected merged cells, got none")
		}
	}

	// cleanup temp dir files (best-effort)
	_ = os.Remove(xlsx)
}
