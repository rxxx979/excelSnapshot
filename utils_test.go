package excelsnapshot

import (
	"testing"

	"github.com/xuri/excelize/v2"
)

// TestCoordinateConversion 测试坐标转换功能
func TestCoordinateConversion(t *testing.T) {
	tests := []struct {
		name string
		col  int
		row  int
		want string
	}{
		{
			name: "A1坐标",
			col:  1,
			row:  1,
			want: "A1",
		},
		{
			name: "B2坐标",
			col:  2,
			row:  2,
			want: "B2",
		},
		{
			name: "Z26坐标",
			col:  26,
			row:  26,
			want: "Z26",
		},
		{
			name: "AA27坐标",
			col:  27,
			row:  27,
			want: "AA27",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := excelize.CoordinatesToCellName(tt.col, tt.row)
			if err != nil {
				t.Errorf("CoordinatesToCellName() error = %v", err)
				return
			}
			if got != tt.want {
				t.Errorf("CoordinatesToCellName() = %v, want %v", got, tt.want)
			}
		})
	}
}

// TestCellNameToCoordinates 测试单元格名称到坐标转换
func TestCellNameToCoordinates(t *testing.T) {
	tests := []struct {
		name     string
		cellName string
		wantCol  int
		wantRow  int
		wantErr  bool
	}{
		{
			name:     "A1",
			cellName: "A1",
			wantCol:  1,
			wantRow:  1,
			wantErr:  false,
		},
		{
			name:     "B2",
			cellName: "B2",
			wantCol:  2,
			wantRow:  2,
			wantErr:  false,
		},
		{
			name:     "Z26",
			cellName: "Z26",
			wantCol:  26,
			wantRow:  26,
			wantErr:  false,
		},
		{
			name:     "AA27",
			cellName: "AA27",
			wantCol:  27,
			wantRow:  27,
			wantErr:  false,
		},
		{
			name:     "无效单元格名",
			cellName: "Invalid",
			wantCol:  -1,
			wantRow:  -1,
			wantErr:  true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			gotCol, gotRow, err := excelize.CellNameToCoordinates(tt.cellName)
			if (err != nil) != tt.wantErr {
				t.Errorf("CellNameToCoordinates() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if gotCol != tt.wantCol {
				t.Errorf("CellNameToCoordinates() col = %v, want %v", gotCol, tt.wantCol)
			}
			if gotRow != tt.wantRow {
				t.Errorf("CellNameToCoordinates() row = %v, want %v", gotRow, tt.wantRow)
			}
		})
	}
}

// TestStringToFloat 测试字符串到浮点数转换
func TestStringToFloat(t *testing.T) {
	tests := []struct {
		name    string
		input   string
		want    float64
		wantErr bool
	}{
		{
			name:    "整数字符串",
			input:   "123",
			want:    123.0,
			wantErr: false,
		},
		{
			name:    "浮点数字符串",
			input:   "123.45",
			want:    123.45,
			wantErr: false,
		},
		{
			name:    "负数",
			input:   "-456.78",
			want:    -456.78,
			wantErr: false,
		},
		{
			name:    "零",
			input:   "0",
			want:    0.0,
			wantErr: false,
		},
		{
			name:    "非数字字符串",
			input:   "不是数字",
			want:    0,
			wantErr: true,
		},
		{
			name:    "空字符串",
			input:   "",
			want:    0,
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			// 模拟Cell的Float64方法测试
			cell := &Cell{Value: tt.input}
			got, err := cell.Float64()

			if (err != nil) != tt.wantErr {
				t.Errorf("Float64() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if got != tt.want {
				t.Errorf("Float64() = %v, want %v", got, tt.want)
			}
		})
	}
}

// TestStringToInt 测试字符串到整数转换
func TestStringToInt(t *testing.T) {
	tests := []struct {
		name    string
		input   string
		want    int
		wantErr bool
	}{
		{
			name:    "正整数",
			input:   "123",
			want:    123,
			wantErr: false,
		},
		{
			name:    "负整数",
			input:   "-456",
			want:    -456,
			wantErr: false,
		},
		{
			name:    "零",
			input:   "0",
			want:    0,
			wantErr: false,
		},
		{
			name:    "浮点数字符串",
			input:   "123.45",
			want:    0,
			wantErr: true,
		},
		{
			name:    "非数字字符串",
			input:   "不是数字",
			want:    0,
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			// 模拟Cell的Int方法测试
			cell := &Cell{Value: tt.input}
			got, err := cell.Int()

			if (err != nil) != tt.wantErr {
				t.Errorf("Int() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if got != tt.want {
				t.Errorf("Int() = %v, want %v", got, tt.want)
			}
		})
	}
}

// TestCellIsEmpty 测试单元格空值判断
func TestCellIsEmpty(t *testing.T) {
	tests := []struct {
		name string
		cell *Cell
		want bool
	}{
		{
			name: "非空文本",
			cell: &Cell{Value: "测试"},
			want: false,
		},
		{
			name: "空字符串",
			cell: &Cell{Value: ""},
			want: true,
		},
		{
			name: "只有空格",
			cell: &Cell{Value: "   "},
			want: true,
		},
		{
			name: "nil单元格",
			cell: nil,
			want: true,
		},
		{
			name: "数字字符串",
			cell: &Cell{Value: "123"},
			want: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := tt.cell.IsEmpty()
			if got != tt.want {
				t.Errorf("IsEmpty() = %v, want %v", got, tt.want)
			}
		})
	}
}

// BenchmarkCoordinateConversion 基准测试坐标转换性能
func BenchmarkCoordinateConversion(b *testing.B) {
	for i := 0; i < b.N; i++ {
		_, _ = excelize.CoordinatesToCellName(i%100+1, i%100+1)
	}
}

// BenchmarkCellNameToCoordinates 基准测试单元格名称转换性能
func BenchmarkCellNameToCoordinates(b *testing.B) {
	cellNames := []string{"A1", "B2", "C3", "D4", "E5"}
	for i := 0; i < b.N; i++ {
		_, _, _ = excelize.CellNameToCoordinates(cellNames[i%len(cellNames)])
	}
}

// BenchmarkCellFloat64 基准测试单元格浮点数转换性能
func BenchmarkCellFloat64(b *testing.B) {
	cell := &Cell{Value: "123.45"}
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_, _ = cell.Float64()
	}
}

// BenchmarkCellIsEmpty 基准测试单元格空值判断性能
func BenchmarkCellIsEmpty(b *testing.B) {
	cell := &Cell{Value: "测试数据"}
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		_ = cell.IsEmpty()
	}
}
