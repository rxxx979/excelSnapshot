package main

import (
	"flag"
	"fmt"
	"image"
	"image/png"
	"os"
	"path/filepath"
	"strings"

	excelsnapshot "github.com/rxxx/excelSnapshot"
)

func main() {
	var (
		inPath  string
		outPath string
		sheet   string
		index   int
		all     bool
	)

	flag.StringVar(&inPath, "i", "", "输入的 Excel 文件路径 (.xlsx)")
	flag.StringVar(&outPath, "o", ".", "输出目录或文件路径（当渲染单个 sheet 时可指定 .png 文件）")
	flag.StringVar(&sheet, "sheet", "", "要渲染的工作表名称（优先于 index）")
	flag.IntVar(&index, "index", -1, "要渲染的工作表索引（0-based）")
	flag.BoolVar(&all, "all", false, "是否渲染所有工作表")
	flag.Parse()

	if inPath == "" {
		fatalf("必须提供 -i 输入文件路径")
	}
	if _, err := os.Stat(inPath); err != nil {
		fatalf("输入文件不存在: %v", err)
	}

	e, err := excelsnapshot.NewExcel(inPath)
	if err != nil {
		fatalf("打开 Excel 失败: %v", err)
	}
	defer e.Close()

	// 简化：使用内嵌字体与默认渲染参数
	imgs, err := e.Render(sheet, index, all)
	if err != nil {
		fatalf("渲染失败: %v", err)
	}

	// 输出
	if len(imgs) == 1 && !all {
		// 单文件输出：允许 -o 指定具体文件
		for name, img := range imgs {
			outFile := outPath
			if !strings.HasSuffix(strings.ToLower(outFile), ".png") {
				// 作为目录处理
				if err := mkdirAll(outFile); err != nil {
					fatalf("创建输出目录失败: %v", err)
				}
				outFile = filepath.Join(outFile, sanitize(name)+".png")
			}
			if err := savePNG(outFile, img); err != nil {
				fatalf("保存 PNG 失败 (%s): %v", outFile, err)
			}
			fmt.Printf("保存: %s\n", outFile)
		}
		return
	}

	// 多 sheet 输出：-o 必须是目录
	if outPath == "" {
		outPath = "."
	}
	if err := mkdirAll(outPath); err != nil {
		fatalf("创建输出目录失败: %v", err)
	}
	for name, img := range imgs {
		file := filepath.Join(outPath, sanitize(name)+".png")
		if err := savePNG(file, img); err != nil {
			fatalf("保存 PNG 失败 (%s): %v", file, err)
		}
		fmt.Printf("保存: %s\n", file)
	}
}

func savePNG(path string, img image.Image) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()
	return png.Encode(f, img)
}

func mkdirAll(dir string) error {
	if dir == "" || dir == "." {
		return nil
	}
	return os.MkdirAll(dir, 0o755)
}

func sanitize(name string) string {
	name = strings.TrimSpace(name)
	name = strings.ReplaceAll(name, "/", "-")
	name = strings.ReplaceAll(name, "\\", "-")
	name = strings.ReplaceAll(name, ":", "-")
	name = strings.ReplaceAll(name, "*", "-")
	name = strings.ReplaceAll(name, "?", "-")
	name = strings.ReplaceAll(name, "\"", "-")
	name = strings.ReplaceAll(name, "<", "-")
	name = strings.ReplaceAll(name, ">", "-")
	name = strings.ReplaceAll(name, "|", "-")
	if name == "" {
		name = "sheet"
	}
	return name
}

func fatalf(format string, a ...any) {
	fmt.Fprintf(os.Stderr, format+"\n", a...)
	os.Exit(1)
}
