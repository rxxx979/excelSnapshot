package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
)

func main() {
	log.SetFlags(0)
	log.SetOutput(os.Stderr)

	var (
		inPath     = flag.String("in", "", "Excel 文件路径 (*.xlsx)")
		sheet      = flag.String("sheet", "", "Sheet 名称（留空则用索引）")
		sheetIndex = flag.Int("sheet-index", 0, "Sheet 索引（从 0 开始）")
		outPath    = flag.String("out", "", "输出文件路径（.png）")
		allSheets  = flag.Bool("all-sheets", false, "导出所有有内容的工作表（输出为多个文件）")
		forceRaw   = flag.Bool("force-raw", false, "强制使用原始未格式化数值渲染")
		genDemo    = flag.String("gen-demo", "", "生成示例 Excel 文件到该路径（如 demo.xlsx），生成后退出")
	)

	flag.Usage = func() {
		exe := filepath.Base(os.Args[0])
		fmt.Fprintf(os.Stderr, "用法:\n  %s -in data.xlsx [-sheet Sheet1|-sheet-index 0] -out out.png\n\n", exe)
		fmt.Fprintf(os.Stderr, "常用参数:\n")
		flag.PrintDefaults()
	}
	flag.Parse()

	forceRawValues := *forceRaw

	// 仅生成示例表并退出
	if strings.TrimSpace(*genDemo) != "" {
		if err := generateDemoExcel(*genDemo); err != nil {
			log.Fatal(err)
		}
		log.Printf("已生成示例: %s", *genDemo)
		return
	}

	if strings.TrimSpace(*inPath) == "" {
		flag.Usage()
		os.Exit(2)
	}

	if strings.TrimSpace(*outPath) == "" {
		*outPath = "excel_screenshot.png"
	} else if strings.ToLower(filepath.Ext(*outPath)) != ".png" {
		*outPath = strings.TrimSuffix(*outPath, filepath.Ext(*outPath)) + ".png"
	}

	initFonts()

	exportAll := *allSheets

	if exportAll {
		names, err := listNonEmptySheets(*inPath)
		if err != nil {
			log.Fatal(err)
		}
		if len(names) == 0 {
			log.Fatal("没有可导出的非空工作表")
		}
		dir := "."
		base := "excel_screenshot"
		if *outPath != "" {
			dir = filepath.Dir(*outPath)
			b := strings.TrimSuffix(filepath.Base(*outPath), filepath.Ext(*outPath))
			if b != "" {
				base = b
			}
		}
		for _, sn := range names {
			out := filepath.Join(dir, fmt.Sprintf("%s_%s.png", base, sanitizeFileComponent(sn)))
			if err := renderPNGFromExcel(*inPath, sn, 0, out, forceRawValues); err != nil {
				log.Printf("导出 PNG 失败: %s: %v", sn, err)
			} else {
				log.Printf("已生成: %s", out)
			}
		}
		return
	}

	if err := renderPNGFromExcel(*inPath, *sheet, *sheetIndex, *outPath, forceRawValues); err != nil {
		log.Fatal(err)
	}
	log.Printf("已生成: %s", *outPath)
}
