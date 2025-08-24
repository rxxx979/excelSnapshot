package main

import (
	"flag"
	"fmt"
	"image"
	"image/png"
	"os"
	"path/filepath"
	"runtime/debug"
	"strings"
	"time"

	excelsnapshot "github.com/rxxx/excelSnapshot"
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
)

// CLI参数结构
type CLIArgs struct {
	inPath  string
	outPath string
	sheet   string
	index   int
	all     bool
	verbose bool
}

// 解析命令行参数
func parseArgs() *CLIArgs {
	args := &CLIArgs{}

	flag.StringVar(&args.inPath, "i", "", "输入的 Excel 文件路径 (.xlsx)")
	flag.StringVar(&args.outPath, "o", ".", "输出目录或文件路径（当渲染单个 sheet 时可指定 .png 文件）")
	flag.StringVar(&args.sheet, "sheet", "", "要渲染的工作表名称（优先于 index）")
	flag.IntVar(&args.index, "index", -1, "要渲染的工作表索引（0-based）")
	flag.BoolVar(&args.all, "all", false, "是否渲染所有工作表")
	flag.BoolVar(&args.verbose, "v", false, "启用调试日志（开发模式）")
	flag.Parse()

	// 参数验证
	if args.inPath == "" {
		fmt.Println("错误: 必须指定输入文件路径 -i")
		flag.Usage()
		os.Exit(1)
	}

	return args
}

// 初始化日志
func setupLogger(verbose bool) (*zap.Logger, func(), error) {
	var level zapcore.Level = zap.InfoLevel
	isDev := false

	if verbose {
		isDev = true
		level = zap.DebugLevel
	} else {
		level = zap.InfoLevel
	}

	return excelsnapshot.SetupLogger("excel_snapshot", level, isDev)
}

// 确定要渲染的工作表名称
func determineTargetSheet(args *CLIArgs, excel *excelsnapshot.Excel) (string, error) {
	if args.sheet != "" {
		return args.sheet, nil
	}

	if args.index >= 0 {
		sheetName := excel.GetSheetNameByIndex(args.index)
		if sheetName == "" {
			return "", fmt.Errorf("工作表索引 %d 超出范围", args.index)
		}
		return sheetName, nil
	}

	// 默认渲染第一个工作表
	return excel.GetSheetNameByIndex(0), nil
}

// 生成输出文件路径：
// - 若 basePath 以 .png 结尾，按文件路径使用；
// - 否则视为目录：目录必须已存在，否则返回错误；存在则在目录内自动命名（excel名_sheet_时间戳.png）。
func generateOutputPath(basePath, sheetName, excelPath string) (string, error) {
	if strings.HasSuffix(strings.ToLower(basePath), ".png") {
		return basePath, nil
	}

	// 视为目录：必须存在
	if basePath == "" {
		basePath = "."
	}
	info, err := os.Stat(basePath)
	if err != nil || !info.IsDir() {
		return "", fmt.Errorf("输出目录不存在或不可用: %s", basePath)
	}

	// 提取Excel文件名（不含扩展名）
	excelFileName := filepath.Base(excelPath)
	excelFileName = strings.TrimSuffix(excelFileName, filepath.Ext(excelFileName))

	// 清理文件名中的特殊字符
	excelFileNameSafe := strings.ReplaceAll(excelFileName, "/", "_")
	excelFileNameSafe = strings.ReplaceAll(excelFileNameSafe, "\\", "_")
	excelFileNameSafe = strings.ReplaceAll(excelFileNameSafe, " ", "_")

	sheetNameSafe := strings.ReplaceAll(sheetName, "/", "_")
	sheetNameSafe = strings.ReplaceAll(sheetNameSafe, "\\", "_")
	sheetNameSafe = strings.ReplaceAll(sheetNameSafe, " ", "_")

	// 生成时间戳
	timestamp := time.Now().Format("20060102_150405")

	// 组合文件名
	filename := fmt.Sprintf("%s_%s_%s.png", excelFileNameSafe, sheetNameSafe, timestamp)
	return filepath.Join(basePath, filename), nil
}

// 保存渲染结果
func saveImage(img image.Image, outputPath string) error {
	outFile, err := os.Create(outputPath)
	if err != nil {
		return err
	}
	defer outFile.Close()

	return png.Encode(outFile, img)
}

// 渲染单个工作表
func renderSingleSheet(args *CLIArgs, excel *excelsnapshot.Excel, renderer *excelsnapshot.SheetRenderer, logger *zap.Logger) error {
	// 确定目标工作表
	targetSheet, err := determineTargetSheet(args, excel)
	if err != nil {
		return err
	}

	logger.Info("开始渲染工作表", zap.String("sheet", targetSheet))

	// 获取工作表
	sheet, err := excel.GetSheet(targetSheet)
	if err != nil {
		return err
	}

	// 渲染
	img, err := renderer.RenderSheet(sheet)
	if err != nil {
		return err
	}

	// 生成输出路径并保存
	outputPath, err := generateOutputPath(args.outPath, targetSheet, args.inPath)
	if err != nil {
		return err
	}
	if err := saveImage(img, outputPath); err != nil {
		return err
	}

	logger.Info("渲染完成", zap.String("output", outputPath))
	return nil
}

// 渲染所有工作表
func renderAllSheets(args *CLIArgs, excel *excelsnapshot.Excel, renderer *excelsnapshot.SheetRenderer, logger *zap.Logger) error {
	logger.Info("开始渲染所有工作表")

	for _, sheet := range excel.Sheets() {
		logger.Info("正在渲染工作表", zap.String("sheet", sheet.Name))

		img, err := renderer.RenderSheet(sheet)
		if err != nil {
			return fmt.Errorf("渲染工作表 %s 失败: %w", sheet.Name, err)
		}

		outputPath, err := generateOutputPath(args.outPath, sheet.Name, args.inPath)
		if err != nil {
			return fmt.Errorf("输出目录校验失败: %w", err)
		}
		if err := saveImage(img, outputPath); err != nil {
			return fmt.Errorf("保存工作表 %s 失败: %w", sheet.Name, err)
		}

		logger.Info("工作表渲染完成", zap.String("sheet", sheet.Name), zap.String("output", outputPath))
	}

	logger.Info("所有工作表渲染完成")
	return nil
}

func main() {
	defer func() {
		if r := recover(); r != nil {
			// 捕获未知 panic，打印并以非零退出
			fmt.Fprintf(os.Stderr, "程序异常退出: %v\n", r)
			debug.PrintStack()
			os.Exit(1)
		}
	}()

	// 解析命令行参数
	args := parseArgs()

	// 初始化日志
	logger, loggerSync, err := setupLogger(args.verbose)
	if err != nil {
		fmt.Fprintf(os.Stderr, "初始化日志失败: %v\n", err)
		os.Exit(1)
	}
	defer loggerSync()

	// 初始化渲染器
	renderer := excelsnapshot.NewSheetRenderer(logger)

	// 加载Excel文件
	logger.Info("加载 Excel 文件", zap.String("path", args.inPath))
	excel, err := excelsnapshot.NewExcel(args.inPath, logger)
	if err != nil {
		fmt.Fprintf(os.Stderr, "加载Excel文件失败: %v\n", err)
		os.Exit(1)
	}

	// 根据参数决定渲染模式
	if args.all {
		// 渲染所有工作表
		if err := excel.ParseAll(); err != nil {
			fmt.Fprintf(os.Stderr, "解析Excel失败: %v\n", err)
			os.Exit(1)
		}

		if err := renderAllSheets(args, excel, renderer, logger); err != nil {
			fmt.Fprintf(os.Stderr, "渲染所有工作表失败: %v\n", err)
			os.Exit(1)
		}
	} else {
		// 渲染单个工作表
		targetSheet, err := determineTargetSheet(args, excel)
		if err != nil {
			fmt.Fprintf(os.Stderr, "参数错误: %v\n", err)
			os.Exit(1)
		}

		if err := excel.Parse(targetSheet); err != nil {
			fmt.Fprintf(os.Stderr, "解析工作表失败: %v\n", err)
			os.Exit(1)
		}

		if err := renderSingleSheet(args, excel, renderer, logger); err != nil {
			fmt.Fprintf(os.Stderr, "渲染失败: %v\n", err)
			os.Exit(1)
		}
	}
}
