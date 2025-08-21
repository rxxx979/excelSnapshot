package main

import (
	"flag"
	"runtime/debug"

	excelsnapshot "github.com/rxxx/excelSnapshot"
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
)

func main() {
	var (
		inPath  string
		outPath string
		sheet   string
		index   int
		all     bool
		verbose bool
	)

	flag.StringVar(&inPath, "i", "", "输入的 Excel 文件路径 (.xlsx)")
	flag.StringVar(&outPath, "o", ".", "输出目录或文件路径（当渲染单个 sheet 时可指定 .png 文件）")
	flag.StringVar(&sheet, "sheet", "", "要渲染的工作表名称（优先于 index）")
	flag.IntVar(&index, "index", -1, "要渲染的工作表索引（0-based）")
	flag.BoolVar(&all, "all", false, "是否渲染所有工作表")
	flag.BoolVar(&verbose, "v", false, "启用调试日志（开发模式）")
	flag.Parse()

	// 判断是否开发环境
	var isDev bool
	_, isBuildInfo := debug.ReadBuildInfo()
	if isBuildInfo {
		isDev = true
	} else {
		isDev = false
	}

	// 判断是否启用调试日志
	var level zapcore.Level
	if verbose {
		level = zap.DebugLevel
	} else {
		level = zap.InfoLevel
	}

	excelSnapshotLogger, excelSnapshotSync, err := excelsnapshot.SetupLogger("excel_snapshot", level, isDev)
	if err != nil {
		panic(err)
	}
	excelSnapshotLogger.Info("加载 Excel 文件", zap.String("path", inPath))
	excel, err := excelsnapshot.NewExcel(inPath, excelSnapshotLogger)
	if err != nil {
		panic(err)
	}

	if err := excel.Parse(); err != nil {
		panic(err)
	}

	// 显示工作表列表
	sheetList, err := excel.GetSheetList()
	if err != nil {
		panic(err)
	}
	excelSnapshotLogger.Info("可用的工作表", zap.Strings("sheets", sheetList))

	if sheet != "" {
		if err := excel.RenderSheet(sheet); err != nil {
			panic(err)
		}
	} else if len(sheetList) > 0 {
		// 如果没有指定工作表，渲染第一个
		if err := excel.RenderSheet(sheetList[0]); err != nil {
			panic(err)
		}
	}

	defer excelSnapshotSync()

}
