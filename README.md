# Excel Snapshot

将 Excel 工作表渲染为 PNG 图片，追求 1:1 样式还原与中文字体完美显示。

## 构建
```bash
go mod download
go build -o excel_snapshot
```

## 使用
```bash
# 指定工作表名
./excel_snapshot -in report.xlsx -sheet 财务报表 -out report.png

# 使用工作表索引（从 0 开始）
./excel_snapshot -in report.xlsx -sheet-index 0 -out first.png

# 导出所有非空工作表（输出多个 PNG）
./excel_snapshot -in report.xlsx -all-sheets -out shots.png
```

参数：
- -in：输入 Excel（.xlsx）
- -sheet：工作表名（留空则用索引）
- -sheet-index：工作表索引（从 0 开始）
- -out：输出 PNG 路径（自动补 .png）
- -all-sheets：导出所有非空工作表
- -force-raw：使用原始未格式化值（跳过公式/格式化）

## 字体
- 字体通过 Go embed 内置自 `static/` 目录（当前包含思源宋体 SC）。
- 运行时无需额外字体文件；如需更换，将字体放入 `static/` 后重新构建。

## 注意
- 输出为 PNG，按 Excel 像素级 1:1 排版。
- 超大工作表会占用更多内存与时间，请根据机器资源评估。