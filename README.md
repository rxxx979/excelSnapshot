# Excel Snapshot

将 Excel 工作表渲染为 PNG 图片

## 构建
```bash
go mod download
go build -o excel_snapshot
```

## 使用
```bash
# 指定工作表名渲染
./excel_snapshot -i report.xlsx -sheet 财务报表 -o ./report.png

# 按索引渲染（从 0 开始）
./excel_snapshot -i report.xlsx -index 0 -o ./first.png

# 渲染所有工作表（输出到当前目录，自动生成文件名）
./excel_snapshot -i report.xlsx -all -o .
```

参数：
- -i string：输入 Excel 文件路径（.xlsx）
- -o string：输出路径
  - 当渲染单个工作表且以 .png 结尾时，作为目标文件
  - 其他情况视为目录，程序自动生成文件名（含时间戳）
- -sheet string：要渲染的工作表名称（优先于 index）
- -index int：要渲染的工作表索引（0-based）
- -all：渲染所有工作表
- -v：启用调试日志（开发模式）

## 字体
- 字体通过 Go embed 内置自 `fonts/` 目录（当前包含思源宋体 SC Regular/Bold）。
- 运行时无需额外字体文件；如需更换字体，将 OTF 放到 `fonts/` 并覆盖同名文件后重新构建。

## 注意
- 输出为 PNG，尽量按 Excel 像素级 1:1 排版。
- 大型工作表会占用较多时间与内存，建议：
  - 仅渲染需要的工作表（使用 -sheet 或 -index）
  - 非调试场景关闭 -v，减少日志开销
  - 输出路径为目录时，程序会自动生成安全文件名与时间戳，避免覆盖