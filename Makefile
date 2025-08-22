VERSION := $(shell git describe --tags --always)
BINARY := excel_snapshot
LDFLAGS := -s -w
GOOS := $(shell go env GOOS)
GOARCH := $(shell go env GOARCH)

.PHONY: build clean build-all \
	darwin-amd64 darwin-arm64 \
	linux-amd64 linux-arm64 \
	windows-amd64 windows-arm64 \
	bench bench-full bench-clean bench-render profile-cpu profile-mem

# 本机构建（当前平台）
build: clean
	CGO_ENABLED=0 go build -ldflags "$(LDFLAGS) -X main.version=$(VERSION)" -o dist/$(BINARY)-v$(VERSION) ./cmd/excelsnapshot
	@echo "Built dist/$(BINARY)-v$(VERSION)"

# macOS
darwin-amd64:
	GOOS=darwin GOARCH=amd64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS) -X main.version=$(VERSION)" -o dist/$(BINARY)-v$(VERSION)-darwin-amd64 ./cmd/excelsnapshot
	@echo "Built dist/$(BINARY)-v$(VERSION)-darwin-amd64"

darwin-arm64:
	GOOS=darwin GOARCH=arm64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS) -X main.version=$(VERSION)" -o dist/$(BINARY)-v$(VERSION)-darwin-arm64 ./cmd/excelsnapshot
	@echo "Built dist/$(BINARY)-v$(VERSION)-darwin-arm64"

# Linux
linux-amd64:
	GOOS=linux GOARCH=amd64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS) -X main.version=$(VERSION)" -o dist/$(BINARY)-v$(VERSION)-linux-amd64 ./cmd/excelsnapshot
	@echo "Built dist/$(BINARY)-v$(VERSION)-linux-amd64"

linux-arm64:
	GOOS=linux GOARCH=arm64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS) -X main.version=$(VERSION)" -o dist/$(BINARY)-v$(VERSION)-linux-arm64 ./cmd/excelsnapshot
	@echo "Built dist/$(BINARY)-v$(VERSION)-linux-arm64"

# Windows
windows-amd64:
	GOOS=windows GOARCH=amd64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS) -X main.version=$(VERSION)" -o dist/$(BINARY)-v$(VERSION)-windows-amd64.exe ./cmd/excelsnapshot
	@echo "Built dist/$(BINARY)-v$(VERSION)-windows-amd64.exe"

windows-arm64:
	GOOS=windows GOARCH=arm64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS) -X main.version=$(VERSION)" -o dist/$(BINARY)-v$(VERSION)-windows-arm64.exe ./cmd/excelsnapshot
	@echo "Built dist/$(BINARY)-v$(VERSION)-windows-arm64.exe"

# 一键全部
build-all: clean darwin-amd64 darwin-arm64 linux-amd64 linux-arm64 windows-amd64 windows-arm64
	@echo "All targets built into dist/"

# 性能测试
bench:
	go test -bench=. -benchmem
	@echo "Benchmark completed"

bench-full:
	go test -bench=. -benchmem -count=5 -benchtime=10s
	@echo "Full benchmark completed"

# 清洁的基准测试（减少日志输出）
bench-clean:
	@echo "Running clean benchmark tests..."
	@go test -bench="BenchmarkSheetRenderer_(Small|Medium|Large)Sheet" -benchmem -run=^$$ | grep -E "(Benchmark|goos|goarch|cpu|PASS|ok)"
	@echo "Clean benchmark completed"

# 渲染性能基准测试（三种尺寸对比）
bench-render:
	@echo "=== Excel渲染性能基准测试 ==="
	@echo "小维度(8x10): 80单元格，随机数据+颜色+合并单元格"
	@go test -bench="BenchmarkSheetRenderer_SmallSheet" -benchmem -run=^$$ 2>/dev/null | grep "Benchmark" | head -1
	@echo "中维度(16x30): 480单元格，随机数据+颜色+合并单元格"
	@go test -bench="BenchmarkSheetRenderer_MediumSheet" -benchmem -run=^$$ 2>/dev/null | grep "Benchmark" | head -1
	@echo "大维度(25x100): 2500单元格，随机数据+颜色+合并单元格"
	@go test -bench="BenchmarkSheetRenderer_LargeSheet" -benchmem -run=^$$ 2>/dev/null | grep "Benchmark" | head -1
	@echo "=== 基准测试完成 ==="

# 性能分析
profile-cpu:
	go test -bench=BenchmarkSheetRenderer_RenderSheet -cpuprofile=cpu.prof
	go tool pprof -http=:8080 cpu.prof

profile-mem:
	go test -bench=BenchmarkSheetRenderer_RenderSheet -memprofile=mem.prof
	go tool pprof -http=:8080 mem.prof

# 清理
clean:
	rm -rf dist *.prof
	@echo "Cleaned dist/ and profiles"
