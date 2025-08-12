# Simple cross-compilation Makefile
# 输出目录：dist/<os>-<arch>/

BINARY := excel_snapshot
LDFLAGS := -s -w

.PHONY: build clean build-all \
	darwin-amd64 darwin-arm64 \
	linux-amd64 linux-arm64 \
	windows-amd64 windows-arm64

# 本机构建（当前平台）
build:
	@mkdir -p dist
	CGO_ENABLED=0 go build -ldflags "$(LDFLAGS)" -o dist/$(BINARY)
	@echo "Built dist/$(BINARY)"

# macOS
darwin-amd64:
	@mkdir -p dist/darwin-amd64
	GOOS=darwin GOARCH=amd64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/darwin-amd64/$(BINARY)
	@echo "Built dist/darwin-amd64/$(BINARY)"

darwin-arm64:
	@mkdir -p dist/darwin-arm64
	GOOS=darwin GOARCH=arm64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/darwin-arm64/$(BINARY)
	@echo "Built dist/darwin-arm64/$(BINARY)"

# Linux
linux-amd64:
	@mkdir -p dist/linux-amd64
	GOOS=linux GOARCH=amd64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/linux-amd64/$(BINARY)
	@echo "Built dist/linux-amd64/$(BINARY)"

linux-arm64:
	@mkdir -p dist/linux-arm64
	GOOS=linux GOARCH=arm64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/linux-arm64/$(BINARY)
	@echo "Built dist/linux-arm64/$(BINARY)"

# Windows
windows-amd64:
	@mkdir -p dist/windows-amd64
	GOOS=windows GOARCH=amd64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/windows-amd64/$(BINARY).exe
	@echo "Built dist/windows-amd64/$(BINARY).exe"

windows-arm64:
	@mkdir -p dist/windows-arm64
	GOOS=windows GOARCH=arm64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/windows-arm64/$(BINARY).exe
	@echo "Built dist/windows-arm64/$(BINARY).exe"

# 一键全部
build-all: darwin-amd64 darwin-arm64 linux-amd64 linux-arm64 windows-amd64 windows-arm64
	@echo "All targets built into dist/"

# 清理
clean:
	rm -rf dist
	@echo "Cleaned dist/"
