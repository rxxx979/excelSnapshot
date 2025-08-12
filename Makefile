# Simple cross-compilation Makefile
# 输出目录：dist/<os>-<arch>/

BINARY := excel_snapshot
LDFLAGS := -s -w

.PHONY: build clean build-all \
	darwin-amd64 darwin-arm64 \
	linux-amd64 linux-arm64 \
	windows-amd64 windows-arm64

# 本机构建（当前平台）
build: clean
	CGO_ENABLED=0 go build -ldflags "$(LDFLAGS)" -o dist/$(BINARY)
	@echo "Built dist/$(BINARY)"

# macOS
darwin-amd64:
	GOOS=darwin GOARCH=amd64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/$(BINARY)_darwin_amd64
	@echo "Built dist/$(BINARY)_darwin_amd64"

darwin-arm64:
	GOOS=darwin GOARCH=arm64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/$(BINARY)_darwin_arm64
	@echo "Built dist/$(BINARY)_darwin_arm64"

# Linux
linux-amd64:
	GOOS=linux GOARCH=amd64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/$(BINARY)_linux_amd64
	@echo "Built dist/$(BINARY)_linux_amd64"

linux-arm64:
	GOOS=linux GOARCH=arm64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/$(BINARY)_linux_arm64
	@echo "Built dist/$(BINARY)_linux_arm64"

# Windows
windows-amd64:
	GOOS=windows GOARCH=amd64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/$(BINARY)_windows_amd64.exe
	@echo "Built dist/$(BINARY)_windows_amd64.exe"

windows-arm64:
	GOOS=windows GOARCH=arm64 CGO_ENABLED=0 \
		go build -ldflags "$(LDFLAGS)" -o dist/$(BINARY)_windows_arm64.exe
	@echo "Built dist/$(BINARY)_windows_arm64.exe"

# 一键全部
build-all: clean darwin-amd64 darwin-arm64 linux-amd64 linux-arm64 windows-amd64 windows-arm64
	@echo "All targets built into dist/"

# 清理
clean:
	rm -rf dist
	@echo "Cleaned dist/"
