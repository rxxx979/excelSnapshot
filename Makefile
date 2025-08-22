VERSION := $(shell git describe --tags --always)
BINARY := excel_snapshot
LDFLAGS := -s -w
GOOS := $(shell go env GOOS)
GOARCH := $(shell go env GOARCH)

.PHONY: build clean build-all \
	darwin-amd64 darwin-arm64 \
	linux-amd64 linux-arm64 \
	windows-amd64 windows-arm64

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

# 清理
clean:
	rm -rf dist
	@echo "Cleaned dist/"
