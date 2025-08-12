package main

import (
	"strings"
)

// sanitizeFileComponent 生成文件名
func sanitizeFileComponent(s string) string {
	if s == "" {
		return "sheet"
	}
	invalid := []string{"<", ">", ":", "\"", "|", "?", "*", "\\", "/"}
	result := s
	for _, inv := range invalid {
		result = strings.ReplaceAll(result, inv, "_")
	}
	result = strings.TrimSpace(result)
	if result == "" {
		result = "sheet"
	}
	return result
}
