package excelsnapshot

import (
	"image"
)

// RenderSheet 渲染工作表
func RenderSheet(sheet *Sheet) (image.Image, error) {
	maxX, maxY := getMaxXAndY(sheet)
	img := image.NewRGBA(image.Rect(0, 0, maxX, maxY))
	return img, nil
}

func getMaxXAndY(sheet *Sheet) (int, int) {
	maxX := 0
	maxY := 0
	for _, cell := range sheet.cells {
		if cell.Col > maxX {
			maxX = cell.Col
		}
		if cell.Row > maxY {
			maxY = cell.Row
		}
	}
	return maxX, maxY
}
