package genppt

import (
	"crypto/rand"
	"encoding/hex"
	"html"
	"strings"
	"time"
)

// generateUUID 生成UUID
func generateUUID() string {
	bytes := make([]byte, 16)
	rand.Read(bytes)
	bytes[6] = (bytes[6] & 0x0f) | 0x40 // Version 4
	bytes[8] = (bytes[8] & 0x3f) | 0x80 // Variant
	return hex.EncodeToString(bytes[:4]) + "-" +
		hex.EncodeToString(bytes[4:6]) + "-" +
		hex.EncodeToString(bytes[6:8]) + "-" +
		hex.EncodeToString(bytes[8:10]) + "-" +
		hex.EncodeToString(bytes[10:])
}

// escapeXML 转义XML特殊字符
func escapeXML(s string) string {
	return html.EscapeString(s)
}

// getDefaultFontFace 返回默认字体
func getDefaultFontFace() string {
	return "微软雅黑"
}

// getDefaultFontSize 返回默认字号（磅）
func getDefaultFontSize() float64 {
	return 18
}

// getDefaultColor 返回默认颜色
func getDefaultColor() string {
	return "000000"
}

// formatDateTime 格式化日期时间为ISO8601
func formatDateTime(t time.Time) string {
	return t.Format("2006-01-02T15:04:05Z")
}

// getImageType 根据图片数据判断类型
func getImageType(data []byte) string {
	if len(data) < 8 {
		return ""
	}
	// PNG
	if data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47 {
		return "png"
	}
	// JPEG
	if data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF {
		return "jpeg"
	}
	// GIF
	if data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46 {
		return "gif"
	}
	// BMP
	if data[0] == 0x42 && data[1] == 0x4D {
		return "bmp"
	}
	// WebP
	if len(data) >= 12 && string(data[0:4]) == "RIFF" && string(data[8:12]) == "WEBP" {
		return "webp"
	}
	return ""
}

// getImageMIME 获取图片MIME类型
func getImageMIME(ext string) string {
	ext = strings.ToLower(ext)
	switch ext {
	case "png":
		return "image/png"
	case "jpg", "jpeg":
		return "image/jpeg"
	case "gif":
		return "image/gif"
	case "bmp":
		return "image/bmp"
	case "webp":
		return "image/webp"
	case "svg":
		return "image/svg+xml"
	default:
		return "image/png"
	}
}

// getMediaMIME 获取媒体MIME类型（图片或视频）
func getMediaMIME(ext string) string {
	ext = strings.ToLower(ext)
	// 先检查视频类型
	switch ext {
	case "mp4", "m4v":
		return "video/mp4"
	case "mov":
		return "video/quicktime"
	case "avi":
		return "video/x-msvideo"
	case "wmv":
		return "video/x-ms-wmv"
	case "mpg", "mpeg":
		return "video/mpeg"
	case "webm":
		return "video/webm"
	}
	// 否则返回图片类型
	return getImageMIME(ext)
}

// getExtFromPath 从路径获取扩展名
func getExtFromPath(path string) string {
	idx := strings.LastIndex(path, ".")
	if idx == -1 {
		return ""
	}
	return strings.ToLower(path[idx+1:])
}

// defaultIfZero 如果值为0则返回默认值
func defaultIfZero(val, defaultVal float64) float64 {
	if val == 0 {
		return defaultVal
	}
	return val
}

// defaultIfEmpty 如果字符串为空则返回默认值
func defaultIfEmpty(val, defaultVal string) string {
	if val == "" {
		return defaultVal
	}
	return val
}

// max 返回两个整数中的较大值
func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}

// min 返回两个整数中的较小值
func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
