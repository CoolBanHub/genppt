package genppt

import (
	"fmt"
	"strings"
)

// EMU (English Metric Units) 常量
// 1英寸 = 914400 EMU
const (
	EMUPerInch  = 914400
	EMUPerPoint = 12700
	EMUPerCM    = 360000
)

// 默认幻灯片尺寸 (标准16:9)
const (
	DefaultSlideWidth  = 9144000 // 10英寸 = 9144000 EMU
	DefaultSlideHeight = 5143500 // 5.625英寸 = 5143500 EMU (16:9)
)

// TextOptions 文本选项
type TextOptions struct {
	X           float64       // X坐标（英寸）
	Y           float64       // Y坐标（英寸）
	Width       float64       // 宽度（英寸）
	Height      float64       // 高度（英寸）
	FontFace    string        // 字体名称
	FontSize    float64       // 字号（磅）
	FontColor   string        // 字体颜色（十六进制，如"#FF0000"或"FF0000"）
	Bold        bool          // 是否粗体
	Italic      bool          // 是否斜体
	Underline   bool          // 是否下划线
	Align       Align         // 水平对齐
	VAlign      VerticalAlign // 垂直对齐
	LineSpacing float64       // 行间距（倍数）
	Rotate      float64       // 旋转角度（度）
	Margin      float64       // 内边距（英寸）
	Fill        string        // 文本框背景色（十六进制），为空则无填充
}

// ShapeOptions 形状选项
type ShapeOptions struct {
	X            float64     // X坐标（英寸）
	Y            float64     // Y坐标（英寸）
	Width        float64     // 宽度（英寸）
	Height       float64     // 高度（英寸）
	Fill         string      // 填充颜色（十六进制）
	LineColor    string      // 边框颜色（十六进制）
	LineWidth    float64     // 边框宽度（磅）
	LineStyle    BorderStyle // 边框样式
	Rotate       float64     // 旋转角度（度）
	Transparency float64     // 透明度（0-100）
	Shadow       bool        // 是否有阴影
}

// TableOptions 表格选项
type TableOptions struct {
	X            float64   // X坐标（英寸）
	Y            float64   // Y坐标（英寸）
	Width        float64   // 总宽度（英寸）
	RowHeights   []float64 // 每行高度（英寸），为空则自动
	ColWidths    []float64 // 每列宽度（英寸），为空则平均分配
	FontFace     string    // 默认字体
	FontSize     float64   // 默认字号
	FontColor    string    // 默认字体颜色
	Fill         string    // 默认单元格背景色
	Border       Border    // 边框设置
	FirstRowBold bool      // 首行是否加粗
	FirstRowFill string    // 首行背景色
}

// Border 边框配置
type Border struct {
	Color string      // 颜色
	Width float64     // 宽度（磅）
	Style BorderStyle // 样式
}

// TableCell 表格单元格
type TableCell struct {
	Text      string        // 文本内容
	FontFace  string        // 字体（覆盖表格默认）
	FontSize  float64       // 字号（覆盖表格默认）
	FontColor string        // 字体颜色
	Bold      bool          // 是否粗体
	Italic    bool          // 是否斜体
	Fill      string        // 背景色
	Align     Align         // 水平对齐
	VAlign    VerticalAlign // 垂直对齐
	ColSpan   int           // 列合并数
	RowSpan   int           // 行合并数
}

// ImageOptions 图片选项
type ImageOptions struct {
	X               float64 // X坐标（英寸）
	Y               float64 // Y坐标（英寸）
	Width           float64 // 宽度（英寸）
	Height          float64 // 高度（英寸）
	Path            string  // 本地文件路径
	Data            []byte  // 图片数据（与Path二选一）
	AltText         string  // 替代文本
	Rotate          float64 // 旋转角度（度）
	Rounding        float64 // 圆角半径（英寸），0为直角
	CodeBackground  string  // 代码背景色
	SlideBackground string  // 幻灯片背景色
	ImageRounding   float64 // 图片圆角（英寸），默认0
}

// BackgroundOptions 背景选项
type BackgroundOptions struct {
	Color string // 纯色背景（十六进制）
	Image string // 背景图片路径
	Data  []byte // 背景图片数据
}

// slideObject 幻灯片对象接口
type slideObject interface {
	getType() string
}

// textObject 文本对象
type textObject struct {
	text    string
	options TextOptions
}

func (t *textObject) getType() string { return "text" }

// shapeObject 形状对象
type shapeObject struct {
	shapeType ShapeType
	options   ShapeOptions
	text      string // 形状内文本（可选）
}

func (s *shapeObject) getType() string { return "shape" }

// tableObject 表格对象
type tableObject struct {
	rows    [][]TableCell
	options TableOptions
}

func (t *tableObject) getType() string { return "table" }

// imageObject 图片对象
type imageObject struct {
	options  ImageOptions
	rID      string // 关系ID
	mediaExt string // 媒体文件扩展名
}

func (i *imageObject) getType() string { return "image" }

// Slide 幻灯片结构
type Slide struct {
	presentation *Presentation
	layout       SlideLayout
	objects      []slideObject
	background   *BackgroundOptions
	notes        string
	number       int // 幻灯片序号
}

// Presentation 演示文稿结构
type Presentation struct {
	title       string
	author      string
	subject     string
	company     string
	revision    string
	layout      SlideLayout
	slides      []*Slide
	slideWidth  int64 // EMU
	slideHeight int64 // EMU
	mediaFiles  []mediaFile
}

// mediaFile 媒体文件
type mediaFile struct {
	path string // 在ZIP中的路径
	data []byte // 文件数据
	ext  string // 扩展名
	rID  string // 关系ID
}

// InchToEMU 将英寸转换为EMU
func InchToEMU(inches float64) int64 {
	return int64(inches * EMUPerInch)
}

// PointToEMU 将磅转换为EMU
func PointToEMU(points float64) int64 {
	return int64(points * EMUPerPoint)
}

// CMToEMU 将厘米转换为EMU
func CMToEMU(cm float64) int64 {
	return int64(cm * EMUPerCM)
}

// ParseColor 解析颜色字符串，处理#前缀和3位简写
func ParseColor(color string) string {
	if color == "" {
		return "000000"
	}

	// 处理颜色名
	switch strings.ToLower(color) {
	case "red":
		return "FF0000"
	case "green":
		return "00FF00"
	case "blue":
		return "0000FF"
	case "black":
		return "000000"
	case "white":
		return "FFFFFF"
	case "yellow":
		return "FFFF00"
	case "orange":
		return "FFA500"
	case "purple":
		return "800080"
	case "gray", "grey":
		return "808080"
	}

	c := color
	if strings.HasPrefix(c, "#") {
		c = c[1:]
	}

	// 处理 3 位 HEX (e.g. "F00" -> "FF0000")
	if len(c) == 3 {
		r := string(c[0])
		g := string(c[1])
		b := string(c[2])
		return r + r + g + g + b + b
	}

	return c
}

// ValidateColor 验证颜色格式
func ValidateColor(color string) error {
	c := ParseColor(color)
	if len(c) != 6 {
		return fmt.Errorf("无效的颜色格式: %s （应为6位十六进制）", color)
	}
	for _, ch := range c {
		if !((ch >= '0' && ch <= '9') || (ch >= 'a' && ch <= 'f') || (ch >= 'A' && ch <= 'F')) {
			return fmt.Errorf("无效的颜色字符: %c", ch)
		}
	}
	return nil
}
