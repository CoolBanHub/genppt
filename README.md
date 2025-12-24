# GenPPT - Go语言PowerPoint生成库

纯Go语言实现的PowerPoint演示文稿生成库，支持创建符合OOXML (ECMA-376)标准的`.pptx`文件。

## 特性

- ✅ **纯Go实现** - 无第三方依赖，仅使用标准库
- ✅ **文本支持** - 多种字体、颜色、对齐方式
- ✅ **形状支持** - 矩形、圆形、箭头等多种形状
- ✅ **表格支持** - 完整的表格功能，支持合并单元格
- ✅ **图片支持** - PNG、JPEG、GIF等格式
- ✅ **图表支持** - 柱状图、折线图、饼图、环形图、面积图
- ✅ **视频支持** - MP4、MOV、AVI等格式
- ✅ **音频支持** - MP3、WAV、M4A等格式，支持背景音乐
- ✅ **Markdown支持** - 从Markdown直接生成PPT
- ✅ **多种导出** - 文件、字节数组、io.Writer

## 安装

```bash
go get github.com/CoolBanHub/genppt
```

## 快速开始

```go
package main

import (
	"github.com/CoolBanHub/genppt"
)

func main() {
	// 1. 创建演示文稿
	pres := genppt.New()
	pres.SetTitle("我的演示文稿")

	// 2. 添加幻灯片
	slide := pres.AddSlide()

	// 3. 添加文本
	slide.AddText("Hello World!", genppt.TextOptions{
		X:         1.0,
		Y:         1.0,
		Width:     8.0,
		Height:    1.0,
		FontSize:  24,
		FontColor: "#363636",
		Bold:      true,
		Align:     genppt.AlignCenter,
	})

	// 4. 保存文件
	pres.WriteFile("output.pptx")
}

```

## API 文档

### 演示文稿

```go
// 创建演示文稿
pres := genppt.New()

// 设置属性
pres.SetTitle("标题")
pres.SetAuthor("作者")
pres.SetSubject("主题")
pres.SetCompany("公司")

// 设置幻灯片尺寸
pres.SetSlideSize16x9()    // 16:9 宽屏（默认）
pres.SetSlideSize4x3()     // 4:3 标准
pres.SetSlideSize(10, 7.5) // 自定义尺寸（英寸）

// 添加幻灯片
slide := pres.AddSlide()

// 导出
pres.WriteFile("output.pptx")    // 保存到文件
data, _ := pres.ToBytes() // 转换为字节数组
pres.Write(writer)        // 写入io.Writer
```

### 文本

```go
slide.AddText("文本内容", genppt.TextOptions{
X:         1.0, // X坐标（英寸）
Y:         1.0, // Y坐标（英寸）
Width:     8.0, // 宽度（英寸）
Height:    1.0, // 高度（英寸）
FontFace:  "微软雅黑",  // 字体
FontSize:  24,          // 字号（磅）
FontColor: "#FF0000",   // 颜色
Bold:      true,  // 粗体
Italic:    false, // 斜体
Underline: false,              // 下划线
Align:     genppt.AlignCenter, // 水平对齐
VAlign:    genppt.VAlignMiddle, // 垂直对齐
Rotate:    45, // 旋转角度
})
```

### 形状

```go
// 基本形状
slide.AddShape(genppt.ShapeRect, genppt.ShapeOptions{
X:           1.0,
Y:           1.0,
Width:       2.0,
Height:      1.0,
Fill:        "#4472C4",
LineColor:   "#2F5496",
LineWidth:   2,
Shadow:      true,
Transparency: 20, // 透明度0-100
})

// 带文本的形状
slide.AddShapeWithText(genppt.ShapeEllipse, "文本", genppt.ShapeOptions{...})

// 支持的形状类型
genppt.ShapeRect      // 矩形
genppt.ShapeRoundRect // 圆角矩形
genppt.ShapeEllipse  // 椭圆
genppt.ShapeTriangle // 三角形
genppt.ShapeDiamond    // 菱形
genppt.ShapeArrowRight // 右箭头
genppt.ShapeArrowLeft  // 左箭头
genppt.ShapeArrowUp   // 上箭头
genppt.ShapeArrowDown // 下箭头
genppt.ShapeStar5 // 五角星
genppt.ShapeHeart // 心形
```

### 表格

```go
slide.AddTable([][]genppt.TableCell{
{{Text: "标题1", Bold: true}, {Text: "标题2", Bold: true}},
{{Text: "数据1"}, {Text: "数据2"}},
{{Text: "数据3", ColSpan: 2}}, // 合并列
}, genppt.TableOptions{
X:            1.0,
Y:            1.0,
Width:        8.0,
FontSize:     14,
FirstRowBold: true,
FirstRowFill: "#4472C4",
Border: genppt.Border{
Color: "#CCCCCC",
Width: 1.0,
},
})
```

### 图片

```go
// 从文件添加
slide.AddImage(genppt.ImageOptions{
Path:   "/path/to/image.png",
X:      1.0,
Y:      1.0,
Width:  4.0,
Height: 3.0,
})

// 从字节数据添加
slide.AddImage(genppt.ImageOptions{
Data:   imageBytes,
X:      1.0,
Y:      1.0,
Width:  4.0,
Height: 3.0,
})
```

### 背景

```go
slide.SetBackground(genppt.BackgroundOptions{
Color: "#1E3A5F", // 纯色背景
})
```

### 音频

```go
// 从文件添加音频
slide.AddAudio(genppt.AudioOptions{
Path:     "/path/to/audio.mp3",
X:        1.0,
Y:        1.0,
Width:    0.5,  // 音频图标大小
Height:   0.5,
AutoPlay: true, // 自动播放
Loop:     true, // 循环播放
})

// 背景音乐（隐藏图标）
slide.AddAudio(genppt.AudioOptions{
Path:     "/path/to/bgm.mp3",
X:        0,
Y:        0,
Hidden:   true, // 隐藏音频图标
AutoPlay: true,
Loop:     true,
})

// 从字节数据添加
slide.AddAudio(genppt.AudioOptions{
Data:     audioBytes,
X:        1.0,
Y:        1.0,
AutoPlay: false,
})
```

**支持的音频格式**: MP3, WAV, WMA, M4A, AAC, OGG, FLAC

## Markdown 支持

GenPPT 支持从 Markdown 直接生成 PPT！

### 基本用法

```go
markdown := `# 第一张幻灯片

这是内容。

# 第二张幻灯片

- 列表项1
- 列表项2
`

pres := genppt.FromMarkdown(markdown)
pres.WriteFile("output.pptx")
```

### Markdown 格式规则

| 格式       | 效果        |
|----------|-----------|
| `# 标题`   | 创建新幻灯片    |
| `## 子标题` | 幻灯片标题     |
| `- 列表`   | 无序列表      |
| `1. 列表`  | 有序列表      |
| ` ``` `  | 代码块       |
| `---`    | 分隔符（新幻灯片） |

### 自定义样式

```go
opts := genppt.DefaultMarkdownOptions()
opts.TitleFontSize = 48
opts.HeadingColor = "#FFFFFF"
opts.BodyColor = "#E0E0E0"
opts.SlideBackground = "#1A1A2E"
opts.CodeBackground = "#16213E"

pres := genppt.FromMarkdownWithOptions(markdown, opts)
```

### 从文件读取

```go
pres, err := genppt.FromMarkdownFile("presentation.md")
if err != nil {
log.Fatal(err)
}
pres.WriteFile("output.pptx")
```

## 运行示例

```bash
cd examples/basic
go run main.go
```

## 运行测试

```bash
go test -v ./...
```

## 许可证

MIT License
