// 基础使用示例
package main

import (
	"fmt"

	"github.com/CoolBanHub/genppt"
)

func main() {
	// 1. 创建演示文稿
	pres := genppt.New()
	pres.SetTitle("我的演示文稿")
	pres.SetAuthor("张三")
	pres.SetCompany("示例公司")

	// 2. 添加第一张幻灯片 - 标题页
	slide1 := pres.AddSlide()
	slide1.SetBackground(genppt.BackgroundOptions{
		Color: "#1E3A5F",
	})

	// 添加大标题
	slide1.AddText("欢迎使用 GenPPT", genppt.TextOptions{
		X:         1.0,
		Y:         2.0,
		Width:     8.0,
		Height:    1.5,
		FontSize:  44,
		FontColor: "#FFFFFF",
		Bold:      true,
		Align:     genppt.AlignCenter,
		VAlign:    genppt.VAlignMiddle,
	})

	// 添加副标题
	slide1.AddText("纯 Go 语言 PowerPoint 生成库", genppt.TextOptions{
		X:         1.0,
		Y:         3.5,
		Width:     8.0,
		Height:    0.5,
		FontSize:  24,
		FontColor: "#A0C4E8",
		Align:     genppt.AlignCenter,
	})

	// 3. 添加第二张幻灯片 - 内容页
	slide2 := pres.AddSlide()

	// 添加标题
	slide2.AddText("功能特性", genppt.TextOptions{
		X:         0.5,
		Y:         0.5,
		Width:     9.0,
		Height:    0.8,
		FontSize:  32,
		FontColor: "#1E3A5F",
		Bold:      true,
	})

	// 添加表格
	slide2.AddTable([][]genppt.TableCell{
		{{Text: "功能", Bold: true}, {Text: "描述", Bold: true}, {Text: "状态", Bold: true}},
		{{Text: "文本框"}, {Text: "支持多种字体、颜色、对齐方式"}, {Text: "✅"}},
		{{Text: "形状"}, {Text: "矩形、圆形、箭头等多种形状"}, {Text: "✅"}},
		{{Text: "表格"}, {Text: "完整的表格支持，包括合并单元格"}, {Text: "✅"}},
		{{Text: "图片"}, {Text: "支持PNG、JPEG、GIF等格式"}, {Text: "✅"}},
	}, genppt.TableOptions{
		X:            0.5,
		Y:            1.5,
		Width:        9.0,
		FontSize:     14,
		FirstRowBold: true,
		FirstRowFill: "#4472C4",
		Border: genppt.Border{
			Color: "#CCCCCC",
			Width: 1.0,
		},
	})

	// 4. 添加第三张幻灯片 - 形状演示
	slide3 := pres.AddSlide()

	slide3.AddText("形状演示", genppt.TextOptions{
		X:         0.5,
		Y:         0.5,
		Width:     9.0,
		Height:    0.8,
		FontSize:  32,
		FontColor: "#1E3A5F",
		Bold:      true,
	})

	// 矩形
	slide3.AddShapeWithText(genppt.ShapeRect, "矩形", genppt.ShapeOptions{
		X:         1.0,
		Y:         1.8,
		Width:     2.0,
		Height:    1.2,
		Fill:      "#4472C4",
		LineColor: "#2F5496",
		LineWidth: 2,
	})

	// 圆角矩形
	slide3.AddShapeWithText(genppt.ShapeRoundRect, "圆角", genppt.ShapeOptions{
		X:         4.0,
		Y:         1.8,
		Width:     2.0,
		Height:    1.2,
		Fill:      "#ED7D31",
		LineColor: "#C55A11",
		LineWidth: 2,
	})

	// 椭圆
	slide3.AddShapeWithText(genppt.ShapeEllipse, "椭圆", genppt.ShapeOptions{
		X:         7.0,
		Y:         1.8,
		Width:     2.0,
		Height:    1.2,
		Fill:      "#70AD47",
		LineColor: "#507E32",
		LineWidth: 2,
	})

	// 箭头
	slide3.AddShape(genppt.ShapeArrowRight, genppt.ShapeOptions{
		X:         2.5,
		Y:         3.5,
		Width:     5.0,
		Height:    0.8,
		Fill:      "#FFC000",
		LineColor: "#BF9000",
		Shadow:    true,
	})

	// 5. 保存文件
	err := pres.WriteFile("demo.pptx")
	if err != nil {
		fmt.Printf("保存失败: %v\n", err)
		return
	}
	fmt.Println("演示文稿已保存: demo.pptx")
	fmt.Printf("共 %d 张幻灯片\n", pres.SlideCount())
}
