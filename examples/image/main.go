// 图片示例 - 演示如何在PPT中添加图片
package main

import (
	"fmt"

	"github.com/CoolBanHub/genppt"
)

func main() {
	pres := genppt.New()
	pres.SetTitle("图片演示")
	pres.SetAuthor("GenPPT")

	// 第一张幻灯片 - 从文件添加图片
	slide1 := pres.AddSlide()
	slide1.AddText("图片嵌入演示", genppt.TextOptions{
		X:         0.5,
		Y:         0.3,
		Width:     9.0,
		Height:    0.6,
		FontSize:  28,
		FontColor: "#1E3A5F",
		Bold:      true,
	})

	slide1.AddText("从本地文件加载图片", genppt.TextOptions{
		X:         0.5,
		Y:         1.0,
		Width:     9.0,
		Height:    0.4,
		FontSize:  18,
		FontColor: "#666666",
	})

	// 使用本地图片文件
	slide1.AddImage(genppt.ImageOptions{
		Path:    "image.png", // 使用当前目录的图片
		X:       1.5,
		Y:       1.6,
		Width:   7.0,
		Height:  4.0,
		AltText: "示例图片",
	})

	// 第二张幻灯片 - 多个图片布局
	slide2 := pres.AddSlide()
	slide2.AddText("多图片布局", genppt.TextOptions{
		X:         0.5,
		Y:         0.3,
		Width:     9.0,
		Height:    0.6,
		FontSize:  28,
		FontColor: "#1E3A5F",
		Bold:      true,
	})

	// 左侧图片
	slide2.AddImage(genppt.ImageOptions{
		Path:   "image.png",
		X:      0.5,
		Y:      1.2,
		Width:  4.3,
		Height: 3.0,
	})

	// 右侧图片（带旋转）
	slide2.AddImage(genppt.ImageOptions{
		Path:   "image.png",
		X:      5.2,
		Y:      1.2,
		Width:  4.3,
		Height: 3.0,
		Rotate: 5, // 轻微旋转
	})

	// 底部小图片
	slide2.AddImage(genppt.ImageOptions{
		Path:   "image.png",
		X:      3.0,
		Y:      4.5,
		Width:  4.0,
		Height: 1.5,
	})

	// 第三张幻灯片 - 图片样式
	slide3 := pres.AddSlide()
	slide3.AddText("图片样式选项", genppt.TextOptions{
		X:         0.5,
		Y:         0.3,
		Width:     9.0,
		Height:    0.6,
		FontSize:  28,
		FontColor: "#1E3A5F",
		Bold:      true,
	})

	slide3.AddText("支持旋转、圆角等效果", genppt.TextOptions{
		X:         0.5,
		Y:         1.0,
		Width:     9.0,
		Height:    0.4,
		FontSize:  16,
		FontColor: "#888888",
	})

	// 原始图片
	slide3.AddText("原始", genppt.TextOptions{
		X: 0.8, Y: 1.5, Width: 3, Height: 0.3, FontSize: 14,
	})
	slide3.AddImage(genppt.ImageOptions{
		Path:   "image.png",
		X:      0.5,
		Y:      1.8,
		Width:  3.0,
		Height: 2.0,
	})

	// 圆角图片
	slide3.AddText("圆角", genppt.TextOptions{
		X: 4.0, Y: 1.5, Width: 3, Height: 0.3, FontSize: 14,
	})
	slide3.AddImage(genppt.ImageOptions{
		Path:     "image.png",
		X:        3.5,
		Y:        1.8,
		Width:    3.0,
		Height:   2.0,
		Rounding: 2000,
	})

	// 旋转图片
	slide3.AddText("旋转15°", genppt.TextOptions{
		X: 7.2, Y: 1.5, Width: 3, Height: 0.3, FontSize: 14,
	})
	slide3.AddImage(genppt.ImageOptions{
		Path:   "image.png",
		X:      6.5,
		Y:      1.8,
		Width:  3.0,
		Height: 2.0,
		Rotate: 15,
	})

	err := pres.WriteFile("image_demo.pptx")
	if err != nil {
		fmt.Printf("保存失败: %v\n", err)
		return
	}
	fmt.Println("演示文稿已保存: image_demo.pptx")
	fmt.Printf("共 %d 张幻灯片\n", pres.SlideCount())
}
