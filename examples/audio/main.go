package main

import (
	"fmt"

	"github.com/CoolBanHub/genppt"
)

func main() {
	// 创建演示文稿
	pres := genppt.New()
	pres.SetTitle("音频示例")
	pres.SetAuthor("GenPPT")

	// 第一张幻灯片 - 普通音频
	slide1 := pres.AddSlide()
	slide1.SetBackground(genppt.BackgroundOptions{
		Color: "#1A1A2E",
	})

	slide1.AddText("音频演示", genppt.TextOptions{
		X:         1.0,
		Y:         0.5,
		Width:     8.0,
		Height:    1.0,
		FontSize:  36,
		FontColor: "#FFFFFF",
		Bold:      true,
		Align:     genppt.AlignCenter,
	})

	slide1.AddText("点击音频图标播放", genppt.TextOptions{
		X:         1.0,
		Y:         2.0,
		Width:     8.0,
		Height:    0.5,
		FontSize:  18,
		FontColor: "#CCCCCC",
		Align:     genppt.AlignCenter,
	})

	// 添加可见的音频图标（如果有音频文件的话）
	// slide1.AddAudio(genppt.AudioOptions{
	// 	Path:     "background.mp3",
	// 	X:        4.5,
	// 	Y:        3.0,
	// 	Width:    1.0,
	// 	Height:   1.0,
	// 	AutoPlay: false,
	// 	Loop:     false,
	// })

	// 第二张幻灯片 - 背景音乐示例说明
	slide2 := pres.AddSlide()
	slide2.SetBackground(genppt.BackgroundOptions{
		Color: "#16213E",
	})

	slide2.AddText("背景音乐", genppt.TextOptions{
		X:         1.0,
		Y:         0.5,
		Width:     8.0,
		Height:    1.0,
		FontSize:  36,
		FontColor: "#FFFFFF",
		Bold:      true,
		Align:     genppt.AlignCenter,
	})

	slide2.AddText("使用 Hidden: true 可以隐藏音频图标", genppt.TextOptions{
		X:         1.0,
		Y:         2.0,
		Width:     8.0,
		Height:    0.5,
		FontSize:  18,
		FontColor: "#CCCCCC",
		Align:     genppt.AlignCenter,
	})

	slide2.AddText("适合用于背景音乐", genppt.TextOptions{
		X:         1.0,
		Y:         2.5,
		Width:     8.0,
		Height:    0.5,
		FontSize:  18,
		FontColor: "#CCCCCC",
		Align:     genppt.AlignCenter,
	})

	// 隐藏的背景音乐（如果有音频文件的话）
	// slide2.AddAudio(genppt.AudioOptions{
	// 	Path:     "bgm.mp3",
	// 	X:        0,
	// 	Y:        0,
	// 	Hidden:   true,
	// 	AutoPlay: true,
	// 	Loop:     true,
	// })

	// 第三张幻灯片 - 支持的格式
	slide3 := pres.AddSlide()
	slide3.SetBackground(genppt.BackgroundOptions{
		Color: "#0F3460",
	})

	slide3.AddText("支持的音频格式", genppt.TextOptions{
		X:         1.0,
		Y:         0.5,
		Width:     8.0,
		Height:    1.0,
		FontSize:  36,
		FontColor: "#FFFFFF",
		Bold:      true,
		Align:     genppt.AlignCenter,
	})

	formats := []string{
		"• MP3 - 最常用的音频格式",
		"• WAV - 无损音频格式",
		"• M4A/AAC - Apple音频格式",
		"• WMA - Windows媒体音频",
		"• OGG - 开源音频格式",
		"• FLAC - 无损压缩格式",
	}

	yPos := 1.8
	for _, format := range formats {
		slide3.AddText(format, genppt.TextOptions{
			X:         2.0,
			Y:         yPos,
			Width:     6.0,
			Height:    0.4,
			FontSize:  16,
			FontColor: "#E0E0E0",
			Align:     genppt.AlignLeft,
		})
		yPos += 0.5
	}

	// 保存文件
	err := pres.WriteFile("audio_demo.pptx")
	if err != nil {
		fmt.Println("保存失败:", err)
		return
	}

	fmt.Println("音频演示PPT已生成: audio_demo.pptx")
	fmt.Println("\n使用说明:")
	fmt.Println("1. 取消注释 AddAudio 代码并提供音频文件路径")
	fmt.Println("2. 重新运行程序生成带音频的PPT")
}
