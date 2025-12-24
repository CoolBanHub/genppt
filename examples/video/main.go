package main

import (
	"fmt"
	"os"

	"github.com/CoolBanHub/genppt"
)

func main() {

	pres := genppt.New()

	pres.SetTitle("视频演示")
	pres.SetAuthor("GenPPT")

	// 读取封面图片（可选）
	poster, _ := os.ReadFile("demo.jpg")

	slide1 := pres.AddSlide()
	slide1.SetBackground(genppt.BackgroundOptions{
		Color: "#1A1A2E",
	})

	slide1.AddText("视频自动播放演示", genppt.TextOptions{
		X:         1.0,
		Y:         0.3,
		Width:     8.0,
		Height:    0.5,
		FontSize:  24,
		FontColor: "#FFFFFF",
		Bold:      true,
		Align:     genppt.AlignCenter,
	})

	// 添加视频（带自动播放和循环）
	slide1.AddVideo(genppt.VideoOptions{
		Path:     "demo.mp4",
		X:        2.0,
		Y:        1.0,
		Width:    6.0,
		Height:   4.0,
		Poster:   poster,
		AutoPlay: true, // 自动播放
		Loop:     true, // 循环播放
		Muted:    false,
	})

	err := pres.WriteFile("video_demo.pptx")
	if err != nil {
		fmt.Printf("保存失败: %v\n", err)
		return
	}
	fmt.Println("演示文稿已保存: video_demo.pptx")
	fmt.Printf("共 %d 张幻灯片\n", pres.SlideCount())
	fmt.Println("\n提示: 请确保 demo.mp4 文件存在于当前目录")
	fmt.Println("自动播放功能需要在幻灯片放映模式下生效")
}
