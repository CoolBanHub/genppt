// 图表示例
package main

import (
	"fmt"

	"github.com/CoolBanHub/genppt"
)

func main() {
	pres := genppt.New()
	pres.SetTitle("图表演示")
	pres.SetAuthor("GenPPT")

	// 第一张幻灯片 - 柱状图
	slide1 := pres.AddSlide()
	slide1.AddText("销售数据分析", genppt.TextOptions{
		X:         0.5,
		Y:         0.3,
		Width:     9.0,
		Height:    0.6,
		FontSize:  28,
		FontColor: "#1E3A5F",
		Bold:      true,
	})

	slide1.AddChart(genppt.ChartBar, []genppt.ChartSeries{
		{
			Name:   "2023年",
			Labels: []string{"Q1", "Q2", "Q3", "Q4"},
			Values: []float64{120, 150, 180, 200},
		},
		{
			Name:   "2024年",
			Labels: []string{"Q1", "Q2", "Q3", "Q4"},
			Values: []float64{150, 180, 220, 260},
		},
	}, genppt.ChartOptions{
		X:          0.5,
		Y:          1.2,
		Width:      9.0,
		Height:     4.0,
		Title:      "季度销售对比",
		ShowTitle:  true,
		ShowLegend: true,
		ShowValues: true,
	})

	// 第二张幻灯片 - 折线图
	slide2 := pres.AddSlide()
	slide2.AddText("趋势分析", genppt.TextOptions{
		X:         0.5,
		Y:         0.3,
		Width:     9.0,
		Height:    0.6,
		FontSize:  28,
		FontColor: "#1E3A5F",
		Bold:      true,
	})

	slide2.AddChart(genppt.ChartLine, []genppt.ChartSeries{
		{
			Name:   "用户增长",
			Labels: []string{"1月", "2月", "3月", "4月", "5月", "6月"},
			Values: []float64{1000, 1500, 2200, 3100, 4200, 5500},
		},
		{
			Name:   "活跃用户",
			Labels: []string{"1月", "2月", "3月", "4月", "5月", "6月"},
			Values: []float64{800, 1200, 1800, 2500, 3400, 4500},
		},
	}, genppt.ChartOptions{
		X:          0.5,
		Y:          1.2,
		Width:      9.0,
		Height:     4.0,
		Title:      "用户增长趋势",
		ShowTitle:  true,
		ShowLegend: true,
	})

	// 第三张幻灯片 - 饼图
	slide3 := pres.AddSlide()
	slide3.AddText("市场份额", genppt.TextOptions{
		X:         0.5,
		Y:         0.3,
		Width:     9.0,
		Height:    0.6,
		FontSize:  28,
		FontColor: "#1E3A5F",
		Bold:      true,
	})

	slide3.AddPieChart("市场份额分布",
		[]string{"产品A", "产品B", "产品C", "产品D", "其他"},
		[]float64{35, 25, 20, 12, 8},
		genppt.ChartOptions{
			X:          1.5,
			Y:          1.2,
			Width:      7.0,
			Height:     4.0,
			ShowLegend: true,
		},
	)

	// 第四张幻灯片 - 环形图
	slide4 := pres.AddSlide()
	slide4.AddText("预算分配", genppt.TextOptions{
		X:         0.5,
		Y:         0.3,
		Width:     9.0,
		Height:    0.6,
		FontSize:  28,
		FontColor: "#1E3A5F",
		Bold:      true,
	})

	slide4.AddChart(genppt.ChartDoughnut, []genppt.ChartSeries{
		{
			Name:   "预算",
			Labels: []string{"研发", "市场", "运营", "人力"},
			Values: []float64{40, 25, 20, 15},
		},
	}, genppt.ChartOptions{
		X:          1.5,
		Y:          1.2,
		Width:      7.0,
		Height:     4.0,
		Title:      "部门预算分配",
		ShowTitle:  true,
		ShowLegend: true,
		HoleSize:   60,
	})

	err := pres.WriteFile("chart_demo.pptx")
	if err != nil {
		fmt.Printf("保存失败: %v\n", err)
		return
	}
	fmt.Println("演示文稿已保存: chart_demo.pptx")
	fmt.Printf("共 %d 张幻灯片\n", pres.SlideCount())
}
