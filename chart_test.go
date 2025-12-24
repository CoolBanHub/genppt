package genppt

import (
	"archive/zip"
	"bytes"
	"strings"
	"testing"
)

// TestAddBarChart 测试柱状图
func TestAddBarChart(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	slide.AddBarChart("测试柱状图",
		[]string{"A", "B", "C"},
		map[string][]float64{
			"系列1": {10, 20, 30},
		},
		DefaultChartOptions(),
	)

	if len(slide.objects) != 1 {
		t.Errorf("添加柱状图后应该有 1 个对象，实际有 %d 个", len(slide.objects))
	}
}

// TestAddLineChart 测试折线图
func TestAddLineChart(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	slide.AddLineChart("测试折线图",
		[]string{"1月", "2月", "3月"},
		map[string][]float64{
			"数据": {100, 200, 150},
		},
		DefaultChartOptions(),
	)

	if len(slide.objects) != 1 {
		t.Errorf("添加折线图后应该有 1 个对象，实际有 %d 个", len(slide.objects))
	}
}

// TestAddPieChart 测试饼图
func TestAddPieChart(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	slide.AddPieChart("测试饼图",
		[]string{"A", "B", "C"},
		[]float64{30, 40, 30},
		DefaultChartOptions(),
	)

	if len(slide.objects) != 1 {
		t.Errorf("添加饼图后应该有 1 个对象，实际有 %d 个", len(slide.objects))
	}
}

// TestChartGeneration 测试图表生成
func TestChartGeneration(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	slide.AddChart(ChartBar, []ChartSeries{
		{
			Name:   "测试系列",
			Labels: []string{"A", "B", "C"},
			Values: []float64{10, 20, 30},
		},
	}, ChartOptions{
		X:         1.0,
		Y:         1.0,
		Width:     8.0,
		Height:    4.0,
		Title:     "测试图表",
		ShowTitle: true,
	})

	data, err := pres.ToBytes()
	if err != nil {
		t.Fatalf("ToBytes() 失败: %v", err)
	}

	// 验证ZIP包含图表文件
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("生成的数据不是有效的ZIP文件: %v", err)
	}

	hasChart := false
	for _, f := range reader.File {
		if strings.Contains(f.Name, "charts/chart") {
			hasChart = true
			break
		}
	}

	if !hasChart {
		t.Error("ZIP中应该包含图表文件")
	}
}

// TestMultipleCharts 测试多个图表
func TestMultipleCharts(t *testing.T) {
	pres := New()

	// 第一张幻灯片有一个图表
	slide1 := pres.AddSlide()
	slide1.AddBarChart("图表1", []string{"A"}, map[string][]float64{"S1": {10}}, DefaultChartOptions())

	// 第二张幻灯片有两个图表
	slide2 := pres.AddSlide()
	slide2.AddLineChart("图表2", []string{"B"}, map[string][]float64{"S2": {20}}, DefaultChartOptions())
	slide2.AddPieChart("图表3", []string{"C", "D"}, []float64{50, 50}, DefaultChartOptions())

	data, err := pres.ToBytes()
	if err != nil {
		t.Fatalf("ToBytes() 失败: %v", err)
	}

	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("生成的数据不是有效的ZIP文件: %v", err)
	}

	chartCount := 0
	for _, f := range reader.File {
		if strings.Contains(f.Name, "charts/chart") && strings.HasSuffix(f.Name, ".xml") {
			chartCount++
		}
	}

	if chartCount != 3 {
		t.Errorf("ZIP中应该有 3 个图表文件，实际有 %d 个", chartCount)
	}
}

// TestChartTypes 测试不同图表类型
func TestChartTypes(t *testing.T) {
	chartTypes := []ChartType{
		ChartBar,
		ChartBarStacked,
		ChartLine,
		ChartLineSmooth,
		ChartPie,
		ChartDoughnut,
		ChartArea,
	}

	for _, ct := range chartTypes {
		pres := New()
		slide := pres.AddSlide()

		slide.AddChart(ct, []ChartSeries{
			{Name: "测试", Labels: []string{"A", "B"}, Values: []float64{10, 20}},
		}, DefaultChartOptions())

		_, err := pres.ToBytes()
		if err != nil {
			t.Errorf("图表类型 %s 生成失败: %v", ct, err)
		}
	}
}
