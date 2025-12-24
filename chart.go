package genppt

import (
	"strings"
)

// ChartType 图表类型
type ChartType string

const (
	// ChartBar 柱状图
	ChartBar ChartType = "bar"
	// ChartBarStacked 堆叠柱状图
	ChartBarStacked ChartType = "barStacked"
	// ChartBar3D 3D柱状图
	ChartBar3D ChartType = "bar3D"
	// ChartLine 折线图
	ChartLine ChartType = "line"
	// ChartLineSmooth 平滑折线图
	ChartLineSmooth ChartType = "lineSmooth"
	// ChartPie 饼图
	ChartPie ChartType = "pie"
	// ChartPie3D 3D饼图
	ChartPie3D ChartType = "pie3D"
	// ChartDoughnut 环形图
	ChartDoughnut ChartType = "doughnut"
	// ChartArea 面积图
	ChartArea ChartType = "area"
	// ChartScatter 散点图
	ChartScatter ChartType = "scatter"
)

// ChartOptions 图表选项
type ChartOptions struct {
	X                float64  // X坐标（英寸）
	Y                float64  // Y坐标（英寸）
	Width            float64  // 宽度（英寸）
	Height           float64  // 高度（英寸）
	Title            string   // 图表标题
	ShowTitle        bool     // 是否显示标题
	ShowLegend       bool     // 是否显示图例
	LegendPos        string   // 图例位置: "r"(右), "l"(左), "t"(上), "b"(下)
	ShowValues       bool     // 是否显示数据标签
	ShowCategoryAxis bool     // 是否显示类别轴
	ShowValueAxis    bool     // 是否显示数值轴
	BarGapWidth      int      // 柱间距百分比
	HoleSize         int      // 环形图空心大小百分比(0-90)
	Colors           []string // 自定义颜色列表
}

// ChartSeries 图表数据系列
type ChartSeries struct {
	Name   string    // 系列名称
	Labels []string  // 类别标签
	Values []float64 // 数据值
	Color  string    // 系列颜色（可选）
}

// DefaultChartOptions 返回默认图表选项
func DefaultChartOptions() ChartOptions {
	return ChartOptions{
		X:                1.0,
		Y:                1.5,
		Width:            8.0,
		Height:           4.0,
		ShowTitle:        true,
		ShowLegend:       true,
		LegendPos:        "r",
		ShowValues:       false,
		ShowCategoryAxis: true,
		ShowValueAxis:    true,
		BarGapWidth:      150,
		HoleSize:         50,
		Colors: []string{
			"4472C4", "ED7D31", "A5A5A5", "FFC000",
			"5B9BD5", "70AD47", "264478", "9E480E",
		},
	}
}

// chartObject 图表对象
type chartObject struct {
	chartType ChartType
	series    []ChartSeries
	options   ChartOptions
	chartIdx  int // 图表索引
}

func (c *chartObject) getType() string { return "chart" }

// AddChart 添加图表
func (s *Slide) AddChart(chartType ChartType, series []ChartSeries, opts ChartOptions) *Slide {
	// 设置默认值
	if opts.Width == 0 {
		opts.Width = 8.0
	}
	if opts.Height == 0 {
		opts.Height = 4.0
	}
	if len(opts.Colors) == 0 {
		opts.Colors = DefaultChartOptions().Colors
	}
	if opts.LegendPos == "" {
		opts.LegendPos = "r"
	}
	if opts.BarGapWidth == 0 {
		opts.BarGapWidth = 150
	}
	if opts.HoleSize == 0 {
		opts.HoleSize = 50
	}

	// 计算图表索引
	chartIdx := 1
	for _, slide := range s.presentation.slides {
		for _, obj := range slide.objects {
			if _, ok := obj.(*chartObject); ok {
				chartIdx++
			}
		}
	}

	obj := &chartObject{
		chartType: chartType,
		series:    series,
		options:   opts,
		chartIdx:  chartIdx,
	}
	s.objects = append(s.objects, obj)
	return s
}

// AddBarChart 添加柱状图（便捷方法）
func (s *Slide) AddBarChart(title string, labels []string, data map[string][]float64, opts ChartOptions) *Slide {
	series := make([]ChartSeries, 0)
	for name, values := range data {
		series = append(series, ChartSeries{
			Name:   name,
			Labels: labels,
			Values: values,
		})
	}
	opts.Title = title
	opts.ShowTitle = true
	return s.AddChart(ChartBar, series, opts)
}

// AddLineChart 添加折线图（便捷方法）
func (s *Slide) AddLineChart(title string, labels []string, data map[string][]float64, opts ChartOptions) *Slide {
	series := make([]ChartSeries, 0)
	for name, values := range data {
		series = append(series, ChartSeries{
			Name:   name,
			Labels: labels,
			Values: values,
		})
	}
	opts.Title = title
	opts.ShowTitle = true
	return s.AddChart(ChartLine, series, opts)
}

// AddPieChart 添加饼图（便捷方法）
func (s *Slide) AddPieChart(title string, labels []string, values []float64, opts ChartOptions) *Slide {
	series := []ChartSeries{{
		Name:   title,
		Labels: labels,
		Values: values,
	}}
	opts.Title = title
	opts.ShowTitle = true
	return s.AddChart(ChartPie, series, opts)
}

// generateChart 生成图表XML引用
func (s *Slide) generateChart(c *chartObject, id int) string {
	var sb strings.Builder

	x := InchToEMU(c.options.X)
	y := InchToEMU(c.options.Y)
	cx := InchToEMU(c.options.Width)
	cy := InchToEMU(c.options.Height)

	rID := "rId" + itoa(200+c.chartIdx) // 使用较大的rId避免冲突

	sb.WriteString(`<p:graphicFrame>`)
	sb.WriteString(`<p:nvGraphicFramePr>`)
	sb.WriteString(`<p:cNvPr id="`)
	sb.WriteString(itoa(id))
	sb.WriteString(`" name="Chart `)
	sb.WriteString(itoa(c.chartIdx))
	sb.WriteString(`"/>`)
	sb.WriteString(`<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>`)
	sb.WriteString(`<p:nvPr/>`)
	sb.WriteString(`</p:nvGraphicFramePr>`)

	sb.WriteString(`<p:xfrm>`)
	sb.WriteString(`<a:off x="`)
	sb.WriteString(itoa(int(x)))
	sb.WriteString(`" y="`)
	sb.WriteString(itoa(int(y)))
	sb.WriteString(`"/>`)
	sb.WriteString(`<a:ext cx="`)
	sb.WriteString(itoa(int(cx)))
	sb.WriteString(`" cy="`)
	sb.WriteString(itoa(int(cy)))
	sb.WriteString(`"/>`)
	sb.WriteString(`</p:xfrm>`)

	sb.WriteString(`<a:graphic>`)
	sb.WriteString(`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">`)
	sb.WriteString(`<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="`)
	sb.WriteString(rID)
	sb.WriteString(`"/>`)
	sb.WriteString(`</a:graphicData>`)
	sb.WriteString(`</a:graphic>`)
	sb.WriteString(`</p:graphicFrame>`)

	return sb.String()
}

// generateChartXML 生成图表XML文件内容
func (c *chartObject) generateChartXML() string {
	var sb strings.Builder

	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)
	sb.WriteString(`<c:date1904 val="0"/>`)
	sb.WriteString(`<c:lang val="zh-CN"/>`)
	sb.WriteString(`<c:roundedCorners val="0"/>`)

	sb.WriteString(`<c:chart>`)

	// 标题
	if c.options.ShowTitle && c.options.Title != "" {
		sb.WriteString(`<c:title>`)
		sb.WriteString(`<c:tx>`)
		sb.WriteString(`<c:rich>`)
		sb.WriteString(`<a:bodyPr/>`)
		sb.WriteString(`<a:lstStyle/>`)
		sb.WriteString(`<a:p>`)
		sb.WriteString(`<a:pPr><a:defRPr sz="1400" b="0"/></a:pPr>`)
		sb.WriteString(`<a:r>`)
		sb.WriteString(`<a:rPr lang="zh-CN" sz="1400" b="1"/>`)
		sb.WriteString(`<a:t>`)
		sb.WriteString(escapeXML(c.options.Title))
		sb.WriteString(`</a:t>`)
		sb.WriteString(`</a:r>`)
		sb.WriteString(`</a:p>`)
		sb.WriteString(`</c:rich>`)
		sb.WriteString(`</c:tx>`)
		sb.WriteString(`<c:overlay val="0"/>`)
		sb.WriteString(`</c:title>`)
	} else {
		sb.WriteString(`<c:autoTitleDeleted val="1"/>`)
	}

	// 绘图区
	sb.WriteString(`<c:plotArea>`)
	sb.WriteString(`<c:layout/>`)

	// 根据图表类型生成相应的图表
	switch c.chartType {
	case ChartBar, ChartBarStacked, ChartBar3D:
		sb.WriteString(c.generateBarChart())
	case ChartLine, ChartLineSmooth:
		sb.WriteString(c.generateLineChart())
	case ChartPie, ChartPie3D:
		sb.WriteString(c.generatePieChart())
	case ChartDoughnut:
		sb.WriteString(c.generateDoughnutChart())
	case ChartArea:
		sb.WriteString(c.generateAreaChart())
	default:
		sb.WriteString(c.generateBarChart())
	}

	// 坐标轴（饼图和环形图不需要）
	if c.chartType != ChartPie && c.chartType != ChartPie3D && c.chartType != ChartDoughnut {
		// 类别轴
		sb.WriteString(`<c:catAx>`)
		sb.WriteString(`<c:axId val="1"/>`)
		sb.WriteString(`<c:scaling><c:orientation val="minMax"/></c:scaling>`)
		sb.WriteString(`<c:delete val="0"/>`)
		sb.WriteString(`<c:axPos val="b"/>`)
		sb.WriteString(`<c:majorTickMark val="out"/>`)
		sb.WriteString(`<c:minorTickMark val="none"/>`)
		sb.WriteString(`<c:tickLblPos val="nextTo"/>`)
		sb.WriteString(`<c:crossAx val="2"/>`)
		sb.WriteString(`<c:crosses val="autoZero"/>`)
		sb.WriteString(`<c:auto val="1"/>`)
		sb.WriteString(`<c:lblAlgn val="ctr"/>`)
		sb.WriteString(`<c:lblOffset val="100"/>`)
		sb.WriteString(`</c:catAx>`)

		// 数值轴
		sb.WriteString(`<c:valAx>`)
		sb.WriteString(`<c:axId val="2"/>`)
		sb.WriteString(`<c:scaling><c:orientation val="minMax"/></c:scaling>`)
		sb.WriteString(`<c:delete val="0"/>`)
		sb.WriteString(`<c:axPos val="l"/>`)
		sb.WriteString(`<c:majorGridlines/>`)
		sb.WriteString(`<c:majorTickMark val="out"/>`)
		sb.WriteString(`<c:minorTickMark val="none"/>`)
		sb.WriteString(`<c:tickLblPos val="nextTo"/>`)
		sb.WriteString(`<c:crossAx val="1"/>`)
		sb.WriteString(`<c:crosses val="autoZero"/>`)
		sb.WriteString(`<c:crossBetween val="between"/>`)
		sb.WriteString(`</c:valAx>`)
	}

	sb.WriteString(`</c:plotArea>`)

	// 图例
	if c.options.ShowLegend {
		sb.WriteString(`<c:legend>`)
		sb.WriteString(`<c:legendPos val="`)
		sb.WriteString(c.options.LegendPos)
		sb.WriteString(`"/>`)
		sb.WriteString(`<c:overlay val="0"/>`)
		sb.WriteString(`</c:legend>`)
	}

	sb.WriteString(`<c:plotVisOnly val="1"/>`)
	sb.WriteString(`<c:dispBlanksAs val="gap"/>`)
	sb.WriteString(`</c:chart>`)

	sb.WriteString(`<c:printSettings>`)
	sb.WriteString(`<c:headerFooter/>`)
	sb.WriteString(`<c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>`)
	sb.WriteString(`<c:pageSetup/>`)
	sb.WriteString(`</c:printSettings>`)

	sb.WriteString(`</c:chartSpace>`)

	return sb.String()
}

// generateBarChart 生成柱状图XML
func (c *chartObject) generateBarChart() string {
	var sb strings.Builder

	sb.WriteString(`<c:barChart>`)

	// 柱状图方向和分组
	sb.WriteString(`<c:barDir val="col"/>`)
	if c.chartType == ChartBarStacked {
		sb.WriteString(`<c:grouping val="stacked"/>`)
	} else {
		sb.WriteString(`<c:grouping val="clustered"/>`)
	}
	sb.WriteString(`<c:varyColors val="0"/>`)

	// 数据系列
	for i, series := range c.series {
		sb.WriteString(c.generateSeries(i, series, "bar"))
	}

	// 数据标签
	if c.options.ShowValues {
		sb.WriteString(`<c:dLbls>`)
		sb.WriteString(`<c:showLegendKey val="0"/>`)
		sb.WriteString(`<c:showVal val="1"/>`)
		sb.WriteString(`<c:showCatName val="0"/>`)
		sb.WriteString(`<c:showSerName val="0"/>`)
		sb.WriteString(`<c:showPercent val="0"/>`)
		sb.WriteString(`</c:dLbls>`)
	}

	sb.WriteString(`<c:gapWidth val="`)
	sb.WriteString(itoa(c.options.BarGapWidth))
	sb.WriteString(`"/>`)
	sb.WriteString(`<c:axId val="1"/>`)
	sb.WriteString(`<c:axId val="2"/>`)
	sb.WriteString(`</c:barChart>`)

	return sb.String()
}

// generateLineChart 生成折线图XML
func (c *chartObject) generateLineChart() string {
	var sb strings.Builder

	sb.WriteString(`<c:lineChart>`)
	sb.WriteString(`<c:grouping val="standard"/>`)
	sb.WriteString(`<c:varyColors val="0"/>`)

	// 数据系列
	for i, series := range c.series {
		sb.WriteString(c.generateSeries(i, series, "line"))
	}

	// 数据标签
	if c.options.ShowValues {
		sb.WriteString(`<c:dLbls>`)
		sb.WriteString(`<c:showLegendKey val="0"/>`)
		sb.WriteString(`<c:showVal val="1"/>`)
		sb.WriteString(`<c:showCatName val="0"/>`)
		sb.WriteString(`<c:showSerName val="0"/>`)
		sb.WriteString(`<c:showPercent val="0"/>`)
		sb.WriteString(`</c:dLbls>`)
	}

	sb.WriteString(`<c:marker val="1"/>`)
	if c.chartType == ChartLineSmooth {
		sb.WriteString(`<c:smooth val="1"/>`)
	} else {
		sb.WriteString(`<c:smooth val="0"/>`)
	}
	sb.WriteString(`<c:axId val="1"/>`)
	sb.WriteString(`<c:axId val="2"/>`)
	sb.WriteString(`</c:lineChart>`)

	return sb.String()
}

// generatePieChart 生成饼图XML
func (c *chartObject) generatePieChart() string {
	var sb strings.Builder

	if c.chartType == ChartPie3D {
		sb.WriteString(`<c:pie3DChart>`)
	} else {
		sb.WriteString(`<c:pieChart>`)
	}
	sb.WriteString(`<c:varyColors val="1"/>`)

	// 只使用第一个系列
	if len(c.series) > 0 {
		sb.WriteString(c.generateSeries(0, c.series[0], "pie"))
	}

	// 数据标签
	sb.WriteString(`<c:dLbls>`)
	sb.WriteString(`<c:showLegendKey val="0"/>`)
	if c.options.ShowValues {
		sb.WriteString(`<c:showVal val="1"/>`)
	} else {
		sb.WriteString(`<c:showVal val="0"/>`)
	}
	sb.WriteString(`<c:showCatName val="1"/>`)
	sb.WriteString(`<c:showSerName val="0"/>`)
	sb.WriteString(`<c:showPercent val="1"/>`)
	sb.WriteString(`<c:showLeaderLines val="1"/>`)
	sb.WriteString(`</c:dLbls>`)

	if c.chartType == ChartPie3D {
		sb.WriteString(`</c:pie3DChart>`)
	} else {
		sb.WriteString(`</c:pieChart>`)
	}

	return sb.String()
}

// generateDoughnutChart 生成环形图XML
func (c *chartObject) generateDoughnutChart() string {
	var sb strings.Builder

	sb.WriteString(`<c:doughnutChart>`)
	sb.WriteString(`<c:varyColors val="1"/>`)

	// 只使用第一个系列
	if len(c.series) > 0 {
		sb.WriteString(c.generateSeries(0, c.series[0], "pie"))
	}

	// 数据标签
	sb.WriteString(`<c:dLbls>`)
	sb.WriteString(`<c:showLegendKey val="0"/>`)
	sb.WriteString(`<c:showVal val="0"/>`)
	sb.WriteString(`<c:showCatName val="1"/>`)
	sb.WriteString(`<c:showSerName val="0"/>`)
	sb.WriteString(`<c:showPercent val="1"/>`)
	sb.WriteString(`</c:dLbls>`)

	sb.WriteString(`<c:firstSliceAng val="0"/>`)
	sb.WriteString(`<c:holeSize val="`)
	sb.WriteString(itoa(c.options.HoleSize))
	sb.WriteString(`"/>`)
	sb.WriteString(`</c:doughnutChart>`)

	return sb.String()
}

// generateAreaChart 生成面积图XML
func (c *chartObject) generateAreaChart() string {
	var sb strings.Builder

	sb.WriteString(`<c:areaChart>`)
	sb.WriteString(`<c:grouping val="standard"/>`)
	sb.WriteString(`<c:varyColors val="0"/>`)

	// 数据系列
	for i, series := range c.series {
		sb.WriteString(c.generateSeries(i, series, "area"))
	}

	sb.WriteString(`<c:axId val="1"/>`)
	sb.WriteString(`<c:axId val="2"/>`)
	sb.WriteString(`</c:areaChart>`)

	return sb.String()
}

// generateSeries 生成数据系列XML
func (c *chartObject) generateSeries(idx int, series ChartSeries, chartKind string) string {
	var sb strings.Builder

	sb.WriteString(`<c:ser>`)
	sb.WriteString(`<c:idx val="`)
	sb.WriteString(itoa(idx))
	sb.WriteString(`"/>`)
	sb.WriteString(`<c:order val="`)
	sb.WriteString(itoa(idx))
	sb.WriteString(`"/>`)

	// 系列名称
	if series.Name != "" {
		sb.WriteString(`<c:tx>`)
		sb.WriteString(`<c:v>`)
		sb.WriteString(escapeXML(series.Name))
		sb.WriteString(`</c:v>`)
		sb.WriteString(`</c:tx>`)
	}

	// 系列颜色
	color := series.Color
	if color == "" && idx < len(c.options.Colors) {
		color = c.options.Colors[idx]
	}
	if color != "" {
		sb.WriteString(`<c:spPr>`)
		sb.WriteString(`<a:solidFill>`)
		sb.WriteString(`<a:srgbClr val="`)
		sb.WriteString(ParseColor(color))
		sb.WriteString(`"/>`)
		sb.WriteString(`</a:solidFill>`)
		if chartKind == "line" {
			sb.WriteString(`<a:ln w="28575">`)
			sb.WriteString(`<a:solidFill>`)
			sb.WriteString(`<a:srgbClr val="`)
			sb.WriteString(ParseColor(color))
			sb.WriteString(`"/>`)
			sb.WriteString(`</a:solidFill>`)
			sb.WriteString(`</a:ln>`)
		}
		sb.WriteString(`</c:spPr>`)
	}

	// 折线图标记点
	if chartKind == "line" {
		sb.WriteString(`<c:marker>`)
		sb.WriteString(`<c:symbol val="circle"/>`)
		sb.WriteString(`<c:size val="5"/>`)
		if color != "" {
			sb.WriteString(`<c:spPr>`)
			sb.WriteString(`<a:solidFill>`)
			sb.WriteString(`<a:srgbClr val="`)
			sb.WriteString(ParseColor(color))
			sb.WriteString(`"/>`)
			sb.WriteString(`</a:solidFill>`)
			sb.WriteString(`</c:spPr>`)
		}
		sb.WriteString(`</c:marker>`)
	}

	// 饼图/环形图数据点颜色
	if chartKind == "pie" && len(series.Values) > 0 {
		sb.WriteString(`<c:dPt>`)
		for i := range series.Values {
			if i < len(c.options.Colors) {
				sb.WriteString(`<c:idx val="`)
				sb.WriteString(itoa(i))
				sb.WriteString(`"/>`)
				sb.WriteString(`<c:spPr>`)
				sb.WriteString(`<a:solidFill>`)
				sb.WriteString(`<a:srgbClr val="`)
				sb.WriteString(ParseColor(c.options.Colors[i]))
				sb.WriteString(`"/>`)
				sb.WriteString(`</a:solidFill>`)
				sb.WriteString(`</c:spPr>`)
			}
		}
		sb.WriteString(`</c:dPt>`)
	}

	// 类别数据
	if len(series.Labels) > 0 {
		sb.WriteString(`<c:cat>`)
		sb.WriteString(`<c:strRef>`)
		sb.WriteString(`<c:strCache>`)
		sb.WriteString(`<c:ptCount val="`)
		sb.WriteString(itoa(len(series.Labels)))
		sb.WriteString(`"/>`)
		for i, label := range series.Labels {
			sb.WriteString(`<c:pt idx="`)
			sb.WriteString(itoa(i))
			sb.WriteString(`"><c:v>`)
			sb.WriteString(escapeXML(label))
			sb.WriteString(`</c:v></c:pt>`)
		}
		sb.WriteString(`</c:strCache>`)
		sb.WriteString(`</c:strRef>`)
		sb.WriteString(`</c:cat>`)
	}

	// 数值数据
	if len(series.Values) > 0 {
		sb.WriteString(`<c:val>`)
		sb.WriteString(`<c:numRef>`)
		sb.WriteString(`<c:numCache>`)
		sb.WriteString(`<c:formatCode>General</c:formatCode>`)
		sb.WriteString(`<c:ptCount val="`)
		sb.WriteString(itoa(len(series.Values)))
		sb.WriteString(`"/>`)
		for i, val := range series.Values {
			sb.WriteString(`<c:pt idx="`)
			sb.WriteString(itoa(i))
			sb.WriteString(`"><c:v>`)
			sb.WriteString(ftoa(val))
			sb.WriteString(`</c:v></c:pt>`)
		}
		sb.WriteString(`</c:numCache>`)
		sb.WriteString(`</c:numRef>`)
		sb.WriteString(`</c:val>`)
	}

	sb.WriteString(`</c:ser>`)

	return sb.String()
}
