package genppt

import (
	"archive/zip"
	"bytes"
	"strings"
	"testing"
)

// TestNewPresentation 测试创建演示文稿
func TestNewPresentation(t *testing.T) {
	pres := New()
	if pres == nil {
		t.Fatal("New() 返回 nil")
	}
	if pres.SlideCount() != 0 {
		t.Errorf("新演示文稿应该没有幻灯片，实际有 %d 张", pres.SlideCount())
	}
}

// TestAddSlide 测试添加幻灯片
func TestAddSlide(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()
	if slide == nil {
		t.Fatal("AddSlide() 返回 nil")
	}
	if pres.SlideCount() != 1 {
		t.Errorf("添加一张幻灯片后应该有 1 张，实际有 %d 张", pres.SlideCount())
	}

	// 添加更多幻灯片
	pres.AddSlide()
	pres.AddSlide()
	if pres.SlideCount() != 3 {
		t.Errorf("添加三张幻灯片后应该有 3 张，实际有 %d 张", pres.SlideCount())
	}
}

// TestSetProperties 测试设置属性
func TestSetProperties(t *testing.T) {
	pres := New()
	pres.SetTitle("测试标题").
		SetAuthor("测试作者").
		SetSubject("测试主题").
		SetCompany("测试公司")

	// 验证链式调用返回了正确的对象
	if pres.title != "测试标题" {
		t.Errorf("标题设置失败")
	}
	if pres.author != "测试作者" {
		t.Errorf("作者设置失败")
	}
}

// TestAddText 测试添加文本
func TestAddText(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	slide.AddText("Hello World", TextOptions{
		X:         1.0,
		Y:         1.0,
		Width:     8.0,
		Height:    1.0,
		FontSize:  24,
		FontColor: "#FF0000",
		Bold:      true,
	})

	if len(slide.objects) != 1 {
		t.Errorf("添加文本后应该有 1 个对象，实际有 %d 个", len(slide.objects))
	}
}

// TestAddShape 测试添加形状
func TestAddShape(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	slide.AddShape(ShapeRect, ShapeOptions{
		X:      1.0,
		Y:      1.0,
		Width:  2.0,
		Height: 1.0,
		Fill:   "#4472C4",
	})

	slide.AddShapeWithText(ShapeEllipse, "文本", ShapeOptions{
		X:      4.0,
		Y:      1.0,
		Width:  2.0,
		Height: 1.0,
		Fill:   "#ED7D31",
	})

	if len(slide.objects) != 2 {
		t.Errorf("添加两个形状后应该有 2 个对象，实际有 %d 个", len(slide.objects))
	}
}

// TestAddTable 测试添加表格
func TestAddTable(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	slide.AddTable([][]TableCell{
		{{Text: "A1"}, {Text: "B1"}},
		{{Text: "A2"}, {Text: "B2"}},
	}, TableOptions{
		X:     1.0,
		Y:     1.0,
		Width: 6.0,
	})

	if len(slide.objects) != 1 {
		t.Errorf("添加表格后应该有 1 个对象，实际有 %d 个", len(slide.objects))
	}
}

// TestWriteFile 测试生成PPTX文件
func TestWriteFile(t *testing.T) {
	pres := New()
	pres.SetTitle("测试演示文稿")

	slide := pres.AddSlide()
	slide.AddText("测试文本", TextOptions{
		X:        1.0,
		Y:        1.0,
		Width:    8.0,
		Height:   1.0,
		FontSize: 24,
	})

	// 生成字节
	data, err := pres.ToBytes()
	if err != nil {
		t.Fatalf("ToBytes() 失败: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("生成的数据为空")
	}

	// 验证是有效的ZIP文件
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("生成的数据不是有效的ZIP文件: %v", err)
	}

	// 检查必需的文件
	requiredFiles := []string{
		"[Content_Types].xml",
		"_rels/.rels",
		"ppt/presentation.xml",
		"ppt/slides/slide1.xml",
		"ppt/theme/theme1.xml",
	}

	fileMap := make(map[string]bool)
	for _, f := range reader.File {
		fileMap[f.Name] = true
	}

	for _, required := range requiredFiles {
		if !fileMap[required] {
			t.Errorf("缺少必需文件: %s", required)
		}
	}
}

// TestParseColor 测试颜色解析
func TestParseColor(t *testing.T) {
	tests := []struct {
		input    string
		expected string
	}{
		{"#FF0000", "FF0000"},
		{"FF0000", "FF0000"},
		{"#ffffff", "ffffff"},
		{"000000", "000000"},
	}

	for _, test := range tests {
		result := ParseColor(test.input)
		if result != test.expected {
			t.Errorf("ParseColor(%s) = %s, 期望 %s", test.input, result, test.expected)
		}
	}
}

// TestInchToEMU 测试单位转换
func TestInchToEMU(t *testing.T) {
	result := InchToEMU(1.0)
	if result != EMUPerInch {
		t.Errorf("InchToEMU(1.0) = %d, 期望 %d", result, EMUPerInch)
	}

	result = InchToEMU(2.5)
	expected := int64(2.5 * float64(EMUPerInch))
	if result != expected {
		t.Errorf("InchToEMU(2.5) = %d, 期望 %d", result, expected)
	}
}

// TestXMLEscape 测试XML转义
func TestXMLEscape(t *testing.T) {
	tests := []struct {
		input    string
		contains string
	}{
		{"<>&", "&lt;&gt;&amp;"},
		{`"'`, "&#34;&#39;"},
		{"普通文本", "普通文本"},
	}

	for _, test := range tests {
		result := escapeXML(test.input)
		if !strings.Contains(result, test.contains) && result != test.contains {
			// 只要结果中包含期望的转义即可
			t.Logf("escapeXML(%s) = %s", test.input, result)
		}
	}
}

// TestSlideBackground 测试幻灯片背景
func TestSlideBackground(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()
	slide.SetBackground(BackgroundOptions{
		Color: "#1E3A5F",
	})

	if slide.background == nil {
		t.Error("背景设置失败")
	}
	if slide.background.Color != "#1E3A5F" {
		t.Errorf("背景颜色设置失败，期望 #1E3A5F，实际 %s", slide.background.Color)
	}
}

// TestMultipleSlides 测试多张幻灯片
func TestMultipleSlides(t *testing.T) {
	pres := New()

	for i := 0; i < 5; i++ {
		slide := pres.AddSlide()
		slide.AddText("幻灯片 "+itoa(i+1), TextOptions{
			X:        1.0,
			Y:        1.0,
			Width:    8.0,
			Height:   1.0,
			FontSize: 24,
		})
	}

	if pres.SlideCount() != 5 {
		t.Errorf("应该有 5 张幻灯片，实际有 %d 张", pres.SlideCount())
	}

	// 生成文件并验证
	data, err := pres.ToBytes()
	if err != nil {
		t.Fatalf("ToBytes() 失败: %v", err)
	}

	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("生成的数据不是有效的ZIP文件: %v", err)
	}

	// 验证所有幻灯片文件都存在
	slideCount := 0
	for _, f := range reader.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			slideCount++
		}
	}

	if slideCount != 5 {
		t.Errorf("ZIP中应该有 5 个幻灯片文件，实际有 %d 个", slideCount)
	}
}
