package genppt

import (
	"testing"
)

// TestFromMarkdown 测试Markdown转PPT
func TestFromMarkdown(t *testing.T) {
	markdown := `# 第一张幻灯片

这是一些内容。

# 第二张幻灯片

- 列表项1
- 列表项2
- 列表项3
`

	pres := FromMarkdown(markdown)
	if pres == nil {
		t.Fatal("FromMarkdown() 返回 nil")
	}

	if pres.SlideCount() != 2 {
		t.Errorf("应该有 2 张幻灯片，实际有 %d 张", pres.SlideCount())
	}
}

// TestMarkdownHeadings 测试标题解析
func TestMarkdownHeadings(t *testing.T) {
	markdown := `# 标题1

## 子标题

内容

# 标题2

更多内容
`

	pres := FromMarkdown(markdown)
	if pres.SlideCount() != 2 {
		t.Errorf("应该有 2 张幻灯片，实际有 %d 张", pres.SlideCount())
	}
}

// TestMarkdownCodeBlock 测试代码块
func TestMarkdownCodeBlock(t *testing.T) {
	markdown := "# 代码示例\n\n```go\npackage main\n\nfunc main() {\n    println(\"hello\")\n}\n```\n"

	pres := FromMarkdown(markdown)
	if pres.SlideCount() != 1 {
		t.Errorf("应该有 1 张幻灯片，实际有 %d 张", pres.SlideCount())
	}
}

// TestMarkdownBulletList 测试列表
func TestMarkdownBulletList(t *testing.T) {
	markdown := `# 列表测试

- 项目1
- 项目2
- 项目3

1. 有序1
2. 有序2
3. 有序3
`

	pres := FromMarkdown(markdown)
	if pres.SlideCount() != 1 {
		t.Errorf("应该有 1 张幻灯片，实际有 %d 张", pres.SlideCount())
	}
}

// TestMarkdownSeparator 测试分隔符
func TestMarkdownSeparator(t *testing.T) {
	markdown := `# 第一页

内容1

---

# 第二页

内容2

---

# 第三页

内容3
`

	pres := FromMarkdown(markdown)
	if pres.SlideCount() != 3 {
		t.Errorf("应该有 3 张幻灯片，实际有 %d 张", pres.SlideCount())
	}
}

// TestMarkdownOptions 测试自定义选项
func TestMarkdownOptions(t *testing.T) {
	opts := DefaultMarkdownOptions()
	opts.TitleFontSize = 48
	opts.SlideBackground = "#000000"

	markdown := `# 测试

内容
`

	pres := FromMarkdownWithOptions(markdown, opts)
	if pres.SlideCount() != 1 {
		t.Errorf("应该有 1 张幻灯片，实际有 %d 张", pres.SlideCount())
	}
}

// TestParseInlineMarkdown 测试行内格式解析
func TestParseInlineMarkdown(t *testing.T) {
	tests := []struct {
		input    string
		expected string
	}{
		{"**粗体**", "粗体"},
		{"*斜体*", "斜体"},
		{"`代码`", "代码"},
		{"[链接](https://example.com)", "链接"},
		{"普通文本", "普通文本"},
		{"混合**粗体**和*斜体*", "混合粗体和斜体"},
	}

	for _, test := range tests {
		result := parseInlineMarkdown(test.input)
		if result != test.expected {
			t.Errorf("parseInlineMarkdown(%s) = %s, 期望 %s", test.input, result, test.expected)
		}
	}
}

// TestMarkdownWriteFile 测试Markdown生成文件
func TestMarkdownWriteFile(t *testing.T) {
	markdown := `# 演示文稿

这是一个测试。

## 功能列表

- 功能1
- 功能2

# 第二页

更多内容
`

	pres := FromMarkdown(markdown)
	data, err := pres.ToBytes()
	if err != nil {
		t.Fatalf("ToBytes() 失败: %v", err)
	}

	if len(data) == 0 {
		t.Fatal("生成的数据为空")
	}
}
