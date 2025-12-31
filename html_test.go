package genppt

import (
	"os"
	"strings"
	"testing"
)

func TestFromHTML_Basic(t *testing.T) {
	html := `
	<h1>第一张幻灯片</h1>
	<p>这是一段普通文本。</p>
	
	<h1>第二张幻灯片</h1>
	<h2>小标题</h2>
	<ul>
		<li>列表项1</li>
		<li>列表项2</li>
		<li>列表项3</li>
	</ul>
	`

	pres := FromHTML(html)

	if pres.SlideCount() != 2 {
		t.Errorf("期望2张幻灯片，实际得到 %d 张", pres.SlideCount())
	}
}

func TestFromHTML_WithCodeBlock(t *testing.T) {
	html := `
	<h1>代码示例</h1>
	<pre><code>func main() {
    fmt.Println("Hello, World!")
}</code></pre>
	`

	pres := FromHTML(html)

	if pres.SlideCount() != 1 {
		t.Errorf("期望1张幻灯片，实际得到 %d 张", pres.SlideCount())
	}
}

func TestFromHTML_WithTable(t *testing.T) {
	html := `
	<h1>表格示例</h1>
	<table>
		<tr><th>姓名</th><th>年龄</th></tr>
		<tr><td>张三</td><td>25</td></tr>
		<tr><td>李四</td><td>30</td></tr>
	</table>
	`

	pres := FromHTML(html)

	if pres.SlideCount() != 1 {
		t.Errorf("期望1张幻灯片，实际得到 %d 张", pres.SlideCount())
	}
}

func TestFromHTML_WithHR(t *testing.T) {
	html := `
	<h1>幻灯片一</h1>
	<p>内容一</p>
	<hr>
	<h1>幻灯片二</h1>
	<p>内容二</p>
	`

	pres := FromHTML(html)

	if pres.SlideCount() != 2 {
		t.Errorf("期望2张幻灯片，实际得到 %d 张", pres.SlideCount())
	}
}

func TestFromHTML_WithSection(t *testing.T) {
	html := `
	<section>
		<h1>第一部分</h1>
		<p>内容</p>
	</section>
	<section>
		<h1>第二部分</h1>
		<p>内容</p>
	</section>
	`

	pres := FromHTML(html)

	if pres.SlideCount() < 2 {
		t.Errorf("期望至少2张幻灯片，实际得到 %d 张", pres.SlideCount())
	}
}

func TestFromHTMLWithOptions(t *testing.T) {
	html := `<h1>测试</h1><p>内容</p>`

	opts := HTMLOptions{
		TitleFontSize:   48,
		HeadingFontSize: 36,
		BodyFontSize:    20,
		TitleColor:      "#000000",
		HeadingColor:    "#333333",
		BodyColor:       "#666666",
		SlideBackground: "#FFFFFF",
	}

	pres := FromHTMLWithOptions(html, opts)

	if pres.SlideCount() != 1 {
		t.Errorf("期望1张幻灯片，实际得到 %d 张", pres.SlideCount())
	}
}

func TestFromHTMLFile(t *testing.T) {
	// 创建临时文件
	content := `<h1>文件测试</h1><p>从文件读取的内容</p>`
	tmpFile, err := os.CreateTemp("", "test_*.html")
	if err != nil {
		t.Fatalf("创建临时文件失败: %v", err)
	}
	defer os.Remove(tmpFile.Name())

	if _, err := tmpFile.WriteString(content); err != nil {
		t.Fatalf("写入临时文件失败: %v", err)
	}
	tmpFile.Close()

	pres, err := FromHTMLFile(tmpFile.Name())
	if err != nil {
		t.Fatalf("FromHTMLFile 失败: %v", err)
	}

	if pres.SlideCount() != 1 {
		t.Errorf("期望1张幻灯片，实际得到 %d 张", pres.SlideCount())
	}
}

func TestDefaultHTMLOptions(t *testing.T) {
	opts := DefaultHTMLOptions()

	if opts.TitleFontSize != 44 {
		t.Errorf("期望标题字号44，实际 %v", opts.TitleFontSize)
	}
	if opts.BodyFontSize != 18 {
		t.Errorf("期望正文字号18，实际 %v", opts.BodyFontSize)
	}
	if !strings.HasPrefix(opts.TitleColor, "#") {
		t.Errorf("期望颜色以#开头，实际 %v", opts.TitleColor)
	}
}

func TestFromHTML_ComplexDocument(t *testing.T) {
	html := `
	<!DOCTYPE html>
	<html>
	<head><title>测试</title></head>
	<body>
		<h1>演示文稿标题</h1>
		<p>副标题或介绍文字</p>
		
		<h1>第二页</h1>
		<h2>要点</h2>
		<ul>
			<li>第一点</li>
			<li>第二点</li>
		</ul>
		<ol>
			<li>步骤一</li>
			<li>步骤二</li>
		</ol>
		
		<h1>代码页</h1>
		<pre>
package main

import "fmt"

func main() {
    fmt.Println("Hello")
}
		</pre>
	</body>
	</html>
	`

	pres := FromHTML(html)

	if pres.SlideCount() != 3 {
		t.Errorf("期望3张幻灯片，实际得到 %d 张", pres.SlideCount())
	}

	// 验证可以生成字节
	_, err := pres.ToBytes()
	if err != nil {
		t.Errorf("ToBytes 失败: %v", err)
	}
}

func TestParseDataURI(t *testing.T) {
	// 1x1 红色 PNG 图片的 Base64
	base64PNG := "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=="

	data, ext := parseDataURI(base64PNG)

	if len(data) == 0 {
		t.Error("Base64 解码失败，数据为空")
	}
	if ext != "png" {
		t.Errorf("期望扩展名 png，实际 %s", ext)
	}
}

func TestParseDataURI_JPEG(t *testing.T) {
	// 简单的 JPEG data URI (minimal valid JPEG)
	base64JPEG := "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwC+KKKAP//Z"

	data, ext := parseDataURI(base64JPEG)

	if len(data) == 0 {
		t.Error("JPEG Base64 解码失败")
	}
	if ext != "jpg" {
		t.Errorf("期望扩展名 jpg，实际 %s", ext)
	}
}

func TestParseImageSrc_LocalFile(t *testing.T) {
	// 创建临时图片文件
	tmpFile, err := os.CreateTemp("", "test_*.png")
	if err != nil {
		t.Fatalf("创建临时文件失败: %v", err)
	}
	defer os.Remove(tmpFile.Name())

	// 写入简单的 PNG 数据
	pngData := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}
	tmpFile.Write(pngData)
	tmpFile.Close()

	data, ext := parseImageSrc(tmpFile.Name())

	if len(data) == 0 {
		t.Error("本地文件读取失败")
	}
	if ext != "png" {
		t.Errorf("期望扩展名 png，实际 %s", ext)
	}
}

func TestFromHTML_WithBase64Image(t *testing.T) {
	// 1x1 红色 PNG
	html := `
	<h1>Base64图片测试</h1>
	<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==" alt="红色像素">
	`

	pres := FromHTML(html)

	if pres.SlideCount() != 1 {
		t.Errorf("期望1张幻灯片，实际得到 %d 张", pres.SlideCount())
	}

	// 验证可以生成 PPT
	_, err := pres.ToBytes()
	if err != nil {
		t.Errorf("生成 PPT 失败: %v", err)
	}
}
