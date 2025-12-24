package genppt

import (
	"bufio"
	"io"
	"os"
	"regexp"
	"strings"
)

// MarkdownOptions Markdown转换选项
type MarkdownOptions struct {
	TitleFontSize   float64 // 标题字号，默认44
	HeadingFontSize float64 // 二级标题字号，默认32
	BodyFontSize    float64 // 正文字号，默认18
	CodeFontSize    float64 // 代码字号，默认14
	TitleColor      string  // 标题颜色
	HeadingColor    string  // 二级标题颜色
	BodyColor       string  // 正文颜色
	CodeBackground  string  // 代码背景色
	SlideBackground string  // 幻灯片背景色
}

// DefaultMarkdownOptions 返回默认Markdown选项
func DefaultMarkdownOptions() MarkdownOptions {
	return MarkdownOptions{
		TitleFontSize:   44,
		HeadingFontSize: 32,
		BodyFontSize:    18,
		CodeFontSize:    14,
		TitleColor:      "#1E3A5F",
		HeadingColor:    "#1E3A5F",
		BodyColor:       "#333333",
		CodeBackground:  "#F5F5F5",
		SlideBackground: "",
	}
}

// markdownSlide 表示解析后的Markdown幻灯片
type markdownSlide struct {
	title    string
	subtitle string
	content  []markdownBlock
}

// markdownBlock 表示Markdown内容块
type markdownBlock struct {
	blockType string   // "text", "bullet", "code", "heading", "image"
	lines     []string // 内容行
	language  string   // 代码块语言（如果是代码块）
	imagePath string   // 图片路径（如果是图片）
	imageAlt  string   // 图片替代文本（如果是图片）
}

// markdownParser Markdown解析器
type markdownParser struct {
	slides    []markdownSlide
	current   *markdownSlide
	inCode    bool
	codeLang  string
	codeLines []string
}

// newMarkdownParser 创建新的Markdown解析器
func newMarkdownParser() *markdownParser {
	return &markdownParser{
		slides: make([]markdownSlide, 0),
	}
}

// parse 解析Markdown文本
func (p *markdownParser) parse(text string) []markdownSlide {
	scanner := bufio.NewScanner(strings.NewReader(text))

	for scanner.Scan() {
		line := scanner.Text()
		p.parseLine(line)
	}

	// 结束代码块（如果有未关闭的）
	if p.inCode {
		p.endCodeBlock()
	}

	// 添加最后一个幻灯片
	if p.current != nil {
		p.slides = append(p.slides, *p.current)
	}

	return p.slides
}

// parseLine 解析单行
func (p *markdownParser) parseLine(line string) {
	// 处理代码块
	if strings.HasPrefix(line, "```") {
		if p.inCode {
			p.endCodeBlock()
		} else {
			p.inCode = true
			p.codeLang = strings.TrimPrefix(line, "```")
			p.codeLines = make([]string, 0)
		}
		return
	}

	if p.inCode {
		p.codeLines = append(p.codeLines, line)
		return
	}

	// 一级标题 - 新幻灯片
	if strings.HasPrefix(line, "# ") {
		p.newSlide(strings.TrimPrefix(line, "# "))
		return
	}

	// 二级标题 - 幻灯片标题或副标题
	if strings.HasPrefix(line, "## ") {
		heading := strings.TrimPrefix(line, "## ")
		if p.current == nil {
			p.newSlide(heading)
		} else if p.current.title == "" {
			p.current.title = heading
		} else {
			// 作为内容标题
			p.addBlock(markdownBlock{
				blockType: "heading",
				lines:     []string{heading},
			})
		}
		return
	}

	// 三级及以下标题 - 作为内容
	if strings.HasPrefix(line, "### ") || strings.HasPrefix(line, "#### ") {
		heading := strings.TrimLeft(line, "#")
		heading = strings.TrimSpace(heading)
		p.addBlock(markdownBlock{
			blockType: "heading",
			lines:     []string{heading},
		})
		return
	}

	// 无序列表
	if strings.HasPrefix(line, "- ") || strings.HasPrefix(line, "* ") || strings.HasPrefix(line, "+ ") {
		text := strings.TrimLeft(line, "-*+ ")
		p.addBullet(text)
		return
	}

	// 有序列表
	if matched, _ := regexp.MatchString(`^\d+\.\s`, line); matched {
		re := regexp.MustCompile(`^\d+\.\s+`)
		text := re.ReplaceAllString(line, "")
		p.addBullet(text)
		return
	}

	// 图片 ![alt](path)
	imageRe := regexp.MustCompile(`^!\[([^\]]*)\]\(([^)]+)\)$`)
	if matches := imageRe.FindStringSubmatch(strings.TrimSpace(line)); len(matches) == 3 {
		p.addBlock(markdownBlock{
			blockType: "image",
			imageAlt:  matches[1],
			imagePath: matches[2],
		})
		return
	}

	// 分隔线 - 新幻灯片
	if line == "---" || line == "***" || line == "___" {
		if p.current != nil {
			p.slides = append(p.slides, *p.current)
			p.current = nil
		}
		return
	}

	// 空行
	if strings.TrimSpace(line) == "" {
		return
	}

	// 普通文本
	p.addText(line)
}

// newSlide 创建新幻灯片
func (p *markdownParser) newSlide(title string) {
	if p.current != nil {
		p.slides = append(p.slides, *p.current)
	}
	p.current = &markdownSlide{
		title:   title,
		content: make([]markdownBlock, 0),
	}
}

// addBlock 添加内容块
func (p *markdownParser) addBlock(block markdownBlock) {
	if p.current == nil {
		p.newSlide("")
	}
	p.current.content = append(p.current.content, block)
}

// addBullet 添加列表项
func (p *markdownParser) addBullet(text string) {
	if p.current == nil {
		p.newSlide("")
	}

	// 检查是否可以合并到上一个bullet块
	if len(p.current.content) > 0 {
		last := &p.current.content[len(p.current.content)-1]
		if last.blockType == "bullet" {
			last.lines = append(last.lines, text)
			return
		}
	}

	p.addBlock(markdownBlock{
		blockType: "bullet",
		lines:     []string{text},
	})
}

// addText 添加普通文本
func (p *markdownParser) addText(text string) {
	if p.current == nil {
		p.newSlide("")
	}

	// 检查是否可以合并到上一个text块
	if len(p.current.content) > 0 {
		last := &p.current.content[len(p.current.content)-1]
		if last.blockType == "text" {
			last.lines = append(last.lines, text)
			return
		}
	}

	p.addBlock(markdownBlock{
		blockType: "text",
		lines:     []string{text},
	})
}

// endCodeBlock 结束代码块
func (p *markdownParser) endCodeBlock() {
	p.inCode = false
	if len(p.codeLines) > 0 {
		p.addBlock(markdownBlock{
			blockType: "code",
			lines:     p.codeLines,
			language:  p.codeLang,
		})
	}
	p.codeLines = nil
	p.codeLang = ""
}

// FromMarkdown 从Markdown字符串创建演示文稿
func FromMarkdown(markdown string) *Presentation {
	return FromMarkdownWithOptions(markdown, DefaultMarkdownOptions())
}

// FromMarkdownWithOptions 从Markdown字符串创建演示文稿（带选项）
func FromMarkdownWithOptions(markdown string, opts MarkdownOptions) *Presentation {
	parser := newMarkdownParser()
	slides := parser.parse(markdown)

	pres := New()

	for _, mdSlide := range slides {
		slide := pres.AddSlide()

		// 设置背景
		if opts.SlideBackground != "" {
			slide.SetBackground(BackgroundOptions{
				Color: opts.SlideBackground,
			})
		}

		yPos := 0.5 // 当前Y位置

		// 添加标题
		if mdSlide.title != "" {
			slide.AddText(mdSlide.title, TextOptions{
				X:         0.5,
				Y:         yPos,
				Width:     9.0,
				Height:    0.8,
				FontSize:  opts.HeadingFontSize,
				FontColor: opts.HeadingColor,
				Bold:      true,
			})
			yPos += 1.0
		}

		// 添加内容
		for _, block := range mdSlide.content {
			switch block.blockType {
			case "heading":
				for _, line := range block.lines {
					slide.AddText(line, TextOptions{
						X:         0.5,
						Y:         yPos,
						Width:     9.0,
						Height:    0.6,
						FontSize:  24,
						FontColor: opts.HeadingColor,
						Bold:      true,
					})
					yPos += 0.7
				}

			case "text":
				text := parseInlineMarkdown(strings.Join(block.lines, "\n"))
				slide.AddText(text, TextOptions{
					X:         0.5,
					Y:         yPos,
					Width:     9.0,
					Height:    0.5,
					FontSize:  opts.BodyFontSize,
					FontColor: opts.BodyColor,
				})
				yPos += 0.6

			case "bullet":
				for _, line := range block.lines {
					text := "• " + parseInlineMarkdown(line)
					slide.AddText(text, TextOptions{
						X:         0.7,
						Y:         yPos,
						Width:     8.5,
						Height:    0.4,
						FontSize:  opts.BodyFontSize,
						FontColor: opts.BodyColor,
					})
					yPos += 0.5
				}

			case "code":
				codeText := strings.Join(block.lines, "\n")
				// 代码块用形状作为背景
				codeHeight := float64(len(block.lines)) * 0.35
				if codeHeight < 0.5 {
					codeHeight = 0.5
				}
				if codeHeight > 3.5 {
					codeHeight = 3.5
				}

				slide.AddShape(ShapeRect, ShapeOptions{
					X:         0.5,
					Y:         yPos,
					Width:     9.0,
					Height:    codeHeight + 0.2,
					Fill:      opts.CodeBackground,
					LineColor: "#CCCCCC",
					LineWidth: 1,
				})

				slide.AddText(codeText, TextOptions{
					X:         0.6,
					Y:         yPos + 0.1,
					Width:     8.8,
					Height:    codeHeight,
					FontSize:  opts.CodeFontSize,
					FontFace:  "Consolas",
					FontColor: "#333333",
				})
				yPos += codeHeight + 0.4

			case "image":
				// 添加图片
				slide.AddImage(ImageOptions{
					Path:    block.imagePath,
					X:       1.0,
					Y:       yPos,
					Width:   8.0,
					Height:  3.0,
					AltText: block.imageAlt,
				})
				yPos += 3.2
			}
		}
	}

	return pres
}

// FromMarkdownFile 从Markdown文件创建演示文稿
func FromMarkdownFile(filename string) (*Presentation, error) {
	return FromMarkdownFileWithOptions(filename, DefaultMarkdownOptions())
}

// FromMarkdownFileWithOptions 从Markdown文件创建演示文稿（带选项）
func FromMarkdownFileWithOptions(filename string, opts MarkdownOptions) (*Presentation, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	data, err := io.ReadAll(file)
	if err != nil {
		return nil, err
	}

	return FromMarkdownWithOptions(string(data), opts), nil
}

// parseInlineMarkdown 解析行内Markdown格式
func parseInlineMarkdown(text string) string {
	// 移除加粗标记 **text** 或 __text__
	boldRe := regexp.MustCompile(`\*\*(.+?)\*\*|__(.+?)__`)
	text = boldRe.ReplaceAllString(text, "$1$2")

	// 移除斜体标记 *text* 或 _text_
	italicRe := regexp.MustCompile(`\*(.+?)\*|_(.+?)_`)
	text = italicRe.ReplaceAllString(text, "$1$2")

	// 移除行内代码标记 `code`
	codeRe := regexp.MustCompile("`(.+?)`")
	text = codeRe.ReplaceAllString(text, "$1")

	// 提取链接文本 [text](url)
	linkRe := regexp.MustCompile(`\[(.+?)\]\(.+?\)`)
	text = linkRe.ReplaceAllString(text, "$1")

	return text
}
