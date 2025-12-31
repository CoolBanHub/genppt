package genppt

import (
	"encoding/base64"
	"io"
	"net/http"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"

	"fmt"

	"golang.org/x/net/html"
)

// HTMLOptions HTML转换选项
type HTMLOptions struct {
	TitleFontSize   float64 // 标题字号，默认44
	HeadingFontSize float64 // 二级标题字号，默认32
	BodyFontSize    float64 // 正文字号，默认18
	CodeFontSize    float64 // 代码字号，默认14
	TitleColor      string  // 标题颜色
	HeadingColor    string  // 二级标题颜色
	BodyColor       string  // 正文颜色
	CodeBackground  string  // 代码背景色
	SlideBackground string  // 幻灯片背景色
	ImageRounding   float64 // 图片圆角（英寸），默认0
}

// DefaultHTMLOptions 返回默认HTML选项
func DefaultHTMLOptions() HTMLOptions {
	return HTMLOptions{
		TitleFontSize:   44,
		HeadingFontSize: 32,
		BodyFontSize:    18,
		CodeFontSize:    14,
		TitleColor:      "#1E3A5F",
		HeadingColor:    "#1E3A5F",
		BodyColor:       "#333333",
		CodeBackground:  "#F5F5F5",
		SlideBackground: "",
		ImageRounding:   0,
	}
}

// htmlSlide 表示解析后的HTML幻灯片
type htmlSlide struct {
	title           string
	titleColor      string // 标题颜色
	titleBackground string // 标题背景色
	titleAlign      string // 标题对齐方式
	backgroundColor string // 幻灯片背景色
	content         []htmlBlock
}

// htmlBlock 表示HTML内容块
type htmlBlock struct {
	blockType   string // "heading", "text", "bullet", "code", "image", "table"
	text        string
	level       int        // for heading (1-6)
	lines       []htmlLine // for bullet
	imageSrc    string
	imageAlt    string
	imageWidth  int
	imageHeight int
	imageLayout string     // "left", "right", "top"
	tableRows   [][]string // for table

	// 样式属性
	styleColor      string
	styleSize       int
	styleX          float64
	styleY          float64
	borderRadius    float64
	styleBackground string // 块级元素背景色
	styleAlign      string // 对齐方式 (left, center, right, justify)
}

// htmlLine 表示列表项
type htmlLine struct {
	text  string
	color string
}

// htmlParser HTML解析器
type htmlParser struct {
	slides                []htmlSlide
	current               *htmlSlide
	globalBackgroundColor string
}

// newHTMLParser 创建新的HTML解析器
func newHTMLParser() *htmlParser {
	return &htmlParser{
		slides: make([]htmlSlide, 0),
	}
}

// parse 解析HTML文本
func (p *htmlParser) parse(htmlContent string) []htmlSlide {
	doc, err := html.Parse(strings.NewReader(htmlContent))
	if err != nil {
		return p.slides
	}

	p.walkNode(doc)

	// 添加最后一个幻灯片
	if p.current != nil {
		p.slides = append(p.slides, *p.current)
	}

	return p.slides
}

// walkNode 遍历HTML节点
func (p *htmlParser) walkNode(n *html.Node) {
	if n.Type == html.ElementNode {
		switch n.Data {
		case "h1":
			// h1 创建新幻灯片
			title := p.extractText(n)
			styleMap := parseStyle(p.getAttr(n, "style"))
			color := p.getStyleOrAttr(n, styleMap, "color", "")
			bgColor := p.getStyleOrAttr(n, styleMap, "background-color", "")
			align := p.getAlign(n, styleMap)
			fmt.Fprintf(os.Stderr, "DEBUG: Parsing Title H1: Text='%s' Color='%s' Bg='%s' Align='%s'\n", title, color, bgColor, align)
			p.newSlideWithStyle(title, color, bgColor, align)
			return

		case "h2", "h3", "h4", "h5", "h6":
			level := int(n.Data[1] - '0')
			text := p.extractText(n)
			styleMap := parseStyle(p.getAttr(n, "style"))
			color := p.getStyleOrAttr(n, styleMap, "color", "")
			bgColor := p.getStyleOrAttr(n, styleMap, "background-color", "")
			align := p.getAlign(n, styleMap)
			fmt.Fprintf(os.Stderr, "DEBUG: Parsing Heading H%d: Text='%s' Color='%s' Bg='%s' Align='%s'\n", level, text, color, bgColor, align)

			size := int(p.getFontSize(n, styleMap))
			// Removed data-size fallback
			x := p.getStyleOrAttrFloat(n, styleMap, "left", "") // Removed data-x
			y := p.getStyleOrAttrFloat(n, styleMap, "top", "")  // Removed data-y
			p.addBlock(htmlBlock{
				blockType:       "heading",
				text:            text,
				level:           level,
				styleColor:      color,
				styleSize:       size,
				styleX:          x,
				styleY:          y,
				styleBackground: bgColor,
				styleAlign:      align,
			})
			return

		case "p":
			text := p.extractText(n)
			if strings.TrimSpace(text) != "" {
				styleMap := parseStyle(p.getAttr(n, "style"))
				size := int(p.getFontSize(n, styleMap))
				bgColor := p.getStyleOrAttr(n, styleMap, "background-color", "")
				align := p.getAlign(n, styleMap)
				// Removed data-size fallback
				p.addBlock(htmlBlock{
					blockType:       "text",
					text:            text,
					styleColor:      p.getStyleOrAttr(n, styleMap, "color", ""),
					styleSize:       size,
					styleX:          p.getStyleOrAttrFloat(n, styleMap, "left", ""),
					styleY:          p.getStyleOrAttrFloat(n, styleMap, "top", ""),
					styleBackground: bgColor,
					styleAlign:      align,
				})
			}
			return

		case "ul", "ol":
			items := p.extractListItems(n)
			if len(items) > 0 {
				p.addBlock(htmlBlock{
					blockType: "bullet",
					lines:     items,
				})
			}
			return

		case "pre":
			code := p.extractText(n)
			p.addBlock(htmlBlock{
				blockType: "code",
				text:      code,
			})
			return

		case "code":
			// 如果code不在pre内，作为行内代码处理，跳过
			if n.Parent == nil || n.Parent.Data != "pre" {
				break
			}

		case "img":
			src := p.getAttr(n, "src")
			alt := p.getAttr(n, "alt")
			styleMap := parseStyle(p.getAttr(n, "style"))

			width := p.getAttrInt(n, "width")
			// 优先使用 style width (单位 inch 转 px，因为 struct 也是 int px)
			// 注意：getStyleOrAttrFloat 返回 inch
			if wInch := p.getStyleOrAttrFloat(n, styleMap, "width", ""); wInch > 0 {
				width = int(wInch * 96)
			}

			height := p.getAttrInt(n, "height")
			if hInch := p.getStyleOrAttrFloat(n, styleMap, "height", ""); hInch > 0 {
				height = int(hInch * 96)
			}

			// Removed data-layout, use style float or similar if needed.
			// Layout logic already checks styleMap["float"]
			layout := "" // Default empty
			// 支持 CSS float
			if f := styleMap["float"]; f != "" {
				if f == "left" {
					layout = "left"
				}
				if f == "right" {
					layout = "right"
				}
			}

			x := p.getStyleOrAttrFloat(n, styleMap, "left", "") // Removed data-x
			y := p.getStyleOrAttrFloat(n, styleMap, "top", "")  // Removed data-y
			borderRadius := p.getStyleOrAttrFloat(n, styleMap, "border-radius", "")

			if src != "" {
				p.addBlock(htmlBlock{
					blockType:    "image",
					imageSrc:     src,
					imageAlt:     alt,
					imageWidth:   width,
					imageHeight:  height,
					imageLayout:  layout,
					styleX:       x,
					styleY:       y,
					borderRadius: borderRadius,
				})
			}
			return

		case "table":
			rows := p.extractTableRows(n)
			if len(rows) > 0 {
				styleMap := parseStyle(p.getAttr(n, "style"))
				x := p.getStyleOrAttrFloat(n, styleMap, "left", "") // Removed data-x
				y := p.getStyleOrAttrFloat(n, styleMap, "top", "")  // Removed data-y
				p.addBlock(htmlBlock{
					blockType: "table",
					tableRows: rows,
					styleX:    x,
					styleY:    y,
				})
			}
			return

		case "style":
			// 解析 <style> 标签中的 body background-color
			// 简单正则匹配: body { ... background-color: #xxx; ... }
			content := p.extractText(n)
			// Normalize generic whitespace
			content = strings.ReplaceAll(content, "\n", " ")
			// Regex for body selector and block
			bodyStyleRegex := regexp.MustCompile(`body\s*\{([^}]+)\}`)
			if match := bodyStyleRegex.FindStringSubmatch(content); len(match) > 1 {
				styleBody := match[1]
				styleMap := parseStyle(styleBody) // Reuse parseStyle logic
				if bg := styleMap["background-color"]; bg != "" {
					p.globalBackgroundColor = bg
				} else if bg := styleMap["background"]; bg != "" {
					p.globalBackgroundColor = extractHexColor(bg)
				}

				if p.globalBackgroundColor != "" {
					fmt.Fprintf(os.Stderr, "DEBUG: Found Global Background in Style: %s\n", p.globalBackgroundColor)
					// 如果当前有正在处理的slide（虽然通常style在head里，但防止万一），更新它
					if p.current != nil && p.current.backgroundColor == "" {
						p.current.backgroundColor = p.globalBackgroundColor
					}
				}
			}
			return

		case "body":
			// 解析 <body style="background-color: ..."> 或 bgcolor
			styleMap := parseStyle(p.getAttr(n, "style"))
			bg := styleMap["background-color"]
			if bg == "" {
				bg = p.getAttr(n, "bgcolor")
			}
			if bg == "" {
				bg = styleMap["background"]
			}
			// 尝试提取 Hex
			bg = extractHexColor(bg)

			if bg != "" {
				p.globalBackgroundColor = bg
				fmt.Fprintf(os.Stderr, "DEBUG: Found Global Background in Body: %s\n", bg)
				if p.current != nil && p.current.backgroundColor == "" {
					p.current.backgroundColor = bg
				}
			}
			// Don't return, recurse children

		case "div":
			// 解析 div 背景
			styleMap := parseStyle(p.getAttr(n, "style"))
			bg := styleMap["background-color"]
			if bg == "" {
				bg = styleMap["background"]
			}
			bg = extractHexColor(bg)

			if bg != "" {
				p.globalBackgroundColor = bg
				fmt.Fprintf(os.Stderr, "DEBUG: Found Global Background in DIV: %s\n", bg)
				if p.current != nil && p.current.backgroundColor == "" {
					p.current.backgroundColor = bg
				}
			}
			// 继续处理子节点
			for c := n.FirstChild; c != nil; c = c.NextSibling {
				p.walkNode(c)
			}
			return

		case "hr", "section":
			// 分隔符或section创建新幻灯片
			if p.current != nil {
				p.slides = append(p.slides, *p.current)
				p.current = nil
			}
			if n.Data == "section" {
				// 继续处理section内的内容
				for c := n.FirstChild; c != nil; c = c.NextSibling {
					p.walkNode(c)
				}
				// section结束后保存
				if p.current != nil {
					p.slides = append(p.slides, *p.current)
					p.current = nil
				}
			}
			return
		}
	}

	// 递归处理子节点
	for c := n.FirstChild; c != nil; c = c.NextSibling {
		p.walkNode(c)
	}
}

// newSlide 创建新幻灯片
func (p *htmlParser) newSlide(title string) {
	p.newSlideWithStyle(title, "", "", "")
}

// newSlideWithStyle 创建新幻灯片（带样式）
func (p *htmlParser) newSlideWithStyle(title string, color, bgColor, align string) {
	if p.current != nil {
		p.slides = append(p.slides, *p.current)
	}
	p.current = &htmlSlide{
		title:           title,
		titleColor:      color,
		titleBackground: bgColor,
		titleAlign:      align,
		content:         make([]htmlBlock, 0),
		backgroundColor: p.globalBackgroundColor,
	}
}

// addBlock 添加内容块
func (p *htmlParser) addBlock(block htmlBlock) {
	if p.current == nil {
		p.newSlide("")
	}
	p.current.content = append(p.current.content, block)
}

// extractText 提取节点纯文本
func (p *htmlParser) extractText(n *html.Node) string {
	var sb strings.Builder
	p.extractTextRecursive(n, &sb)
	return strings.TrimSpace(sb.String())
}

// extractTextRecursive 递归提取文本
func (p *htmlParser) extractTextRecursive(n *html.Node, sb *strings.Builder) {
	if n.Type == html.TextNode {
		sb.WriteString(n.Data)
	}
	for c := n.FirstChild; c != nil; c = c.NextSibling {
		p.extractTextRecursive(c, sb)
	}
}

// extractListItems 提取列表项
func (p *htmlParser) extractListItems(n *html.Node) []htmlLine {
	var items []htmlLine
	for c := n.FirstChild; c != nil; c = c.NextSibling {
		if c.Type == html.ElementNode && c.Data == "li" {
			text := p.extractText(c)
			styleMap := parseStyle(p.getAttr(c, "style"))
			color := p.getStyleOrAttr(c, styleMap, "color", "") // Removed data-color
			if text != "" {
				fmt.Fprintf(os.Stderr, "DEBUG: Parsing LI: Text='%s' Color='%s'\n", text, color)
				items = append(items, htmlLine{
					text:  text,
					color: color,
				})
			}
		}
	}
	return items
}

// extractTableRows 提取表格数据
func (p *htmlParser) extractTableRows(n *html.Node) [][]string {
	var rows [][]string

	var walkTable func(*html.Node)
	walkTable = func(node *html.Node) {
		if node.Type == html.ElementNode && node.Data == "tr" {
			var row []string
			for c := node.FirstChild; c != nil; c = c.NextSibling {
				if c.Type == html.ElementNode && (c.Data == "td" || c.Data == "th") {
					row = append(row, p.extractText(c))
				}
			}
			if len(row) > 0 {
				rows = append(rows, row)
			}
			return
		}
		for c := node.FirstChild; c != nil; c = c.NextSibling {
			walkTable(c)
		}
	}

	walkTable(n)
	return rows
}

// getAttr 获取属性值
func (p *htmlParser) getAttr(n *html.Node, key string) string {
	for _, attr := range n.Attr {
		if attr.Key == key {
			return attr.Val
		}
	}
	return ""
}

// getAttrInt 获取整数属性值
func (p *htmlParser) getAttrInt(n *html.Node, key string) int {
	val := p.getAttr(n, key)
	if val == "" {
		return 0
	}
	var result int
	for _, c := range val {
		if c >= '0' && c <= '9' {
			result = result*10 + int(c-'0')
		} else {
			break
		}
	}
	return result
}

// getAttrFloat 获取浮点数属性值
func (p *htmlParser) getAttrFloat(n *html.Node, key string) float64 {
	val := p.getAttr(n, key)
	if val == "" {
		return 0
	}
	var result float64
	var decimal float64 = 0
	var divisor float64 = 1
	afterDot := false
	for _, c := range val {
		if c == '.' {
			afterDot = true
			continue
		}
		if c >= '0' && c <= '9' {
			if afterDot {
				divisor *= 10
				decimal += float64(c-'0') / divisor
			} else {
				result = result*10 + float64(c-'0')
			}
		} else {
			break
		}
	}
	return result + decimal
}

// parseImageSrc 解析图片来源，支持本地路径、URL和Base64
// 返回图片数据和扩展名
func parseImageSrc(src string) (data []byte, ext string) {
	// 处理 Base64 data URI
	if strings.HasPrefix(src, "data:") {
		return parseDataURI(src)
	}

	// 处理 HTTP/HTTPS URL
	if strings.HasPrefix(src, "http://") || strings.HasPrefix(src, "https://") {
		return downloadImage(src)
	}

	// 本地文件路径
	fileData, err := os.ReadFile(src)
	if err != nil {
		return nil, ""
	}
	return fileData, getExtFromPath(src)
}

// parseDataURI 解析 data URI 格式的图片
// 格式: data:image/png;base64,iVBORw0KGgo...
func parseDataURI(dataURI string) (data []byte, ext string) {
	// 去掉 "data:" 前缀
	content := strings.TrimPrefix(dataURI, "data:")

	// 查找分号和逗号
	semicolonIdx := strings.Index(content, ";")
	commaIdx := strings.Index(content, ",")

	if commaIdx == -1 {
		return nil, ""
	}

	// 提取 MIME 类型
	mimeType := ""
	if semicolonIdx > 0 && semicolonIdx < commaIdx {
		mimeType = content[:semicolonIdx]
	} else {
		mimeType = content[:commaIdx]
	}

	// 根据 MIME 类型确定扩展名
	switch mimeType {
	case "image/png":
		ext = "png"
	case "image/jpeg", "image/jpg":
		ext = "jpg"
	case "image/gif":
		ext = "gif"
	case "image/webp":
		ext = "webp"
	case "image/svg+xml":
		ext = "svg"
	case "image/bmp":
		ext = "bmp"
	default:
		ext = "png" // 默认当作 PNG
	}

	// 提取 Base64 数据
	base64Data := content[commaIdx+1:]

	// 解码 Base64
	decoded, err := base64.StdEncoding.DecodeString(base64Data)
	if err != nil {
		// 尝试 URL 安全的 Base64
		decoded, err = base64.URLEncoding.DecodeString(base64Data)
		if err != nil {
			// 尝试带 padding 校正
			decoded, err = base64.RawStdEncoding.DecodeString(base64Data)
			if err != nil {
				return nil, ""
			}
		}
	}

	return decoded, ext
}

// downloadImage 从 URL 下载图片
func downloadImage(url string) (data []byte, ext string) {
	fmt.Fprintf(os.Stderr, "DEBUG: Downloading image: %s\n", url)
	client := &http.Client{
		Timeout: 30 * time.Second,
	}

	req, err := http.NewRequest("GET", url, nil)
	if err != nil {
		fmt.Fprintf(os.Stderr, "DEBUG: NewRequest failed: %v\n", err)
		return nil, ""
	}
	// Add User-Agent to avoid some bot blocking
	req.Header.Set("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

	resp, err := client.Do(req)
	if err != nil {
		fmt.Fprintf(os.Stderr, "DEBUG: Download failed: %v\n", err)
		return nil, ""
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		fmt.Fprintf(os.Stderr, "DEBUG: Download status error: %d for %s\n", resp.StatusCode, url)
		return nil, ""
	}

	data, err = io.ReadAll(resp.Body)
	if err != nil {
		fmt.Fprintf(os.Stderr, "DEBUG: Read body failed: %v\n", err)
		return nil, ""
	}

	// 从 URL 或 Content-Type 推断扩展名
	ext = getExtFromPath(url)
	if ext == "" {
		contentType := resp.Header.Get("Content-Type")
		switch {
		case strings.Contains(contentType, "png"):
			ext = "png"
		case strings.Contains(contentType, "jpeg"), strings.Contains(contentType, "jpg"):
			ext = "jpg"
		case strings.Contains(contentType, "gif"):
			ext = "gif"
		case strings.Contains(contentType, "webp"):
			ext = "webp"
		default:
			ext = "png"
		}
	}
	fmt.Fprintf(os.Stderr, "DEBUG: Download success: %d bytes, ext: %s\n", len(data), ext)
	return data, ext
}

// FromHTML 从HTML字符串创建演示文稿
func FromHTML(htmlContent string) *Presentation {

	return FromHTMLWithOptions(htmlContent, DefaultHTMLOptions())
}

// FromHTMLWithOptions 从HTML字符串创建演示文稿（带选项）
func FromHTMLWithOptions(htmlContent string, opts HTMLOptions) *Presentation {
	parser := newHTMLParser()
	slides := parser.parse(htmlContent)

	pres := New()

	for _, htmlSlide := range slides {
		slide := pres.AddSlide()

		// 设置背景
		// 优先使用 Slide 自身的背景色（来自 body/style 解析），其次用 Options
		bg := opts.SlideBackground
		if htmlSlide.backgroundColor != "" {
			bg = htmlSlide.backgroundColor
		}

		if bg != "" {
			slide.SetBackground(BackgroundOptions{
				Color: bg,
			})
		}

		state := &renderState{
			yPos: 0.5,
		}

		// 辅助函数：新起一页
		newSlide := func() {
			slide = pres.AddSlide()
			// 同样应用背景
			bg := opts.SlideBackground
			if htmlSlide.backgroundColor != "" {
				bg = htmlSlide.backgroundColor
			}

			if bg != "" {
				slide.SetBackground(BackgroundOptions{
					Color: bg,
				})
			}
			state.yPos = 0.5
			state.hasSideImg = false
			state.sideImgY = 0
			state.sideImgH = 0
		}

		// 布局状态管理
		// var (
		// 	hasSideImg bool    // 是否有侧边图片
		// 	sideImgPos string  // "left" 或 "right"
		// 	sideImgY   float64 // 侧边图片开始Y
		// 	sideImgH   float64 // 侧边图片高度
		// )

		// 添加标题
		// 添加标题
		if htmlSlide.title != "" {
			titleColor := opts.HeadingColor
			if htmlSlide.titleColor != "" {
				titleColor = htmlSlide.titleColor
			}
			slide.AddText(htmlSlide.title, TextOptions{
				X:         0.5,
				Y:         state.yPos,
				Width:     9.0,
				Height:    0.8,
				FontSize:  opts.HeadingFontSize,
				FontColor: titleColor,
				Bold:      true,
				Fill:      htmlSlide.titleBackground,
				Align:     Align(htmlSlide.titleAlign),
			})
			state.yPos += 1.0
		}

		// Layout Optimization: Hoist floating images to the top
		// This ensures that the image is rendered first, setting the 'hasSideImg' state,
		// so that subsequent text blocks are correctly width-constrained.
		var floats []htmlBlock
		var flow []htmlBlock
		hasFloats := false
		for _, block := range htmlSlide.content {
			if block.blockType == "image" && (block.imageLayout == "left" || block.imageLayout == "right") {
				floats = append(floats, block)
				hasFloats = true
			} else {
				flow = append(flow, block)
			}
		}
		if hasFloats {
			// Reconstruct content with floats first
			htmlSlide.content = append(floats, flow...)
		}

		// Two-Pass Strategy
		// 1. Dry Run / Measure
		// Two-Pass Strategy
		// 1. Dry Run / Measure
		scale := 1.0

		// Measure total height
		// Use a COPY of state or temp state initialized with current yPos
		dryRunState := &renderState{
			yPos:       state.yPos,
			hasSideImg: state.hasSideImg, // Should be false initially usually
			sideImgY:   state.sideImgY,
			sideImgH:   state.sideImgH,
		}
		newSlideNoOp := func() {} // Do nothing on dry run split (won't trigger if disabled anyway)

		for _, block := range htmlSlide.content {
			renderBlock(pres, slide, block, opts, dryRunState, 1.0, true, newSlideNoOp)
		}

		totalHeight := dryRunState.yPos
		// Debug
		fmt.Fprintf(os.Stderr, "DEBUG: Slide '%s' Total Height=%.2f\n", htmlSlide.title, totalHeight)

		if totalHeight > 5.6 {
			// Calculate scale
			// Target is ~5.5 to leave some margin
			scale = 5.5 / totalHeight
			fmt.Fprintf(os.Stderr, "DEBUG: Auto-Scaling trigger! Scale=%.2f\n", scale)
		}

		// 2. Render Pass
		// Reset state for actual render
		// If title was added, yPos should be updated?
		// Wait, title logic (lines 716+) happens BEFORE this block.
		// If we reset yPos here, we overwrite the title space or draw over it?
		// The original code reset yPos to 0.5 INSIDE the content loop?
		// No, original code:
		// newSlide sets yPos=0.5
		// AddTitle sets yPos+=1.0 => yPos=1.5
		// Then content loop starts with yPos=1.5

		// My Two-Pass logic in Step 121:
		// 1. Dry Run (starts with yPos=0.5, simulates title?)
		// Wait, the dry run loop `for _, block := range htmlSlide.content` checks CONTENT height.
		// It assumes yPos starts where title left off?
		// Yes, I need to init dry run state with current state.yPos (which includes Title space).
		// AND for Render Pass, I should NOT reset to 0.5, but keep the current state.yPos!

		// Fixing logic:

		// Update helper to use state
		newSlideWrapper := func() {
			newSlide()
			// Sync state back after new slide
			state.yPos = 0.5
			state.hasSideImg = false
			state.sideImgY = 0
			state.sideImgH = 0
		}

		for _, block := range htmlSlide.content {
			renderBlock(pres, slide, block, opts, state, scale, false, newSlideWrapper)
		}

		// Create new slides if needed (handled by newSlideWrapper inside renderBlock if not disabled)
	}

	return pres
}

// FromHTMLFile 从HTML文件创建演示文稿
func FromHTMLFile(filename string) (*Presentation, error) {
	return FromHTMLFileWithOptions(filename, DefaultHTMLOptions())
}

// FromHTMLFileWithOptions 从HTML文件创建演示文稿（带选项）
func FromHTMLFileWithOptions(filename string, opts HTMLOptions) (*Presentation, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	data, err := io.ReadAll(file)
	if err != nil {
		return nil, err
	}

	return FromHTMLWithOptions(string(data), opts), nil
}

// estimateLines 估算文本行数和高度
func estimateLines(text string, fontSize float64, width float64) (float64, int) {
	if width <= 0 {
		return 0.5, 1
	}

	// 转换字号由 point 到 inch
	// 1 pt = 1/72 inch = 0.0138 inch
	// 稍微保守一点，假设中文占 1.1 倍字号宽度
	charW_CN := (fontSize / 72.0) * 1.1
	charW_EN := (fontSize / 72.0) * 0.6

	totalLen := 0.0
	for _, r := range text {
		if r > 127 {
			totalLen += charW_CN
		} else {
			totalLen += charW_EN
		}
	}

	lines := int(totalLen/width) + 1

	// 行高：字号 * 1.4 (行间距)
	lineHeight := (fontSize / 72.0) * 1.4
	totalHeight := float64(lines) * lineHeight

	// 至少 0.4 inch 高度 (单行)
	if totalHeight < 0.4 {
		totalHeight = 0.4
	}

	return totalHeight, lines
}

// parseStyle 解析 CSS 样式字符串
func parseStyle(style string) map[string]string {
	m := make(map[string]string)
	parts := strings.Split(style, ";")
	for _, part := range parts {
		kv := strings.SplitN(part, ":", 2)
		if len(kv) == 2 {
			k := strings.ToLower(strings.TrimSpace(kv[0]))
			v := strings.TrimSpace(kv[1])
			m[k] = v
		}
	}
	return m
}

// getStyleOrAttr 获取样式或属性值 (优先样式)
func (p *htmlParser) getStyleOrAttr(n *html.Node, styleMap map[string]string, styleKey, attrKey string) string {
	if v, ok := styleMap[styleKey]; ok && v != "" {
		return v
	}
	// 尝试非 data- 属性 (兼容 width/height/src/alt 等标准属性)
	if !strings.HasPrefix(attrKey, "data-") && attrKey != "" {
		return p.getAttr(n, attrKey)
	}
	return ""
}

// getStyleOrAttrFloat 获取浮点数值（支持 px, pt, in 单位转换）
// 默认单位：英寸 (为了兼容 data-x/y)
// 如果是 font-size，通常当作 point 处理
func (p *htmlParser) getStyleOrAttrFloat(n *html.Node, styleMap map[string]string, styleKey, attrKey string) float64 {
	valStr := p.getStyleOrAttr(n, styleMap, styleKey, attrKey)
	if valStr == "" {
		return 0
	}

	valStr = strings.ToLower(valStr)
	scale := 1.0

	// 解析单位
	if strings.HasSuffix(valStr, "px") {
		valStr = strings.TrimSuffix(valStr, "px")
		// 96 px = 1 inch
		scale = 1.0 / 96.0
	} else if strings.HasSuffix(valStr, "pt") {
		valStr = strings.TrimSuffix(valStr, "pt")
		// 72 pt = 1 inch
		scale = 1.0 / 72.0
	} else if strings.HasSuffix(valStr, "in") {
		valStr = strings.TrimSuffix(valStr, "in")
		scale = 1.0
	} else if strings.HasSuffix(valStr, "%") {
		valStr = strings.TrimSuffix(valStr, "%")
		f, err := strconv.ParseFloat(valStr, 64)
		if err != nil {
			return 0
		}
		// 返回负值表示百分比
		return -f / 100.0
	}

	f, err := strconv.ParseFloat(valStr, 64)
	if err != nil {
		return 0
	}

	// 如果是 layout 相关的 (left, top, width, height)，我们需要返回 inch 单位
	if styleKey == "left" || styleKey == "top" || styleKey == "width" || styleKey == "height" {
		return f * scale
	}

	return f * scale
}

// getAlign 解析对齐方式
func (p *htmlParser) getAlign(n *html.Node, styleMap map[string]string) string {
	valStr := p.getStyleOrAttr(n, styleMap, "text-align", "align")
	switch strings.ToLower(valStr) {
	case "left":
		return "l"
	case "center":
		return "ctr"
	case "right":
		return "r"
	case "justify":
		return "just"
	}
	return ""
}

// getFontSize 解析字号 (返回 Point)
func (p *htmlParser) getFontSize(n *html.Node, styleMap map[string]string) float64 {
	valStr := p.getStyleOrAttr(n, styleMap, "font-size", "")
	if valStr == "" {
		return 0
	}

	valStr = strings.ToLower(valStr)
	if strings.HasSuffix(valStr, "px") {
		f, _ := strconv.ParseFloat(strings.TrimSuffix(valStr, "px"), 64)
		// 1 px = 0.75 pt
		return f * 0.75
	}
	if strings.HasSuffix(valStr, "pt") {
		f, _ := strconv.ParseFloat(strings.TrimSuffix(valStr, "pt"), 64)
		return f
	}
	f, _ := strconv.ParseFloat(valStr, 64)
	return f
}

// extractHexColor 从字符串中提取 Hex 颜色 (例如从 gradient 中提取第一个颜色)
func extractHexColor(s string) string {
	if s == "" {
		return ""
	}
	s = strings.TrimSpace(s)
	// 如果已经是 #xxxxxx 格式
	if strings.HasPrefix(s, "#") {
		// 可能是简单的 hex color
		return s
	}

	// 使用正则查找 #xxxxxx 或 #xxx
	// 优先找 6位 hex
	hexRegex6 := regexp.MustCompile(`#([0-9a-fA-F]{6})`)
	if match := hexRegex6.FindString(s); match != "" {
		return match
	}
	// 其次找 3位 hex
	hexRegex3 := regexp.MustCompile(`#([0-9a-fA-F]{3})`)
	if match := hexRegex3.FindString(s); match != "" {
		return match
	}

	// 如果没有 hex，也许是 rgb/rgba? 暂时不支持解析 rgb 字符串到 hex
	// 如果本来就是 blue/red 这种名称，直接返回
	if !strings.Contains(s, "(") && !strings.Contains(s, " ") {
		return s
	}

	return ""
}
