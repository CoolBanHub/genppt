package genppt

import (
	"encoding/base64"
	"io"
	"net/http"
	"os"
	"strings"
	"time"

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
	}
}

// htmlSlide 表示解析后的HTML幻灯片
type htmlSlide struct {
	title   string
	content []htmlBlock
}

// htmlBlock 表示HTML内容块
type htmlBlock struct {
	blockType string     // "text", "bullet", "code", "heading", "image", "table"
	text      string     // 文本内容
	lines     []string   // 列表项内容
	imageSrc  string     // 图片路径
	imageAlt  string     // 图片替代文本
	level     int        // 标题级别 (2-6)
	tableRows [][]string // 表格数据
}

// htmlParser HTML解析器
type htmlParser struct {
	slides  []htmlSlide
	current *htmlSlide
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
			p.newSlide(title)
			return

		case "h2", "h3", "h4", "h5", "h6":
			level := int(n.Data[1] - '0')
			text := p.extractText(n)
			p.addBlock(htmlBlock{
				blockType: "heading",
				text:      text,
				level:     level,
			})
			return

		case "p":
			text := p.extractText(n)
			if strings.TrimSpace(text) != "" {
				p.addBlock(htmlBlock{
					blockType: "text",
					text:      text,
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
			if src != "" {
				p.addBlock(htmlBlock{
					blockType: "image",
					imageSrc:  src,
					imageAlt:  alt,
				})
			}
			return

		case "table":
			rows := p.extractTableRows(n)
			if len(rows) > 0 {
				p.addBlock(htmlBlock{
					blockType: "table",
					tableRows: rows,
				})
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
	if p.current != nil {
		p.slides = append(p.slides, *p.current)
	}
	p.current = &htmlSlide{
		title:   title,
		content: make([]htmlBlock, 0),
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
func (p *htmlParser) extractListItems(n *html.Node) []string {
	var items []string
	for c := n.FirstChild; c != nil; c = c.NextSibling {
		if c.Type == html.ElementNode && c.Data == "li" {
			text := p.extractText(c)
			if text != "" {
				items = append(items, text)
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
	client := &http.Client{
		Timeout: 30 * time.Second,
	}

	resp, err := client.Get(url)
	if err != nil {
		return nil, ""
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, ""
	}

	data, err = io.ReadAll(resp.Body)
	if err != nil {
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
		if opts.SlideBackground != "" {
			slide.SetBackground(BackgroundOptions{
				Color: opts.SlideBackground,
			})
		}

		yPos := 0.5 // 当前Y位置

		// 添加标题
		if htmlSlide.title != "" {
			slide.AddText(htmlSlide.title, TextOptions{
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
		for _, block := range htmlSlide.content {
			switch block.blockType {
			case "heading":
				fontSize := 24.0
				if block.level == 2 {
					fontSize = 28.0
				} else if block.level >= 5 {
					fontSize = 20.0
				}
				slide.AddText(block.text, TextOptions{
					X:         0.5,
					Y:         yPos,
					Width:     9.0,
					Height:    0.6,
					FontSize:  fontSize,
					FontColor: opts.HeadingColor,
					Bold:      true,
				})
				yPos += 0.7

			case "text":
				slide.AddText(block.text, TextOptions{
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
					text := "• " + line
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
				lines := strings.Split(block.text, "\n")
				codeHeight := float64(len(lines)) * 0.35
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

				slide.AddText(block.text, TextOptions{
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
				// 解析图片来源（支持本地路径、URL、Base64）
				imageData, imageExt := parseImageSrc(block.imageSrc)
				if len(imageData) > 0 {
					slide.AddImage(ImageOptions{
						Data:    imageData,
						X:       1.0,
						Y:       yPos,
						Width:   8.0,
						Height:  3.0,
						AltText: block.imageAlt,
					})
					_ = imageExt // 扩展名由 AddImage 自动检测
					yPos += 3.2
				}

			case "table":
				if len(block.tableRows) > 0 {
					// 构建表格单元格
					var tableCells [][]TableCell
					for i, row := range block.tableRows {
						var cellRow []TableCell
						for _, cell := range row {
							tc := TableCell{
								Text: cell,
							}
							// 第一行加粗
							if i == 0 {
								tc.Bold = true
							}
							cellRow = append(cellRow, tc)
						}
						tableCells = append(tableCells, cellRow)
					}

					rowCount := len(tableCells)
					tableHeight := float64(rowCount) * 0.4
					if tableHeight > 3.0 {
						tableHeight = 3.0
					}

					slide.AddTable(tableCells, TableOptions{
						X:            0.5,
						Y:            yPos,
						Width:        9.0,
						FirstRowBold: true,
						FirstRowFill: "#E6E6E6",
					})
					yPos += tableHeight + 0.3
				}
			}
		}
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
