package main

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"flag"
	"os"
	"strings"

	"github.com/CoolBanHub/genppt"
	"github.com/cloudwego/eino-ext/components/model/gemini"
	"github.com/cloudwego/eino/schema"
	"github.com/golang/glog"
	"github.com/joho/godotenv"
	"google.golang.org/genai"
)

// Data Structures for JSON Parsing
type GeneratedPPT struct {
	Title  string         `json:"title"`
	Theme  Theme          `json:"theme"`
	Slides []SlideContent `json:"slides"`
}

type Theme struct {
	BackgroundColor string `json:"background_color"`
	TitleColor      string `json:"title_color"`
	BodyColor       string `json:"body_color"`
	FontFamily      string `json:"font_family"`
}

type SlideContent struct {
	Title           string    `json:"title"`
	Layout          string    `json:"layout"` // e.g., "title_body", "two_column", "title_only"
	BackgroundColor string    `json:"background_color,omitempty"`
	Elements        []Element `json:"elements"`
	Notes           string    `json:"notes"`
}

type Element struct {
	Type        string  `json:"type"`    // "text", "bullet", "image", "shape"
	Content     string  `json:"content"` // For text/bullet
	Prompt      string  `json:"prompt"`  // For image generation
	X           float64 `json:"x"`       // Inches
	Y           float64 `json:"y"`
	Width       float64 `json:"width"`
	Height      float64 `json:"height"`
	Color       string  `json:"color,omitempty"`
	Size        float64 `json:"size,omitempty"`
	Align       string  `json:"align,omitempty"`        // "left", "center", "right"
	LineSpacing float64 `json:"line_spacing,omitempty"` // Line spacing multiplier
	// Image Props
	AspectRatio string `json:"aspect_ratio,omitempty"` // "16:9", "1:1", "4:3"
	MimeType    string `json:"mime_type,omitempty"`    // "image/png", "image/jpeg"
}

// Step 1: Design System Prompt
const designSystemPrompt = `你是一位创意总监和演示文稿专家。
分析用户的需求："{{.UserReq}}"

创建一个全面的 **演示文稿设计策略**。
1. **视觉风格**：定义基调（例如：商务、活泼、未来感）。
2. **调色板**：指定以下颜色的十六进制代码：
   - **主背景色**（用于大多数幻灯片）
   - **强调背景色**（用于封面或章节页）
   - **主标题色**
   - **正文色**
3. **结构**：概述幻灯片流程（大约 5-10 张幻灯片）。
4. **关键意象**：描述视觉母题。

**重要**：所有生成的内容必须使用 **中文**。
请以清晰的 Markdown 格式输出。暂时不要输出 JSON。`

// Step 2: JSON System Prompt
const jsonSystemPrompt = `你是一位 PPT 实现工程师。
将以下的 **设计策略** 转换为严格的 JSON 结构。

╔═══════════════════════════════════════════════════╗
║ 关键方法：使用预定义模板                             	║
║                                                   ║
║ 1. 不要手动计算位置                                  ║
║ 2. 严格复制模板坐标                                  ║
║ 3. 保持内容简短:每张胶片最多 4 个要点，每个要点 8-10 个字 ║
║                                                   ║
║ 内容太长 = 文本溢出 = PPT 排版混乱！                   ║
╚═══════════════════════════════════════════════════╝

## 设计策略
{{.DesignPlan}}

## 可用模板（每张胶片选择一个）

你有 7 个预先设计的模板。选择最适合内容的模板，然后仅提供每个插槽的内容。

### 模板 ID：
1. **"title_slide"** - 用于封面/第一张胶片
2. **"content_1col"** - 带要点的单列
3. **"content_2col_text_image"** - 左文右图
4. **"content_2col_image_text"** - 左图右文
5. **"content_3col"** - 三列等宽
6. **"big_fact"** - 大数字/统计
7. **"section_header"** - 章节分隔页

## 模板如何工作

重要：位置在模板中是 **预定义** 的。你只提供内容，而不是 x/y/width/height。

每个模板都有 **固定插槽**。以 "title_slide" 为例：
- 插槽 1：标题（居中，字号 48）
- 插槽 2：副标题（居中，字号 28）

你只需要为每个插槽提供 **文本**。**所有生成的内容必须使用中文**。

## 模板坐标参考（严格复制这些！）

### Template 1: "title_slide" - Cover slide
"elements": [
  {"type": "text", "content": "YOUR TITLE", "x": 0.5, "y": 2.0, "width": 9.0, "height": 1.5, "size": 48, "align": "center"},
  {"type": "text", "content": "YOUR SUBTITLE", "x": 0.5, "y": 3.5, "width": 9.0, "height": 0.8, "size": 28, "align": "center"}
]

### Template 2: "content_1col" - Title + bullets
"elements": [
  {"type": "text", "content": "TITLE", "x": 0.5, "y": 0.3, "width": 9.0, "height": 0.8, "size": 36, "align": "left"},
  {"type": "bullet", "content": "Point 1\nPoint 2\nPoint 3\nPoint 4", "x": 0.5, "y": 1.3, "width": 9.0, "height": 3.8, "size": 20, "line_spacing": 1.5}
]
注意：最多 4 个要点，每个要点最多 8-10 个字！

### Template 3: "content_2col_text_image" - Text left, image right
"elements": [
  {"type": "text", "content": "TITLE", "x": 0.5, "y": 0.3, "width": 9.0, "height": 0.8, "size": 36, "align": "left"},
  {"type": "bullet", "content": "Point 1\nPoint 2\nPoint 3", "x": 0.5, "y": 1.3, "width": 4.2, "height": 3.8, "size": 18, "line_spacing": 1.5},
  {"type": "image", "prompt": "IMAGE DESCRIPTION", "x": 5.3, "y": 1.3, "width": 4.2, "height": 3.8, "aspect_ratio": "1:1", "mime_type": "image/jpeg"}
]

### Template 4: "content_2col_image_text" - Image left, text right
"elements": [
  {"type": "text", "content": "TITLE", "x": 0.5, "y": 0.3, "width": 9.0, "height": 0.8, "size": 36, "align": "left"},
  {"type": "image", "prompt": "IMAGE DESCRIPTION", "x": 0.5, "y": 1.3, "width": 4.2, "height": 3.8, "aspect_ratio": "1:1", "mime_type": "image/jpeg"},
  {"type": "bullet", "content": "Point 1\nPoint 2\nPoint 3", "x": 5.3, "y": 1.3, "width": 4.2, "height": 3.8, "size": 18, "line_spacing": 1.5}
]

### Template 5: "content_3col" - Three columns
"elements": [
  {"type": "text", "content": "TITLE", "x": 0.5, "y": 0.3, "width": 9.0, "height": 0.8, "size": 36, "align": "left"},
  {"type": "text", "content": "Column 1", "x": 0.5, "y": 1.3, "width": 2.8, "height": 3.8, "size": 18, "line_spacing": 1.3},
  {"type": "text", "content": "Column 2", "x": 3.6, "y": 1.3, "width": 2.8, "height": 3.8, "size": 18, "line_spacing": 1.3},
  {"type": "text", "content": "Column 3", "x": 6.7, "y": 1.3, "width": 2.8, "height": 3.8, "size": 18, "line_spacing": 1.3}
]

### Template 6: "big_fact" - Large number/stat
"elements": [
  {"type": "text", "content": "CONTEXT TITLE", "x": 0.5, "y": 0.3, "width": 9.0, "height": 0.8, "size": 36, "align": "left"},
  {"type": "text", "content": "75%", "x": 2.0, "y": 1.8, "width": 6.0, "height": 1.5, "size": 72, "align": "center"},
  {"type": "text", "content": "Description", "x": 2.0, "y": 3.5, "width": 6.0, "height": 0.8, "size": 24, "align": "center"}
]

### Template 7: "section_header" - Section divider
"elements": [
  {"type": "text", "content": "SECTION NAME", "x": 0.5, "y": 2.3, "width": 9.0, "height": 1.2, "size": 54, "align": "center"}
]

## 输出格式 - 严格复制模板

重要：对于每张胶片，选择一个模板并 **严格复制** 其坐标。只更改 "content"。

仅返回有效的 JSON。不要 Markdown 块，不要代码栅栏。

{
  "title": "演示文稿标题",
  "theme": {
    "background_color": "#FFFFFF",
    "title_color": "#000000",
    "body_color": "#333333",
    "font_family": "Arial"
  },
  "slides": [
    {
      "title": "描述性标题",
      "layout": "title_slide",
      "elements": [
        {
          "type": "text",
          "content": "你的标题",
          "x": 0.5, "y": 2.0, "width": 9.0, "height": 1.5,
          "size": 48, "align": "center"
        },
        {
          "type": "text",
          "content": "你的副标题",
          "x": 0.5, "y": 3.5, "width": 9.0, "height": 0.8,
          "size": 28, "align": "center"
        }
      ],
      "notes": "演讲者备注"
    }
  ]
}

记住：严格复制模板坐标，只修改 "content" 字段！**确保所有内容（content）都是中文**。

## 强制规则

### 必须做：
1. **从上面的 7 个选项中选择一个模板**
2. **严格复制** 模板示例中的坐标
3. **仅更改 "content"** 字段为你的文本
4. **使用 "bullet" 类型** 用于列表（而不是 "text"）
5. **保持列表超简洁**：
   - 每张胶片最多 4 个要点（不能是 5 个或 6 个！）
   - 每个要点最多 8-10 个字（仅一行！）
   - 好例子："具有触控响应的交互式数字墙"
   - 坏例子："数字交互式墙：具有多种功能和能力的触控反应艺术装置"
6. **编写详细的图像提示词**：为了 AI 图像生成的质量（图像提示词可以用英文或中文）

### 不要这样做：
1. ❌ 不要手动计算 x, y, width, height 值
2. ❌ 不要修改模板坐标 "为了让它看起来更好"
3. ❌ 不要在模板插槽之外添加额外元素
4. ❌ 不要使用微小字体（正文最小 18pt，标题 32pt）
5. ❌ 不要创建自定义布局 - 仅使用 7 个模板

### 颜色控制（严格执行）：
1. **严禁发明颜色**：只使用设计策略中定义的“主背景色”和“强调背景色”。
2. **背景分配**：
   - 封面页(title_slide) 和 章节页(section_header) -> 使用 **强调背景色**
   - 所有其他内容页 -> 使用 **主背景色**
3. **一致性**：整个演示文稿中背景颜色必须统一，不要随机更改。

### 为什么模板很重要：
- 模板经过测试并保证有效
- 自定义坐标会导致重叠/混乱的布局
- AI 不应该猜测位置 - 人类设计了模板

## 验证清单
在输出 JSON 之前，请验证：
- [ ] 坐标完全复制自模板（未修改！）
- [ ] 每个元素都有 x, y, width, height 属性
- [ ] 每个元素都有 "align" 和 "line_spacing" (如果需要)
- [ ] 标题胶片严格使用模板 1 格式
- [ ] 内容胶片严格使用模板 2-7
- [ ] **内容长度**：计算要点 - 每张胶片最多 4 个
- [ ] **内容长度**：计算每个要点的字数 - 每个最多 8-10 个字
- [ ] **内容语言**：确保所有 content 都是中文
- [ ] 文本元素在需要时有 "align" 属性
- [ ] 列表元素使用 "bullet" 类型而不是 "text"
- [ ] JSON 有效（正确的逗号、括号、引号）
`

var (
	textModel  *gemini.ChatModel
	imageModel *gemini.ChatModel
)

// parseAlign converts string alignment to genppt.Align
func parseAlign(align string) genppt.Align {
	switch strings.ToLower(align) {
	case "center":
		return genppt.AlignCenter
	case "right":
		return genppt.AlignRight
	case "left":
		return genppt.AlignLeft
	default:
		return genppt.AlignLeft // default
	}
}

func main() {
	userReq := flag.String("req", "Human Colonization of Mars", "User Requirement for presentation")
	outPath := flag.String("out", "output_json.pptx", "Path to output PPTX file")
	flag.Parse()
	defer glog.Flush()

	// Initialize Env
	if err := godotenv.Load(); err != nil {
		glog.Warningf("No .env file found: %v", err)
	}

	apiKey := os.Getenv("GOOGLE_API_KEY")
	if apiKey == "" {
		glog.Error("GOOGLE_API_KEY not set")
		return
	}

	ctx := context.Background()
	client, err := genai.NewClient(ctx, &genai.ClientConfig{
		APIKey:  apiKey,
		Backend: genai.BackendVertexAI,
	})
	if err != nil {
		glog.Errorf("Failed to create Gemini client: %v", err)
		return
	}

	// Initialize Models
	textModel, err = gemini.NewChatModel(ctx, &gemini.Config{
		Client: client, Model: "gemini-3-flash-preview",
	})
	if err != nil {
		glog.Errorf("Failed to create text model: %v", err)
		return
	}

	imageModel, err = gemini.NewChatModel(ctx, &gemini.Config{
		Client:             client,
		Model:              "gemini-2.5-flash-image",
		ResponseModalities: []gemini.GeminiResponseModality{gemini.GeminiResponseModalityImage},
	})
	if err != nil {
		glog.Errorf("Failed to create image model: %v", err)
		return
	}

	// Step 1: Design Phase
	glog.Infof("Step 1: Analyzing Requirement & Designing: %s", *userReq)
	designPlan := generateDesignPlan(ctx, *userReq)
	if designPlan == "" {
		return
	}
	os.WriteFile("design_plan.md", []byte(designPlan), 0644)

	// Step 2: Implementation Phase (JSON)
	glog.Infof("Step 2: Generating JSON Plan...")
	jsonContent := generateJSON(ctx, designPlan)
	if jsonContent == "" {
		return
	}
	os.WriteFile("plan.json", []byte(jsonContent), 0644)

	var pptPlan GeneratedPPT
	if err := json.Unmarshal([]byte(jsonContent), &pptPlan); err != nil {
		glog.Errorf("Failed to parse JSON: %v", err)
		return
	}

	// Step 2: Render PPT
	glog.Info("Rendering PPT from JSON plan...")
	pres := genppt.New()

	// Default Theme Colors
	bgColor := pptPlan.Theme.BackgroundColor
	if bgColor == "" {
		bgColor = "FFFFFF"
	}
	titleColor := pptPlan.Theme.TitleColor
	if titleColor == "" {
		titleColor = "000000"
	}
	bodyColor := pptPlan.Theme.BodyColor
	if bodyColor == "" {
		bodyColor = "333333"
	}

	for i, slideData := range pptPlan.Slides {
		glog.Infof("Processing slide %d: %s", i+1, slideData.Title)
		slide := pres.AddSlide()

		// Set Background
		bg := bgColor
		if slideData.BackgroundColor != "" {
			bg = slideData.BackgroundColor
		}
		slide.SetBackground(genppt.BackgroundOptions{Color: bg})

		// Process elements from JSON - let JSON fully control slide content
		for _, el := range slideData.Elements {
			switch el.Type {
			case "text":
				size := el.Size
				if size == 0 {
					size = 18
				}
				color := el.Color
				if color == "" {
					color = bodyColor
				}

				// Use align from JSON, default to left
				align := parseAlign(el.Align)

				// Use line spacing from JSON, default to 1.2
				lineSpacing := el.LineSpacing
				if lineSpacing == 0 {
					lineSpacing = 1.2
				}

				slide.AddText(el.Content, genppt.TextOptions{
					X: el.X, Y: el.Y, Width: el.Width, Height: el.Height,
					FontSize: size, FontColor: color,
					Align:       align,
					LineSpacing: lineSpacing,
				})

			case "bullet":
				// Process bullet points as a single text box with proper formatting
				size := el.Size
				if size == 0 {
					size = 18
				}
				color := el.Color
				if color == "" {
					color = bodyColor
				}

				// Use align from JSON, default to left
				align := parseAlign(el.Align)

				// Use line spacing from JSON, default to 1.5 for better bullet readability
				lineSpacing := el.LineSpacing
				if lineSpacing == 0 {
					lineSpacing = 1.5
				}

				// Format bullet points - ensure each line has a bullet
				lines := strings.Split(el.Content, "\n")
				var formattedLines []string
				for _, line := range lines {
					line = strings.TrimSpace(line)
					if line == "" {
						continue
					}
					// Add bullet if not present
					if !strings.HasPrefix(line, "•") {
						line = "• " + line
					}
					formattedLines = append(formattedLines, line)
				}

				// Use single text box for all bullets - this prevents overlap
				slide.AddText(strings.Join(formattedLines, "\n"), genppt.TextOptions{
					X: el.X, Y: el.Y, Width: el.Width, Height: el.Height,
					FontSize: size, FontColor: color,
					Align:       align,
					LineSpacing: lineSpacing,
				})

			case "image":
				// Generate Image
				glog.Infof("Generating image: %s (Ratio: %s)", el.Prompt, el.AspectRatio)

				imgConfig := &genai.ImageConfig{
					OutputMIMEType: el.MimeType,
				}
				if el.AspectRatio != "" {
					imgConfig.AspectRatio = el.AspectRatio
				}

				imgDataURI := generateImage(ctx, el.Prompt, imgConfig)
				if imgDataURI != "" {
					// Decode data URI
					b64 := imgDataURI[strings.Index(imgDataURI, ",")+1:]
					data, _ := base64.StdEncoding.DecodeString(b64)

					slide.AddImage(genppt.ImageOptions{
						X: el.X, Y: el.Y, Width: el.Width, Height: el.Height,
						Data:     data,
						Rounding: 0.1,
					})
				}
			}
		}

		if slideData.Notes != "" {
			slide.SetNotes(slideData.Notes)
		}
	}

	if err := pres.WriteFile(*outPath); err != nil {
		glog.Errorf("Failed to save PPT: %v", err)
	} else {
		glog.Infof("✅ Success! Encoded to %s", *outPath)
	}
}

func generateDesignPlan(ctx context.Context, req string) string {
	prompt := strings.Replace(designSystemPrompt, "{{.UserReq}}", req, 1)
	msg, err := textModel.Generate(ctx, []*schema.Message{
		{Role: schema.User, Content: prompt},
	})
	if err != nil {
		glog.Errorf("Design Gen Error: %v", err)
		return ""
	}
	return msg.Content
}

func generateJSON(ctx context.Context, designPlan string) string {
	prompt := strings.Replace(jsonSystemPrompt, "{{.DesignPlan}}", designPlan, 1)
	msg, err := textModel.Generate(ctx, []*schema.Message{
		{Role: schema.User, Content: prompt},
	})
	if err != nil {
		glog.Errorf("JSON Gen Error: %v", err)
		return ""
	}
	content := msg.Content
	// Cleanup json
	content = strings.TrimPrefix(content, "```json")
	content = strings.TrimPrefix(content, "```")
	content = strings.TrimSuffix(content, "```")
	return strings.TrimSpace(content)
}

func generateImage(ctx context.Context, prompt string, cfg *genai.ImageConfig) string {
	if imageModel == nil {
		return ""
	}

	var msg *schema.Message
	var err error

	if cfg != nil {
		msg, err = imageModel.Generate(ctx, []*schema.Message{
			{
				Role: schema.User,
				UserInputMultiContent: []schema.MessageInputPart{
					{Type: "text", Text: prompt},
				},
			},
		}, gemini.WithImageConfig(cfg))
	} else {
		msg, err = imageModel.Generate(ctx, []*schema.Message{
			{
				Role: schema.User,
				UserInputMultiContent: []schema.MessageInputPart{
					{Type: "text", Text: prompt},
				},
			},
		})
	}
	if err != nil {
		glog.Warningf("Image Gen Error: %v", err)
		return ""
	}
	for _, part := range msg.AssistantGenMultiContent {
		if part.Image != nil && part.Image.Base64Data != nil {
			return "data:image/png;base64," + *part.Image.Base64Data
		}
	}
	return ""
}
