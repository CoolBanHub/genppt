package main

import (
	"context"
	"encoding/base64"
	"flag"
	"fmt"
	"os"
	"regexp"
	"strings"

	"github.com/CoolBanHub/genppt"
	"github.com/cloudwego/eino-ext/components/model/gemini"
	"github.com/cloudwego/eino/schema"
	"github.com/golang/glog"
	"github.com/joho/godotenv"
	"google.golang.org/genai"
)

// 设计方案生成提示词模板
const designPromptTemplate = `你是一位顶尖的演示文稿设计师。用户的需求是：“{{.UserReq}}”。

请为这个主题制订一份详尽的【PPT设计方案】。你的输出必须包含以下四个部分，并使用 Markdown 格式：

1. **设计风格定位**：描述整体视觉氛围（如：极简商务、温馨童趣、赛博朋克等）。
2. **色彩方案**：
   - 提供 1 个 **主背景色**（Hex code，用于大多数内容页）。
   - 提供 1 个 **强调背景色**（Hex code，仅用于封面、目录、章节页）。
   - 提供 1 个标题主色（Hex code）。
   - 提供 1 个正文色（Hex code）。
3. **字体与排版规范**：
   - 建议标题和正文的字体族。
   - 建议标题和正文的字号（像素）。
   - 建议边距（Margin）。
4. **内容结构规划**：
   - 规划约 10-15 页的具体章节（如：封面、目录、核心章节1、2、3...、致谢）。

你的输出将直接作为下一阶段 HTML 生成的指导手册。请确保方案专业、高级且极具视觉美感。`

// HTML 生成系统提示词模板
const htmlSystemPromptTemplate = `你是一位 HTML PPT 生成器。你的任务是使用 **极简 HTML** 生成演示文稿。

╔══════════════════════════════════════════════════════════════╗
║ 核心规则（违反将导致排版混乱）：                               ║
║                                                              ║
║ 1. 只能使用下方提供的 4 个 HTML 模板                          ║
║ 2. 禁止使用 flexbox, grid, float, absolute 等复杂布局        ║
║ 3. 禁止使用 inline style（除了颜色和字号）                    ║
║ 4. 每页最多 4 个要点，每个要点最多 15 字                      ║
║                                                              ║
║ 模板是唯一正确答案！不要自由发挥！                             ║
╚══════════════════════════════════════════════════════════════╝

## 设计方案参考（仅用于主题和颜色）
{{.DesignPlan}}

## 基本约束
- 所有内容必须使用中文
- 只有 <h1> 触发新幻灯片
- 页面尺寸：960px x 540px

## 允许使用的 HTML 标签（仅限以下）
- <h1>: 幻灯片标题（最多 12 字）
- <p>: 段落（最多 40 字）
- <ul>, <li>: 列表（最多 4 项，每项最多 15 字）
- <img>: 图片（必须有 data-prompt）
- <table>, <tr>, <td>, <th>: 表格（最多 4 行 x 3 列）

## 禁止使用的内容（会导致排版混乱）
❌ 禁止 flexbox (display: flex)
❌ 禁止 grid (display: grid)
❌ 禁止 float
❌ 禁止 position: absolute
❌ 禁止复杂的 inline style
❌ 禁止超过 15 字的列表项
❌ 禁止超过 4 项的列表

## HTML 结构模板（严格遵循这些格式）

### 模板 1：封面页（复制这个结构！）
<h1>活动主题</h1>
<img src="" data-prompt="封面图描述" alt="封面">
<p>副标题</p>

### 模板 2：列表页（复制这个结构！）
<h1>页面标题</h1>
<ul>
<li>要点1</li>
<li>要点2</li>
<li>要点3</li>
<li>要点4</li>
</ul>

### 模板 3：段落页（复制这个结构！）
<h1>页面标题</h1>
<p>简短说明（最多40字）。</p>

### 模板 4：表格页（复制这个结构！）
<h1>页面标题</h1>
<table>
<tr><th>列1</th><th>列2</th><th>列3</th></tr>
<tr><td>数据1</td><td>数据2</td><td>数据3</td></tr>
<tr><td>数据4</td><td>数据5</td><td>数据6</td></tr>
</table>

## 正确示例（严格遵循模板）

封面：
<h1>春节活动方案</h1>
<img src="" data-prompt="新年场景，红色金色主调" alt="封面">
<p>传统文化新体验</p>

列表页：
<h1>活动亮点</h1>
<ul>
<li>传统文化体验</li>
<li>创意手工制作</li>
<li>亲子互动游戏</li>
<li>新年祝福墙</li>
</ul>

## 错误示例（内容过长）

❌ 错误 - 段落太长：
<h1>活动目标</h1>
<p>本次活动旨在通过丰富多彩的新年主题活动，让孩子们深入了解中国传统文化中关于马年的寓意和习俗，同时通过各种互动游戏和手工制作，培养孩子们的动手能力、创造力和团队协作精神，为孩子们创造一个难忘的新年体验。</p>

❌ 错误 - 列表太多太长：
<h1>活动内容</h1>
<ul>
<li>传统文化体验区：包括剪纸、书法、投壶等传统游戏</li>
<li>创意手工制作区：制作新年贺卡、灯笼、福字等</li>
<li>亲子互动游戏区：家长和孩子一起参与的趣味游戏</li>
<li>新年祝福墙：孩子们写下新年愿望</li>
<li>美食品尝区：传统小吃和新年糕点</li>
<li>摄影留念区：专业摄影师为家庭拍照</li>
</ul>`

var (
	textModel  *gemini.ChatModel
	imageModel *gemini.ChatModel
)

func main() {
	// CLI Flags
	htmlPath := flag.String("html", "", "Path to input HTML file (if provided, AI generation is skipped)")
	outPath := flag.String("out", "output.pptx", "Path to output PPTX file")
	flag.Parse()
	defer glog.Flush()
	glog.V(2).Info("DEBUG: Program started with html:", *htmlPath)

	// 加载环境变量 (使用 Overload 确保 .env 中的配置覆盖系统已有的空值)
	if err := godotenv.Load(); err != nil {
		glog.V(2).Infof("未发现 .env 文件或加载失败 (将使用系统环境变量): %v", err)
	} else {
		glog.V(2).Info(".env 配置文件加载成功")
	}

	ctx := context.Background()
	apiKey := os.Getenv("GOOGLE_API_KEY")
	if apiKey == "" {
		glog.Error("错误: 未发现 GOOGLE_API_KEY，请确保 .env 文件配置正确或已设置环境变量")
		if *htmlPath == "" {
			return
		}
	}

	// 创建 Gemini 客户端
	client, err := genai.NewClient(ctx, &genai.ClientConfig{
		APIKey:  apiKey,
		Backend: genai.BackendVertexAI,
	})
	if err != nil {
		glog.Errorf("创建gemini客户端失败: %v", err)
		if *htmlPath == "" {
			return
		}
	}

	// 创建文字模型
	if client != nil {
		textModel, err = gemini.NewChatModel(ctx, &gemini.Config{
			Client: client,
			Model:  "gemini-3-flash-preview",
		})
		if err != nil {
			glog.Errorf("创建文字模型失败: %v", err)
			if *htmlPath == "" {
				return
			}
		}
	}

	// 创建图片生成模型
	if client != nil {
		imageModel, err = gemini.NewChatModel(ctx, &gemini.Config{
			Client: client,
			Model:  "gemini-2.5-flash-image",
			ResponseModalities: []gemini.GeminiResponseModality{
				gemini.GeminiResponseModalityImage,
			},
		})
		if err != nil {
			glog.Errorf("创建图片模型失败: %v", err)
			if *htmlPath == "" {
				return
			}
		}
	}

	// 用户需求描述

	var finalHTML string

	if *htmlPath != "" {
		glog.V(1).Infof("从文件读取 HTML: %s", *htmlPath)
		content, err := os.ReadFile(*htmlPath)
		if err != nil {
			glog.Errorf("读取 HTML 文件失败: %v", err)
			return
		}
		htmlContent := string(content)
		// Process images even if reading from file, to ensure src overwrites work if alt exists
		// But usually local HTML might already have valid src.
		// For our test case, we want to test image replacement.
		glog.V(1).Info("Step 2: 提取图片需求并生成图片...")
		finalHTML = processImages(ctx, htmlContent)
	} else {
		// Dynamic Generation Flow
		userReq := "2026年春节活动方案"
		glog.Infof("开始为主题生成 PPT: %s", userReq)

		// Step 1: 生成设计方案
		glog.Info("Step 1: 正在生成 PPT 设计方案...")
		designPlan := generateDesignPlan(ctx, userReq)
		if designPlan == "" {
			glog.Error("设计方案生成失败")
			return
		}
		os.WriteFile("design_plan.md", []byte(designPlan), 0644)
		glog.V(1).Info("设计方案已保存: design_plan.md")

		// Step 2: 根据设计方案生成 HTML
		glog.Info("Step 2: 正在根据设计方案生成 HTML 内容...")
		htmlContent := generateHTML(ctx, designPlan)
		if htmlContent == "" {
			glog.Error("HTML 生成失败")
			return
		}
		// 保存原始HTML
		os.WriteFile("output_original.html", []byte(htmlContent), 0644)
		glog.V(2).Info("已保存原始HTML: output_original.html")

		// Step 3: 提取图片需求并生成图片
		glog.Info("Step 3: 提取图片需求并生成图片...")
		finalHTML = processImages(ctx, htmlContent)

		// 保存最终HTML
		os.WriteFile("output_final.html", []byte(finalHTML), 0644)
	}

	// Step 4: 转换 HTML 为 PPT
	glog.V(1).Info("Step 4: 转换 HTML 为 PPT...")
	opts := genppt.DefaultHTMLOptions()
	pres := genppt.FromHTMLWithOptions(finalHTML, opts)

	// 保存PPT文件
	outputFile := *outPath
	if err := pres.WriteFile(outputFile); err != nil {
		glog.Errorf("保存PPT失败: %v", err)
		return
	}

	glog.Infof("✅ PPT生成成功: %s (共%d页)", outputFile, pres.SlideCount())
}

// generateDesignPlan 生成 PPT 设计方案
func generateDesignPlan(ctx context.Context, userReq string) string {
	prompt := strings.Replace(designPromptTemplate, "{{.UserReq}}", userReq, 1)
	msg, err := textModel.Generate(ctx, []*schema.Message{
		{Role: schema.User, Content: prompt},
	})
	if err != nil {
		glog.Errorf("设计方案生成失败: %v", err)
		return ""
	}
	return msg.Content
}

// generateHTML 根据设计方案生成 HTML 代码
func generateHTML(ctx context.Context, designPlan string) string {
	systemPrompt := strings.Replace(htmlSystemPromptTemplate, "{{.DesignPlan}}", designPlan, 1)
	msg, err := textModel.Generate(ctx, []*schema.Message{
		{Role: schema.System, Content: systemPrompt},
		{Role: schema.User, Content: "请根据上述设计方案生成一份演示文稿 HTML。"},
	})
	if err != nil {
		glog.Errorf("HTML生成失败: %v", err)
		return ""
	}

	content := msg.Content
	// 清理可能的 markdown 代码块标记
	content = strings.TrimPrefix(content, "```html")
	content = strings.TrimPrefix(content, "```")
	content = strings.TrimSuffix(content, "```")
	content = strings.TrimSpace(content)

	return content
}

// processImages 提取 alt 属性作为提示词，生成图片并替换 src
func processImages(ctx context.Context, html string) string {
	// 匹配所有 img 标签
	imgTagRegex := regexp.MustCompile(`<img\s+[^>]*>`)
	matches := imgTagRegex.FindAllString(html, -1)

	// 预扫描：计算需要生成的图片数量
	validMatches := 0
	for _, tag := range matches {
		// 只要有 alt 属性且不为空，我们就认为是 prompt
		// 我们稍微放宽条件，允许任意引号或无引号，但通常 AI 生成的都是标准双引号
		if strings.Contains(tag, `alt="`) {
			validMatches++
		}
	}
	glog.V(2).Infof("找到 %d 张图片需要生成", validMatches)

	result := html
	count := 0

	for _, tag := range matches {
		// 优先检查 src 是否已存在且非空
		// 提取 src
		var srcVal string
		srcRegex := regexp.MustCompile(`src=["']([^"']*)["']`)
		if m := srcRegex.FindStringSubmatch(tag); len(m) >= 2 {
			srcVal = m[1]
		}

		// 如果 src 不为空，则跳过生成（用户要求：src存在图片地址就不需要生成）
		if srcVal != "" {
			glog.V(3).Infof("图片已存在 src，跳过生成: %s", truncateString(srcVal, 30))
			continue
		}

		// 提取 prompt (data-prompt 优先，其次 alt)
		// 支持 data-prompt="..." 和 alt="..."
		var prompt string

		dpRegexDouble := regexp.MustCompile(`data-prompt="([^"]+)"`)
		dpRegexSingle := regexp.MustCompile(`data-prompt='([^']+)'`)

		if m := dpRegexDouble.FindStringSubmatch(tag); len(m) >= 2 {
			prompt = m[1]
		} else if m := dpRegexSingle.FindStringSubmatch(tag); len(m) >= 2 {
			prompt = m[1]
		}

		if prompt == "" {
			altRegexDouble := regexp.MustCompile(`alt="([^"]+)"`)
			altRegexSingle := regexp.MustCompile(`alt='([^']+)'`)

			if m := altRegexDouble.FindStringSubmatch(tag); len(m) >= 2 {
				prompt = m[1]
			} else if m := altRegexSingle.FindStringSubmatch(tag); len(m) >= 2 {
				prompt = m[1]
			}
		}

		if prompt == "" {
			continue
		}

		count++
		glog.V(2).Infof("正在生成图片 %d/%d: %s", count, validMatches, truncateString(prompt, 30))

		// 生成图片
		imageData := generateImage(ctx, prompt)
		if imageData != "" {
			// 保存单独的图片文件
			saveImageFile(imageData, count)

			// 替换 src
			// 我们需要替换整个 src="..." 部分，不管它原来是什么
			// 正则匹配 src="...", src='...', 或 src=无引号 (较少见但HTML允许)
			// 为了安全，我们假设标准格式，或者尽可能匹配宽泛

			newTag := tag
			srcRegex := regexp.MustCompile(`src=["']([^"']*)["']`)

			if srcRegex.MatchString(tag) {
				// 替换现有的 src
				newTag = srcRegex.ReplaceAllString(tag, `src="`+imageData+`"`)
			} else {
				// 如果没有 src，插入一个 (虽然 invalid HTML 但为了容错)
				// 插入在 img 后
				newTag = strings.Replace(tag, "<img", `<img src="`+imageData+`"`, 1)
			}

			// 替换原 HTML 中的标签
			// 注意：如果多个标签完全一样（内容完全一样），Replace 会替换所有
			// 这在一般 PPT 生成场景下概率较低（alt 通常不同），但为了严谨可以使用 Replace(..., 1) 配合 split 处理
			// 简单起见，这里假设 AI 生成的 alt 不重复，或者重复时替换也没问题
			result = strings.Replace(result, tag, newTag, 1)
			glog.V(2).Infof("图片 %d 生成成功", count)
		} else {
			glog.Warningf("图片 %d 生成失败", count)
		}
	}

	return result
}

// generateImage 使用 AI 生成图片，返回 data URI
func generateImage(ctx context.Context, prompt string) string {
	if imageModel == nil {
		glog.Warning("Image model not initialized, skipping image generation")
		return ""
	}
	msg, err := imageModel.Generate(ctx, []*schema.Message{
		{
			Role: schema.User,
			UserInputMultiContent: []schema.MessageInputPart{
				{Type: "text", Text: prompt},
			},
		},
	})
	if err != nil {
		glog.Warningf("图片生成API调用失败: %v", err)
		return ""
	}

	// 提取生成的图片数据
	for _, part := range msg.AssistantGenMultiContent {
		if part.Image != nil && part.Image.Base64Data != nil {
			// 验证 base64 是否有效
			_, err := base64.StdEncoding.DecodeString(*part.Image.Base64Data)
			if err != nil {
				glog.Warningf("Base64数据无效: %v", err)
				continue
			}

			// 构建 data URI
			mimeType := "image/png"
			if part.Image.MIMEType != "" {
				mimeType = part.Image.MIMEType
			}
			return "data:" + mimeType + ";base64," + *part.Image.Base64Data
		}
	}

	return ""
}

// truncateString 截断字符串用于日志显示
func truncateString(s string, maxLen int) string {
	runes := []rune(s)
	if len(runes) <= maxLen {
		return s
	}
	return string(runes[:maxLen]) + "..."
}

// saveImageFile 保存图片文件
func saveImageFile(dataURI string, index int) {
	// 解析 data URI: data:image/png;base64,xxxxx
	if !strings.HasPrefix(dataURI, "data:") {
		return
	}

	commaIdx := strings.Index(dataURI, ",")
	if commaIdx == -1 {
		return
	}

	// 获取 MIME 类型和扩展名
	header := dataURI[5:commaIdx] // 去掉 "data:"
	ext := "png"
	if strings.Contains(header, "jpeg") || strings.Contains(header, "jpg") {
		ext = "jpg"
	} else if strings.Contains(header, "gif") {
		ext = "gif"
	} else if strings.Contains(header, "webp") {
		ext = "webp"
	}

	// 解码 base64
	b64Data := dataURI[commaIdx+1:]
	data, err := base64.StdEncoding.DecodeString(b64Data)
	if err != nil {
		glog.Warningf("图片解码失败: %v", err)
		return
	}

	// 保存文件
	filename := fmt.Sprintf("image_%d.%s", index, ext)
	if err := os.WriteFile(filename, data, 0644); err != nil {
		glog.Warningf("保存图片失败: %v", err)
		return
	}
	glog.V(2).Infof("已保存图片: %s", filename)
}
