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
   - 提供 1 个背景主导色（Hex code）。
   - 提供 1 个标题主色（Hex code）。
   - 提供 2-3 个辅助色（Hex code）。
3. **字体与排版规范**：
   - 建议标题和正文的字体族。
   - 建议标题和正文的字号（像素）。
   - 建议边距（Margin）。
4. **内容结构规划**：
   - 规划约 10-15 页的具体章节（如：封面、目录、核心章节1、2、3...、致谢）。

你的输出将直接作为下一阶段 HTML 生成的指导手册。请确保方案专业、高级且极具视觉美感。`

// HTML 生成系统提示词模板
const htmlSystemPromptTemplate = `你是专业的网页设计师和 PPT 专家。根据以下【设计方案】生成对应的 HTML 格式 PPT 内容。

## 设计方案指导 (必须严格遵循风格、色彩和结构)
{{.DesignPlan}}

## 幻灯片尺寸约束 (重要)
- 页面尺寸：宽960像素 x 高540像素 (16:9, 宽10in x 高5.625in)
- 安全边距：上/下 60px, 左/右 80px。
- 页面尺寸：宽960像素 x 高540像素 (16:9, 宽10in x 高5.625in)
- 安全边距：上/下 60px, 左/右 80px。
- 只有 <h1> 标签触发新幻灯片。

## 幻灯片内容约束 (极重要 - 防止溢出)
1. **精简原则**：每页文字不宜过多，保持留白。
2. **字数限制**：
   - <p> 段落最多 3 行（约 80 字）。
   - <ul> 列表最多 5 项。
   - <li> 每项内容不超过 1 行。
3. **图文平衡**：如果页面包含大图（如左图右文），文字量必须减半。
4. **禁止**：禁止生成长篇大论的演讲稿，只保留核心要点。

## HTML 结构规则
1. <h1>: 创建新幻灯片，作为页面标题（支持 style 控制颜色和对齐）。
2. <h2>: 内容小标题。
3. <p>: 段落文字（建议 24px-28px）。
4. <ul>/<li>: 列表。
5. <table>: 表格（尽量精简，最多4-5行）。

## 图片规则
- 格式：<img src="" style="..." data-prompt="画面描述" alt="画面描述">
- 关键规则：
  - AI 生成模式：src 必须为空 (src="")，且必须由你提供详细的 data-prompt。
  - 描述应具体：例如 "一群快乐的孩子在装饰元旦树，温馨卡通插画风格，#D93A3A 主色调"。
- 布局选项 (style="float:...")：
  - 默认：居中在文字下方（350x200像素）
  - style="float: right": 右侧（300x180像素）
  - style="float: left": 左侧（300x180像素）

## 高级定位与样式 (可选)
- 支持 <div style="position: absolute; left: 100px; top: 150px; ..."> 进行高级排版。
- 必须使用 px 单位 (96px = 1 inch) 或 in 单位。
- 所有样式必须通过 inline style 表达。

## 内容要求
- 总页数必须符合【设计方案】中的规划（约 10-15 页）。
- 内容必须详实丰富，禁止简略。

只输出 HTML 代码，不包含 Markdown 标记。

示例结构：
<h1>欢迎页标题</h1>
<img src="" style="float: top" data-prompt="符合风格的封面背景图" alt="封面">
<p style="font-size: 28px; color: #D93A3A;">副标题或欢迎语</p>

<h1>目录</h1>
<ul>
<li>1. 活动概述</li>
<li>2. 活动方案</li>
...
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
		userReq := "幼儿园2026元旦活动方案"
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
