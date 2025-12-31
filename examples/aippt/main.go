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
	"github.com/gookit/slog"
	"github.com/joho/godotenv"
	"google.golang.org/genai"
)

// 系统提示词：让 AI 生成适合转 PPT 的 HTML 代码
const systemPrompt = `你是专业的演示文稿设计师。根据用户主题生成HTML格式的PPT。

## 幻灯片尺寸约束（重要）
- 页面尺寸：宽960像素 x 高540像素（16:9）
- 标题区域：顶部80像素
- 内容区域：剩余约460像素高度

## HTML结构规则
1. <h1> 创建新幻灯片，作为页面标题
2. <h2> 内容小标题
3. <p> 段落文字（简短）
4. <ul>/<li> 列表
5. <table> 表格（最多4行）

## 图片规则
- 格式：<img src="" style="..." data-prompt="画面描述" alt="画面描述">
- 关键规则（src与data-prompt二选一）：
  - 场景A（AI生成图片）：src 必须为空 (src="")，且必须提供 data-prompt 描述画面。
  - 场景B（使用现有图片）：src 填入有效 URL，此时 data-prompt 应为空或省略。
- 布局选项 (style="float:...")：
  - 默认：居中在文字下方（350x200像素）
  - style="float: right"：右侧（300x180像素）
  - style="float: left"：左侧（300x180像素）

## 高级定位（可选）：
- 所有元素（图片、文字）均支持 style 属性进行绝对定位
- 必须使用 px 单位 (96px = 1 inch)，或者 in 单位
- 页面尺寸：960x540像素 (宽10in x 高5.625in)
- 示例：<img src="" style="position: absolute; left: 576px; top: 144px; width: 288px;" ...>

## 文字样式（可选）：
- 使用 standard CSS style 属性
- 示例：<p style="color: #FF0000; font-size: 24px;">强调文字</p>

## 设计原则（关键）：
- 只有 <h1> 标签才会触发新幻灯片。
- 如果内容较多，请使用多个 <h1> 分成多页展示。
- 混合使用不同布局，避免单调。

只输出HTML代码。

示例：
<h1>欢迎页</h1>
<img src="" style="float: top" data-prompt="背景图片" alt="背景">
<p style="font-size: 24px">欢迎参加！</p>

<h1>内容页</h1>
<!-- 右侧布局示例 -->
<img src="" style="float: right" data-prompt="插图" alt="插图">
<h2>副标题</h2>
<ul>
<li>要点一</li>
<li>要点二</li>
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
	fmt.Println("DEBUG: Program started with html:", *htmlPath)

	// 加载环境变量
	if err := godotenv.Load(); err != nil {
		slog.Warnf("警告: 无法加载 .env 文件: %v", err)
	}

	ctx := context.Background()

	// 创建 Gemini 客户端
	client, err := genai.NewClient(ctx, &genai.ClientConfig{
		APIKey:  os.Getenv("GOOGLE_API_KEY"),
		Backend: genai.BackendVertexAI,
	})
	if err != nil {
		slog.Error("创建gemini客户端失败", "err", err)
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
			slog.Error("创建文字模型失败", "err", err)
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
			slog.Error("创建图片模型失败", "err", err)
			if *htmlPath == "" {
				return
			}
		}
	}

	// 用户需求描述

	var finalHTML string

	if *htmlPath != "" {
		slog.Infof("从文件读取 HTML: %s", *htmlPath)
		content, err := os.ReadFile(*htmlPath)
		if err != nil {
			slog.Error("读取 HTML 文件失败", "err", err)
			return
		}
		htmlContent := string(content)
		// Process images even if reading from file, to ensure src overwrites work if alt exists
		// But usually local HTML might already have valid src.
		// For our test case, we want to test image replacement.
		slog.Info("Step 2: 提取图片需求并生成图片...")
		finalHTML = processImages(ctx, htmlContent)
	} else {
		// AI Generation Flow
		userReq := "帮我生成一个幼儿园2026元旦活动的PPT。要求：\n1. 内容适合幼儿园小朋友，语言生动有趣。\n2. 每页内容不要太多，建议包含标题、最多4个简短要点和1张相关的插图。\n3. 内容要丰富，包含活动安排、手工制作、美食分享和安全注意事项。"

		slog.Infof("开始生成 PPT: %s", userReq)

		// Step 1: 让 AI 生成 HTML 代码
		slog.Info("Step 1: AI 生成 HTML 内容...")
		htmlContent := generateHTML(ctx, userReq)
		if htmlContent == "" {
			slog.Error("HTML 生成失败")
			return
		}
		// 保存原始HTML
		os.WriteFile("output_original.html", []byte(htmlContent), 0644)
		slog.Info("已保存原始HTML: output_original.html")

		// Step 2: 提取图片需求并生成图片
		slog.Info("Step 2: 提取图片需求并生成图片...")
		finalHTML = processImages(ctx, htmlContent)

		// 保存最终HTML
		os.WriteFile("output_final.html", []byte(finalHTML), 0644)
	}

	// Step 4: 转换 HTML 为 PPT
	slog.Info("Step 4: 转换 HTML 为 PPT...")
	pres := genppt.FromHTML(finalHTML)

	// 保存PPT文件
	outputFile := *outPath
	if err := pres.WriteFile(outputFile); err != nil {
		slog.Error("保存PPT失败", "err", err)
		return
	}

	slog.Infof("✅ PPT生成成功: %s (共%d页)", outputFile, pres.SlideCount())
}

// generateHTML 让 AI 生成 HTML 代码
func generateHTML(ctx context.Context, topic string) string {
	msg, err := textModel.Generate(ctx, []*schema.Message{
		{Role: schema.System, Content: systemPrompt},
		{Role: schema.User, Content: "请为以下主题生成一份演示文稿HTML：" + topic},
	})
	if err != nil {
		slog.Errorf("HTML生成失败: %v", err)
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
	slog.Infof("找到 %d 张图片需要生成", validMatches)

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
			slog.Infof("图片已存在 src，跳过生成: %s", truncateString(srcVal, 30))
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
		slog.Infof("正在生成图片 %d/%d: %s", count, validMatches, truncateString(prompt, 30))

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
			slog.Infof("图片 %d 生成成功", count)
		} else {
			slog.Warnf("图片 %d 生成失败", count)
		}
	}

	return result
}

// generateImage 使用 AI 生成图片，返回 data URI
func generateImage(ctx context.Context, prompt string) string {
	if imageModel == nil {
		slog.Warn("Image model not initialized, skipping image generation")
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
		slog.Warnf("图片生成API调用失败: %v", err)
		return ""
	}

	// 提取生成的图片数据
	for _, part := range msg.AssistantGenMultiContent {
		if part.Image != nil && part.Image.Base64Data != nil {
			// 验证 base64 是否有效
			_, err := base64.StdEncoding.DecodeString(*part.Image.Base64Data)
			if err != nil {
				slog.Warnf("Base64数据无效: %v", err)
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
		slog.Warnf("图片解码失败: %v", err)
		return
	}

	// 保存文件
	filename := fmt.Sprintf("image_%d.%s", index, ext)
	if err := os.WriteFile(filename, data, 0644); err != nil {
		slog.Warnf("保存图片失败: %v", err)
		return
	}
	slog.Infof("已保存图片: %s", filename)
}
