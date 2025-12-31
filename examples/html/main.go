package main

import (
	"fmt"
	"log"

	"github.com/CoolBanHub/genppt"
)

func main() {
	// HTML 内容示例
	htmlContent := `
<!DOCTYPE html>
<html>
<head>
    <title>产品发布会</title>
</head>
<body>
    <h1>产品发布会</h1>
    <p>2024年度新品发布</p>

    <h1>产品亮点</h1>
    <h2>核心功能</h2>
    <ul>
        <li>高性能处理器</li>
        <li>超长续航</li>
        <li>轻薄设计</li>
        <li>智能互联</li>
    </ul>

    <h1>技术规格</h1>
    <table>
        <tr>
            <th>参数</th>
            <th>规格</th>
        </tr>
        <tr>
            <td>处理器</td>
            <td>8核心 3.2GHz</td>
        </tr>
        <tr>
            <td>内存</td>
            <td>16GB DDR5</td>
        </tr>
        <tr>
            <td>存储</td>
            <td>512GB SSD</td>
        </tr>
        <tr>
            <td>电池</td>
            <td>5000mAh</td>
        </tr>
    </table>

    <h1>代码示例</h1>
    <h2>快速开始</h2>
    <pre><code>
// 初始化产品
product := NewProduct()
product.Configure()
product.Start()

fmt.Println("产品已启动！")
    </code></pre>
    <p>更多详情请访问官网。</p>

    <h1>联系我们</h1>
    <h2>获取更多信息</h2>
    <ul>
        <li>官网: www.example.com</li>
        <li>邮箱: contact@example.com</li>
        <li>电话: 400-xxx-xxxx</li>
    </ul>
    <p>感谢您的关注！</p>
</body>
</html>
`

	// 方式1: 使用默认选项
	pres := genppt.FromHTML(htmlContent)
	err := pres.WriteFile("output_default.pptx")
	if err != nil {
		log.Fatalf("保存失败: %v", err)
	}
	fmt.Println("已生成: output_default.pptx")

	// 方式2: 使用自定义选项
	opts := genppt.HTMLOptions{
		TitleFontSize:   48,
		HeadingFontSize: 36,
		BodyFontSize:    20,
		CodeFontSize:    16,
		TitleColor:      "#2E4057",
		HeadingColor:    "#2E4057",
		BodyColor:       "#333333",
		CodeBackground:  "#F0F0F0",
		SlideBackground: "#FFFFFF",
	}

	pres2 := genppt.FromHTMLWithOptions(htmlContent, opts)
	err = pres2.WriteFile("output_custom.pptx")
	if err != nil {
		log.Fatalf("保存失败: %v", err)
	}
	fmt.Println("已生成: output_custom.pptx")

	fmt.Printf("共生成 %d 张幻灯片\n", pres.SlideCount())
}
