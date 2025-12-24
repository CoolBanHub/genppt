package genppt

import (
	"strings"
	"time"
)

// generatePresentation 生成 ppt/presentation.xml
func (p *Presentation) generatePresentation() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">`)

	// 幻灯片母版ID列表
	sb.WriteString(`<p:sldMasterIdLst>`)
	sb.WriteString(`<p:sldMasterId id="2147483648" r:id="rId1"/>`)
	sb.WriteString(`</p:sldMasterIdLst>`)

	// 幻灯片ID列表
	sb.WriteString(`<p:sldIdLst>`)
	for i := range p.slides {
		sb.WriteString(`<p:sldId id="`)
		sb.WriteString(itoa(256 + i))
		sb.WriteString(`" r:id="rId`)
		sb.WriteString(itoa(2 + i)) // rId2开始是幻灯片
		sb.WriteString(`"/>`)
	}
	sb.WriteString(`</p:sldIdLst>`)

	// 幻灯片尺寸
	sb.WriteString(`<p:sldSz cx="`)
	sb.WriteString(itoa(int(p.slideWidth)))
	sb.WriteString(`" cy="`)
	sb.WriteString(itoa(int(p.slideHeight)))
	sb.WriteString(`" type="custom"/>`)

	// 备注尺寸
	sb.WriteString(`<p:notesSz cx="6858000" cy="9144000"/>`)

	// 默认文本样式
	sb.WriteString(`<p:defaultTextStyle>`)
	sb.WriteString(`<a:defPPr>`)
	sb.WriteString(`<a:defRPr lang="zh-CN"/>`)
	sb.WriteString(`</a:defPPr>`)
	sb.WriteString(`</p:defaultTextStyle>`)

	sb.WriteString(`</p:presentation>`)
	return sb.String()
}

// generatePresProps 生成 ppt/presProps.xml
func generatePresProps() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:extLst>
<p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}">
<p14:discardImageEditData xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="0"/>
</p:ext>
</p:extLst>
</p:presentationPr>`
}

// generateViewProps 生成 ppt/viewProps.xml
func generateViewProps() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:normalViewPr>
<p:restoredLeft sz="15620"/>
<p:restoredTop sz="94660"/>
</p:normalViewPr>
<p:slideViewPr>
<p:cSldViewPr>
<p:cViewPr varScale="1">
<p:scale>
<a:sx n="100" d="100"/>
<a:sy n="100" d="100"/>
</p:scale>
<p:origin x="0" y="0"/>
</p:cViewPr>
</p:cSldViewPr>
</p:slideViewPr>
</p:viewPr>`
}

// generateTableStyles 生成 ppt/tableStyles.xml
func generateTableStyles() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`
}

// generateCoreProps 生成 docProps/core.xml
func (p *Presentation) generateCoreProps() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`)

	if p.title != "" {
		sb.WriteString(`<dc:title>`)
		sb.WriteString(escapeXML(p.title))
		sb.WriteString(`</dc:title>`)
	}
	if p.subject != "" {
		sb.WriteString(`<dc:subject>`)
		sb.WriteString(escapeXML(p.subject))
		sb.WriteString(`</dc:subject>`)
	}
	if p.author != "" {
		sb.WriteString(`<dc:creator>`)
		sb.WriteString(escapeXML(p.author))
		sb.WriteString(`</dc:creator>`)
		sb.WriteString(`<cp:lastModifiedBy>`)
		sb.WriteString(escapeXML(p.author))
		sb.WriteString(`</cp:lastModifiedBy>`)
	}

	now := formatDateTime(time.Now())
	sb.WriteString(`<dcterms:created xsi:type="dcterms:W3CDTF">`)
	sb.WriteString(now)
	sb.WriteString(`</dcterms:created>`)
	sb.WriteString(`<dcterms:modified xsi:type="dcterms:W3CDTF">`)
	sb.WriteString(now)
	sb.WriteString(`</dcterms:modified>`)

	sb.WriteString(`</cp:coreProperties>`)
	return sb.String()
}

// generateAppProps 生成 docProps/app.xml
func (p *Presentation) generateAppProps() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">`)
	sb.WriteString(`<TotalTime>0</TotalTime>`)
	sb.WriteString(`<Words>0</Words>`)
	sb.WriteString(`<Application>genppt/1.0</Application>`)
	sb.WriteString(`<PresentationFormat>自定义</PresentationFormat>`)
	sb.WriteString(`<Paragraphs>0</Paragraphs>`)
	sb.WriteString(`<Slides>`)
	sb.WriteString(itoa(len(p.slides)))
	sb.WriteString(`</Slides>`)
	sb.WriteString(`<Notes>0</Notes>`)
	sb.WriteString(`<HiddenSlides>0</HiddenSlides>`)
	sb.WriteString(`<MMClips>0</MMClips>`)
	sb.WriteString(`<ScaleCrop>false</ScaleCrop>`)

	if p.company != "" {
		sb.WriteString(`<Company>`)
		sb.WriteString(escapeXML(p.company))
		sb.WriteString(`</Company>`)
	}

	sb.WriteString(`<LinksUpToDate>false</LinksUpToDate>`)
	sb.WriteString(`<SharedDoc>false</SharedDoc>`)
	sb.WriteString(`<HyperlinksChanged>false</HyperlinksChanged>`)
	sb.WriteString(`<AppVersion>16.0000</AppVersion>`)
	sb.WriteString(`</Properties>`)
	return sb.String()
}
