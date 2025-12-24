package genppt

import (
	"strings"
)

// generateSlideMaster 生成 ppt/slideMasters/slideMaster1.xml
func (p *Presentation) generateSlideMaster() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">`)

	// 通用幻灯片数据
	sb.WriteString(`<p:cSld>`)
	sb.WriteString(`<p:bg>`)
	sb.WriteString(`<p:bgRef idx="1001">`)
	sb.WriteString(`<a:schemeClr val="bg1"/>`)
	sb.WriteString(`</p:bgRef>`)
	sb.WriteString(`</p:bg>`)
	sb.WriteString(`<p:spTree>`)
	sb.WriteString(`<p:nvGrpSpPr>`)
	sb.WriteString(`<p:cNvPr id="1" name=""/>`)
	sb.WriteString(`<p:cNvGrpSpPr/>`)
	sb.WriteString(`<p:nvPr/>`)
	sb.WriteString(`</p:nvGrpSpPr>`)
	sb.WriteString(`<p:grpSpPr>`)
	sb.WriteString(`<a:xfrm>`)
	sb.WriteString(`<a:off x="0" y="0"/>`)
	sb.WriteString(`<a:ext cx="0" cy="0"/>`)
	sb.WriteString(`<a:chOff x="0" y="0"/>`)
	sb.WriteString(`<a:chExt cx="0" cy="0"/>`)
	sb.WriteString(`</a:xfrm>`)
	sb.WriteString(`</p:grpSpPr>`)
	sb.WriteString(`</p:spTree>`)
	sb.WriteString(`</p:cSld>`)

	// 颜色映射
	sb.WriteString(`<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`)

	// 幻灯片布局ID列表
	sb.WriteString(`<p:sldLayoutIdLst>`)
	sb.WriteString(`<p:sldLayoutId id="2147483649" r:id="rId1"/>`)
	sb.WriteString(`</p:sldLayoutIdLst>`)

	// 文本样式
	sb.WriteString(`<p:txStyles>`)
	sb.WriteString(`<p:titleStyle>`)
	sb.WriteString(`<a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`)
	sb.WriteString(`<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>`)
	sb.WriteString(`<a:spcBef><a:spcPct val="0"/></a:spcBef>`)
	sb.WriteString(`<a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr>`)
	sb.WriteString(`</a:lvl1pPr>`)
	sb.WriteString(`</p:titleStyle>`)
	sb.WriteString(`<p:bodyStyle>`)
	sb.WriteString(`<a:lvl1pPr marL="228600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`)
	sb.WriteString(`<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>`)
	sb.WriteString(`<a:spcBef><a:spcPts val="1000"/></a:spcBef>`)
	sb.WriteString(`<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/>`)
	sb.WriteString(`<a:buChar char="•"/>`)
	sb.WriteString(`<a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>`)
	sb.WriteString(`</a:lvl1pPr>`)
	sb.WriteString(`</p:bodyStyle>`)
	sb.WriteString(`<p:otherStyle>`)
	sb.WriteString(`<a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr>`)
	sb.WriteString(`</p:otherStyle>`)
	sb.WriteString(`</p:txStyles>`)

	sb.WriteString(`</p:sldMaster>`)
	return sb.String()
}

// generateSlideLayout 生成 ppt/slideLayouts/slideLayout1.xml (空白布局)
func (p *Presentation) generateSlideLayout() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">`)

	sb.WriteString(`<p:cSld name="空白">`)
	sb.WriteString(`<p:spTree>`)
	sb.WriteString(`<p:nvGrpSpPr>`)
	sb.WriteString(`<p:cNvPr id="1" name=""/>`)
	sb.WriteString(`<p:cNvGrpSpPr/>`)
	sb.WriteString(`<p:nvPr/>`)
	sb.WriteString(`</p:nvGrpSpPr>`)
	sb.WriteString(`<p:grpSpPr>`)
	sb.WriteString(`<a:xfrm>`)
	sb.WriteString(`<a:off x="0" y="0"/>`)
	sb.WriteString(`<a:ext cx="0" cy="0"/>`)
	sb.WriteString(`<a:chOff x="0" y="0"/>`)
	sb.WriteString(`<a:chExt cx="0" cy="0"/>`)
	sb.WriteString(`</a:xfrm>`)
	sb.WriteString(`</p:grpSpPr>`)
	sb.WriteString(`</p:spTree>`)
	sb.WriteString(`</p:cSld>`)

	sb.WriteString(`<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`)

	sb.WriteString(`</p:sldLayout>`)
	return sb.String()
}
