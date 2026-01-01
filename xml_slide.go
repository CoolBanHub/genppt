package genppt

import (
	"strings"
)

// generateSlide 生成幻灯片XML
func (s *Slide) generateSlide() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">`)

	sb.WriteString(`<p:cSld>`)

	// 背景
	if s.background != nil {
		sb.WriteString(s.generateBackground())
	}

	// 形状树
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

	// 生成各个对象
	objectId := 2 // 从2开始，1已被组使用
	for _, obj := range s.objects {
		switch o := obj.(type) {
		case *textObject:
			sb.WriteString(s.generateTextBox(o, objectId))
			objectId++
		case *shapeObject:
			sb.WriteString(s.generateShape(o, objectId))
			objectId++
		case *tableObject:
			sb.WriteString(s.generateTable(o, objectId))
			objectId++
		case *imageObject:
			sb.WriteString(s.generateImage(o, objectId))
			objectId++
		case *chartObject:
			sb.WriteString(s.generateChart(o, objectId))
			objectId++
		case *videoObject:
			sb.WriteString(s.generateVideo(o, objectId))
			objectId++
		case *audioObject:
			sb.WriteString(s.generateAudio(o, objectId))
			objectId++
		}
	}

	sb.WriteString(`</p:spTree>`)
	sb.WriteString(`</p:cSld>`)
	sb.WriteString(`<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`)

	// 生成时间轴（用于自动播放媒体）
	timing := s.generateTiming()
	if timing != "" {
		sb.WriteString(timing)
	}

	sb.WriteString(`</p:sld>`)

	return sb.String()
}

// generateTiming 生成时间轴XML（用于自动播放媒体）
func (s *Slide) generateTiming() string {
	// 收集需要自动播放的媒体对象
	var autoPlayMedia []struct {
		id      int
		isVideo bool
		loop    bool
	}

	objectId := 2
	for _, obj := range s.objects {
		switch o := obj.(type) {
		case *videoObject:
			if o.options.AutoPlay {
				autoPlayMedia = append(autoPlayMedia, struct {
					id      int
					isVideo bool
					loop    bool
				}{objectId, true, o.options.Loop})
			}
			objectId++
		case *audioObject:
			if o.options.AutoPlay {
				autoPlayMedia = append(autoPlayMedia, struct {
					id      int
					isVideo bool
					loop    bool
				}{objectId, false, o.options.Loop})
			}
			objectId++
		default:
			objectId++
		}
	}

	if len(autoPlayMedia) == 0 {
		return ""
	}

	var sb strings.Builder
	sb.WriteString(`<p:timing>`)
	sb.WriteString(`<p:tnLst>`)
	sb.WriteString(`<p:par>`)
	sb.WriteString(`<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">`)
	sb.WriteString(`<p:childTnLst>`)
	sb.WriteString(`<p:seq concurrent="1" nextAc="seek">`)
	sb.WriteString(`<p:cTn id="2" dur="indefinite" nodeType="mainSeq">`)
	sb.WriteString(`<p:childTnLst>`)

	tnId := 3
	for i, media := range autoPlayMedia {
		sb.WriteString(`<p:par>`)
		sb.WriteString(`<p:cTn id="`)
		sb.WriteString(itoa(tnId))
		tnId++
		sb.WriteString(`" fill="hold">`)
		sb.WriteString(`<p:stCondLst>`)
		if i == 0 {
			// 第一个媒体在幻灯片显示时开始
			sb.WriteString(`<p:cond delay="0"/>`)
		} else {
			// 后续媒体也同时开始
			sb.WriteString(`<p:cond delay="0"/>`)
		}
		sb.WriteString(`</p:stCondLst>`)
		sb.WriteString(`<p:childTnLst>`)
		sb.WriteString(`<p:par>`)
		sb.WriteString(`<p:cTn id="`)
		sb.WriteString(itoa(tnId))
		tnId++
		sb.WriteString(`" fill="hold">`)
		sb.WriteString(`<p:stCondLst>`)
		sb.WriteString(`<p:cond delay="0"/>`)
		sb.WriteString(`</p:stCondLst>`)
		sb.WriteString(`<p:childTnLst>`)
		sb.WriteString(`<p:par>`)
		sb.WriteString(`<p:cTn id="`)
		sb.WriteString(itoa(tnId))
		tnId++
		sb.WriteString(`" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">`)
		sb.WriteString(`<p:stCondLst>`)
		sb.WriteString(`<p:cond delay="0"/>`)
		sb.WriteString(`</p:stCondLst>`)
		sb.WriteString(`<p:childTnLst>`)

		// 媒体命令
		sb.WriteString(`<p:cmd type="call" cmd="playFrom(0.0)">`)
		sb.WriteString(`<p:cBhvr>`)
		sb.WriteString(`<p:cTn id="`)
		sb.WriteString(itoa(tnId))
		tnId++
		sb.WriteString(`" dur="1" fill="hold"/>`)
		sb.WriteString(`<p:tgtEl>`)
		sb.WriteString(`<p:spTgt spid="`)
		sb.WriteString(itoa(media.id))
		sb.WriteString(`"/>`)
		sb.WriteString(`</p:tgtEl>`)
		sb.WriteString(`</p:cBhvr>`)
		sb.WriteString(`</p:cmd>`)

		sb.WriteString(`</p:childTnLst>`)
		sb.WriteString(`</p:cTn>`)
		sb.WriteString(`</p:par>`)
		sb.WriteString(`</p:childTnLst>`)
		sb.WriteString(`</p:cTn>`)
		sb.WriteString(`</p:par>`)
		sb.WriteString(`</p:childTnLst>`)
		sb.WriteString(`</p:cTn>`)
		sb.WriteString(`</p:par>`)
	}

	sb.WriteString(`</p:childTnLst>`)
	sb.WriteString(`</p:cTn>`)
	sb.WriteString(`<p:prevCondLst>`)
	sb.WriteString(`<p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>`)
	sb.WriteString(`</p:prevCondLst>`)
	sb.WriteString(`<p:nextCondLst>`)
	sb.WriteString(`<p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>`)
	sb.WriteString(`</p:nextCondLst>`)
	sb.WriteString(`</p:seq>`)
	sb.WriteString(`</p:childTnLst>`)
	sb.WriteString(`</p:cTn>`)
	sb.WriteString(`</p:par>`)
	sb.WriteString(`</p:tnLst>`)
	sb.WriteString(`</p:timing>`)

	return sb.String()
}

// generateBackground 生成背景
func (s *Slide) generateBackground() string {
	var sb strings.Builder
	sb.WriteString(`<p:bg>`)
	sb.WriteString(`<p:bgPr>`)

	if s.background.Color != "" {
		sb.WriteString(`<a:solidFill>`)
		sb.WriteString(`<a:srgbClr val="`)
		sb.WriteString(ParseColor(s.background.Color))
		sb.WriteString(`"/>`)
		sb.WriteString(`</a:solidFill>`)
	}

	sb.WriteString(`<a:effectLst/>`)
	sb.WriteString(`</p:bgPr>`)
	sb.WriteString(`</p:bg>`)
	return sb.String()
}

// generateTextBox 生成文本框
func (s *Slide) generateTextBox(t *textObject, id int) string {
	var sb strings.Builder

	x := InchToEMU(t.options.X)
	y := InchToEMU(t.options.Y)
	cx := InchToEMU(defaultIfZero(t.options.Width, 4))
	cy := InchToEMU(defaultIfZero(t.options.Height, 0.5))

	sb.WriteString(`<p:sp>`)
	sb.WriteString(`<p:nvSpPr>`)
	sb.WriteString(`<p:cNvPr id="`)
	sb.WriteString(itoa(id))
	sb.WriteString(`" name="TextBox `)
	sb.WriteString(itoa(id))
	sb.WriteString(`"/>`)
	sb.WriteString(`<p:cNvSpPr txBox="1"/>`)
	sb.WriteString(`<p:nvPr/>`)
	sb.WriteString(`</p:nvSpPr>`)

	// 形状属性
	sb.WriteString(`<p:spPr>`)
	sb.WriteString(`<a:xfrm`)
	if t.options.Rotate != 0 {
		sb.WriteString(` rot="`)
		sb.WriteString(itoa(int(t.options.Rotate * 60000)))
		sb.WriteString(`"`)
	}
	sb.WriteString(`>`)
	sb.WriteString(`<a:off x="`)
	sb.WriteString(itoa(int(x)))
	sb.WriteString(`" y="`)
	sb.WriteString(itoa(int(y)))
	sb.WriteString(`"/>`)
	sb.WriteString(`<a:ext cx="`)
	sb.WriteString(itoa(int(cx)))
	sb.WriteString(`" cy="`)
	sb.WriteString(itoa(int(cy)))
	sb.WriteString(`"/>`)
	sb.WriteString(`</a:xfrm>`)
	sb.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`)
	if t.options.Fill != "" {
		sb.WriteString(`<a:solidFill>`)
		sb.WriteString(`<a:srgbClr val="`)
		sb.WriteString(ParseColor(t.options.Fill))
		sb.WriteString(`"/>`)
		sb.WriteString(`</a:solidFill>`)
	} else {
		sb.WriteString(`<a:noFill/>`)
	}
	sb.WriteString(`</p:spPr>`)

	// 文本框
	sb.WriteString(`<p:txBody>`)
	sb.WriteString(`<a:bodyPr wrap="square" rtlCol="0"`)
	// 垂直对齐
	if t.options.VAlign != "" {
		sb.WriteString(` anchor="`)
		sb.WriteString(string(t.options.VAlign))
		sb.WriteString(`"`)
	}
	sb.WriteString(`/>`)
	sb.WriteString(`<a:lstStyle/>`)

	// 段落
	sb.WriteString(`<a:p>`)
	sb.WriteString(`<a:pPr`)
	if t.options.Align != "" {
		sb.WriteString(` algn="`)
		sb.WriteString(string(t.options.Align))
		sb.WriteString(`"`)
	}
	sb.WriteString(`>`)
	if t.options.LineSpacing > 0 {
		sb.WriteString(`<a:lnSpc><a:spcPct val="`)
		sb.WriteString(itoa(int(t.options.LineSpacing * 100000))) // 100% = 100000
		sb.WriteString(`"/></a:lnSpc>`)
	}
	sb.WriteString(`</a:pPr>`)

	// 文本运行
	sb.WriteString(`<a:r>`)
	sb.WriteString(`<a:rPr lang="zh-CN"`)
	if t.options.FontSize > 0 {
		sb.WriteString(` sz="`)
		sb.WriteString(itoa(int(t.options.FontSize * 100)))
		sb.WriteString(`"`)
	}
	if t.options.Bold {
		sb.WriteString(` b="1"`)
	}
	if t.options.Italic {
		sb.WriteString(` i="1"`)
	}
	if t.options.Underline {
		sb.WriteString(` u="sng"`)
	}
	if t.options.CharSpacing != 0 {
		sb.WriteString(` spc="`)
		sb.WriteString(itoa(int(t.options.CharSpacing * 100))) // 1pt = 100
		sb.WriteString(`"`)
	}
	sb.WriteString(`>`)

	// 字体颜色
	if t.options.FontColor != "" {
		sb.WriteString(`<a:solidFill>`)
		sb.WriteString(`<a:srgbClr val="`)
		sb.WriteString(ParseColor(t.options.FontColor))
		sb.WriteString(`"/>`)
		sb.WriteString(`</a:solidFill>`)
	}

	// 字体
	if t.options.FontFace != "" {
		sb.WriteString(`<a:latin typeface="`)
		sb.WriteString(escapeXML(t.options.FontFace))
		sb.WriteString(`"/>`)
		sb.WriteString(`<a:ea typeface="`)
		sb.WriteString(escapeXML(t.options.FontFace))
		sb.WriteString(`"/>`)
	}

	sb.WriteString(`</a:rPr>`)
	sb.WriteString(`<a:t>`)
	sb.WriteString(escapeXML(t.text))
	sb.WriteString(`</a:t>`)
	sb.WriteString(`</a:r>`)

	sb.WriteString(`<a:endParaRPr lang="zh-CN"/>`)
	sb.WriteString(`</a:p>`)
	sb.WriteString(`</p:txBody>`)
	sb.WriteString(`</p:sp>`)

	return sb.String()
}

// getShapePreset 获取形状预设名称
func getShapePreset(shapeType ShapeType) string {
	switch shapeType {
	case ShapeRect:
		return "rect"
	case ShapeRoundRect:
		return "roundRect"
	case ShapeEllipse:
		return "ellipse"
	case ShapeTriangle:
		return "triangle"
	case ShapeDiamond:
		return "diamond"
	case ShapeArrowRight:
		return "rightArrow"
	case ShapeArrowLeft:
		return "leftArrow"
	case ShapeArrowUp:
		return "upArrow"
	case ShapeArrowDown:
		return "downArrow"
	case ShapeStar5:
		return "star5"
	case ShapeHeart:
		return "heart"
	case ShapeLine:
		return "line"
	default:
		return "rect"
	}
}

// generateShape 生成形状
func (s *Slide) generateShape(sh *shapeObject, id int) string {
	var sb strings.Builder

	x := InchToEMU(sh.options.X)
	y := InchToEMU(sh.options.Y)
	cx := InchToEMU(defaultIfZero(sh.options.Width, 2))
	cy := InchToEMU(defaultIfZero(sh.options.Height, 1))

	sb.WriteString(`<p:sp>`)
	sb.WriteString(`<p:nvSpPr>`)
	sb.WriteString(`<p:cNvPr id="`)
	sb.WriteString(itoa(id))
	sb.WriteString(`" name="Shape `)
	sb.WriteString(itoa(id))
	sb.WriteString(`"/>`)
	sb.WriteString(`<p:cNvSpPr/>`)
	sb.WriteString(`<p:nvPr/>`)
	sb.WriteString(`</p:nvSpPr>`)

	// 形状属性
	sb.WriteString(`<p:spPr>`)
	sb.WriteString(`<a:xfrm`)
	if sh.options.Rotate != 0 {
		sb.WriteString(` rot="`)
		sb.WriteString(itoa(int(sh.options.Rotate * 60000)))
		sb.WriteString(`"`)
	}
	sb.WriteString(`>`)
	sb.WriteString(`<a:off x="`)
	sb.WriteString(itoa(int(x)))
	sb.WriteString(`" y="`)
	sb.WriteString(itoa(int(y)))
	sb.WriteString(`"/>`)
	sb.WriteString(`<a:ext cx="`)
	sb.WriteString(itoa(int(cx)))
	sb.WriteString(`" cy="`)
	sb.WriteString(itoa(int(cy)))
	sb.WriteString(`"/>`)
	sb.WriteString(`</a:xfrm>`)

	// 形状类型
	sb.WriteString(`<a:prstGeom prst="`)
	sb.WriteString(getShapePreset(sh.shapeType))
	sb.WriteString(`"><a:avLst/></a:prstGeom>`)

	// 填充
	if sh.options.Fill != "" {
		sb.WriteString(`<a:solidFill>`)
		sb.WriteString(`<a:srgbClr val="`)
		sb.WriteString(ParseColor(sh.options.Fill))
		sb.WriteString(`"`)
		if sh.options.Transparency > 0 {
			sb.WriteString(`><a:alpha val="`)
			sb.WriteString(itoa(int((100 - sh.options.Transparency) * 1000)))
			sb.WriteString(`"/></a:srgbClr>`)
		} else {
			sb.WriteString(`/>`)
		}
		sb.WriteString(`</a:solidFill>`)
	} else {
		sb.WriteString(`<a:noFill/>`)
	}

	// 边框
	if sh.options.LineColor != "" && sh.options.LineWidth > 0 {
		sb.WriteString(`<a:ln w="`)
		sb.WriteString(itoa(int(sh.options.LineWidth * 12700)))
		sb.WriteString(`">`)
		sb.WriteString(`<a:solidFill>`)
		sb.WriteString(`<a:srgbClr val="`)
		sb.WriteString(ParseColor(sh.options.LineColor))
		sb.WriteString(`"/>`)
		sb.WriteString(`</a:solidFill>`)
		sb.WriteString(`</a:ln>`)
	}

	// 阴影效果
	if sh.options.Shadow {
		sb.WriteString(`<a:effectLst>`)
		sb.WriteString(`<a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0">`)
		sb.WriteString(`<a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr>`)
		sb.WriteString(`</a:outerShdw>`)
		sb.WriteString(`</a:effectLst>`)
	}

	sb.WriteString(`</p:spPr>`)

	// 如果有文本
	if sh.text != "" {
		sb.WriteString(`<p:txBody>`)
		sb.WriteString(`<a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/>`)
		sb.WriteString(`<a:lstStyle/>`)
		sb.WriteString(`<a:p>`)
		sb.WriteString(`<a:pPr algn="ctr"/>`)
		sb.WriteString(`<a:r>`)
		sb.WriteString(`<a:rPr lang="zh-CN" sz="1800">`)
		sb.WriteString(`<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`)
		sb.WriteString(`</a:rPr>`)
		sb.WriteString(`<a:t>`)
		sb.WriteString(escapeXML(sh.text))
		sb.WriteString(`</a:t>`)
		sb.WriteString(`</a:r>`)
		sb.WriteString(`<a:endParaRPr lang="zh-CN"/>`)
		sb.WriteString(`</a:p>`)
		sb.WriteString(`</p:txBody>`)
	}

	sb.WriteString(`</p:sp>`)

	return sb.String()
}

// generateTable 生成表格
func (s *Slide) generateTable(t *tableObject, id int) string {
	var sb strings.Builder

	x := InchToEMU(t.options.X)
	y := InchToEMU(t.options.Y)
	cx := InchToEMU(defaultIfZero(t.options.Width, 8))

	// 计算行高和列宽
	numRows := len(t.rows)
	numCols := 0
	if numRows > 0 {
		numCols = len(t.rows[0])
	}
	if numCols == 0 {
		return ""
	}

	// 默认行高
	rowHeight := InchToEMU(0.4)
	if len(t.options.RowHeights) > 0 {
		rowHeight = InchToEMU(t.options.RowHeights[0])
	}

	// 计算总高度
	cy := rowHeight * int64(numRows)

	// 列宽
	colWidths := make([]int64, numCols)
	if len(t.options.ColWidths) >= numCols {
		for i := 0; i < numCols; i++ {
			colWidths[i] = InchToEMU(t.options.ColWidths[i])
		}
	} else {
		avgWidth := cx / int64(numCols)
		for i := 0; i < numCols; i++ {
			colWidths[i] = avgWidth
		}
	}

	sb.WriteString(`<p:graphicFrame>`)
	sb.WriteString(`<p:nvGraphicFramePr>`)
	sb.WriteString(`<p:cNvPr id="`)
	sb.WriteString(itoa(id))
	sb.WriteString(`" name="Table `)
	sb.WriteString(itoa(id))
	sb.WriteString(`"/>`)
	sb.WriteString(`<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>`)
	sb.WriteString(`<p:nvPr/>`)
	sb.WriteString(`</p:nvGraphicFramePr>`)

	sb.WriteString(`<p:xfrm>`)
	sb.WriteString(`<a:off x="`)
	sb.WriteString(itoa(int(x)))
	sb.WriteString(`" y="`)
	sb.WriteString(itoa(int(y)))
	sb.WriteString(`"/>`)
	sb.WriteString(`<a:ext cx="`)
	sb.WriteString(itoa(int(cx)))
	sb.WriteString(`" cy="`)
	sb.WriteString(itoa(int(cy)))
	sb.WriteString(`"/>`)
	sb.WriteString(`</p:xfrm>`)

	sb.WriteString(`<a:graphic>`)
	sb.WriteString(`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">`)
	sb.WriteString(`<a:tbl>`)

	// 表格属性
	sb.WriteString(`<a:tblPr firstRow="1" bandRow="1">`)
	sb.WriteString(`<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>`)
	sb.WriteString(`</a:tblPr>`)

	// 表格网格
	sb.WriteString(`<a:tblGrid>`)
	for _, w := range colWidths {
		sb.WriteString(`<a:gridCol w="`)
		sb.WriteString(itoa(int(w)))
		sb.WriteString(`"/>`)
	}
	sb.WriteString(`</a:tblGrid>`)

	// 表格行
	for rowIdx, row := range t.rows {
		rh := rowHeight
		if rowIdx < len(t.options.RowHeights) {
			rh = InchToEMU(t.options.RowHeights[rowIdx])
		}

		sb.WriteString(`<a:tr h="`)
		sb.WriteString(itoa(int(rh)))
		sb.WriteString(`">`)

		for colIdx, cell := range row {
			if colIdx >= numCols {
				break
			}
			sb.WriteString(s.generateTableCell(&cell, t, rowIdx))
		}

		sb.WriteString(`</a:tr>`)
	}

	sb.WriteString(`</a:tbl>`)
	sb.WriteString(`</a:graphicData>`)
	sb.WriteString(`</a:graphic>`)
	sb.WriteString(`</p:graphicFrame>`)

	return sb.String()
}

// generateTableCell 生成表格单元格
func (s *Slide) generateTableCell(cell *TableCell, t *tableObject, rowIdx int) string {
	var sb strings.Builder

	sb.WriteString(`<a:tc`)
	if cell.ColSpan > 1 {
		sb.WriteString(` gridSpan="`)
		sb.WriteString(itoa(cell.ColSpan))
		sb.WriteString(`"`)
	}
	if cell.RowSpan > 1 {
		sb.WriteString(` rowSpan="`)
		sb.WriteString(itoa(cell.RowSpan))
		sb.WriteString(`"`)
	}
	sb.WriteString(`>`)

	// 文本体
	sb.WriteString(`<a:txBody>`)
	sb.WriteString(`<a:bodyPr/>`)
	sb.WriteString(`<a:lstStyle/>`)
	sb.WriteString(`<a:p>`)

	// 段落属性
	align := cell.Align
	if align == "" {
		align = AlignLeft
	}
	sb.WriteString(`<a:pPr algn="`)
	sb.WriteString(string(align))
	sb.WriteString(`"/>`)

	// 文本运行
	sb.WriteString(`<a:r>`)
	sb.WriteString(`<a:rPr lang="zh-CN"`)

	// 字号
	fontSize := cell.FontSize
	if fontSize == 0 {
		fontSize = t.options.FontSize
	}
	if fontSize > 0 {
		sb.WriteString(` sz="`)
		sb.WriteString(itoa(int(fontSize * 100)))
		sb.WriteString(`"`)
	}

	// 粗体
	isBold := cell.Bold || (rowIdx == 0 && t.options.FirstRowBold)
	if isBold {
		sb.WriteString(` b="1"`)
	}
	if cell.Italic {
		sb.WriteString(` i="1"`)
	}
	sb.WriteString(`>`)

	// 字体颜色
	fontColor := cell.FontColor
	if fontColor == "" {
		fontColor = t.options.FontColor
	}
	if fontColor != "" {
		sb.WriteString(`<a:solidFill>`)
		sb.WriteString(`<a:srgbClr val="`)
		sb.WriteString(ParseColor(fontColor))
		sb.WriteString(`"/>`)
		sb.WriteString(`</a:solidFill>`)
	}

	// 字体
	fontFace := cell.FontFace
	if fontFace == "" {
		fontFace = t.options.FontFace
	}
	if fontFace != "" {
		sb.WriteString(`<a:latin typeface="`)
		sb.WriteString(escapeXML(fontFace))
		sb.WriteString(`"/>`)
		sb.WriteString(`<a:ea typeface="`)
		sb.WriteString(escapeXML(fontFace))
		sb.WriteString(`"/>`)
	}

	sb.WriteString(`</a:rPr>`)
	sb.WriteString(`<a:t>`)
	sb.WriteString(escapeXML(cell.Text))
	sb.WriteString(`</a:t>`)
	sb.WriteString(`</a:r>`)
	sb.WriteString(`<a:endParaRPr lang="zh-CN"/>`)
	sb.WriteString(`</a:p>`)
	sb.WriteString(`</a:txBody>`)

	// 单元格属性
	sb.WriteString(`<a:tcPr`)
	// 垂直对齐
	vAlign := cell.VAlign
	if vAlign == "" {
		vAlign = VAlignMiddle
	}
	sb.WriteString(` anchor="`)
	sb.WriteString(string(vAlign))
	sb.WriteString(`">`)

	// 单元格填充
	fillColor := cell.Fill
	if fillColor == "" {
		if rowIdx == 0 && t.options.FirstRowFill != "" {
			fillColor = t.options.FirstRowFill
		} else if t.options.Fill != "" {
			fillColor = t.options.Fill
		}
	}
	if fillColor != "" {
		sb.WriteString(`<a:solidFill>`)
		sb.WriteString(`<a:srgbClr val="`)
		sb.WriteString(ParseColor(fillColor))
		sb.WriteString(`"/>`)
		sb.WriteString(`</a:solidFill>`)
	}

	// 边框
	if t.options.Border.Color != "" {
		borderWidth := int(t.options.Border.Width * 12700)
		borderColor := ParseColor(t.options.Border.Color)

		borders := []string{"lnL", "lnR", "lnT", "lnB"}
		for _, b := range borders {
			sb.WriteString(`<a:`)
			sb.WriteString(b)
			sb.WriteString(` w="`)
			sb.WriteString(itoa(borderWidth))
			sb.WriteString(`" cap="flat" cmpd="sng" algn="ctr">`)
			sb.WriteString(`<a:solidFill><a:srgbClr val="`)
			sb.WriteString(borderColor)
			sb.WriteString(`"/></a:solidFill>`)
			sb.WriteString(`<a:prstDash val="solid"/>`)
			sb.WriteString(`</a:`)
			sb.WriteString(b)
			sb.WriteString(`>`)
		}
	}

	sb.WriteString(`</a:tcPr>`)
	sb.WriteString(`</a:tc>`)

	return sb.String()
}

// generateImage 生成图片
func (s *Slide) generateImage(img *imageObject, id int) string {
	var sb strings.Builder

	x := InchToEMU(img.options.X)
	y := InchToEMU(img.options.Y)
	cx := InchToEMU(defaultIfZero(img.options.Width, 4))
	cy := InchToEMU(defaultIfZero(img.options.Height, 3))

	sb.WriteString(`<p:pic>`)
	sb.WriteString(`<p:nvPicPr>`)
	sb.WriteString(`<p:cNvPr id="`)
	sb.WriteString(itoa(id))
	sb.WriteString(`" name="Picture `)
	sb.WriteString(itoa(id))
	sb.WriteString(`"`)
	if img.options.AltText != "" {
		sb.WriteString(` descr="`)
		sb.WriteString(escapeXML(img.options.AltText))
		sb.WriteString(`"`)
	}
	sb.WriteString(`/>`)
	sb.WriteString(`<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>`)
	sb.WriteString(`<p:nvPr/>`)
	sb.WriteString(`</p:nvPicPr>`)

	// 图片填充
	sb.WriteString(`<p:blipFill>`)
	sb.WriteString(`<a:blip r:embed="`)
	sb.WriteString(img.rID)
	sb.WriteString(`"/>`)
	sb.WriteString(`<a:stretch><a:fillRect/></a:stretch>`)
	sb.WriteString(`</p:blipFill>`)

	// 形状属性
	sb.WriteString(`<p:spPr>`)
	sb.WriteString(`<a:xfrm`)
	if img.options.Rotate != 0 {
		sb.WriteString(` rot="`)
		sb.WriteString(itoa(int(img.options.Rotate * 60000)))
		sb.WriteString(`"`)
	}
	sb.WriteString(`>`)
	sb.WriteString(`<a:off x="`)
	sb.WriteString(itoa(int(x)))
	sb.WriteString(`" y="`)
	sb.WriteString(itoa(int(y)))
	sb.WriteString(`"/>`)
	sb.WriteString(`<a:ext cx="`)
	sb.WriteString(itoa(int(cx)))
	sb.WriteString(`" cy="`)
	sb.WriteString(itoa(int(cy)))
	sb.WriteString(`"/>`)
	sb.WriteString(`</a:xfrm>`)

	// 如果有圆角
	if img.options.Rounding != 0 {
		sb.WriteString(`<a:prstGeom prst="roundRect">`)
		sb.WriteString(`<a:avLst>`)

		// PowerPoint 的圆角 adj 值是相对于最小边的
		// 100000 表示 100% 的最小边（即完全圆形）
		// 实际上 adj 值通常在 0-50000 之间，50000 代表半径为短边的一半
		minSide := cx
		if cy < cx {
			minSide = cy
		}
		adj := 0
		if img.options.Rounding < 0 {
			// 百分比情况：-0.5 = 50% -> 50000
			adj = int(-img.options.Rounding * 100000)
		} else {
			// 英寸情况
			radiusEMU := InchToEMU(img.options.Rounding)
			if minSide > 0 {
				adj = int((radiusEMU * 100000) / minSide)
			}
		}
		if adj > 50000 {
			adj = 50000
		}

		sb.WriteString(`<a:gd name="adj" fmla="val `)
		sb.WriteString(itoa(adj))
		sb.WriteString(`"/>`)
		sb.WriteString(`</a:avLst>`)
		sb.WriteString(`</a:prstGeom>`)
	} else {
		sb.WriteString(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`)
	}

	sb.WriteString(`</p:spPr>`)
	sb.WriteString(`</p:pic>`)

	return sb.String()
}
