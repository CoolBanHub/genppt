package genppt

import (
	"strings"
)

// generateContentTypes 生成 [Content_Types].xml
func (p *Presentation) generateContentTypes() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`)

	// 默认类型
	sb.WriteString(`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
	sb.WriteString(`<Default Extension="xml" ContentType="application/xml"/>`)

	// 媒体类型（图片和视频）
	mediaExts := make(map[string]bool)
	for _, media := range p.mediaFiles {
		if !mediaExts[media.ext] {
			mediaExts[media.ext] = true
			sb.WriteString(`<Default Extension="`)
			sb.WriteString(media.ext)
			sb.WriteString(`" ContentType="`)
			sb.WriteString(getMediaMIME(media.ext))
			sb.WriteString(`"/>`)
		}
	}

	// Override类型
	sb.WriteString(`<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`)
	sb.WriteString(`<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>`)
	sb.WriteString(`<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>`)
	sb.WriteString(`<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>`)
	sb.WriteString(`<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`)
	sb.WriteString(`<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`)
	sb.WriteString(`<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`)
	sb.WriteString(`<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`)
	sb.WriteString(`<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`)

	// 幻灯片
	for i := range p.slides {
		sb.WriteString(`<Override PartName="/ppt/slides/slide`)
		sb.WriteString(itoa(i + 1))
		sb.WriteString(`.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`)
	}

	// 图表
	chartIdx := 1
	for _, slide := range p.slides {
		for _, obj := range slide.objects {
			if _, ok := obj.(*chartObject); ok {
				sb.WriteString(`<Override PartName="/ppt/charts/chart`)
				sb.WriteString(itoa(chartIdx))
				sb.WriteString(`.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`)
				chartIdx++
			}
		}
	}

	sb.WriteString(`</Types>`)
	return sb.String()
}

// generateRootRels 生成 _rels/.rels
func generateRootRels() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`
}

// generatePresentationRels 生成 ppt/_rels/presentation.xml.rels
func (p *Presentation) generatePresentationRels() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)

	rId := 1
	// 幻灯片母版
	sb.WriteString(`<Relationship Id="rId`)
	sb.WriteString(itoa(rId))
	sb.WriteString(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>`)
	rId++

	// 幻灯片
	for i := range p.slides {
		sb.WriteString(`<Relationship Id="rId`)
		sb.WriteString(itoa(rId))
		sb.WriteString(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide`)
		sb.WriteString(itoa(i + 1))
		sb.WriteString(`.xml"/>`)
		rId++
	}

	// 其他关系
	sb.WriteString(`<Relationship Id="rId`)
	sb.WriteString(itoa(rId))
	sb.WriteString(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>`)
	rId++

	sb.WriteString(`<Relationship Id="rId`)
	sb.WriteString(itoa(rId))
	sb.WriteString(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>`)
	rId++

	sb.WriteString(`<Relationship Id="rId`)
	sb.WriteString(itoa(rId))
	sb.WriteString(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>`)
	rId++

	sb.WriteString(`<Relationship Id="rId`)
	sb.WriteString(itoa(rId))
	sb.WriteString(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`)

	sb.WriteString(`</Relationships>`)
	return sb.String()
}

// generateSlideRels 生成幻灯片关系文件
func (s *Slide) generateSlideRels() string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	sb.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)

	// 幻灯片布局关系
	sb.WriteString(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`)

	// 图片关系
	for _, obj := range s.objects {
		if img, ok := obj.(*imageObject); ok {
			sb.WriteString(`<Relationship Id="`)
			sb.WriteString(img.rID)
			sb.WriteString(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image`)
			// 从rID提取索引
			idx := 1
			for i, m := range s.presentation.mediaFiles {
				if m.rID == img.rID {
					idx = i + 1
					break
				}
			}
			sb.WriteString(itoa(idx))
			sb.WriteString(".")
			sb.WriteString(img.mediaExt)
			sb.WriteString(`"/>`)
		}
	}

	// 图表关系
	for _, obj := range s.objects {
		if chart, ok := obj.(*chartObject); ok {
			rID := "rId" + itoa(200+chart.chartIdx)
			sb.WriteString(`<Relationship Id="`)
			sb.WriteString(rID)
			sb.WriteString(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart`)
			sb.WriteString(itoa(chart.chartIdx))
			sb.WriteString(`.xml"/>`)
		}
	}

	// 视频关系
	for _, obj := range s.objects {
		if video, ok := obj.(*videoObject); ok {
			// 找到视频在media列表中的索引
			videoIdx := 1
			for i, m := range s.presentation.mediaFiles {
				if m.rID == video.rID {
					videoIdx = i + 1
					break
				}
			}
			sb.WriteString(`<Relationship Id="`)
			sb.WriteString(video.rID)
			sb.WriteString(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/video`)
			sb.WriteString(itoa(videoIdx))
			sb.WriteString(".")
			sb.WriteString(video.mediaExt)
			sb.WriteString(`"/>`)

			// 封面图片关系
			if video.posterRID != "" {
				sb.WriteString(`<Relationship Id="`)
				sb.WriteString(video.posterRID)
				sb.WriteString(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/poster`)
				// 找封面图片索引
				for i, m := range s.presentation.mediaFiles {
					if m.rID == video.posterRID {
						sb.WriteString(itoa(i + 1))
						sb.WriteString(".")
						sb.WriteString(m.ext)
						break
					}
				}
				sb.WriteString(`"/>`)
			}
		}
	}

	sb.WriteString(`</Relationships>`)
	return sb.String()
}

// generateSlideMasterRels 生成幻灯片母版关系文件
func generateSlideMasterRels() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`
}

// generateSlideLayoutRels 生成幻灯片布局关系文件
func generateSlideLayoutRels() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`
}
