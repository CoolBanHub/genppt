package genppt

import (
	"archive/zip"
	"bytes"
	"io"
)

// pptxWriter 负责将演示文稿打包为PPTX文件
type pptxWriter struct {
	pres *Presentation
}

// newPptxWriter 创建新的PPTX写入器
func newPptxWriter(pres *Presentation) *pptxWriter {
	return &pptxWriter{pres: pres}
}

// write 写入到io.Writer
func (w *pptxWriter) write(writer io.Writer) error {
	zipWriter := zip.NewWriter(writer)
	defer zipWriter.Close()

	// [Content_Types].xml
	if err := w.addFile(zipWriter, "[Content_Types].xml", w.pres.generateContentTypes()); err != nil {
		return err
	}

	// _rels/.rels
	if err := w.addFile(zipWriter, "_rels/.rels", generateRootRels()); err != nil {
		return err
	}

	// docProps/core.xml
	if err := w.addFile(zipWriter, "docProps/core.xml", w.pres.generateCoreProps()); err != nil {
		return err
	}

	// docProps/app.xml
	if err := w.addFile(zipWriter, "docProps/app.xml", w.pres.generateAppProps()); err != nil {
		return err
	}

	// ppt/presentation.xml
	if err := w.addFile(zipWriter, "ppt/presentation.xml", w.pres.generatePresentation()); err != nil {
		return err
	}

	// ppt/_rels/presentation.xml.rels
	if err := w.addFile(zipWriter, "ppt/_rels/presentation.xml.rels", w.pres.generatePresentationRels()); err != nil {
		return err
	}

	// ppt/presProps.xml
	if err := w.addFile(zipWriter, "ppt/presProps.xml", generatePresProps()); err != nil {
		return err
	}

	// ppt/viewProps.xml
	if err := w.addFile(zipWriter, "ppt/viewProps.xml", generateViewProps()); err != nil {
		return err
	}

	// ppt/tableStyles.xml
	if err := w.addFile(zipWriter, "ppt/tableStyles.xml", generateTableStyles()); err != nil {
		return err
	}

	// ppt/theme/theme1.xml
	if err := w.addFile(zipWriter, "ppt/theme/theme1.xml", generateTheme()); err != nil {
		return err
	}

	// ppt/slideMasters/slideMaster1.xml
	if err := w.addFile(zipWriter, "ppt/slideMasters/slideMaster1.xml", w.pres.generateSlideMaster()); err != nil {
		return err
	}

	// ppt/slideMasters/_rels/slideMaster1.xml.rels
	if err := w.addFile(zipWriter, "ppt/slideMasters/_rels/slideMaster1.xml.rels", generateSlideMasterRels()); err != nil {
		return err
	}

	// ppt/slideLayouts/slideLayout1.xml
	if err := w.addFile(zipWriter, "ppt/slideLayouts/slideLayout1.xml", w.pres.generateSlideLayout()); err != nil {
		return err
	}

	// ppt/slideLayouts/_rels/slideLayout1.xml.rels
	if err := w.addFile(zipWriter, "ppt/slideLayouts/_rels/slideLayout1.xml.rels", generateSlideLayoutRels()); err != nil {
		return err
	}

	// 幻灯片
	for i, slide := range w.pres.slides {
		slideNum := i + 1

		// ppt/slides/slideN.xml
		if err := w.addFile(zipWriter, "ppt/slides/slide"+itoa(slideNum)+".xml", slide.generateSlide()); err != nil {
			return err
		}

		// ppt/slides/_rels/slideN.xml.rels
		if err := w.addFile(zipWriter, "ppt/slides/_rels/slide"+itoa(slideNum)+".xml.rels", slide.generateSlideRels()); err != nil {
			return err
		}
	}

	// 媒体文件
	for _, media := range w.pres.mediaFiles {
		if err := w.addBytes(zipWriter, media.path, media.data); err != nil {
			return err
		}
	}

	// 图表文件
	for _, slide := range w.pres.slides {
		for _, obj := range slide.objects {
			if chart, ok := obj.(*chartObject); ok {
				chartPath := "ppt/charts/chart" + itoa(chart.chartIdx) + ".xml"
				if err := w.addFile(zipWriter, chartPath, chart.generateChartXML()); err != nil {
					return err
				}
			}
		}
	}

	return nil
}

// toBytes 转换为字节数组
func (w *pptxWriter) toBytes() ([]byte, error) {
	var buf bytes.Buffer
	if err := w.write(&buf); err != nil {
		return nil, err
	}
	return buf.Bytes(), nil
}

// addFile 添加文本文件到ZIP
func (w *pptxWriter) addFile(zipWriter *zip.Writer, path, content string) error {
	writer, err := zipWriter.Create(path)
	if err != nil {
		return err
	}
	_, err = writer.Write([]byte(content))
	return err
}

// addBytes 添加二进制文件到ZIP
func (w *pptxWriter) addBytes(zipWriter *zip.Writer, path string, data []byte) error {
	writer, err := zipWriter.Create(path)
	if err != nil {
		return err
	}
	_, err = writer.Write(data)
	return err
}
