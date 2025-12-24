package genppt

import (
	"io"
	"os"
)

// New 创建新的演示文稿
func New() *Presentation {
	return &Presentation{
		slideWidth:  DefaultSlideWidth,
		slideHeight: DefaultSlideHeight,
		layout:      LayoutBlank,
		slides:      make([]*Slide, 0),
		mediaFiles:  make([]mediaFile, 0),
	}
}

// SetTitle 设置演示文稿标题
func (p *Presentation) SetTitle(title string) *Presentation {
	p.title = title
	return p
}

// SetAuthor 设置作者
func (p *Presentation) SetAuthor(author string) *Presentation {
	p.author = author
	return p
}

// SetSubject 设置主题
func (p *Presentation) SetSubject(subject string) *Presentation {
	p.subject = subject
	return p
}

// SetCompany 设置公司
func (p *Presentation) SetCompany(company string) *Presentation {
	p.company = company
	return p
}

// SetLayout 设置默认布局
func (p *Presentation) SetLayout(layout SlideLayout) *Presentation {
	p.layout = layout
	return p
}

// SetSlideSize 设置幻灯片尺寸（英寸）
func (p *Presentation) SetSlideSize(width, height float64) *Presentation {
	p.slideWidth = InchToEMU(width)
	p.slideHeight = InchToEMU(height)
	return p
}

// SetSlideSize4x3 设置为4:3比例（10"x7.5"）
func (p *Presentation) SetSlideSize4x3() *Presentation {
	p.slideWidth = 9144000  // 10英寸
	p.slideHeight = 6858000 // 7.5英寸
	return p
}

// SetSlideSize16x9 设置为16:9比例（10"x5.625"）
func (p *Presentation) SetSlideSize16x9() *Presentation {
	p.slideWidth = DefaultSlideWidth
	p.slideHeight = DefaultSlideHeight
	return p
}

// SetSlideSize16x10 设置为16:10比例
func (p *Presentation) SetSlideSize16x10() *Presentation {
	p.slideWidth = 9144000  // 10英寸
	p.slideHeight = 5715000 // 6.25英寸
	return p
}

// AddSlide 添加新幻灯片
func (p *Presentation) AddSlide() *Slide {
	slide := &Slide{
		presentation: p,
		layout:       p.layout,
		objects:      make([]slideObject, 0),
		number:       len(p.slides) + 1,
	}
	p.slides = append(p.slides, slide)
	return slide
}

// GetSlide 获取指定索引的幻灯片（从0开始）
func (p *Presentation) GetSlide(index int) *Slide {
	if index < 0 || index >= len(p.slides) {
		return nil
	}
	return p.slides[index]
}

// SlideCount 返回幻灯片数量
func (p *Presentation) SlideCount() int {
	return len(p.slides)
}

// WriteFile 将演示文稿保存到文件
func (p *Presentation) WriteFile(filename string) error {
	file, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer file.Close()
	return p.Write(file)
}

// Write 将演示文稿写入io.Writer
func (p *Presentation) Write(w io.Writer) error {
	writer := newPptxWriter(p)
	return writer.write(w)
}

// ToBytes 将演示文稿转换为字节数组
func (p *Presentation) ToBytes() ([]byte, error) {
	writer := newPptxWriter(p)
	return writer.toBytes()
}
