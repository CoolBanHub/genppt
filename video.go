package genppt

import (
	"os"
	"strings"
)

// VideoOptions 视频选项
type VideoOptions struct {
	X        float64 // X坐标（英寸）
	Y        float64 // Y坐标（英寸）
	Width    float64 // 宽度（英寸）
	Height   float64 // 高度（英寸）
	Path     string  // 本地视频文件路径
	Data     []byte  // 视频数据（与Path二选一）
	Poster   []byte  // 封面图片数据（可选）
	AutoPlay bool    // 是否自动播放
	Loop     bool    // 是否循环播放
	Muted    bool    // 是否静音
}

// videoObject 视频对象
type videoObject struct {
	options   VideoOptions
	rID       string // 关系ID
	mediaExt  string // 媒体文件扩展名
	posterRID string // 封面图片关系ID
}

func (v *videoObject) getType() string { return "video" }

// AddVideo 添加视频
func (s *Slide) AddVideo(opts VideoOptions) *Slide {
	var data []byte
	var ext string
	var err error

	if opts.Path != "" {
		// 从文件读取
		data, err = os.ReadFile(opts.Path)
		if err != nil {
			// 视频读取失败，跳过
			return s
		}
		ext = getVideoExtFromPath(opts.Path)
	} else if len(opts.Data) > 0 {
		data = opts.Data
		ext = getVideoType(data)
	} else {
		return s
	}

	if ext == "" {
		ext = "mp4"
	}

	// 设置默认尺寸
	if opts.Width == 0 {
		opts.Width = 6.0
	}
	if opts.Height == 0 {
		opts.Height = 4.0
	}

	// 生成媒体文件关系ID
	mediaIndex := len(s.presentation.mediaFiles) + 1
	rID := "rId" + itoa(mediaIndex+300) // 使用较大的rId避免冲突
	mediaPath := "ppt/media/video" + itoa(mediaIndex) + "." + ext

	// 添加到演示文稿的媒体文件列表
	s.presentation.mediaFiles = append(s.presentation.mediaFiles, mediaFile{
		path: mediaPath,
		data: data,
		ext:  ext,
		rID:  rID,
	})

	obj := &videoObject{
		options:  opts,
		rID:      rID,
		mediaExt: ext,
	}

	// 如果有封面图片
	if len(opts.Poster) > 0 {
		posterExt := getImageType(opts.Poster)
		if posterExt == "" {
			posterExt = "png"
		}
		posterIndex := len(s.presentation.mediaFiles) + 1
		posterRID := "rId" + itoa(posterIndex+300)
		posterPath := "ppt/media/poster" + itoa(posterIndex) + "." + posterExt

		s.presentation.mediaFiles = append(s.presentation.mediaFiles, mediaFile{
			path: posterPath,
			data: opts.Poster,
			ext:  posterExt,
			rID:  posterRID,
		})
		obj.posterRID = posterRID
	}

	s.objects = append(s.objects, obj)
	return s
}

// getVideoExtFromPath 从路径获取视频扩展名
func getVideoExtFromPath(path string) string {
	idx := strings.LastIndex(path, ".")
	if idx == -1 {
		return ""
	}
	ext := strings.ToLower(path[idx+1:])
	switch ext {
	case "mp4", "m4v", "mov", "avi", "wmv", "mpg", "mpeg", "webm":
		return ext
	default:
		return "mp4"
	}
}

// getVideoType 根据视频数据判断类型
func getVideoType(data []byte) string {
	if len(data) < 12 {
		return ""
	}
	// MP4/M4V (ftyp box)
	if string(data[4:8]) == "ftyp" {
		return "mp4"
	}
	// WebM
	if data[0] == 0x1A && data[1] == 0x45 && data[2] == 0xDF && data[3] == 0xA3 {
		return "webm"
	}
	// AVI
	if string(data[0:4]) == "RIFF" && string(data[8:12]) == "AVI " {
		return "avi"
	}
	return "mp4"
}

// getVideoMIME 获取视频MIME类型
func getVideoMIME(ext string) string {
	ext = strings.ToLower(ext)
	switch ext {
	case "mp4", "m4v":
		return "video/mp4"
	case "mov":
		return "video/quicktime"
	case "avi":
		return "video/x-msvideo"
	case "wmv":
		return "video/x-ms-wmv"
	case "mpg", "mpeg":
		return "video/mpeg"
	case "webm":
		return "video/webm"
	default:
		return "video/mp4"
	}
}

// generateVideo 生成视频XML
func (s *Slide) generateVideo(v *videoObject, id int) string {
	var sb strings.Builder

	x := InchToEMU(v.options.X)
	y := InchToEMU(v.options.Y)
	cx := InchToEMU(v.options.Width)
	cy := InchToEMU(v.options.Height)

	sb.WriteString(`<p:pic>`)
	sb.WriteString(`<p:nvPicPr>`)
	sb.WriteString(`<p:cNvPr id="`)
	sb.WriteString(itoa(id))
	sb.WriteString(`" name="Video `)
	sb.WriteString(itoa(id))
	sb.WriteString(`">`)

	// 视频扩展
	sb.WriteString(`<a:hlinkClick r:id="" action="ppaction://media"/>`)
	sb.WriteString(`</p:cNvPr>`)
	sb.WriteString(`<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>`)
	sb.WriteString(`<p:nvPr>`)

	// 视频媒体
	sb.WriteString(`<a:videoFile r:link="`)
	sb.WriteString(v.rID)
	sb.WriteString(`"/>`)

	// 播放设置
	sb.WriteString(`<p:extLst>`)
	sb.WriteString(`<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">`)
	sb.WriteString(`<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="`)
	sb.WriteString(v.rID)
	sb.WriteString(`"/>`)
	sb.WriteString(`</p:ext>`)
	sb.WriteString(`</p:extLst>`)

	sb.WriteString(`</p:nvPr>`)
	sb.WriteString(`</p:nvPicPr>`)

	// 图片填充（封面或占位符）
	sb.WriteString(`<p:blipFill>`)
	if v.posterRID != "" {
		sb.WriteString(`<a:blip r:embed="`)
		sb.WriteString(v.posterRID)
		sb.WriteString(`"/>`)
	} else {
		// 使用空白填充
		sb.WriteString(`<a:blip/>`)
	}
	sb.WriteString(`<a:stretch><a:fillRect/></a:stretch>`)
	sb.WriteString(`</p:blipFill>`)

	// 形状属性
	sb.WriteString(`<p:spPr>`)
	sb.WriteString(`<a:xfrm>`)
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
	sb.WriteString(`</p:spPr>`)

	sb.WriteString(`</p:pic>`)

	return sb.String()
}
