package genppt

import (
	"os"
	"strings"
)

// AudioOptions 音频选项
type AudioOptions struct {
	X        float64 // X坐标（英寸），音频图标位置
	Y        float64 // Y坐标（英寸），音频图标位置
	Width    float64 // 宽度（英寸），音频图标大小
	Height   float64 // 高度（英寸），音频图标大小
	Path     string  // 本地音频文件路径
	Data     []byte  // 音频数据（与Path二选一）
	AutoPlay bool    // 是否自动播放
	Loop     bool    // 是否循环播放
	Hidden   bool    // 是否隐藏音频图标（用于背景音乐）
}

// audioObject 音频对象
type audioObject struct {
	options  AudioOptions
	rID      string // 关系ID
	mediaExt string // 媒体文件扩展名
}

func (a *audioObject) getType() string { return "audio" }

// AddAudio 添加音频
func (s *Slide) AddAudio(opts AudioOptions) *Slide {
	var data []byte
	var ext string
	var err error

	if opts.Path != "" {
		// 从文件读取
		data, err = os.ReadFile(opts.Path)
		if err != nil {
			// 音频读取失败，跳过
			return s
		}
		ext = getAudioExtFromPath(opts.Path)
	} else if len(opts.Data) > 0 {
		data = opts.Data
		ext = getAudioType(data)
	} else {
		return s
	}

	if ext == "" {
		ext = "mp3"
	}

	// 设置默认尺寸（音频图标大小）
	if opts.Width == 0 {
		opts.Width = 0.5
	}
	if opts.Height == 0 {
		opts.Height = 0.5
	}

	// 生成媒体文件关系ID
	mediaIndex := len(s.presentation.mediaFiles) + 1
	rID := "rId" + itoa(mediaIndex+400) // 使用较大的rId避免冲突
	mediaPath := "ppt/media/audio" + itoa(mediaIndex) + "." + ext

	// 添加到演示文稿的媒体文件列表
	s.presentation.mediaFiles = append(s.presentation.mediaFiles, mediaFile{
		path: mediaPath,
		data: data,
		ext:  ext,
		rID:  rID,
	})

	obj := &audioObject{
		options:  opts,
		rID:      rID,
		mediaExt: ext,
	}

	s.objects = append(s.objects, obj)
	return s
}

// getAudioExtFromPath 从路径获取音频扩展名
func getAudioExtFromPath(path string) string {
	idx := strings.LastIndex(path, ".")
	if idx == -1 {
		return ""
	}
	ext := strings.ToLower(path[idx+1:])
	switch ext {
	case "mp3", "wav", "wma", "m4a", "aac", "ogg", "flac":
		return ext
	default:
		return "mp3"
	}
}

// getAudioType 根据音频数据判断类型
func getAudioType(data []byte) string {
	if len(data) < 12 {
		return ""
	}
	// MP3 (ID3 header or sync word)
	if (data[0] == 0x49 && data[1] == 0x44 && data[2] == 0x33) || // ID3
		(data[0] == 0xFF && (data[1]&0xE0) == 0xE0) { // Sync word
		return "mp3"
	}
	// WAV
	if string(data[0:4]) == "RIFF" && string(data[8:12]) == "WAVE" {
		return "wav"
	}
	// OGG
	if string(data[0:4]) == "OggS" {
		return "ogg"
	}
	// FLAC
	if string(data[0:4]) == "fLaC" {
		return "flac"
	}
	// M4A/AAC (ftyp box)
	if len(data) >= 8 && string(data[4:8]) == "ftyp" {
		return "m4a"
	}
	return "mp3"
}

// getAudioMIME 获取音频MIME类型
func getAudioMIME(ext string) string {
	ext = strings.ToLower(ext)
	switch ext {
	case "mp3":
		return "audio/mpeg"
	case "wav":
		return "audio/wav"
	case "wma":
		return "audio/x-ms-wma"
	case "m4a", "aac":
		return "audio/mp4"
	case "ogg":
		return "audio/ogg"
	case "flac":
		return "audio/flac"
	default:
		return "audio/mpeg"
	}
}

// generateAudio 生成音频XML
func (s *Slide) generateAudio(a *audioObject, id int) string {
	var sb strings.Builder

	x := InchToEMU(a.options.X)
	y := InchToEMU(a.options.Y)
	cx := InchToEMU(a.options.Width)
	cy := InchToEMU(a.options.Height)

	sb.WriteString(`<p:pic>`)
	sb.WriteString(`<p:nvPicPr>`)
	sb.WriteString(`<p:cNvPr id="`)
	sb.WriteString(itoa(id))
	sb.WriteString(`" name="Audio `)
	sb.WriteString(itoa(id))
	sb.WriteString(`">`)

	// 音频操作链接
	sb.WriteString(`<a:hlinkClick r:id="" action="ppaction://media"/>`)
	sb.WriteString(`</p:cNvPr>`)
	sb.WriteString(`<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>`)
	sb.WriteString(`<p:nvPr>`)

	// 音频媒体
	sb.WriteString(`<a:audioFile r:link="`)
	sb.WriteString(a.rID)
	sb.WriteString(`"/>`)

	// 播放设置扩展
	sb.WriteString(`<p:extLst>`)
	sb.WriteString(`<p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">`)
	sb.WriteString(`<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="`)
	sb.WriteString(a.rID)
	sb.WriteString(`"/>`)
	sb.WriteString(`</p:ext>`)
	sb.WriteString(`</p:extLst>`)

	sb.WriteString(`</p:nvPr>`)
	sb.WriteString(`</p:nvPicPr>`)

	// 图片填充（使用空白或音频图标占位）
	sb.WriteString(`<p:blipFill>`)
	sb.WriteString(`<a:blip/>`)
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

	// 如果隐藏，设置不可见
	if a.options.Hidden {
		sb.WriteString(`<a:noFill/>`)
	}

	sb.WriteString(`</p:spPr>`)
	sb.WriteString(`</p:pic>`)

	return sb.String()
}
