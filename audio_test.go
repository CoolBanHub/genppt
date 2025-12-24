package genppt

import (
	"os"
	"strings"
	"testing"
)

func TestAddAudio(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	// 测试使用数据添加音频
	slide.AddAudio(AudioOptions{
		Data:     []byte{0x49, 0x44, 0x33, 0x00}, // ID3 header (MP3)
		X:        1.0,
		Y:        1.0,
		Width:    0.5,
		Height:   0.5,
		AutoPlay: true,
		Loop:     true,
	})

	if len(slide.objects) != 1 {
		t.Errorf("Expected 1 object, got %d", len(slide.objects))
	}

	audio, ok := slide.objects[0].(*audioObject)
	if !ok {
		t.Error("Expected audioObject type")
	}

	if audio.getType() != "audio" {
		t.Errorf("Expected type 'audio', got '%s'", audio.getType())
	}

	if audio.mediaExt != "mp3" {
		t.Errorf("Expected extension 'mp3', got '%s'", audio.mediaExt)
	}
}

func TestAddAudioHidden(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	slide.AddAudio(AudioOptions{
		Data:   []byte{0x49, 0x44, 0x33, 0x00},
		Hidden: true,
	})

	audio := slide.objects[0].(*audioObject)
	if !audio.options.Hidden {
		t.Error("Expected Hidden to be true")
	}
}

func TestAudioGeneration(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	slide.AddAudio(AudioOptions{
		Data:     []byte{0x49, 0x44, 0x33, 0x00},
		X:        2.0,
		Y:        3.0,
		Width:    1.0,
		Height:   1.0,
		AutoPlay: true,
	})

	// 生成幻灯片XML
	xml := slide.generateSlide()

	// 检查是否包含音频元素
	if !strings.Contains(xml, "audioFile") {
		t.Error("Expected XML to contain audioFile element")
	}

	if !strings.Contains(xml, "ppaction://media") {
		t.Error("Expected XML to contain media action")
	}
}

func TestAudioWriteFile(t *testing.T) {
	pres := New()
	pres.SetTitle("Audio Test")
	slide := pres.AddSlide()

	slide.AddAudio(AudioOptions{
		Data:     []byte{0x49, 0x44, 0x33, 0x00, 0x00, 0x00, 0x00, 0x00},
		X:        1.0,
		Y:        1.0,
		AutoPlay: true,
	})

	// 写入临时文件
	tmpFile := "test_audio_output.pptx"
	err := pres.WriteFile(tmpFile)
	if err != nil {
		t.Errorf("Failed to write file: %v", err)
	}

	// 检查文件是否存在
	if _, err := os.Stat(tmpFile); os.IsNotExist(err) {
		t.Error("Output file was not created")
	}

	// 清理
	os.Remove(tmpFile)
}

func TestGetAudioType(t *testing.T) {
	tests := []struct {
		name     string
		data     []byte
		expected string
	}{
		{"MP3 ID3", []byte{0x49, 0x44, 0x33, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, "mp3"},
		{"MP3 Sync", []byte{0xFF, 0xFB, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, "mp3"},
		{"WAV", []byte{'R', 'I', 'F', 'F', 0x00, 0x00, 0x00, 0x00, 'W', 'A', 'V', 'E'}, "wav"},
		{"OGG", []byte{'O', 'g', 'g', 'S', 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, "ogg"},
		{"FLAC", []byte{'f', 'L', 'a', 'C', 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, "flac"},
		{"M4A", []byte{0x00, 0x00, 0x00, 0x00, 'f', 't', 'y', 'p', 0x00, 0x00, 0x00, 0x00}, "m4a"},
		{"Unknown", []byte{0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00}, "mp3"},
		{"Short", []byte{0x00}, ""},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := getAudioType(tt.data)
			if result != tt.expected {
				t.Errorf("getAudioType() = %s, expected %s", result, tt.expected)
			}
		})
	}
}

func TestGetAudioMIME(t *testing.T) {
	tests := []struct {
		ext      string
		expected string
	}{
		{"mp3", "audio/mpeg"},
		{"wav", "audio/wav"},
		{"wma", "audio/x-ms-wma"},
		{"m4a", "audio/mp4"},
		{"aac", "audio/mp4"},
		{"ogg", "audio/ogg"},
		{"flac", "audio/flac"},
		{"unknown", "audio/mpeg"},
	}

	for _, tt := range tests {
		t.Run(tt.ext, func(t *testing.T) {
			result := getAudioMIME(tt.ext)
			if result != tt.expected {
				t.Errorf("getAudioMIME(%s) = %s, expected %s", tt.ext, result, tt.expected)
			}
		})
	}
}

func TestGetAudioExtFromPath(t *testing.T) {
	tests := []struct {
		path     string
		expected string
	}{
		{"/path/to/file.mp3", "mp3"},
		{"/path/to/file.wav", "wav"},
		{"/path/to/file.m4a", "m4a"},
		{"/path/to/file.MP3", "mp3"},
		{"/path/to/file.xyz", "mp3"}, // 未知格式默认为mp3
		{"noextension", ""},
	}

	for _, tt := range tests {
		t.Run(tt.path, func(t *testing.T) {
			result := getAudioExtFromPath(tt.path)
			if result != tt.expected {
				t.Errorf("getAudioExtFromPath(%s) = %s, expected %s", tt.path, result, tt.expected)
			}
		})
	}
}

func TestAudioDefaultSize(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	slide.AddAudio(AudioOptions{
		Data: []byte{0x49, 0x44, 0x33, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00},
		X:    1.0,
		Y:    1.0,
		// Width和Height不设置，应该使用默认值
	})

	audio := slide.objects[0].(*audioObject)
	if audio.options.Width != 0.5 {
		t.Errorf("Expected default width 0.5, got %f", audio.options.Width)
	}
	if audio.options.Height != 0.5 {
		t.Errorf("Expected default height 0.5, got %f", audio.options.Height)
	}
}

func TestAudioEmptyData(t *testing.T) {
	pres := New()
	slide := pres.AddSlide()

	// 没有提供Path或Data，应该不添加任何对象
	slide.AddAudio(AudioOptions{
		X: 1.0,
		Y: 1.0,
	})

	if len(slide.objects) != 0 {
		t.Error("Expected no objects when no path or data provided")
	}
}
