package excelsior

import (
	"fmt"
	"net/http"

	"github.com/xuri/excelize/v2"
)

const (
	// StyleHeaderKey is key for the hf style header
	StyleHeaderKey = "excelsior-header"
	// StyleHeaderValue is value for the hf style header.
	StyleHeaderValue = `{"font":{"bold":true}}`
	// DefaultStyleID is Excel without style.
	DefaultStyleID = 0
	// HeaderContentDisposition is header key for content disposition
	HeaderContentDisposition = `Content-Disposition`
	// HeaderContentDispositionValue is header value for content disposition
	HeaderContentDispositionValue = `attachment; filename=%[1]s.xlsx`
	// HeaderContentType  is header key for content type
	HeaderContentType = `Content-Type`
	// HeaderContentTypeValue is header value for content type
	HeaderContentTypeValue = `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
)

// Style holds the style name with style value.
type Style map[string]int

// Header returns the header style.
func (s Style) Header() int {
	i, _ := s[StyleHeaderKey]
	return i
}

// SerializeFn is signature serializer function.
type SerializeFn func(file *excelize.File, style Style) ([]byte, error)

// Serialize is the serializer framework to serialize data to excel bytes.
func Serialize(f SerializeFn) ([]byte, error) {
	file, style, err := NewExcelizeFile()
	if err != nil {
		return nil, err
	}

	return f(file, style)
}

// DefaultStyleSetter is helper for the default excelsior row style.
func DefaultStyleSetter(_ int) int {
	return DefaultStyleID
}

// ExcelizeStyle holds the excelize.Style.
type ExcelizeStyle map[string]interface{}

// Style converts ExcelizeStyle to Style.
func (s ExcelizeStyle) Style(file *excelize.File) (Style, error) {
	styles := make(Style, len(s))

	for key, raw := range s {
		style, err := file.NewStyle(raw)
		if err != nil {
			return nil, fmt.Errorf("fail to create style %s: %w", key, err)
		}

		styles[key] = style
	}

	return styles, nil
}

// NewExcelizeFile creates excelize.File with default excelsior.Style.
func NewExcelizeFile() (*excelize.File, Style, error) {
	file := excelize.NewFile()

	styles, err := ExcelizeStyle{
		StyleHeaderKey: StyleHeaderValue,
	}.Style(file)
	if err != nil {
		return nil, nil, err
	}

	return file, styles, nil
}

// Byte converts excelize.File to bytes.
func Byte(file *excelize.File) ([]byte, error) {
	buf, err := file.WriteToBuffer()
	if err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

// SetHeader sets http.Header for Excel file.
func SetHeader(header http.Header, filename string) {
	header.Set(HeaderContentDisposition, fmt.Sprintf(HeaderContentDispositionValue, filename))
	header.Set(HeaderContentType, HeaderContentTypeValue)
}

// SetDefaultSheetName changes the Excel tab name.
func SetDefaultSheetName(file *excelize.File, name string) {
	const (
		index   = 0
		tabName = "Sheet1"
	)

	file.SetSheetName(tabName, name)
	file.SetActiveSheet(index)
}

// WriteByte writes byte to http.ResponseWriter.
func WriteByte(b []byte, filename string, w http.ResponseWriter) error {
	header := w.Header()
	SetHeader(header, filename)
	_, err := w.Write(b)
	return err
}
