package excelsior_test

import (
	"net/http"
	"testing"

	"github.com/kakilangit/excelsior"
	"github.com/xuri/excelize/v2"
)

type nopResponseWriter struct{}

func (nopResponseWriter) Header() http.Header {
	return http.Header{}
}

func (nopResponseWriter) Write([]byte) (int, error) {
	return 0, nil
}

func (nopResponseWriter) WriteHeader(int) {}

func TestSerialize(t *testing.T) {
	data, err := excelsior.Serialize(func(file *excelize.File, style excelsior.Style) ([]byte, error) {
		const sheetName = "not found alphabet"

		excelsior.SetDefaultSheetName(file, sheetName)

		headers := []any{"alphabet"}
		data := nopSheetData{row: excelsior.Row{"a"}}

		sheet := excelsior.NewSheet(headers, excelsior.DefaultGetStyleFn, style.Header(), data)
		if err := sheet.Generate(file, sheetName); err != nil {
			return nil, err
		}

		return excelsior.Byte(file)
	})

	if err != nil {
		t.Errorf("failed to serialize: %v", err)
	}

	if len(data) == 0 {
		t.Errorf("empty data")
	}
}

func TestSetHeader(t *testing.T) {
	header := http.Header{}
	excelsior.SetHeader(header, "test")

	if header.Get("Content-Disposition") != "attachment; filename=test.xlsx" {
		t.Errorf("failed to set header")
	}
}

func TestWriteByt(t *testing.T) {
	data := []byte("test")
	rw := nopResponseWriter{}
	err := excelsior.WriteByte(data, "test", rw)
	if err != nil {
		t.Errorf("failed to write byte: %v", err)
	}
}
