package excelsior_test

import (
	"testing"

	"github.com/kakilangit/excelsior"
)

type nopSheetData struct {
	row excelsior.Row
}

func (nopSheetData) Total() int {
	return 1
}

func (d nopSheetData) Row(int) excelsior.Row {
	return d.row
}

func TestSheet(t *testing.T) {
	header := excelsior.Row{"Name", "Age"}
	getRowStyle := func(int) int { return 0 }
	headStyleID := 1
	data := nopSheetData{row: excelsior.Row{"John Doe", 30}}
	s := excelsior.NewSheet(header, getRowStyle, headStyleID, data)

	if s.TotalColumn() != 2 {
		t.Errorf("TotalColumn() = %d; want 2", s.TotalColumn())
	}

	if s.HeaderRow()[0] != "Name" {
		t.Errorf("HeaderRow()[0] = %s; want Name", s.HeaderRow()[0])
	}

	if s.HeaderRow()[1] != "Age" {
		t.Errorf("HeaderRow()[1] = %s; want Age", s.HeaderRow()[1])
	}

	if s.HeaderRowStyle()(0) != 1 {
		t.Errorf("HeaderRowStyle()(0) = %d; want 1", s.HeaderRowStyle()(0))
	}

	if s.HeaderRowStyle()(1) != 1 {
		t.Errorf("HeaderRowStyle()(1) = %d; want 1", s.HeaderRowStyle()(1))
	}

	if s.RowStyle()(0) != 0 {
		t.Errorf("RowStyle()(0) = %d; want 0", s.RowStyle()(0))
	}

	if s.RowStyle()(1) != 0 {
		t.Errorf("RowStyle()(1) = %d; want 0", s.RowStyle()(1))
	}

	if s.Row(0)[0] != "John Doe" {
		t.Errorf("Row(0)[0] = %s; want John Doe", s.Row(0)[0])
	}

	if s.Row(0)[1] != 30 {
		t.Errorf("Row(0)[1] = %d; want 30", s.Row(0)[1])
	}
}
