package excelsior

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// SetRow sets a single row.
func SetRow(writer *excelize.StreamWriter, colID, rowID int, data []any, styleSetter StyleSetter) error {
	row := make([]any, len(data))
	for i, value := range data {
		row[i] = excelize.Cell{Value: value, StyleID: styleSetter(i)}
	}

	cell, err := excelize.CoordinatesToCellName(colID, rowID)
	if err != nil {
		return fmt.Errorf("failed to coordinate cell name: %w", err)
	}

	if err := writer.SetRow(cell, row); err != nil {
		return fmt.Errorf("failed to set row: %w", err)
	}

	return nil
}

// GenerateSheet generates a single sheet in Excel sheet.
func GenerateSheet(file *excelize.File, name string, sheet SheetProvider) error {
	writer, err := file.NewStreamWriter(name)
	if err != nil {
		return fmt.Errorf("failed to create new streamer writer: %w", err)
	}

	if err := GenerateSheetContent(writer, sheet); err != nil {
		return fmt.Errorf("failed to generate report: %w", err)
	}

	if err := writer.Flush(); err != nil {
		return fmt.Errorf("failed to flush writer: %w", err)
	}

	return nil
}

// GenerateSheetContent generates sheet content.
func GenerateSheetContent(writer *excelize.StreamWriter, sheet SheetProvider) error {
	const (
		colID = 1
		rowID = 1
	)

	if err := SetRow(writer, colID, rowID, sheet.HeaderRow(), sheet.HeaderRowStyle()); err != nil {
		return fmt.Errorf("failed to set header: %w", err)
	}

	// next row index
	const rowIndex = rowID + 1
	for index := 0; index < sheet.Total(); index++ {
		if err := SetRow(writer, colID, rowIndex+index, sheet.Row(index), sheet.RowStyle()); err != nil {
			return fmt.Errorf("failed to set content index %d: %w", index, err)
		}
	}

	return nil
}
