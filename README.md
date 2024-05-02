# Excelsior

[![PkgGoDev](https://pkg.go.dev/badge/github.com/kakilangit/excelsior)](https://pkg.go.dev/github.com/kakilangit/excelsior)
[![Build Status](https://app.travis-ci.com/kakilangit/excelsior.svg?branch=main)](https://app.travis-ci.com/github/kakilangit/excelsior)

An excelize wrapper to separate the presentation and business logic.

```shell

go get -u -v github.com/kakilangit/excelsior

```

Example:

```go
package main

import (
	"log"
	"os"

	"github.com/kakilangit/excelsior"
	"github.com/xuri/excelize/v2"
)

// notFound implements excelsior.SheetData.
type notFound []string

// Total return total data.
func (f notFound) Total() int {
	return len(f)
}

// Row represents excel row.
func (f notFound) Row(i int) excelsior.Row {
	return []any{f[i]}
}

func main() {
	data, err := excelsior.Serialize(func(file *excelize.File, style excelsior.Style) ([]byte, error) {
		const sheetName = "not found alphabet"

		excelsior.SetDefaultSheetName(file, sheetName)

		headers := []any{"alphabet"}
		rows := []string{"a", "b", "c"}

		sheet := excelsior.NewSheet(headers, excelsior.DefaultStyleSetter, style.Header(), notFound(rows))
		if err := sheet.Generate(file, sheetName); err != nil {
			return nil, err
		}

		return excelsior.Byte(file)
	})
	if err != nil {
		log.Fatal(err)
	}

	f, err := os.Create("alphabet.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer func() { _ = f.Close() }()

	if _, err := f.Write(data); err != nil {
		log.Fatal(err)
	}

	if err := f.Sync(); err != nil {
		log.Fatal(err)
	}
}

```
