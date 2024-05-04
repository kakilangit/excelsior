[<img alt="github" src="https://img.shields.io/badge/github-kakilangit/excelsior-37a8e0?style=for-the-badge&labelColor=555555&logo=github" height="20">](https://github.com/kakilangit/excelsior)
[<img alt="pkg.go.dev" src="https://img.shields.io/badge/go-%2300ADD8.svg?style=for-the-badge&logo=go&logoColor=white" height="20">](https://pkg.go.dev/github.com/kakilangit/excelsior)

# Excelsior

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

## License

MIT
Copyright (c) 2022 kakilangit
