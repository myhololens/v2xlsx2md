package main

import (
	"flag"
	"fmt"
	"os"
	"strings"

	excelize "github.com/xuri/excelize/v2"
)

var (
	oFile     = flag.String("i", "test.xslx", "filename")
	oSheet    = flag.String("sheet", "Sheet1", "name of sheet")
	oMarkdown = flag.String("o", "output.md", "Markdown file to create")
	oHeader   = flag.String("header", "A1:A1", "range of header titles")
	oBody     = flag.String("body", "A2:A2", "range of data")
)

func main() {
	flag.Parse()

	f, err := excelize.OpenFile(*oFile)
	if err != nil {
		fmt.Println(err)
		return
	}
	// waiting for 2.5.0
	// defer func() {
	//     // Close the spreadsheet.
	//     if err := f.Close(); err != nil {
	//         fmt.Println(err)
	//     }
	// }()
	out, err := os.Create(*oMarkdown)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer out.Close()

	header := strings.Split(*oHeader, ":")
	hleft, htop, err := excelize.CellNameToCoordinates(header[0])
	if err != nil {
		fmt.Println(err)
		return
	}
	hright, _, err := excelize.CellNameToCoordinates(header[1]) // TODO us hbottom
	if err != nil {
		fmt.Println(err)
		return
	}
	cols := 0
	for h := hleft; h <= hright; h++ {
		s := ValueAt(f, h, htop)
		fmt.Fprintf(out, "| %s ", s)
		cols++
	}
	fmt.Fprintln(out, "|")
	for c := 0; c < cols; c++ {
		fmt.Fprintf(out, "| --- ")
	}
	fmt.Fprintln(out, "|")

	body := strings.Split(*oBody, ":")
	left, top, err := excelize.CellNameToCoordinates(body[0])
	if err != nil {
		fmt.Println(err)
		return
	}
	right, bottom, err := excelize.CellNameToCoordinates(body[1])
	if err != nil {
		fmt.Println(err)
		return
	}

	for y := top; y <= bottom; y++ {
		fmt.Fprint(out, "|")
		for x := left; x <= right; x++ {
			s := ValueAt(f, x, y)
			fmt.Fprintf(out, " %s |", s)
		}
		fmt.Fprintln(out)
	}
}

func ValueAt(f *excelize.File, col, row int) string {
	axis, _ := excelize.CoordinatesToCellName(col, row)
	s, _ := f.GetCellValue(*oSheet, axis)
	return Sanitize(s)
}

func Sanitize(s string) string {
	return strings.ReplaceAll(strings.ReplaceAll(s, "\n", " "), "|", "\\|")
}
