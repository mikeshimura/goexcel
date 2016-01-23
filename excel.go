package goexcel

import (
	"errors"
	"fmt"
	"github.com/tealeg/xlsx"
	"io"
	"strings"
	"time"
)

//border thin medium double など
type Goexcel struct {
	File  *xlsx.File
	Sheet *xlsx.Sheet
	Row   int
	Style map[string]*xlsx.Style
}

func CreateGoexcel() *Goexcel {
	excel := new(Goexcel)
	excel.File = xlsx.NewFile()
	excel.Style = make(map[string]*xlsx.Style)
	return excel
}
func (e *Goexcel) OpenFile(filename string) error {
	file, err := xlsx.OpenFile(filename)
	if err != nil {
		return err
	}
	e.File = file
	return nil
}
func (e *Goexcel) Save(filename string) error {
	return e.File.Save(filename)
}
func (e *Goexcel) Write(writer io.Writer) error {
	return e.File.Write(writer)
}
func (e *Goexcel) CopyStyle(oldkey string, newkey string) {
	sold := e.Style[oldkey]
	if sold.Font.Name == "" {
		panic("Style " + oldkey + " not found or Font Name not defined")
	}
	style := new(xlsx.Style)
	style.Alignment = sold.Alignment
	style.ApplyAlignment = sold.ApplyAlignment
	style.ApplyBorder = sold.ApplyBorder
	style.ApplyFill = sold.ApplyFill
	style.ApplyFont = sold.ApplyFont
	style.Border = sold.Border
	style.Fill = sold.Fill
	style.Font = sold.Font
	e.Style[newkey] = style
}
func (e *Goexcel) SetColWidth(startCol int, endCol int, width float64) {
	e.GetCell(0, endCol)
	e.Sheet.SetColWidth(startCol, endCol, width)
}
func (e *Goexcel) GetRow(no int) *xlsx.Row {
	addRowNo := no - e.Sheet.MaxRow + 1
	//fmt.Printf("addRowNo %v \n", addRowNo)
	for i := 0; i < addRowNo; i++ {
		e.Sheet.AddRow()
		fmt.Printf("i %v \n", i)
	}

	return e.Sheet.Rows[no]
}
func (e *Goexcel) GetCell(rowno int, colno int) *xlsx.Cell {
	row := e.GetRow(rowno)
	addColNo := colno - len(row.Cells) + 1
	//fmt.Printf("addColNo %v \n", addColNo)
	for i := 0; i < addColNo; i++ {
		row.AddCell()
		fmt.Printf("i %v \n", i)
	}

	return row.Cells[colno]
}
func (e *Goexcel) AddSheet(name string) error {
	sheet, err := e.File.AddSheet(name)
	if err != nil {
		return err
	}
	e.Sheet = sheet
	return nil
}
func (e *Goexcel) AddSheetPanic(name string) {
	sheet, err := e.File.AddSheet(name)
	if err != nil {
		panic(err)
	}
	e.Sheet = sheet
}

func (e *Goexcel) SetSheet(name string) error {
	sheet := e.File.Sheet[name]
	if sheet != nil {
		e.Sheet = sheet
		return nil
	} else {
		return errors.New("Sheet :" + name + " が見つかりません")
	}
}
func (e *Goexcel) SetSheetPanic(name string) {
	sheet := e.File.Sheet[name]
	if sheet != nil {
		e.Sheet = sheet
	} else {
		panic("Sheet :" + name + " が見つかりません")
	}
}
func (e *Goexcel) SetFormat(rowno int, colno int, fmt string) {
	cell := e.GetCell(rowno, colno)
	cell.NumFmt = fmt
}
func (e *Goexcel) SetString(rowno int, colno int, s string) {
	cell := e.GetCell(rowno, colno)
	cell.SetString(s)
}
func (e *Goexcel) SetInt(rowno int, colno int, i int) {
	cell := e.GetCell(rowno, colno)
	fmt := cell.GetNumberFormat()
	cell.SetInt(i)
	if fmt != "" {
		cell.NumFmt = fmt
	}
}
func (e *Goexcel) SetIntFormat(rowno int, colno int, i int, fmt string) {
	cell := e.GetCell(rowno, colno)
	cell.SetInt(i)
	cell.NumFmt = fmt
}
func (e *Goexcel) SetFloat(rowno int, colno int, f float64) {
	cell := e.GetCell(rowno, colno)
	fmt := cell.GetNumberFormat()
	cell.SetFloat(f)
	if fmt != "" {
		cell.NumFmt = fmt
	}
}
func (e *Goexcel) SetFloatFormat(rowno int, colno int, f float64, fmt string) {
	cell := e.GetCell(rowno, colno)
	cell.SetFloatWithFormat(f, fmt)
}
func (e *Goexcel) SetDate(rowno int, colno int, t time.Time) {
	cell := e.GetCell(rowno, colno)
	fmt := cell.GetNumberFormat()
	cell.SetDate(t)
	if fmt != "" {
		cell.NumFmt = fmt
	}
}
func (e *Goexcel) SetDateFormat(rowno int, colno int, t time.Time, fmt string) {
	cell := e.GetCell(rowno, colno)
	cell.SetDate(t)
	if fmt != "" {
		cell.NumFmt = fmt
	}
}
func (e *Goexcel) SetDateTime(rowno int, colno int, t time.Time) {
	cell := e.GetCell(rowno, colno)
	fmt := cell.GetNumberFormat()
	cell.SetDateTime(t)
	if fmt != "" {
		cell.NumFmt = fmt
	}
}
func (e *Goexcel) SetDateTimeFormat(rowno int, colno int, t time.Time, fmt string) {
	cell := e.GetCell(rowno, colno)
	cell.SetDateTime(t)
	if fmt != "" {
		cell.NumFmt = fmt
	}
}

func (e *Goexcel) SetStyleByKey(rowno int, colno int, style string) {
	cell := e.GetCell(rowno, colno)
	cell.SetStyle(e.Style[style])
}
func (e *Goexcel) SetStyle(rowno int, colno int, style *xlsx.Style) {
	cell := e.GetCell(rowno, colno)
	cell.SetStyle(style)
}
func (e *Goexcel) GetStyle(rowno int, colno int) *xlsx.Style {
	cell := e.GetCell(rowno, colno)
	return cell.GetStyle()
}

func (e *Goexcel) Merge(rowno int, colno int, toRowNo int, toColNo int) {
	rows := toRowNo - rowno + 1
	cols := toColNo - colno + 1
	if rows < 1 {
		panic("toRowNo must be greater than rowno")
	}
	if cols < 1 {
		panic("toColNo must be greater than colno")
	}
	for row := rowno; row < rowno+cols-1; row++ {
		e.GetCell(row, colno+cols-1)
	}
	cell := e.GetCell(rowno, colno)
	cell.Merge(cols-1, rows-1)
}

func (e *Goexcel) SetFormula(rowno int, colno int, formula string) {
	cell := e.GetCell(rowno, colno)
	cell.SetFormula(formula)
}
func (e *Goexcel) CreateStyleByKey(keyname string, font string, size int, border string,
	borderPattern string) {
	style := e.CreateStyle(font, size, border, borderPattern)
	e.Style[keyname] = style
}
func (e *Goexcel) CreateStyle(fontName string, fontSize int, border string,
	borderPattern string) *xlsx.Style {
	style := new(xlsx.Style)
	style.Font.Name = fontName
	style.Font.Size = fontSize
	setBorderSub(style, border, borderPattern)
	return style
}
func (e *Goexcel) SetFontName(keyname string, fontName string) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	style.Font.Name = fontName
}
func (e *Goexcel) SetFontSize(keyname string, fontSize int) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	style.Font.Size = fontSize
}
func (e *Goexcel) SetFontColor(keyname string, color string) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	style.Font.Color = Color(color)
}
func (e *Goexcel) SetItalic(keyname string, italic bool) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	style.Font.Italic = italic
}
func (e *Goexcel) SetBold(keyname string, bold bool) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	style.Font.Bold = bold
}
func (e *Goexcel) SetUnderline(keyname string, underline bool) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	style.Font.Underline = underline
}
func (e *Goexcel) SetBorder(keyname string, border string,
	borderPattern string) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	setBorderSub(style, border, borderPattern)
}
func (e *Goexcel) SetBorderColor(keyname string, border string,
	color string) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	border = strings.ToUpper(border)
	if strings.Contains(border, "T") {
		style.Border.TopColor = Color(color)
	}
	if strings.Contains(border, "B") {
		style.Border.BottomColor = Color(color)
	}
	if strings.Contains(border, "R") {
		style.Border.RightColor = Color(color)
	}
	if strings.Contains(border, "L") {
		style.Border.LeftColor = Color(color)
	}
}
func (e *Goexcel) SetFill(keyname string, pattern string,
	fgColor string, bgColor string) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	style.Fill.PatternType = pattern
	style.Fill.FgColor = Color(fgColor)
	style.Fill.BgColor = Color(bgColor)
}
func (e *Goexcel) SetHorizontalAlign(keyname string, alignment string) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	style.Alignment.Horizontal = alignment
}
func (e *Goexcel) SetVerticalAlign(keyname string, alignment string) {
	style := e.Style[keyname]
	if style.Font.Name == "" {
		panic("Style " + keyname + " not found or Font Name not defined")
	}
	style.Alignment.Vertical = alignment
}
func setBorderSub(style *xlsx.Style, border string,
	borderPattern string) {
	if len(border) > 0 {
		style.ApplyBorder = true
	}
	border = strings.ToUpper(border)
	if strings.Contains(border, "T") {
		style.Border.Top = borderPattern
	}
	if strings.Contains(border, "B") {
		style.Border.Bottom = borderPattern
	}
	if strings.Contains(border, "R") {
		style.Border.Right = borderPattern
	}
	if strings.Contains(border, "L") {
		style.Border.Left = borderPattern
	}
}
