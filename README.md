# goexcel
Excel Generator for Go
[日本語の説明はこちら](https://github.com/mikeshimura/goexcel/wiki/%E6%97%A5%E6%9C%AC%E8%AA%9E%E8%AA%AC%E6%98%8E)

This program use github.com/tealeg/xlsx for excel generation.

Generated Excel Sample
![Simple1](https://bytebucket.org/mikeshimura/goreport/wiki/image/simple1excel.jpg "Simple1")

[excel](https://bytebucket.org/mikeshimura/goreport/wiki/excel/simple1.xlsx)

##Installation
```
go get -u github.com/mikeshimura/goexcel
```

##Concept

- Easy to generate Excel

- Color names are already defined as constant. (used github.com/golang/image/colornames)

- Color density can be used.

- Border patterns are already definedas constant.

- Fill patterns are already defined as constant.

- Base style (Font Name etc.) can be inherited easily.

Pre defined Border pattern and Fill pattern Sample
![Sample](https://bytebucket.org/mikeshimura/goreport/wiki/image/sample.jpg "Sample")

[excel](https://bytebucket.org/mikeshimura/goreport/wiki/excel/parameter.xlsx)

##Setup Command

- New Excel  
 CreateGoexcel()

- Open Existing Excel  
 OpenFile(filename string) error  

- Sheet handling
 AddSheet(name string) error  
 AddSheetPanic(name string)  
 SetSheet(name string) error  
 SetSheetPanic(name string)  

##File Write Command  
 Save(filename string) error  
 Write(writer io.Writer) error  

##Predefined Color  
Colors are predefined like CLR_Blue  

- Color Set  
 Color(color string) string  
 // use like Color(ge.CLR_Blue)  

 ColorDencity(color string, dencity int) string  
 // use like ColorDencity(ge.CLR_Blue, 20)

##Style Command  
Styles are stored as map[string]*xlsx.Style

- Create Style  
 CreateStyleByKey(keyname string, font string, size int, border string,
 borderPattern string)

- Copy Style  
  CopyStyle(oldkey string, newkey string)
  //Using this command, style can be inherited.

- Set Font
 SetFontName(keyname string, fontName string)  
 SetFontSize(keyname string, fontSize int)  
 SetFontColor(keyname string, color string)  
 SetItalic(keyname string, italic bool)  
 SetBold(keyname string, bold bool)   
 SetUnderline(keyname string, underline bool)  

- Set Border  
SetBorder(keyname string, border string,borderPattern string)  
// border "T" top  "B" bottom  "L" left  "R" right  

- Set Fill  
SetFill(keyname string, pattern string,	fgColor string,bgColor string)  

- Set Alignment  
 SetHorizontalAlign(keyname string, alignment string)  
 //alignment like H_Right

 SetVerticalAlign(keyname string, alignment string)   
 //alignment like V_Center  

- Set Column Width  
 SetColWidth(startCol int, endCol int, width float64)  

- Merge Cells  
 Merge(rowno int, colno int, toRowNo int, toColNo int)

##Value / Format Set Command  
 SetInt(rowno int, colno int, i int)  
 SetIntFormat(rowno int, colno int, i int, fmt string)   
 SetFloat(rowno int, colno int, f float64)  
 SetFloatFormat(rowno int, colno int, f float64, fmt string)  
 SetDate(rowno int, colno int, t time.Time)
 SetDateFormat(rowno int, colno int, t time.Time, fmt string)  
 SetDateTime(rowno int, colno int, t time.Time)  
 SetDateTimeFormat(rowno int, colno int, t time.Time, fmt string)
 SetString(rowno int, colno int, s string)
 SetFormula(rowno int, colno int, formula string)  
 SetFormat(rowno int, colno int, fmt string)  

 ##License  

 goexcel is released under the MIT License. It is copyrighted by Masanobu Shimura. (Gmail mikeshimura)

 ##Sample program

 [sample source](https://github.com/mikeshimura/goexcel/blob/master/example/simple1.go)
 ```go
 package example

 import (
 	"fmt"
 	ge "github.com/mikeshimura/goexcel"
 	"io/ioutil"
 	"strconv"
 	"strings"
 	"time"
 )

 func Simple1() {
 	list := ReadTextFile("simple1.txt", 8)
 	excel := ge.CreateGoexcel()
 	excel.AddSheetPanic("Sheet1")
 	excel.CreateStyleByKey("base", "Verdana", 10, "", "")

 	//Style TBLR
 	excel.CopyStyle("base", "TBLR")
 	excel.SetBorder("TBLR", "TBLR", ge.BDR_Hair)
 	excel.SetFill("TBLR", ge.PTN_Solid, ge.ColorDencity(ge.CLR_Blue, 20), ge.Color(ge.CLR_Yellow))

 	//Style TBLR_Right
 	excel.CopyStyle("TBLR", "TBLR_R")
 	excel.SetHorizontalAlign("TBLR_R", ge.H_Right)

 	//Style TITLE
 	excel.CreateStyleByKey("TITLE", "Arial", 24, "TBLR", ge.BDR_Double)
 	excel.SetFill("TITLE", ge.PTN_Gray125, ge.ColorDencity(ge.CLR_Black, 50),
 		"CCCCFF")
 	excel.SetItalic("TITLE", true)
 	excel.SetHorizontalAlign("TITLE", ge.H_Center)
 	//Style DATE
 	excel.CreateStyleByKey("DATE", "Arial", 11, "", "")
 	excel.SetFontColor("DATE", ge.ColorDencity(ge.CLR_Black, 60))

 	//Style Header
 	excel.CopyStyle("TBLR", "HEADER")
 	excel.SetBold("HEADER", true)
 	excel.SetFill("HEADER", ge.PTN_Solid, ge.ColorDencity(ge.CLR_Blue, 40), ge.Color(ge.CLR_Yellow))
 	excel.SetBorder("HEADER", "TB", ge.BDR_Medium)
 	excel.SetHorizontalAlign("HEADER", ge.H_Center)

 	//Style Header Left
 	excel.CopyStyle("HEADER", "HEADER_L")
 	excel.SetBorder("HEADER_L", "L", ge.BDR_Medium)

 	//Style Header Right
 	excel.CopyStyle("HEADER", "HEADER_R")
 	excel.SetBorder("HEADER_R", "R", ge.BDR_Medium)
 	excel.SetColWidth(0, 0, 9)
 	excel.SetColWidth(1, 1, 6)
 	excel.SetColWidth(2, 2, 11)
 	excel.SetColWidth(3, 3, 8)
 	excel.SetColWidth(4, 4, 13)
 	excel.SetColWidth(5, 5, 7)
 	excel.SetColWidth(6, 6, 9)
 	excel.SetColWidth(7, 7, 6)
 	excel.SetColWidth(8, 8, 11)
 	for col := 1; col < 7; col++ {
 		excel.SetStyleByKey(1, col, "TITLE")
 	}
 	excel.Merge(1, 1, 1, 6)
 	excel.SetString(1, 1, "Goexcel Sample Report-Sample1")

 	excel.SetStyleByKey(0, 8, "DATE")
 	excel.SetDateFormat(0, 8, time.Now(), "yyyy/mm/dd")

 	excel.SetStyleByKey(4, 0, "HEADER_L")
 	for col := 1; col < 8; col++ {
 		excel.SetStyleByKey(4, col, "HEADER")
 	}
 	excel.SetStyleByKey(4, 8, "HEADER_R")
 	excel.SetString(4, 0, "Date")
 	excel.SetString(4, 1, "Deptno")
 	excel.SetString(4, 2, "Dept")
 	excel.SetString(4, 3, "Slip")
 	excel.SetString(4, 4, "Stock Code")
 	excel.SetString(4, 5, "Name")
 	excel.SetString(4, 6, "Unit Price")
 	excel.SetString(4, 7, "Qty")
 	excel.SetString(4, 8, "Amount")
 	for i, l := range list {
 		cols := l.([]string)
 		excel.SetStyleByKey(5+i, 0, "TBLR")
 		excel.SetStyleByKey(5+i, 1, "TBLR")
 		excel.SetStyleByKey(5+i, 2, "TBLR")
 		excel.SetStyleByKey(5+i, 3, "TBLR")
 		excel.SetStyleByKey(5+i, 4, "TBLR")
 		excel.SetStyleByKey(5+i, 5, "TBLR")
 		excel.SetStyleByKey(5+i, 6, "TBLR_R")
 		excel.SetStyleByKey(5+i, 7, "TBLR_R")
 		excel.SetStyleByKey(5+i, 8, "TBLR_R")
 		t, err := time.Parse("2006/01/02", cols[0])
 		if err != nil {
 			panic(cols[0] + " is not Date")
 		}
 		excel.SetDateFormat(5+i, 0, t, "yyyy/mm/dd")
 		excel.SetString(5+i, 1, cols[1])
 		excel.SetString(5+i, 2, cols[2])
 		excel.SetString(5+i, 3, cols[3])
 		excel.SetString(5+i, 4, cols[4])
 		excel.SetString(5+i, 5, cols[5])
 		excel.SetFloatFormat(5+i, 6, ge.ParseFloatPanic(cols[6]), "#,##0.00")
 		excel.SetIntFormat(5+i, 7, ge.AtoiPanic(cols[7]), "#,##0")
 		excel.SetFormat(5+i, 8, "#,##0.00")
 		excel.SetFormula(5+i, 8, "G"+strconv.Itoa(6+i)+"*H"+strconv.Itoa(6+i))
 	}

 	err := excel.Save("simple1.xlsx")
 	if err != nil {
 		panic(err)
 	}
 	fmt.Println("Simple1")
 	//	excel2 := ge.CreateGoexcel()
 	//	excel2.OpenFile("simple1.xlsx")
 	//	excel2.Save("simple1_test.xlsx")

 }
 func ReadTextFile(filename string, colno int) []interface{} {
 	res, _ := ioutil.ReadFile(filename)
 	lines := strings.Split(string(res), "\r\n")
 	list := make([]interface{}, 0, 100)
 	for _, line := range lines {
 		cols := strings.Split(line, "\t")
 		if len(cols) < colno {
 			continue
 		}
 		list = append(list, cols)
 	}
 	return list
 }
 ```
