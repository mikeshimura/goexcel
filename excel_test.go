package goexcel

import (
  "testing"
 "fmt"
  "time"
  "github.com/tealeg/xlsx"
  //"image/color"
  
)

func TestExcel(t *testing.T) {
	var err error
//	style:=new(xlsx.Style)
//	style.Font.Name="ＭＳ Ｐ明朝"
//	style.Font.Size=14
//	style.ApplyBorder=true
//	style.Border.Top="thin"
//	style.Border.Left="thin"
//	style.Border.Right="thin"
//	style.Border.Bottom="thin"
	svc:=new(Goexcel)
	style:=svc.CreateStyle("ＭＳ Ｐ明朝",12,"BTLR","medium")
	style.Font.Color=Color(CLR_Black)
	fill := *xlsx.NewFill(PTN_LightHorizontal, Color(CLR_Aqua), Color(CLR_Yellow))
	style.Fill = fill
//	svc.File,err=xlsx.OpenFile("d:\\temp\\test.xlsx")
//	svc.SetSheetPanic("sheet1")
//	if err != nil {
//        fmt.Printf(err.Error())
//    }
//	fmt.Printf("top %v \n",svc.GetStyle(0,0).Border.Top)
	svc.File=xlsx.NewFile()
	svc.AddSheetPanic("sheet1")
	svc.GetCell(0,1)
	svc.Sheet.SetColWidth(1,1,10.9)
	svc.SetDateFormat(1,1,time.Now(),"mm/dd/yyyy")
	svc.SetDate(1,1,time.Now())
	svc.SetStyle(1,1,style)
	
	err=svc.File.Save("exceltest.xlsx")
	if err != nil {
        fmt.Printf(err.Error())
    }

//	var file *xlsx.File
//    var sheet *xlsx.Sheet
//    var row *xlsx.Row
//    var cell *xlsx.Cell
//    var err error
//	file=xlsx.NewFile()
//	sheet,err=file.AddSheet("Sheet1")
//	if err != nil {
//        fmt.Printf(err.Error())
//    }
//	row=sheet.AddRow()
//	cell=row.AddCell()
//	cell.Value="テスト"

	
}