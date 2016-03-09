package goexcel

import (
	"fmt"
	"testing"
	//"image/color"
)

func TestExcel(t *testing.T) {
	var err error
	svc := new(Goexcel)
	err=svc.OpenFile("d:\\temp\\test.xlsm")
	if err!=nil{
		panic(err)
	}
	svc.SetSheetPanic("sheet1")
	fmt.Printf("3 2 %v \n", svc.String(2, 1))
	fmt.Printf("3 3 %v\n", svc.String(2, 2))
	fmt.Printf("3 4 %v\n", svc.String(2, 3))
	svc = nil

}
