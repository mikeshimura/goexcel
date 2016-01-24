package example

import (
	"fmt"
	ge "github.com/mikeshimura/goexcel"
)

func ParameterSample2() {

	excel := ge.CreateGoexcel()
	excel.AddSheetPanic("Sheet1")
	excel.SetColWidth(0, 0, 20)
	excel.SetColWidth(2, 2, 20)
	excel.CreateStyleByKey("base", "Verdana", 10, "", "")
	excel.SetStyleByKey(0, 0, "base")
	excel.SetString(0, 0, "LINE STYLE")
	i := 0
	for s := range ge.BorderMap {
		ss := ge.BorderMap[s]
		fmt.Println(ss)
		excel.CopyStyle("base", ss)
		excel.SetBorder(ss, "TBLR", ss)
		excel.SetStyleByKey(2+i*2, 0, ss)
		excel.SetString(2+i*2, 0, s)
		i++
	}

	excel.SetStyleByKey(0, 2, "base")
	excel.SetString(0, 2, "FILL STYLE")
	i = 0
	for s := range ge.PatternMap {
		ss := ge.PatternMap[s]
		fmt.Println(ss)
		excel.CopyStyle("base", ss)
		excel.SetFill(ss, ss, ge.CLR_Blue, ge.CLR_Yellow)
		excel.SetStyleByKey(2+i*2, 2, ss)
		excel.SetString(2+i*2, 2, s)
		i++
	}

	err := excel.Save("parameter2.xlsx")
	if err != nil {
		panic(err)
	}
}
