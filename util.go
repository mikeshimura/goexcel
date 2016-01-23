package goexcel

import (
	"encoding/hex"
	"github.com/golang/image/colornames"
	"strconv"
	"strings"
)

func Color(color string) string {
	if len(color) < 2 {
		panic(color + ": length of Color < 2 ")

	}
	colorm := ColorMap[strings.ToUpper(color[0:1])+color[1:]]
	//if not found on Map, use as-is. ("FF0000" etc)
	if colorm == "" {
		if len(color) != 6 {
			panic(color + ": not defined ")
		}
		for i:=0;i < 6;i++{
			cx:=color[i:i+1]
			switch cx {
				case "0","1","2","3","4","5",
				"6","7","8","9",
				"A","B","C","D","E","F":
				default:panic(color + ": not defined ")
			}
		}
		return color
	}
	colorx := colornames.Map[colorm]
	cs := hex.EncodeToString([]byte{colorx.R, colorx.G, colorx.B})
	return strings.ToUpper(cs)
}
func ColorDencity(color string, density int) string {
	colorm := ColorMap[strings.ToUpper(color[0:1])+color[1:]]
	if colorm == "" {
		panic(color + ": not defined ")
	}
	colorx := colornames.Map[colorm]
	cs := hex.EncodeToString([]byte{getColorElement(colorx.R, density),
		getColorElement(colorx.G, density),
		getColorElement(colorx.B, density)})
	return strings.ToUpper(cs)
}
func getColorElement(color uint8, dencity int) uint8 {
	newColor := (255 - float64(color)) * float64(dencity) / 100.00
	return uint8(255 - newColor)
}
func AtoiPanic(s string) int {
	i, err := strconv.Atoi(s)
	if err != nil {
		panic(s + " not Integer")
	}
	return i
}
func ParseFloatPanic(s string) float64 {
	f, err := strconv.ParseFloat(s, 64)
	if err != nil {
		panic(s + " not Numeric")
	}
	return f
}
