package goexcel

import (
	"encoding/hex"
	"github.com/golang/image/colornames"
	"strings"
	"strconv"
)

func Color(color string) string {
	if len(color)<2 {
		panic("length of Color < 2 :"+color)
	}
	colorm:=ColorMap[strings.ToUpper(color[0:1])+color[1:]]
	//if not found on Map, use as-is. ("FF0000" etc)
 	if colorm==""{
		return color
	}
	colorx := colornames.Map[colorm]
	cs := hex.EncodeToString([]byte{colorx.R, colorx.G, colorx.B})
	return strings.ToUpper(cs)
}
func ColorDencity(color string, dencity int) string {
	colorx := colornames.Map[color]
	cs := hex.EncodeToString([]byte{getColorElement(colorx.R, dencity),
		getColorElement(colorx.G, dencity),
		getColorElement(colorx.B, dencity)})
	return strings.ToUpper(cs)
}
func getColorElement(color uint8, dencity int) uint8 {
	newColor := (255-float64(color)) * float64(dencity) / 100.00
	return uint8(255-newColor)
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