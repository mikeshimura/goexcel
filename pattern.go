package goexcel

import ()

const (
	PTN_DarkTrellis     = "darkTrellis"
	PTN_LightGray       = "lightGray"
	PTN_MediumGray      = "mediumGray"
	PTN_LightGrid       = "lightGrid"
	PTN_LightDown       = "lightDown"
	PTN_Gray0625        = "gray0625"
	PTN_LightTrellis    = "lightTrellis"
	PTN_LightUp         = "lightUp"
	PTN_DarkUp          = "darkUp"
	PTN_DarkVertical    = "darkVertical"
	PTN_Gray125         = "gray125"
	PTN_Solid           = "solid"
	PTN_DarkDown        = "darkDown"
	PTN_DarkHorizontal  = "darkHorizontal"
	PTN_LightVertical   = "lightVertical"
	PTN_DarkGrid        = "darkGrid"
	PTN_LightHorizontal = "lightHorizontal"
	PTN_DarkGray        = "darkGray"
)

var PatternMap map[string]string

func MakePatternMap() {
	PatternMap = make(map[string]string)
	PatternMap["DarkTrellis"] = "darkTrellis"
	PatternMap["LightGray"] = "lightGray"
	PatternMap["MediumGray"] = "mediumGray"
	PatternMap["LightGrid"] = "lightGrid"
	PatternMap["LightDown"] = "lightDown"
	PatternMap["Gray0625"] = "gray0625"
	PatternMap["LightTrellis"] = "lightTrellis"
	PatternMap["LightUp"] = "lightUp"
	PatternMap["DarkUp"] = "darkUp"
	PatternMap["DarkVertical"] = "darkVertical"
	PatternMap["Gray125"] = "gray125"
	PatternMap["Solid"] = "solid"
	PatternMap["DarkDown"] = "darkDown"
	PatternMap["DarkHorizontal"] = "darkHorizontal"
	PatternMap["LightVertical"] = "lightVertical"
	PatternMap["DarkGrid"] = "darkGrid"
	PatternMap["LightHorizontal"] = "lightHorizontal"
	PatternMap["DarkGray"] = "darkGray"
}
