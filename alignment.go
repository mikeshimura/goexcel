package goexcel

import ()

const (
	H_CenterContinuous = "centerContinuous"
	H_Left             = "left"
	H_Distributed      = "distributed"
	H_Justify          = "justify"
	H_Fill             = "fill"
	H_Center           = "center"
	H_General          = "general"
	H_Right            = "right"
	V_Bottom           = "bottom"
	V_Distributed      = "distributed"
	V_Justify          = "justify"
	V_Top              = "top"
	V_Center           = "center"
)

var HAlingnMap map[string]string

func MakeHAlingnMap() {
	HAlingnMap = make(map[string]string)
	HAlingnMap["CenterContinuous"] = "centerContinuous"
	HAlingnMap["Left"] = "left"
	HAlingnMap["Distributed"] = "distributed"
	HAlingnMap["Justify"] = "justify"
	HAlingnMap["Fill"] = "fill"
	HAlingnMap["Center"] = "center"
	HAlingnMap["General"] = "general"
	HAlingnMap["Right"] = "right"
}

var VAlingnMap map[string]string

func MakeVAlingnMap() {
	VAlingnMap = make(map[string]string)
	VAlingnMap["Bottom"] = "bottom"
	VAlingnMap["Distributed"] = "distributed"
	VAlingnMap["Justify"] = "justify"
	VAlingnMap["Top"] = "top"
	VAlingnMap["Center"] = "center"
}
