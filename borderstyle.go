package goexcel

import ()

const (
	BDR_None	=""
	BDR_DashDot          = "dashDot"
	BDR_MediumDashDot    = "mediumDashDot"
	BDR_Dotted           = "dotted"
	BDR_SlantDashDot     = "slantDashDot"
	BDR_Thick            = "thick"
	BDR_DashDotDot       = "dashDotDot"
	BDR_Hair             = "hair"
	BDR_MediumDashDotDot = "mediumDashDotDot"
	BDR_MediumDashed     = "mediumDashed"
	BDR_Thin             = "thin"
	BDR_Double           = "double"
	BDR_Medium           = "medium"
	BDR_Dashed           = "dashed"
)

var BorderMap map[string]string

func MakeBorderMap() {
	BorderMap = make(map[string]string)
	BorderMap["None"] = ""
	BorderMap["DashDot"] = "dashDot"
	BorderMap["MediumDashDot"] = "mediumDashDot"
	BorderMap["Dotted"] = "dotted"
	BorderMap["SlantDashDot"] = "slantDashDot"
	BorderMap["Thick"] = "thick"
	BorderMap["DashDotDot"] = "dashDotDot"
	BorderMap["Hair"] = "hair"
	BorderMap["MediumDashDotDot"] = "mediumDashDotDot"
	BorderMap["MediumDashed"] = "mediumDashed"
	BorderMap["Thin"] = "thin"
	BorderMap["Double"] = "double"
	BorderMap["Medium"] = "medium"
	BorderMap["Dashed"] = "dashed"
}
