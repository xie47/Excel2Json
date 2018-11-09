#encoding:utf-8

AutoEnum ={
	"Enum1"=>{
		["Enum1", 0] =>"枚举1",
		["Enum2", 1] =>"枚举2",
		["Enum3", 2] =>"枚举3",
		["Enum4", 3] =>"枚举4",
		["Enum5", 4] =>"枚举5",
	},

	"type"=>{
		["type1", 0] =>"type1",
		["type2", 1] =>"type2",
		["type3", 2] =>"type3",
		["type4", 3] =>"type4",
		["type5", 4] =>"type5",
	},

}

AutoStructs ={
	"TEXT_CONFIG"=>{
		"ID" =>"_U32",
		"attr1" =>"string",
		"attr2" =>"string",
		"num1" =>"_U32",
		"num2" =>"_F32",
		"enum" =>{"Enum"=>"Enum1"},
	},

	"CONST"=>{
		"ID" =>"_U32",
		"type" =>"_U32",
		"value" =>"_U32",
	},

}

