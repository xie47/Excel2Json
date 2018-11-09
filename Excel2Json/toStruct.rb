#encoding:utf-8
require 'win32ole'
require 'fileutils'
$Language = Encoding::UTF_8
load 'DictGame.rb'

$ExcelDir = "#{Dir.pwd}/Excels"
$OutputPath = "./bin"
$EnumExcel ={}

load 'AutoConfig.rb'
load 'Commons.rb'
LocalStructs={}
LocalEnums = {}
$ExcelData = {}

def checkDir(tree)
	curDir = Dir.pwd
	dirList = tree.split("/")
	dirList.each{|name|
		Dir.mkdir(name) if !Dir.exist?(name)
		Dir.chdir(name)
	}
	Dir.chdir(curDir)
end

def saveToFile(fileName)
	temp="#encoding:utf-8\n\n"
	temp +="AutoEnum ={\n"
	LocalEnums.each{|key, value|
		temp += "\t\"#{key}\"=>{\n"
		value.each{|enum_name, enum_desc|
			temp += "\t\t[\"#{enum_name[0]}\", #{enum_name[1]}] =>\"#{enum_desc}\",\n"
		}
		temp += "\t},\n\n"
	}
	temp += "}\n\n"
	
	temp +="AutoStructs ={\n"
	LocalStructs.each{|key, value|
		temp += "\t\"#{key}\"=>{\n"
		value.each{|name, type|
			if type["Enum"] != nil then
				temp += "\t\t\"#{name}\" =>#{type},\n"
			else
				temp += "\t\t\"#{name}\" =>\"#{type}\",\n"
			end
		}
		temp += "\t},\n\n"
	}
	temp += "}\n\n"
	
	checkDir($OutputPath)
	luaTempFile = "#{$OutputPath}/#{fileName}.temp"
	f = File.open(luaTempFile, "w")
	f.write(temp.encode($Language))
	f.close
	
	luaOldFile = "#{$OutputPath}/#{fileName}.rb"
	FileUtils.install luaTempFile, luaOldFile
	FileUtils.rm luaTempFile
	logInfo("save #{luaOldFile} over!\n\n") 
end

class ExcelParser
	def initialize(excelFile, sheetname, structName)
		@fileName=excelFile
		@sheetName=sheetname
		@structName=structName
		@rowsValue=$ExcelData[@fileName][@sheetName]
	end
	def parseStruct()
		structMap={}
		for colIndex in 0..65535
			cell = @rowsValue[0][colIndex]
			break if cell == nil or cell == ""
			name = cell.to_s.encode($Language).to_s
			
			value = @rowsValue[2][colIndex].to_s.encode($Language).to_s
			structMap[name] = getType(value)
		end
		LocalStructs[@structName] = structMap
	end
	def parseConst()
		parseEnumDetail(1);
	end
	def parseEnum()
		for colIndex in 0..65535
			parseEnumDetail(colIndex)
		end
	end

	def parseEnumDetail(colIndex)
		cell = @rowsValue[0][colIndex]
		return if cell == nil or cell == ""
		name = cell.to_s.encode($Language).to_s
		valueEnum = {}
		cellEnum={}
		for rowIndex in 4..65535
			break if @rowsValue[rowIndex] == nil or @rowsValue[rowIndex][colIndex] == nil
			value = @rowsValue[rowIndex][colIndex].to_s.encode($Language).to_s
			next if (value =~ /\d/) == 0
			valueEnum[value]=rowIndex-4
			cell=[]
			cell.push(translate(value))
			cell.push(rowIndex-4)
			cellEnum[cell] = value
		end
		return if cellEnum.length < 1
		if $EnumExcel[name] != nil 
			logError("枚举名字重复\t#{name}\t#{@structName}\t#{@fileName}\t#{@sheetName}")
		else 
			$EnumExcel[name]= valueEnum
			LocalEnums[name]= cellEnum
		end
	end
	def getType(type) 
		if (type == "string")
			return "string"
		elsif (type == "int")
			return  "_U32"
		elsif (type == "float")
			return "_F32"
		elsif (type == "enum")
			return "_U32"
		elsif (type == "double")
			return "_F32"
		elsif $EnumExcel.has_key?type
			enumType={}
			enumType["Enum"] = type
			return enumType
		else
			logError("Error Type\t#{type}\t#{@structName}\t#{@fileName}\t#{@sheetName}")
			return "string"
		end
	end	
end

def loadExecl(file, sheets)
	# return if $ExcelData[file] == nil
	sheetRows={}
	begin
		@excel = WIN32OLE::new('excel.Application')
		@excel.visible = false
		@workbook = @excel.Workbooks.Open("#{$ExcelDir}/#{file}")
		sheets.each{|sheet, config|
			@worksheet = @workbook.Worksheets(sheet.to_s) 
			@worksheet.Select
			rowsValue = @worksheet.UsedRange.Rows.Value;
			sheetRows[sheet.to_s] = rowsValue
			# print sheet.to_s, " ,", rowsValue[0], " ,", rowsValue[1]
		}
		@excel.ActiveWorkbook.Close(0);
		@excel.Quit()
	rescue NoMethodError
		@excel.ActiveWorkbook.Close(0);
		@excel.Quit()
	
		logError(file+ " Error!")
		exit
	end

	$ExcelData[file]=sheetRows
	logDebug(file+ " Load OK!") 
end

def parseConfigFile()
	timeStart = Time.now
	logInfo("Started At #{timeStart}")
	Files.each{|file, sheets|
		loadExecl(file, sheets)
		sheets.each{|sheet, config|
			if config.index("Enum") != nil
				excel = ExcelParser.new(file, sheet, config)
				excel.parseEnum()
			end
			if config.index("CONST") != nil
				excel = ExcelParser.new(file, sheet, config)
				excel.parseConst()
			end
		}
	}
	timeEnd = Time.now
	logInfo("loadExecl Over At #{timeEnd}, eclips: #{timeEnd - timeStart}")
	Files.each{|file, sheets|
		sheets.each{|sheet, config|
			if config.index("Enum") == nil
				excel = ExcelParser.new(file, sheet, config)
				excel.parseStruct()
			end
		}
	}
	saveToFile("AutoStruct")
	timeEnd = Time.now
	logInfo("Stoped At #{timeEnd}, eclips: #{timeEnd - timeStart}")
end

parseConfigFile()

