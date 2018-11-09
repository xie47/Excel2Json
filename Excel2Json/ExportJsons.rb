#encoding:utf-8

require 'win32ole'
require 'win32API'
require 'json'
require 'digest/md5'
load 'Commons.rb'


def exitWait()
	Process.wait(spawn("pause"))
end

class Excel
	def initialize(filename, path, sheetname, structdata, propdata, iForceExport)
		@filename=filename
		@Path=path.gsub(/\//,"\\")
		@sheetname=sheetname
		@json=""
		@structdata=structdata
		@typeName=propdata.keys
		@typeString=propdata.values
		@excel=nil
		@md5=nil
		@jsonmd5=nil
		@jsonFile=nil
		@update=iForceExport
		WIN32OLE.codepage = WIN32OLE::CP_UTF8
	end
	def ProcessExcel()
		if @update == 0
			if @md5 == @jsonmd5
				return
			end
		end
		OpenJson()
		
		@excel = WIN32OLE::new('excel.Application')
		@excel.visible = false
		workbook = @excel.Workbooks.Open(@Path)  
		begin
			worksheet = workbook.Worksheets(@sheetname.to_s) 
		rescue WIN32OLERuntimeError
			dialog = Win32API.new("user32", "MessageBox", "LPPL", "I")
			result = dialog.call(0, 
					    "\"#{@filename.to_s.encode(Encoding::UTF_8)}-#{@sheetname.to_s.encode(Encoding::UTF_8)}\" 不存在！".encode(Encoding::GBK),
					    "错误".encode(Encoding::GBK),
					    1)
			if result == 2
				@excel.ActiveWorkbook.Close(0);
				@excel.Quit()
				exit
			end
		end
		worksheet.Select
		
		@json += "{\n"
		@json += "\t\"type\":\"#{@structdata}\",\n"
		@json += "\t\"data\":\n"
		@json += "\t[\n"
			
		@cols = []
		'B'.upto('Z').to_a.each{|col| @cols.push(col)}
		'A'.upto('Z').to_a.each{|col| @cols.push('A'+col)}
		'A'.upto('Z').to_a.each{|col| @cols.push('B'+col)}
		
		end_col = ""
		numColum=0
		@cols.each{|col|
			cell = worksheet.Range("#{col}#{1}").value
			numColum += 1
			if cell == nil
				end_col = col
				break
			end
		}
		
		if (@typeName.length != numColum-1)
			dialog = Win32API.new("user32", "MessageBox", "LPPL", "I")
			result = dialog.call(0, 
					    "\"#{@filename.to_s.encode(Encoding::UTF_8)}-#{@sheetname.to_s.encode(Encoding::UTF_8)}\" 与structs.rb里列数量不匹配！".encode(Encoding::GBK),
					    "错误".encode(Encoding::GBK),
					    1)
			if result == 2
				@excel.ActiveWorkbook.Close(0);
				@excel.Quit()
				exit
			end
		end
		
		logDebug("#{$arg} to_a.each{|row| At #{Time.now}, eclips: #{Time.now - $timeStart}")
		2.upto(65535).to_a.each{|row|
			break if worksheet.Range("A#{row}").value == nil
			valuedata = []
			@cols.each{|col|
				break if col == end_col
				cell = worksheet.Range("#{col}#{row}")
				cell = "" if cell == nil
				valuedata.push(cell)
			}

			# ProcessRow(row,valuedata, worksheet.Range("A#{row}").value, worksheet.Range("A#{row+1}").value)				
		}
		logDebug("#{$arg} ProcessRow end At #{Time.now}, eclips: #{Time.now - $timeStart}")
		@json += "\t],\n"
		@json += "\t\"md5\":\"#{@md5}\"\n"
		@json += "}"

		@excel.ActiveWorkbook.Close(0);
		@excel.Quit()
		
		WriteJson()
		CloseJson()
	end
	def ReportError(row,col,data,type,msg)
		logError("#{row}, #{col}, #{data}, #{type}, #{msg}, #{@filename}, #{@sheetname}")
		logError('')
		logError(@json)
		exitWait()
		exit
	end
	def ProcessRow(rowIndex,data)
		@json += "\t\t{"
		@json += "\"ID\":#{data[0].to_i},"
		# for i in 0..data.length-1 do
		0.upto(@typeString.length-1).each{|i|
			begin
				j = i + 1
				curData = data[i + 1]
				curData = "" if curData == nil
				case @typeString[i].to_s
				when "_F32", "_F64"
					@json += "\"" + @typeName[i].to_s + "\":" + (curData == '' ? '0.0' : curData.to_s)
				when "_U8", "_U16", "_U32", "_U64", "_S8", "_S16", "_S32", "_S64"
					@json += "\"" + @typeName[i].to_s + "\":" + (curData == '' ? '0' : curData.to_i.to_s)
				when "bool"
					v = curData
					if v == 0 then
						v = "0"
					elsif v == 1 then
						v = "1"
					elsif v == '' then
						v = "0"
					elsif v == nil then
						v = "0"
					else
						v = Bools[@typeName[i]].key(curData.to_s.encode(Encoding::UTF_8)).to_s
					end
					@json += "\"" + @typeName[i].to_s + "\":" +  v
				when "Enum"
					@json += "\"" + @typeName[i].to_s + "\":" +  (curData == '' ? '0' : $EnumsAll[@typeName[i]].key(curData.to_s.encode(Encoding::UTF_8))[1].to_s)
				when "string"
					if curData.class == Float then
						if curData == curData.to_i then
							curData = curData.to_i.to_s
						end
					end
					@json += @typeName[i].to_s.to_json + ":"+  (curData == '' ? "\"\"" : curData.to_s.encode(Encoding::UTF_8).to_s.to_json )
				else
					if @typeString[i]["Enum"] != nil then
						if curData.class == Float then
							if curData == curData.to_i then
								curData = curData.to_i.to_s
							end
						end
						if curData == "0" and !$EnumsAll[@typeString[i]["Enum"]].invert.has_key?(0)
							curData = ""
						end
						@json += "\"" + @typeName[i].to_s + "\":" +  (curData == '' ? '0' : $EnumsAll[@typeString[i]["Enum"]].key(curData.to_s.encode(Encoding::UTF_8))[1].to_s)
					else
						ReportError(rowIndex, @rowsValue[0][i], curData, @typeString[i], "类型错误")
					end
				end
			rescue NoMethodError 
				ReportError(rowIndex, @rowsValue[0][i], curData, @typeString[i], "匹配错误")
			end
			(@json += ",") if i < @typeString.length-1
		# end
		}
		
		if @rowsValue[rowIndex+1] == nil or @rowsValue[rowIndex+1][0] == nil
			@json += "}\n"
		else
			@json += "},\n"
		end
		WriteJson()
	end
	def ParserExcel()
		if @update == 0
			if @md5 == @jsonmd5
				return
			end
		end
		LoadExcel()
		OpenJson()
		ids = Array.new
		@json += "{\n"
		@json += "\t\"type\":\"#{@structdata}\",\n"
		@json += "\t\"data\":\n"
		@json += "\t[\n"
		4.upto(65535).each{|rowIndex|
			break if @rowsValue[rowIndex] == nil or @rowsValue[rowIndex][0] == nil
			id = @rowsValue[rowIndex][0].to_i
			if ids.include?(id)
				ReportError(rowIndex, @rowsValue[0][0], "", "", "ID 重复！")
				dialog = Win32API.new("user32", "MessageBox", "LPPL", "I")
				dialog.call(0, 
					    "\"#{rowIndex.to_s.encode(Encoding::UTF_8)}-#{@@rowsValue[0][0]}\" ID 重复！".encode(Encoding::GBK),
					    "错误".encode(Encoding::GBK),
					    1)
				return
			else
				ids.push(id)
			end
			ProcessRow(rowIndex, @rowsValue[rowIndex])
		}
		@json += "\t],\n"
		@json += "\t\"md5\":\"#{@md5}\"\n"
		@json += "}"
		WriteJson()
		CloseJson()
	end
	def LoadExcel
		# @sheetRows={}
		begin
			excel = WIN32OLE::new('excel.Application')
			excel.visible = false
			workbook = excel.Workbooks.Open(@Path)
			worksheet = workbook.Worksheets(@sheetname.to_s) 
			worksheet.Select
			@rowsValue = worksheet.UsedRange.Rows.Value;
			excel.ActiveWorkbook.Close(0);
			excel.Quit()
		rescue NoMethodError
			excel.ActiveWorkbook.Close(0);
			excel.Quit()
			@rowsValue = nil
			logError( file+ " Error!")
			exit
		end
	end
	def OpenJson
		json_dir = "#{Dir.pwd}\\Jsons"
		Dir.mkdir(json_dir) if !Dir.exist?(json_dir)
		json_dir += "\\#{@structdata}.json"
		@jsonFile = File.open(json_dir, "w")
	end
	def WriteJson
		if @jsonFile != nil 
			@jsonFile.write(@json.encode(Encoding::UTF_8))
		end
		@json=""
	end
	def CloseJson
		if @jsonFile != nil 
			@jsonFile.close
		end
	end
	def SaveToJson
		if @update == 0
			if @md5 == @jsonmd5
				return
			end
		end
		json_dir = "#{Dir.pwd}\\Jsons"
		Dir.mkdir(json_dir) if !Dir.exist?(json_dir)
		json_dir += "\\#{@structdata}.json"
		f = File.open(json_dir, "w")
		f.write(@json.encode(Encoding::UTF_8))
		
		#~ json_dir = "#{$GameDir}\\Content\\JsonHO"
		#~ Dir.mkdir(json_dir) if !Dir.exist?(json_dir)
		#~ json_dir += "\\#{@structdata}.json"
		#~ f = File.open(json_dir, "w")
		#~ f.write(@json.encode(Encoding::UTF_8))
	end
	
	def GetMD5FromExcel	
		@md5 = Digest::MD5.open(@Path).hexdigest
	end
	
	def ReadMD5FromJson
		json_dir = "#{Dir.pwd}\\Jsons"
		Dir.mkdir(json_dir) if !Dir.exist?(json_dir)
		json_dir += "\\#{@structdata}.json"
		
		if File.exists?(json_dir)
			f = File.open(json_dir, "r")
			f.each do |line|
				begin			
					/\"md5\":\"(\w+)\"/.match(line.encode(Encoding::UTF_8)) do
						@jsonmd5 = $1
					end
				rescue Encoding::InvalidByteSequenceError
				rescue Encoding::UndefinedConversionError
				end
			end
		end
	end
end

class Digest::MD5
	def self.open(path)
		o = new
		File.open(path) { |f|
			buf = ""
			while f.read(256,buf)
				o << buf
			end
		}
		o
	end
end

def Export(structName, fileName, sheetName, iForceExport)
	logInfo("##### Export:#{fileName}=>#{sheetName}=>#{structName}####\n") 
	excel = Excel.new(fileName, "#{Dir.pwd}\\Excels\\#{fileName}", sheetName, structName, $StructsAll[structName], iForceExport)
	excel.GetMD5FromExcel()
	excel.ReadMD5FromJson()
	excel.ParserExcel()
	# excel.ProcessExcel()
	# excel.SaveToJson()

end

## 自动生成

begin
$timeStart = Time.now
$arg = ARGV.length < 1 ? 'All' : ARGV[0].to_s
logInfo("#{$arg} Started At #{Time.now} #{ARGV}") 

logInfo("##### ARGV.length: #{ARGV.length} ####\n")

load "./bin/AutoStruct.rb"
load 'AutoConfig.rb'
$StructsAll = AutoStructs
$EnumsAll = AutoEnum


if ARGV.length >3
	Export(ARGV[0], ARGV[1], ARGV[2], ARGV[3])
	
else
	forceUpdate=0
	if(ARGV.length==1)
		forceUpdate=1
	end

	Files.each{|file, sheets|
		sheets.each{|sheet, config|
			if config.index("Enum") == nil
				Export(config, file, sheet, forceUpdate)
			end
		}
	}
	
end

logInfo("#{$arg} Stoped At #{Time.now}, eclips: #{Time.now - $timeStart}")
exitWait()
rescue Exception => e
logInfo(e.to_s.force_encoding("UTF-8"))
system('pause')

end