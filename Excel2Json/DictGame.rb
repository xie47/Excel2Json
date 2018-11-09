#encoding:utf-8

Dict =
{

"枚举"=>"Enum",
}

def getEnglish(word)
	if Dict.has_key?(word) then
		return Dict[word]
	end
end

def translate(all)
	trans = getEnglish(all)
	if trans != nil
		return trans
	end
	trans=""
	untrans = ""
	
	for i in (0...all.size) 
		unless all[i].valid_encoding? 
		all[i] = "Z" 
		end
	end
	while all.length > 0 do
		# engLen = all =~ /\p{Han}/u
		# if engLen == nil then
			# break
		# end
		if (all =~ /\w/) == 0 then
			trans += all[0]
			all = all[1..-1]
			next
		end
		chWord = all
		eng = nil
		while chWord.length > 1 do
			eng = getEnglish(chWord)
			break if eng != nil
			chWord = all[0..(chWord.length-2)]
		end
		if eng != nil
			trans += eng
			all = all[chWord.length..-1]
			next
		end

		if all.length >= 1 then
			eng = getEnglish(all[0])
			if eng != nil then
				trans += eng
			else
				untrans += all[0]
				trans += "X"
			end
		end
		all = all[1..-1]
	end
	if untrans.size > 1
#		puts "can't translate #{untrans}"
	end
	Dict[all] = trans
	return trans
end

