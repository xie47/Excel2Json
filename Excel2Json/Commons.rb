#encoding:utf-8


def logError(message)
    puts "   \033[31m #{message}\033[0m\n"
end

def logInfo(message)
    puts "\033[32m #{message}\033[0m\n"
end

def logDebug(message)
    puts message
end

