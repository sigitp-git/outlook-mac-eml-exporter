set logFile to (path to desktop as string) & "message_count.log"

tell application "Microsoft Outlook"
	try
		set allMessages to every message
		set totalCount to count of allMessages
		my writeLog("Total messages found: " & totalCount, logFile)
		
		if totalCount > 0 then
			set firstMessage to item 1 of allMessages
			set msgSubject to subject of firstMessage
			my writeLog("First message subject: " & msgSubject, logFile)
		else
			my writeLog("No messages found", logFile)
		end if
		
	on error errMsg
		my writeLog("Error: " & errMsg, logFile)
	end try
end tell

on writeLog(logText, logFile)
	try
		set logEntry to logText & return
		set fileRef to open for access file logFile with write permission
		write logEntry to fileRef starting at eof
		close access fileRef
	on error
		try
			close access file logFile
		end try
	end try
end writeLog
