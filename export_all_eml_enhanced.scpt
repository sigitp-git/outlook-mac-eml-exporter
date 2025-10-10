set exportFolder to (path to desktop as string) & "Outlook_EML_Complete:"
set logFile to (path to desktop as string) & "export_log.txt"

tell application "Finder"
	if not (exists folder exportFolder) then
		make new folder at desktop with properties {name:"Outlook_EML_Complete"}
	end if
end tell

-- Clear previous log
try
	set fileRef to open for access file logFile with write permission
	set eof fileRef to 0
	write "Starting export of all messages..." & return to fileRef
	close access fileRef
end try

tell application "Microsoft Outlook"
	set allMessages to every message
	set totalCount to count of allMessages
	set exportedCount to 0
	set failedCount to 0
	
	my writeLog("Total messages to export: " & totalCount, logFile)
	
	repeat with i from 1 to totalCount
		try
			set currentMessage to item i of allMessages
			
			-- Get message properties first to check if accessible
			try
				set messageSubject to subject of currentMessage
				set messageID to id of currentMessage
			on error propErr
				set messageSubject to "Unknown_Subject"
				set messageID to "msg_" & i
				my writeLog("Warning: Could not get properties for message " & i & ": " & propErr, logFile)
			end try
			
			-- Clean filename more thoroughly
			set cleanSubject to my cleanFileName(messageSubject)
			
			-- Create unique filename using message ID if available
			set fileName to (cleanSubject & "_" & messageID & ".eml") as string
			set filePath to exportFolder & fileName
			
			-- Try multiple export methods
			set exported to false
			
			-- Method 1: Standard eml export
			if not exported then
				try
					save currentMessage in file filePath as "eml"
					set exported to true
					set exportedCount to exportedCount + 1
				on error emlErr
					-- Try method 2
				end try
			end if
			
			-- Method 2: Try msg format then rename
			if not exported then
				try
					set msgPath to exportFolder & (cleanSubject & "_" & messageID & ".msg")
					save currentMessage in file msgPath as "msg"
					-- Rename to .eml
					tell application "Finder"
						set name of file msgPath to (cleanSubject & "_" & messageID & ".eml")
					end tell
					set exported to true
					set exportedCount to exportedCount + 1
				on error msgErr
					-- Try method 3
				end try
			end if
			
			-- Method 3: Export as text with email headers
			if not exported then
				try
					set messageContent to content of currentMessage
					set messageSender to sender of currentMessage
					set messageDate to time received of currentMessage
					
					set emailText to "Subject: " & messageSubject & return
					set emailText to emailText & "From: " & (name of messageSender) & " <" & (address of messageSender) & ">" & return
					set emailText to emailText & "Date: " & messageDate & return & return
					set emailText to emailText & messageContent
					
					set txtPath to exportFolder & (cleanSubject & "_" & messageID & ".eml")
					set fileRef to open for access file txtPath with write permission
					write emailText to fileRef
					close access fileRef
					
					set exported to true
					set exportedCount to exportedCount + 1
				on error txtErr
					set failedCount to failedCount + 1
					my writeLog("FAILED message " & i & " (" & messageSubject & "): " & txtErr, logFile)
				end try
			end if
			
			-- Progress indicator
			if (i mod 100) = 0 then
				display notification "Processed " & i & " of " & totalCount & " (" & exportedCount & " exported)"
				my writeLog("Progress: " & i & "/" & totalCount & " processed, " & exportedCount & " exported, " & failedCount & " failed", logFile)
			end if
			
		on error mainErr
			set failedCount to failedCount + 1
			my writeLog("MAJOR ERROR with message " & i & ": " & mainErr, logFile)
		end try
	end repeat
	
	my writeLog("FINAL RESULTS: " & exportedCount & " exported, " & failedCount & " failed out of " & totalCount & " total", logFile)
	display dialog "Export complete! " & exportedCount & " of " & totalCount & " emails exported (" & failedCount & " failed). Check export_log.txt for details."
end tell

on cleanFileName(fileName)
	set cleanName to fileName
	set cleanName to my replaceText(cleanName, ":", "-")
	set cleanName to my replaceText(cleanName, "/", "-")
	set cleanName to my replaceText(cleanName, "\\", "-")
	set cleanName to my replaceText(cleanName, "?", "")
	set cleanName to my replaceText(cleanName, "*", "")
	set cleanName to my replaceText(cleanName, "<", "")
	set cleanName to my replaceText(cleanName, ">", "")
	set cleanName to my replaceText(cleanName, "|", "")
	set cleanName to my replaceText(cleanName, "\"", "")
	set cleanName to my replaceText(cleanName, "'", "")
	
	if length of cleanName > 50 then
		set cleanName to text 1 thru 50 of cleanName
	end if
	
	return cleanName
end cleanFileName

on replaceText(sourceText, findText, replaceText)
	set AppleScript's text item delimiters to findText
	set textItems to text items of sourceText
	set AppleScript's text item delimiters to replaceText
	set resultText to textItems as string
	set AppleScript's text item delimiters to ""
	return resultText
end replaceText

on writeLog(logText, logFile)
	try
		set logEntry to (current date) & ": " & logText & return
		set fileRef to open for access file logFile with write permission
		write logEntry to fileRef starting at eof
		close access fileRef
	on error
		try
			close access file logFile
		end try
	end try
end writeLog
