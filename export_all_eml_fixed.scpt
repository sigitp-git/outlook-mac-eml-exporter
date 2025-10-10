set exportFolder to (path to desktop as string) & "Outlook_EML_Export:"

tell application "Finder"
	if not (exists folder exportFolder) then
		make new folder at desktop with properties {name:"Outlook_EML_Export"}
	end if
end tell

tell application "Microsoft Outlook"
	set allMessages to every message
	set totalCount to count of allMessages
	set exportedCount to 0
	
	repeat with i from 1 to totalCount
		try
			set currentMessage to item i of allMessages
			set messageSubject to subject of currentMessage
			
			-- Clean filename
			set cleanSubject to my replaceText(messageSubject, ":", "-")
			set cleanSubject to my replaceText(cleanSubject, "/", "-")
			set cleanSubject to my replaceText(cleanSubject, "\\", "-")
			set cleanSubject to my replaceText(cleanSubject, "?", "")
			set cleanSubject to my replaceText(cleanSubject, "*", "")
			set cleanSubject to my replaceText(cleanSubject, "<", "")
			set cleanSubject to my replaceText(cleanSubject, ">", "")
			set cleanSubject to my replaceText(cleanSubject, "|", "")
			
			-- Limit filename length
			if length of cleanSubject > 50 then
				set cleanSubject to text 1 thru 50 of cleanSubject
			end if
			
			set fileName to (cleanSubject & "_" & i & ".eml") as string
			set filePath to exportFolder & fileName
			
			-- Use string format "eml" instead of eml constant
			save currentMessage in file filePath as "eml"
			set exportedCount to exportedCount + 1
			
			-- Progress indicator every 100 messages
			if (exportedCount mod 100) = 0 then
				display notification "Exported " & exportedCount & " of " & totalCount & " messages"
			end if
			
		on error errMsg
			-- Skip messages that can't be exported
		end try
	end repeat
	
	display dialog "Export complete! " & exportedCount & " of " & totalCount & " emails exported to Desktop/Outlook_EML_Export"
end tell

on replaceText(sourceText, findText, replaceText)
	set AppleScript's text item delimiters to findText
	set textItems to text items of sourceText
	set AppleScript's text item delimiters to replaceText
	set resultText to textItems as string
	set AppleScript's text item delimiters to ""
	return resultText
end replaceText
