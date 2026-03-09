-- Usage:
-- Export single month:  osascript ~/export_eml_by_month.scpt 2025 1
-- Export range:         osascript ~/export_eml_by_month.scpt 2025 1 2025 12
-- Export all:           osascript ~/export_eml_by_month.scpt

on run argv
	set baseFolder to (POSIX path of (path to desktop)) & "Outlook_EML_Export/"
	set logFile to (POSIX path of (path to desktop)) & "export_log.txt"

	-- Parse arguments
	set argCount to count of argv
	set filterByDate to false
	set startYear to 0
	set startMonth to 0
	set endYear to 0
	set endMonth to 0

	if argCount = 2 then
		set filterByDate to true
		set startYear to (item 1 of argv) as integer
		set startMonth to (item 2 of argv) as integer
		set endYear to startYear
		set endMonth to startMonth
	else if argCount = 4 then
		set filterByDate to true
		set startYear to (item 1 of argv) as integer
		set startMonth to (item 2 of argv) as integer
		set endYear to (item 3 of argv) as integer
		set endMonth to (item 4 of argv) as integer
	end if

	if filterByDate then
		set startYM to startYear * 100 + startMonth
		set endYM to endYear * 100 + endMonth
	else
		set startYM to 0
		set endYM to 999999
	end if

	-- Clear log
	do shell script "echo 'Starting export...' > " & quoted form of logFile

	set exportedCount to 0
	set skippedCount to 0
	set failedCount to 0
	set totalScanned to 0

	tell application "Microsoft Outlook"
		if filterByDate then
			my logBoth("Filtering: " & startYear & "-" & my zeroPad(startMonth) & " to " & endYear & "-" & my zeroPad(endMonth), logFile)
		else
			my logBoth("Exporting ALL messages (no date filter)", logFile)
		end if

		set allMessages to every message
		set totalCount to count of allMessages
		my logBoth("Total messages found: " & totalCount, logFile)

		repeat with i from 1 to totalCount
			set totalScanned to totalScanned + 1
			try
				set currentMessage to item i of allMessages
				set msgDate to time received of currentMessage
				set msgYear to year of msgDate
				set msgMonth to month of msgDate as integer
				set msgYM to msgYear * 100 + msgMonth

				-- Get folder name
				try
					set folderName to name of (folder of currentMessage)
				on error
					set folderName to "Unknown"
				end try

				if msgYM < startYM or msgYM > endYM then
					set skippedCount to skippedCount + 1
					if (totalScanned mod 500) = 0 then
						my logBoth("Scanned " & totalScanned & "/" & totalCount & " (" & exportedCount & " exported, " & skippedCount & " skipped)", logFile)
					end if
				else
					set exportResult to my exportMessage(currentMessage, totalScanned, msgYear, msgMonth, folderName, baseFolder, logFile)
					if exportResult then
						set exportedCount to exportedCount + 1
					else
						set failedCount to failedCount + 1
					end if
				end if
			on error msgErr
				set failedCount to failedCount + 1
				my logBoth("ERROR message " & totalScanned & ": " & msgErr, logFile)
			end try
		end repeat

		set summary to "Done! " & exportedCount & " exported, " & failedCount & " failed, " & skippedCount & " skipped (out of " & totalScanned & " scanned)"
		my logBoth(summary, logFile)
		display dialog summary
	end tell
end run

on exportMessage(currentMessage, idx, msgYear, msgMonth, folderName, baseFolder, logFile)
	tell application "Microsoft Outlook"
		-- Build output folder: baseFolder/YYYY/MM/
		set monthFolder to baseFolder & (msgYear as string) & "/" & my zeroPad(msgMonth) & "/"
		do shell script "mkdir -p " & quoted form of monthFolder

		-- Get subject
		try
			set messageSubject to subject of currentMessage
			set messageID to id of currentMessage
		on error
			set messageSubject to "Unknown_Subject"
			set messageID to "msg_" & idx
		end try

		set cleanSubject to my cleanFileName(messageSubject)
		set cleanFolder to my cleanFileName(folderName)
		set fileName to cleanFolder & "_" & cleanSubject & "_" & messageID & ".eml"
		set hfsPath to POSIX file (monthFolder & fileName) as string

		-- Method 1: save as eml
		try
			save currentMessage in file hfsPath as "eml"
			my logBoth("  [" & idx & "] " & fileName, logFile)
			return true
		end try

		-- Method 2: save as msg, rename
		try
			set msgFileName to cleanFolder & "_" & cleanSubject & "_" & messageID & ".msg"
			set msgHfsPath to POSIX file (monthFolder & msgFileName) as string
			save currentMessage in file msgHfsPath as "msg"
			do shell script "mv " & quoted form of (monthFolder & msgFileName) & " " & quoted form of (monthFolder & fileName)
			my logBoth("  [" & idx & "] (via msg) " & fileName, logFile)
			return true
		end try

		-- Method 3: construct text eml
		try
			set messageContent to content of currentMessage
			set messageSender to sender of currentMessage
			set messageDate to time received of currentMessage
			set emailText to "Subject: " & messageSubject & return
			set emailText to emailText & "From: " & (name of messageSender) & " <" & (address of messageSender) & ">" & return
			set emailText to emailText & "Date: " & messageDate & return & return
			set emailText to emailText & messageContent
			do shell script "cat > " & quoted form of (monthFolder & fileName) & " <<'EMLEOF'" & return & emailText & return & "EMLEOF"
			my logBoth("  [" & idx & "] (via text) " & fileName, logFile)
			return true
		on error txtErr
			my logBoth("FAILED " & idx & " (" & messageSubject & "): " & txtErr, logFile)
			return false
		end try
	end tell
end exportMessage

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

on replaceText(sourceText, findText, replText)
	set AppleScript's text item delimiters to findText
	set textItems to text items of sourceText
	set AppleScript's text item delimiters to replText
	set resultText to textItems as string
	set AppleScript's text item delimiters to ""
	return resultText
end replaceText

on zeroPad(n)
	if n < 10 then
		return "0" & (n as string)
	else
		return n as string
	end if
end zeroPad

on logBoth(logText, logFile)
	log logText
	try
		do shell script "echo " & quoted form of ((current date) as string) & "': '" & quoted form of logText & " >> " & quoted form of logFile
	end try
end logBoth
