use scripting additions
use AppleScript version "2.4" -- Yosemite or later

-- Append Paid out Transactions from CSV to Numbers Actuals Table

set csvFile to choose file with prompt "Select the Downloaded Actuals CSV file:"
set {csvLines, isCC} to loadCSV(csvFile)
set headerIndex to findHeader(csvLines)
if headerIndex = 0 then
	display alert "Could not locate the CSV header row containing required columns." buttons {"OK"}
	return
end if
set cols to parseHeaderColumns(item headerIndex of csvLines, isCC)
if cols is false then
	display alert "Missing one or more required columns in header." buttons {"OK"}
	return
end if

repeat with k from (headerIndex + 1) to (count of csvLines)
	set csvLine to item k of csvLines
	if csvLine is not "" then
		set oldDelims to AppleScript's text item delimiters
		set AppleScript's text item delimiters to ","
		set parts to text items of csvLine
		set AppleScript's text item delimiters to oldDelims

		set pOut to stripQuotes(item (paidOutIdx of cols) of parts)
		set pOutNum to numericValue(pOut)

		if pOutNum > 0 then
			set dVal to stripQuotes(item (dateIdx of cols) of parts)
			set rawDesc to stripQuotes(item (descIdx of cols) of parts)
			set defaultComment to (commentPrefix of cols) & titleCase(rawDesc)

			set result to promptForTransaction(dVal, pOutNum, defaultComment)
			if action of result is "cancel" then return
			if action of result is "continue" then
				set theComment to theComment of result
				set theCategory to selectCategory(dVal)
				appendRowToNumbers(date dVal, theCategory, pOutNum, theComment)
			end if
		end if
	end if
end repeat

display dialog "All input CSV records have been read." buttons {"OK"}


-- Handler: read CSV file, return {lines, isCC}
on loadCSV(csvFile)
	set csvContent to do shell script "cat " & quoted form of POSIX path of csvFile
	set csvLines to paragraphs of csvContent
	set isCC to false
	if (item 1 of csvLines) contains "Select Credit Card" then set isCC to true
	return {csvLines, isCC}
end loadCSV

-- Handler: find the header row index (returns 0 if not found)
on findHeader(csvLines)
	repeat with i from 1 to (count of csvLines)
		set csvLine to item i of csvLines
		if csvLine contains "\"Date\"" and csvLine contains "\"Paid out\"" and csvLine contains "\"Paid in\"" then
			return i
		end if
	end repeat
	return 0
end findHeader

-- Handler: parse column indices from header row
-- Returns a record {dateIdx, paidOutIdx, descIdx, commentPrefix} or false if columns are missing
on parseHeaderColumns(headerLine, isCC)
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to ","
	set headerItems to text items of headerLine
	set AppleScript's text item delimiters to oldDelims

	set dateIdx to 0
	set paidOutIdx to 0
	set descIdx to 0
	set commentPrefix to ""

	repeat with j from 1 to (count of headerItems)
		set colName to item j of headerItems
		if colName starts with "\"" and colName ends with "\"" then set colName to text 2 thru -2 of colName
		if colName is "Date" then set dateIdx to j
		if colName is "Paid out" then set paidOutIdx to j
		if isCC and colName is "Transactions" then
			set descIdx to j
			set commentPrefix to "Visa: "
		end if
		if not isCC and colName is "Description" then
			set descIdx to j
			set commentPrefix to "Flex: "
		end if
	end repeat

	if dateIdx = 0 or paidOutIdx = 0 or descIdx = 0 then return false
	return {dateIdx:dateIdx, paidOutIdx:paidOutIdx, descIdx:descIdx, commentPrefix:commentPrefix}
end parseHeaderColumns

-- Handler: show dialog to edit comment; returns {action, theComment}
-- action is "cancel", "skip", or "continue"
on promptForTransaction(dVal, parsedAmount, defaultComment)
	set poundSign to character id 163
	set promptText to "Edit comment for transaction on " & dVal & " (" & poundSign & parsedAmount & "):"
	set userInput to display dialog promptText default answer defaultComment buttons {"Cancel Script", "Skip", "Continue"} default button "Continue"
	set btn to button returned of userInput
	if btn is "Cancel Script" then
		display dialog "Script cancelled by user." buttons {"OK"}
		return {action:"cancel", theComment:""}
	end if
	if btn is "Skip" then return {action:"skip", theComment:""}
	set theComment to text returned of userInput
	if theComment is "" then set theComment to defaultComment
	return {action:"continue", theComment:theComment}
end promptForTransaction

-- Handler: prompt user to pick a budget category
on selectCategory(dVal)
	set theCategory to choose from list {"Home", "Insurance", "Eats", "Transport & Travel", "Savings", "Family", "Projects & Pastimes", "Health & Beauty", "Clothes", "Big One-off", "Charitable & Other"} with prompt "Select Category for " & dVal & ":"
	if theCategory is false then return ""
	return item 1 of theCategory
end selectCategory

-- Handler: append one row to the Actuals table in Numbers
on appendRowToNumbers(parsedDate, theCategory, parsedAmount, theComment)
	tell application "Numbers Creator Studio"
		activate
		set doc to front document
		set tbl to table "Actuals" of sheet "Actual" of doc
		tell tbl
			make new row at end of rows
			set newRow to last row
			try
				tell newRow
					set value of cell 1 to parsedDate
					set value of cell 2 to theCategory
					set value of cell 3 to parsedAmount
					tell cell 3 to set format to currency
					set value of cell 4 to theComment
				end tell
			on error
				display alert "Numbers got an error: ensure Actuals table is not organised by Category while using this script" buttons {"OK"} as warning
			end try
		end tell
	end tell
end appendRowToNumbers


-- Handler: strip surrounding quotes
on stripQuotes(s)
	if s starts with "\"" and s ends with "\"" then return text 2 thru -2 of s
	return s
end stripQuotes

-- Handler: convert currency string to number
-- Skips any leading non-numeric characters (handles £, Â£, ï¿½ etc.)
on numericValue(s)
	try
		set t to s
		repeat while length of t > 0
			if text 1 of t is in "0123456789" then exit repeat
			set t to text 2 thru -1 of t
		end repeat
		if t is "" then return 0
		set AppleScript's text item delimiters to ","
		set parts to text items of t
		set AppleScript's text item delimiters to ""
		set cleanStr to parts as string
		return cleanStr as number
	on error
		return 0
	end try
end numericValue

-- Handler: uppercasing via awk
on makeUpper2(inString)
	return do shell script "awk '{ print toupper($0) }' <<< " & quoted form of inString
end makeUpper2

-- Handler: lowercasing via awk
on makeLower2(inString)
	return do shell script "awk '{ print tolower($0) }' <<< " & quoted form of inString
end makeLower2

-- Handler: title-case each word
on titleCase(inputText)
	set wordList to words of inputText
	set newList to {}
	repeat with w in wordList
		set w to w as text
		if length of w > 0 then
			set firstLetter to text 1 thru 1 of w
			set restLetters to ""
			if length of w > 1 then set restLetters to text 2 thru -1 of w
			set end of newList to my makeUpper2(firstLetter) & my makeLower2(restLetters)
		end if
	end repeat
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to " "
	set resultText to newList as text
	set AppleScript's text item delimiters to oldDelims
	return resultText
end titleCase
