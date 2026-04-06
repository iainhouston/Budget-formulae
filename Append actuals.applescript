use scripting additions
use AppleScript version "2.4" -- Yosemite or later

-- Append Paid out Transactions from CSV to Numbers Actuals Table

-- Prompt user to select the source CSV file
set csvFile to choose file with prompt "Select the Downloaded Actuals CSV file:"

-- Read CSV content and split into lines
set csvContent to do shell script "cat " & quoted form of POSIX path of csvFile
set csvLines to paragraphs of csvContent

-- Determine account type based on first line
set firstLine to item 1 of csvLines
set isCC to false
if firstLine contains "Select Credit Card" then set isCC to true

-- Find the header row containing required columns
set headerIndex to 0
repeat with i from 1 to (count of csvLines)
	set csvLine to item i of csvLines
	if csvLine contains "\"Date\"" and csvLine contains "\"Paid out\"" and csvLine contains "\"Paid in\"" then
		set headerIndex to i
		exit repeat
	end if
end repeat

if headerIndex = 0 then
	display alert "Could not locate the CSV header row containing required columns." buttons {"OK"}
	return
end if

-- Determine column indices from header
set oldDelims to AppleScript's text item delimiters
set AppleScript's text item delimiters to ","
set headerItems to text items of item headerIndex of csvLines
set dateIndex to 0
set paidOutIndex to 0
set paidInIndex to 0
if isCC then
	set transIndex to 0
	repeat with j from 1 to (count of headerItems)
		set colName to item j of headerItems
		if colName starts with "\"" and colName ends with "\"" then set colName to text 2 thru -2 of colName
		if colName is "Date" then set dateIndex to j
		if colName is "Transactions" then set transIndex to j
		if colName is "Paid out" then set paidOutIndex to j
		if colName is "Paid in" then set paidInIndex to j
	end repeat
else
	set descIndex to 0
	repeat with j from 1 to (count of headerItems)
		set colName to item j of headerItems
		if colName starts with "\"" and colName ends with "\"" then set colName to text 2 thru -2 of colName
		if colName is "Date" then set dateIndex to j
		if colName is "Description" then set descIndex to j
		if colName is "Paid out" then set paidOutIndex to j
		if colName is "Paid in" then set paidInIndex to j
	end repeat
end if
set AppleScript's text item delimiters to oldDelims

if dateIndex = 0 or paidOutIndex = 0 or paidInIndex = 0 or (isCC and transIndex = 0) or (not isCC and descIndex = 0) then
	display alert "Missing one or more required columns in header." buttons {"OK"}
	return
end if

-- Process each data row following the header
repeat with k from (headerIndex + 1) to (count of csvLines)
	set csvLine to item k of csvLines
	if csvLine is not "" then
		-- Parse fields
		set oldDelims to AppleScript's text item delimiters
		set AppleScript's text item delimiters to ","
		set parts to text items of csvLine
		set AppleScript's text item delimiters to oldDelims

		set dRaw to item dateIndex of parts
		set pOutRaw to item paidOutIndex of parts
		if isCC then
			set commentRaw to item transIndex of parts
			set commentSource to "Visa: "
		else
			set commentRaw to item descIndex of parts
			set commentSource to "Flex: "
		end if

		set pOut to stripQuotes(pOutRaw)
		set pOutNum to numericValue(pOut)

		-- Only process rows where Paid out > 0
		if pOutNum > 0 then
			set dVal to stripQuotes(dRaw)
			set parsedDate to date dVal
			set parsedAmount to pOutNum
			set defaultComment to commentSource & titleCase(stripQuotes(commentRaw))

			-- Edit comment and decide whether to process
			set poundSign to character id 163
			set promptText to "Edit comment for transaction on " & dVal & " (" & poundSign & parsedAmount & "):"
			set userInput to display dialog promptText default answer defaultComment buttons {"Cancel Script", "Skip", "Continue"} default button "Continue"
			set skipAnswer to button returned of userInput
			if skipAnswer is "Cancel Script" then
				display dialog "Script cancelled by user." buttons {"OK"}
				return
			end if
			if skipAnswer is "Continue" then
				set theComment to text returned of userInput
				if theComment is "" then set theComment to defaultComment
				-- Category
				set theCategory to choose from list {"Home", "Insurance", "Eats", "Transport & Travel", "Savings", "Family", "Projects & Pastimes", "Health & Beauty", "Clothes", "Big One-off", "Charitable & Other"} with prompt "Select Category for " & dVal & ":"
				if theCategory is false then set theCategory to {""}
				set theCategory to item 1 of theCategory

				-- Append to Numbers and format
				tell application "Numbers Creator Studio"
					activate
					set doc to front document
					set tbl to table "Actuals" of sheet "Actual" of doc
					tell tbl
						make new row at end of rows
						set newRow to last row
						-- Wrap cell-setting in a try block to catch 'Can't set row' errors
						try
							tell newRow
								set value of cell 1 to parsedDate
								set value of cell 2 to theCategory
								set value of cell 3 to parsedAmount
								tell cell 3 to set format to currency
								set value of cell 4 to theComment
								set value of cell 5 to ((month of parsedDate as string) & " " & (year of parsedDate as string))
							end tell
						on error errMsg number errNum
							display alert "Numbers got an error: ensure Actuals table is not organised by Category while using this script" buttons {"OK"} as warning
							return
						end try
					end tell
				end tell
			end if
		end if
	end if
end repeat

-- Notify when done processing all CSV records
display dialog "All input CSV records have been read." buttons {"OK"}

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
		-- Strip leading non-digit characters (currency symbols regardless of encoding)
		repeat while length of t > 0
			if text 1 of t is in "0123456789" then exit repeat
			set t to text 2 thru -1 of t
		end repeat
		if t is "" then return 0
		-- Remove comma thousands separators
		set AppleScript's text item delimiters to ","
		set parts to text items of t
		set AppleScript's text item delimiters to ""
		set cleanStr to parts as string
		return cleanStr as number
	on error
		return 0
	end try
end numericValue

-- Uppercasing via awk
on makeUpper2(inString)
	return do shell script "awk '{ print toupper($0) }' <<< " & quoted form of inString
end makeUpper2

-- Lowercasing via awk
on makeLower2(inString)
	return do shell script "awk '{ print tolower($0) }' <<< " & quoted form of inString
end makeLower2

-- Title-case each word and preserve spaces
on titleCase(inputText)
	set wordList to words of inputText
	set newList to {}
	repeat with w in wordList
		set w to w as text
		if length of w > 0 then
			set firstLetter to text 1 thru 1 of w
			if length of w > 1 then
				set restLetters to text 2 thru -1 of w
			else
				set restLetters to ""
			end if
			set firstUpper to my makeUpper2(firstLetter)
			set restLower to my makeLower2(restLetters)
			set end of newList to firstUpper & restLower
		end if
	end repeat
	-- Join list items with a single space
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to " "
	set resultText to newList as text
	set AppleScript's text item delimiters to oldDelims
	return resultText
end titleCase
