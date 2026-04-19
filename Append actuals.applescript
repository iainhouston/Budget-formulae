use scripting additions
use AppleScript version "2.4" -- Yosemite or later

-- Append Paid out Transactions from CSV to Numbers Actuals Table

-- Prompt user to select a CSV file and determine whether it is a credit card export
set csvFile to choose file with prompt "Select the Downloaded Actuals CSV file:"
set {csvLines, isCC} to loadCSV(csvFile)

-- Locate the header row; abort if not found
set headerIndex to findHeader(csvLines)
if headerIndex = 0 then
	display alert "Could not locate the CSV header row containing required columns." buttons {"OK"}
	return
end if

-- Extract column indices from the header; abort if any required column is missing
set cols to parseHeaderColumns(item headerIndex of csvLines, isCC)
if cols is false then
	display alert "Missing one or more required columns in header." buttons {"OK"}
	return
end if

-- Iterate over every data row below the header
repeat with k from (headerIndex + 1) to (count of csvLines)
	set csvLine to item k of csvLines
	if csvLine is not "" then

		-- Split the CSV row into individual field values
		set oldDelims to AppleScript's text item delimiters
		set AppleScript's text item delimiters to ","
		set parts to text items of csvLine
		set AppleScript's text item delimiters to oldDelims

		-- Extract the Paid-out amount and convert it to a number
		set pOut to stripQuotes(item (paidOutIdx of cols) of parts)
		set pOutNum to numericValue(pOut)

		-- Only process rows that represent a payment (positive paid-out value)
		if pOutNum > 0 then

			-- Extract the date and description; build a default comment with account prefix
			set dVal to stripQuotes(item (dateIdx of cols) of parts)
			set rawDesc to stripQuotes(item (descIdx of cols) of parts)
			set defaultComment to (commentPrefix of cols) & titleCase(rawDesc)

			-- Ask the user to confirm, edit, or skip this transaction
			set txnResult to promptForTransaction(dVal, pOutNum, defaultComment)
			if action of txnResult is "cancel" then return
			if action of txnResult is "continue" then

				-- Ask the user to assign a budget category, then write the row to Numbers
				set theComment to theComment of txnResult
				set theCategory to selectCategory(dVal)
				appendRowToNumbers(date dVal, theCategory, pOutNum, theComment)
			end if
		end if
	end if
end repeat

display dialog "All input CSV records have been read." buttons {"OK"}


-- Handler: read CSV file, return {lines, isCC}
on loadCSV(csvFile)
	-- Read the entire file and split into lines
	set csvContent to do shell script "cat " & quoted form of POSIX path of csvFile
	set csvLines to paragraphs of csvContent

	-- Detect credit-card exports by the presence of a header sentinel on the first line
	set isCC to false
	if (item 1 of csvLines) contains "Select Credit Card" then set isCC to true
	return {csvLines, isCC}
end loadCSV

-- Handler: find the header row index (returns 0 if not found)
on findHeader(csvLines)
	-- Scan lines until we find one containing all expected column headings
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
	-- Split the header line into individual column name tokens
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to ","
	set headerItems to text items of headerLine
	set AppleScript's text item delimiters to oldDelims

	-- Initialise index trackers and the account-type prefix for comments
	set dateIdx to 0
	set paidOutIdx to 0
	set descIdx to 0
	set commentPrefix to ""

	-- Map each column name to its 1-based index position
	repeat with j from 1 to (count of headerItems)
		set colName to item j of headerItems
		if colName starts with "\"" and colName ends with "\"" then set colName to text 2 thru -2 of colName
		if colName is "Date" then set dateIdx to j
		if colName is "Paid out" then set paidOutIdx to j
		-- Credit-card exports use "Transactions" as the description column
		if isCC and colName is "Transactions" then
			set descIdx to j
			set commentPrefix to "Visa: "
		end if
		-- Flex (current) account exports use "Description"
		if not isCC and colName is "Description" then
			set descIdx to j
			set commentPrefix to "Flex: "
		end if
	end repeat

	-- Return false if any required column was not found
	if dateIdx = 0 or paidOutIdx = 0 or descIdx = 0 then return false
	return {dateIdx:dateIdx, paidOutIdx:paidOutIdx, descIdx:descIdx, commentPrefix:commentPrefix}
end parseHeaderColumns

-- Handler: show dialog to edit comment; returns {action, theComment}
-- action is "cancel", "skip", or "continue"
on promptForTransaction(dVal, parsedAmount, defaultComment)
	-- Build the prompt text, using the £ character for the currency symbol
	set poundSign to character id 163
	set promptText to "Edit comment for transaction on " & dVal & " (" & poundSign & parsedAmount & "):"

	-- Display the editable dialog; user can cancel the whole script, skip this row, or proceed
	set userInput to display dialog promptText default answer defaultComment buttons {"Cancel Script", "Skip", "Continue"} default button "Continue"
	set btn to button returned of userInput

	-- Propagate a full cancellation back to the main script
	if btn is "Cancel Script" then
		display dialog "Script cancelled by user." buttons {"OK"}
		return {action:"cancel", theComment:""}
	end if

	-- Skip this row without writing anything to Numbers
	if btn is "Skip" then return {action:"skip", theComment:""}

	-- Use the user-edited comment, falling back to the default if left blank
	set theComment to text returned of userInput
	if theComment is "" then set theComment to defaultComment
	return {action:"continue", theComment:theComment}
end promptForTransaction

-- Handler: prompt user to pick a budget category
on selectCategory(dVal)
	-- Present the fixed list of budget categories for the user to choose from
	set theCategory to choose from list {"Home", "Insurance", "Eats", "Transport & Travel", "Savings", "Family", "Projects & Pastimes", "Health & Beauty", "Clothes", "Big One-off", "Charitable & Other"} with prompt "Select Category for " & dVal & ":"
	if theCategory is false then return ""
	return item 1 of theCategory
end selectCategory

-- Handler: append one row to the Actuals table in Numbers
on appendRowToNumbers(parsedDate, theCategory, parsedAmount, theComment)
	tell application "Numbers Creator Studio"
		activate
		set doc to front document
		-- Target the Actuals table on the Actual sheet
		set tbl to table "Actuals" of sheet "Actual" of doc
		tell tbl
			-- Add a new row at the bottom of the table
			make new row at end of rows
			set newRow to last row
			try
				-- Populate date, category, amount (formatted as currency), and comment
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
	-- Remove a single layer of enclosing double-quote characters if present
	if s starts with "\"" and s ends with "\"" then return text 2 thru -2 of s
	return s
end stripQuotes

-- Handler: convert currency string to number
-- Skips any leading non-numeric characters (handles £, Â£, ï¿½ etc.)
on numericValue(s)
	try
		set t to s
		-- Advance past any currency symbol or other non-digit prefix
		repeat while length of t > 0
			if text 1 of t is in "0123456789" then exit repeat
			set t to text 2 thru -1 of t
		end repeat
		if t is "" then return 0
		-- Remove thousands-separator commas so the string can be coerced to a number
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
	-- Split the input into individual words
	set wordList to words of inputText
	set newList to {}
	-- Capitalise the first letter of each word and lowercase the rest
	repeat with w in wordList
		set w to w as text
		if length of w > 0 then
			set firstLetter to text 1 thru 1 of w
			set restLetters to ""
			if length of w > 1 then set restLetters to text 2 thru -1 of w
			set end of newList to my makeUpper2(firstLetter) & my makeLower2(restLetters)
		end if
	end repeat
	-- Rejoin words with a single space separator
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to " "
	set resultText to newList as text
	set AppleScript's text item delimiters to oldDelims
	return resultText
end titleCase
