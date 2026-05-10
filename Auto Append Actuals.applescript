use scripting additions
use AppleScript version "2.4"

-- Automated ingestion of Nationwide CSV exports into the Numbers Actuals table.
-- Finds all unprocessed "Statement Download" CSVs in Downloads, deduplicates
-- against a persistent fingerprint log, infers categories from Actuals history,
-- and appends new paid-out transactions to Numbers.

property storeDir : "/Users/iainhouston/Documents/Money/ActualsStore"
property currentSpreadsheetFile : "/Users/iainhouston/Documents/Money/ActualsStore/CurrentSpreadsheet.txt"
property processedFilePath : "/Users/iainhouston/Documents/Money/ActualsStore/processed.txt"
property greylistPath : "/Users/iainhouston/Documents/Money/ActualsStore/greylist.txt"
property downloadsDir : "/Users/iainhouston/Downloads"
property confidenceThreshold : 2 -- minimum significant-word overlap for silent auto-assign

-- Guarantee the store directory exists
do shell script "mkdir -p " & quoted form of storeDir

-- Load fingerprints of all previously processed transactions
set processedFPs to loadProcessed(processedFilePath)

-- Load greylist terms (transactions that must always go to the UI regardless of match confidence)
set greylist to loadGreylist(greylistPath)

-- Find all Statement Download CSV files currently in Downloads
set csvPaths to findCSVFiles(downloadsDir)
if (count of csvPaths) = 0 then
	display dialog "No 'Statement Download' CSV files found in Downloads." buttons {"OK"}
	return
end if

-- Open the budget spreadsheet from CurrentSpreadsheet.txt and read the Actuals history
set budgetDocPath to loadSpreadsheetPath(currentSpreadsheetFile)
ensureNumbersOpen(budgetDocPath)
set {wordLists, categories} to readActualsData()

-- Process each CSV in sequence, accumulating new fingerprints
set accumulatedFPs to {}
set csvFullyProcessed to {} -- only CSVs completed without a user cancel
repeat with csvPath in csvPaths
	set csvPath to csvPath as text
	set {csvLines, isCC} to loadCSVPath(csvPath)
	set headerIndex to findHeader(csvLines)
	if headerIndex = 0 then
		display alert "Header row not found; skipping: " & csvPath buttons {"OK"}
	else
		set cols to parseHeaderColumns(item headerIndex of csvLines, isCC)
		if cols is false then
			display alert "Required columns missing; skipping: " & csvPath buttons {"OK"}
		else
			set csvResult to processCSV(csvLines, headerIndex, cols, processedFPs, wordLists, categories, accumulatedFPs, greylist)
			set wasCancelled to item 1 of csvResult
			set accumulatedFPs to item 2 of csvResult
			if wasCancelled then
				-- Save progress so far, leave this CSV in Downloads for re-run
				exit repeat
			end if
			set end of csvFullyProcessed to csvPath
		end if
	end if
end repeat

-- Archive only the CSVs that were fully processed
repeat with processedPath in csvFullyProcessed
	archiveCSV(processedPath as text, storeDir)
end repeat

-- Persist all new fingerprints (including those from a cancelled mid-CSV run)
appendToProcessed(processedFilePath, accumulatedFPs)

set archivedCount to count of csvFullyProcessed
set newTxnCount to count of accumulatedFPs
display dialog (newTxnCount as text) & " new transaction(s) added. " & (archivedCount as text) & " CSV file(s) archived." buttons {"OK"}


-- ── Handlers ──────────────────────────────────────────────────────────────────

-- Read the fingerprint log and return it as a list of strings
on loadProcessed(filePath)
	try
		set content to do shell script "cat " & quoted form of filePath
		if content is "" then return {}
		set rawFPs to paragraphs of content
		set fps to {}
		repeat with fp in rawFPs
			set fp to fp as text
			if fp is not "" then set end of fps to fp
		end repeat
		return fps
	on error
		return {}
	end try
end loadProcessed

-- Return a list of POSIX paths for every "Statement Download*.csv" in the given directory
on findCSVFiles(dir)
	try
		set listing to do shell script "find " & quoted form of dir & " -maxdepth 1 -name 'Statement Download*.csv' 2>/dev/null"
		if listing is "" then return {}
		-- 'paragraphs' handles \r, \n and \r\n — do shell script returns \r-separated lines
		set rawPaths to paragraphs of listing
		set cleanPaths to {}
		repeat with p in rawPaths
			set p to p as text
			if p is not "" then set end of cleanPaths to p
		end repeat
		return cleanPaths
	on error
		return {}
	end try
end findCSVFiles

-- Read a CSV from a POSIX path string; detect Flex vs credit-card account
on loadCSVPath(csvPath)
	set content to do shell script "cat " & quoted form of csvPath
	set csvLines to paragraphs of content
	set isCC to false
	if (item 1 of csvLines) contains "Select Credit Card" then set isCC to true
	return {csvLines, isCC}
end loadCSVPath

-- Scan lines for the row containing Date / Paid out / Paid in; return its 1-based index
on findHeader(csvLines)
	repeat with i from 1 to (count of csvLines)
		set ln to item i of csvLines
		if ln contains "\"Date\"" and ln contains "\"Paid out\"" and ln contains "\"Paid in\"" then return i
	end repeat
	return 0
end findHeader

-- Map column names to 1-based indices; return a record or false if any required column is absent
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
		-- Credit-card exports label the description column "Transactions"
		if isCC and colName is "Transactions" then
			set descIdx to j
			set commentPrefix to "Visa: "
		end if
		-- Flex account exports label the description column "Description"
		if not isCC and colName is "Description" then
			set descIdx to j
			set commentPrefix to "Flex: "
		end if
	end repeat
	if dateIdx = 0 or paidOutIdx = 0 or descIdx = 0 then return false
	return {dateIdx:dateIdx, paidOutIdx:paidOutIdx, descIdx:descIdx, commentPrefix:commentPrefix}
end parseHeaderColumns

-- Open the budget spreadsheet in Numbers if it is not already open, then bring it to front
on ensureNumbersOpen(docPath)
	tell application "Numbers Creator Studio"
		set docIsOpen to false
		try
			repeat with d in documents
				try
					if (POSIX path of (path of d)) is docPath then
						set docIsOpen to true
						exit repeat
					end if
				end try
			end repeat
		end try
		if not docIsOpen then open POSIX file docPath
		activate
	end tell
end ensureNumbersOpen

-- Read the budget spreadsheet path from the CurrentSpreadsheet file
on loadSpreadsheetPath(pathFile)
	try
		set fileContent to do shell script "tr -d '\\r' < " & quoted form of pathFile & " | tr -d '\\n'"
		set fileContent to my trim(fileContent)
		if fileContent is "" then error "CurrentSpreadsheet.txt is empty"
		return fileContent
	on error errMsg number errNum
		display alert "Unable to read spreadsheet path from " & pathFile buttons {"OK"}
		error number errNum
	end try
end loadSpreadsheetPath

on trim(str)
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to {space, tab, linefeed, return}
	repeat while str starts with space or str starts with tab or str starts with linefeed or str starts with return
		set str to text 2 thru -1 of str
	end repeat
	repeat while str ends with space or str ends with tab or str ends with linefeed or str ends with return
		set str to text 1 thru -2 of str
	end repeat
	set AppleScript's text item delimiters to oldDelims
	return str
end trim

-- Read every Actuals row and return two parallel lists:
--   wordLists  — significant uppercased words precomputed for each comment
--   categories — the category assigned to each corresponding row
on readActualsData()
	set wordLists to {}
	set categories to {}
	tell application "Numbers Creator Studio"
		set tbl to table "Actuals" of sheet "Actual" of front document
		set rowCount to count of rows of tbl
		repeat with r from 2 to rowCount
			set theComment to value of cell 4 of row r of tbl
			set theCategory to value of cell 2 of row r of tbl
			if theComment is not missing value and theCategory is not missing value then
				set commentStr to theComment as text
				if commentStr is not "" then
					-- Precompute significant words once per row to avoid repeated shell calls during matching
					set end of wordLists to my significantWords(commentStr)
					set end of categories to theCategory as text
				end if
			end if
		end repeat
	end tell
	return {wordLists, categories}
end readActualsData

-- Iterate paid-out rows in one CSV; skip duplicates; assign or prompt for category;
-- append accepted transactions to Numbers.
-- Returns {wasCancelled, updatedFPs}. wasCancelled is true if the user clicked Cancel Script.
-- Fingerprints are recorded for both accepted and skipped transactions so they are
-- not re-presented on future runs; only fully cancelled transactions are left unrecorded.
on processCSV(csvLines, headerIndex, cols, processedFPs, wordLists, categories, newFPs, greylist)
	repeat with k from (headerIndex + 1) to (count of csvLines)
		set csvLine to item k of csvLines
		if csvLine is not "" then
			-- Split the row into fields
			set oldDelims to AppleScript's text item delimiters
			set AppleScript's text item delimiters to ","
			set parts to text items of csvLine
			set AppleScript's text item delimiters to oldDelims

			-- Only process rows with a positive paid-out amount
			set pOut to stripQuotes(item (paidOutIdx of cols) of parts)
			set pOutNum to numericValue(pOut)

			if pOutNum > 0 then
				set dVal to stripQuotes(item (dateIdx of cols) of parts)
				set rawDesc to stripQuotes(item (descIdx of cols) of parts)

				-- Fingerprint uniquely identifies this transaction across downloads
				set fp to (commentPrefix of cols) & "|" & dVal & "|" & pOut & "|" & rawDesc

				-- Skip if seen in a previous run or earlier in this run
				if processedFPs does not contain fp and newFPs does not contain fp then
					set theComment to (commentPrefix of cols) & my titleCase(rawDesc)
					set matchResult to my findBestCategory(wordLists, categories, rawDesc)
					set bestCat to item 1 of matchResult
					set bestScore to item 2 of matchResult

					-- Greylisted transactions always go to the UI regardless of match confidence
					set forcePrompt to my isOnGreylist(theComment, greylist)

					if bestScore >= confidenceThreshold and not forcePrompt then
						-- Confident match, not greylisted: silently assign and record
						my appendRowToNumbers(date dVal, bestCat, pOutNum, theComment)
						set end of newFPs to fp
					else
						-- Uncertain or greylisted: ask the user
						set userChoice to my promptForCategory(dVal, pOutNum, bestCat, theComment)
						if userChoice is false then
							-- User cancelled the entire script; return with FPs saved so far
							return {true, newFPs}
						end if
						-- Record the fingerprint whether the user accepted a category or skipped,
						-- so this transaction is not re-presented on the next run
						set end of newFPs to fp
						if userChoice is not "" then
							my appendRowToNumbers(date dVal, userChoice, pOutNum, theComment)
						end if
					end if
				end if
			end if
		end if
	end repeat
	return {false, newFPs}
end processCSV

-- Score rawDesc against all Actuals entries; return {bestCategory, bestScore}
on findBestCategory(wordLists, categories, rawDesc)
	if (count of categories) = 0 then return {"", 0}
	set descWords to my significantWords(rawDesc)
	if (count of descWords) = 0 then return {"", 0}
	set bestScore to 0
	set bestCat to ""
	repeat with i from 1 to (count of categories)
		set s to my wordOverlap(descWords, item i of wordLists)
		-- >= so the most recent row (highest index) wins ties, giving updated categories priority
		if s >= bestScore then
			set bestScore to s
			set bestCat to item i of categories
		end if
	end repeat
	return {bestCat, bestScore}
end findBestCategory

-- Uppercase a string via tr and return words of four or more characters
on significantWords(str)
	set upperStr to do shell script "echo " & quoted form of str & " | tr '[:lower:]' '[:upper:]'"
	set allWords to words of upperStr
	set sigWords to {}
	repeat with w in allWords
		set w to w as text
		-- Short words (GB, UK, etc.) and two-letter abbreviations are excluded as noise
		if length of w >= 4 then set end of sigWords to w
	end repeat
	return sigWords
end significantWords

-- Count how many words from listA appear in listB
on wordOverlap(listA, listB)
	set score to 0
	repeat with w in listA
		if listB contains (w as text) then set score to score + 1
	end repeat
	return score
end wordOverlap

-- Present a category prompt when automatic matching is uncertain or the transaction is greylisted.
-- Returns: chosen category string | "" (skip, record FP) | false (cancel whole script)
-- Buttons: Cancel Script | Skip | Accept… (Accept… opens the picker pre-selected to the suggestion)
on promptForCategory(dVal, amount, suggestedCat, theComment)
	set poundSign to character id 163
	set categoryList to {"Home", "Insurance", "Eats", "Transport & Travel", "Savings", "Family", "Projects & Pastimes", "Health & Beauty", "Clothes", "Big One-off", "Charitable & Other"}

	-- Build the prompt; include the suggestion when one exists
	set promptMsg to theComment & " — " & poundSign & amount & " on " & dVal
	if suggestedCat is not "" then set promptMsg to promptMsg & return & "Suggested: " & suggestedCat

	set btn to button returned of (display dialog promptMsg buttons {"Cancel Script", "Skip", "Accept…"} default button "Accept…")
	if btn is "Cancel Script" then
		display dialog "Script cancelled." buttons {"OK"}
		return false
	end if
	-- Skip: record the fingerprint so this transaction is not re-presented, but don't add to Actuals
	if btn is "Skip" then return ""

	-- Accept…: open the category picker pre-selected to the suggestion (user can confirm or change)
	if suggestedCat is not "" then
		set picked to choose from list categoryList with prompt "Confirm or change category for " & dVal & ":" default items {suggestedCat}
	else
		set picked to choose from list categoryList with prompt "Category for " & theComment & " (" & poundSign & amount & ") on " & dVal & ":"
	end if
	-- Dismissing the picker also counts as Skip
	if picked is false then return ""
	return item 1 of picked
end promptForCategory

-- Add one row to the Actuals table in the open Numbers document
on appendRowToNumbers(parsedDate, theCategory, parsedAmount, theComment)
	tell application "Numbers Creator Studio"
		set tbl to table "Actuals" of sheet "Actual" of front document
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
				display alert "Numbers error: ensure the Actuals table is not grouped by Category." buttons {"OK"} as warning
			end try
		end tell
	end tell
end appendRowToNumbers

-- Move a CSV file into the archive directory
on archiveCSV(csvPath, archiveDir)
	set fileName to do shell script "basename " & quoted form of csvPath
	do shell script "mv " & quoted form of csvPath & " " & quoted form of (archiveDir & "/" & fileName)
end archiveCSV

-- Append each new fingerprint as a line to the persistent log
on appendToProcessed(filePath, newFPs)
	repeat with fp in newFPs
		do shell script "printf '%s\n' " & quoted form of (fp as text) & " >> " & quoted form of filePath
	end repeat
end appendToProcessed

-- Load greylist.txt and return a list of uppercased match terms
on loadGreylist(filePath)
	try
		set content to do shell script "cat " & quoted form of filePath
		if content is "" then return {}
		set rawTerms to paragraphs of content
		set greylist to {}
		repeat with t in rawTerms
			set t to t as text
			if t is not "" then
				set end of greylist to do shell script "echo " & quoted form of t & " | tr '[:lower:]' '[:upper:]'"
			end if
		end repeat
		return greylist
	on error
		return {}
	end try
end loadGreylist

-- Return true if theComment contains any greylist term (case-insensitive substring match)
on isOnGreylist(theComment, greylist)
	if (count of greylist) = 0 then return false
	set upperComment to do shell script "echo " & quoted form of theComment & " | tr '[:lower:]' '[:upper:]'"
	repeat with term in greylist
		if upperComment contains (term as text) then return true
	end repeat
	return false
end isOnGreylist

-- Remove a single layer of enclosing double-quote characters if present
on stripQuotes(s)
	if s starts with "\"" and s ends with "\"" then return text 2 thru -2 of s
	return s
end stripQuotes

-- Strip leading non-numeric characters (handles £, encoding artefacts) then coerce to number
on numericValue(s)
	try
		set t to s
		repeat while length of t > 0
			if text 1 of t is in "0123456789" then exit repeat
			set t to text 2 thru -1 of t
		end repeat
		if t is "" then return 0
		-- Remove thousands-separator commas before coercion
		set AppleScript's text item delimiters to ","
		set parts to text items of t
		set AppleScript's text item delimiters to ""
		return (parts as text) as number
	on error
		return 0
	end try
end numericValue

-- Title-case a string in a single awk shell call
on titleCase(inputText)
	return do shell script "echo " & quoted form of inputText & " | awk '{for(i=1;i<=NF;i++) $i=toupper(substr($i,1,1)) tolower(substr($i,2)); print}'"
end titleCase
