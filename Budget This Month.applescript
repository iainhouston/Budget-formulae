use scripting additions
use AppleScript version "2.4"

property currentSpreadsheetFile : "/Users/iainhouston/Documents/Money/ActualsStore/CurrentSpreadsheet.txt"

-- Budget This Month: copy filtered Spending Template rows into Monthly Budget for a chosen month.

-- ── Month selection ────────────────────────────────────────────────────────────

set today to current date
copy today to nextMonth
set month of nextMonth to (month of nextMonth) + 1

set thisMonthLabel to monthLabel(today)
set nextMonthLabel to monthLabel(nextMonth)

set chosenMonth to button returned of (display dialog ¬
	"Which budget month are you working on?

** Don't forget to turn Categories OFF, Table Sorting and Filters OFF **
** in both Monthly Budget AND Budget Template. **" ¬
	buttons {"Cancel", nextMonthLabel, thisMonthLabel} default button thisMonthLabel)

if chosenMonth is "Cancel" then error number -128

if chosenMonth is thisMonthLabel then
	set targetDate to today
else
	set targetDate to nextMonth
end if

-- ── Open spreadsheet ──────────────────────────────────────────────────────────

set budgetSpreadsheetPath to loadSpreadsheetPath(currentSpreadsheetFile)
set budgetDoc to ensureBudgetDocumentOpen(budgetSpreadsheetPath)

-- ── Read Spending Template (column-by-column to avoid row-access restrictions) ─

set rowData to {}
tell application "Numbers Creator Studio"
	tell table "Spending Template" of sheet "Budget Template" of budgetDoc
		set col1 to value of every cell of column 1
		set col2 to value of every cell of column 2
		set col3 to value of every cell of column 3
		set col4 to value of every cell of column 4
		set col5 to value of every cell of column 5
		set col6 to value of every cell of column 6
	end tell
end tell

repeat with i from 2 to (count of col1)
	set budgetItem to item i of col1
	set category to item i of col2
	if budgetItem is missing value or budgetItem = "" then
		-- skip blank rows
	else if category is missing value or category = "" or category is "Enter Category" then
		-- skip placeholder rows
	else
		set end of rowData to {¬
			budgetItem, ¬
			category, ¬
			item i of col3, ¬
			item i of col4, ¬
			item i of col5, ¬
			item i of col6}
	end if
end repeat

-- ── Filter and append rows to Monthly Budget ─────────────────────────────────

set addedCount to 0
repeat with rowVals in rowData
	set budgetItem to item 1 of rowVals
	set category to item 2 of rowVals
	set paymentInterval to item 3 of rowVals
	set budgetedAmount to item 4 of rowVals
	set estimatedMonthlyPayment to item 5 of rowVals
	set firstPayment to item 6 of rowVals

	set amountToUse to budgetedAmount
	set shouldAdd to false
	set showDialog to false

	if paymentInterval is "Monthly" then
		-- Always include; use budgeted amount
		set shouldAdd to true
	else if paymentInterval is "Weekly" then
		-- Always include; use pre-calculated monthly equivalent
		set amountToUse to estimatedMonthlyPayment
		set shouldAdd to true
	else if paymentInterval is "Yearly" then
		-- Only include in the month the payment falls due
		if firstPayment is not missing value then
			if (month of firstPayment) = (month of targetDate) then
				set showDialog to true
			end if
		end if
	else
		-- Three-yearly, Discretionary, Fortnightly, etc. — always prompt
		set showDialog to true
	end if

	if showDialog then
		set firstPaymentText to "unknown"
		if firstPayment is not missing value then
			set firstPaymentText to (month of firstPayment as text) & " " & (year of firstPayment as text)
		end if

		-- Discretionary and Three-yearly default to estimated monthly; others to full budgeted amount
		if paymentInterval is "Discretionary" or paymentInterval is "Three-yearly" then
			set dialogDefault to estimatedMonthlyPayment
		else
			set dialogDefault to budgetedAmount
		end if
		if dialogDefault is missing value then set dialogDefault to 0

		set amountNum to (round (dialogDefault * 100)) / 100
		set amountText to (character id 163) & (amountNum as text)

		set infoText to "Item: " & budgetItem & "
Category: " & category & "
Interval: " & paymentInterval & "
First Payment: " & firstPaymentText & "

Enter the amount to budget this month:"

		set amountAnswer to display dialog infoText default answer amountText ¬
			buttons {"Cancel Script", "Skip", "Add to Budget"} default button "Add to Budget"

		set clickedButton to button returned of amountAnswer
		if clickedButton is "Cancel Script" then
			error number -128
		else if clickedButton is "Add to Budget" then
			set enteredText to text returned of amountAnswer
			if enteredText starts with (character id 163) then
				set enteredText to text 2 thru -1 of enteredText
			end if
			if enteredText is "" then
				set shouldAdd to false
			else
				try
					set amountToUse to enteredText as number
					set shouldAdd to true
				on error
					display alert "Invalid amount for \"" & budgetItem & "\" — skipping."
					set shouldAdd to false
				end try
			end if
		end if
		-- if Skip: shouldAdd remains false
	end if

	if shouldAdd then
		my appendBudgetRow(budgetDoc, chosenMonth, budgetItem, category, paymentInterval, amountToUse)
		set addedCount to addedCount + 1
	end if
end repeat

display dialog (addedCount as text) & " rows added to Monthly Budget for " & chosenMonth & "." buttons {"OK"}

-- ── Handlers ──────────────────────────────────────────────────────────────────

on monthLabel(d)
	return (month of d as text) & " " & (year of d as text)
end monthLabel

on appendBudgetRow(budgetDoc, chosenMonth, budgetItem, category, paymentInterval, amountToUse)
	tell application "Numbers Creator Studio"
		try
			tell table "Monthly Budget" of sheet "Monthly Budget" of budgetDoc
				make new row at end of rows
				tell last row
					set value of cell 1 to chosenMonth
					set value of cell 2 to budgetItem
					set value of cell 3 to category
					set value of cell 4 to paymentInterval
					if amountToUse is not missing value then
						set value of cell 5 to amountToUse
					end if
					tell cell 5 to set format to currency
				end tell
			end tell
		on error errMsg number errNum
			if errMsg contains "Unable to retrieve existing cell value" then
				display alert "Turn Categories OFF in both Monthly Budget and Budget Template, then re-run."
				error number -128
			else if errNum is -128 then
				error number -128
			else
				display alert "Skipped row for: " & budgetItem & return & errMsg
			end if
		end try
	end tell
end appendBudgetRow

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

on ensureBudgetDocumentOpen(docPath)
	tell application "Numbers Creator Studio"
		set targetDoc to missing value
		repeat with d in documents
			try
				if (POSIX path of (path of d)) is docPath then
					set targetDoc to d
					exit repeat
				end if
			end try
		end repeat
		if targetDoc is missing value then
			open POSIX file docPath
			set targetDoc to front document
		else
			set index of targetDoc to 1
		end if
		activate
		return targetDoc
	end tell
end ensureBudgetDocumentOpen

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
