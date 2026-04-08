use scripting additions
use AppleScript version "2.4" -- Yosemite or later

-- Budget This Month: populate Monthly Budget table from Budget Template

-- Compute month name options
set today to current date
copy today to nextMonth
set month of nextMonth to (month of nextMonth) + 1

set thisMonthLabel to monthLabel(today)
set nextMonthLabel to monthLabel(nextMonth)

-- Ask which month to budget for
set chosenMonth to button returned of (display dialog ¬
	"Which budget month are you working on?
Don't forget to turn categories OFF in both Monthly Budget AND Budget Template" ¬
	buttons {"Cancel", nextMonthLabel, thisMonthLabel} default button thisMonthLabel)

if chosenMonth is "Cancel" then error number -128

set targetDate to today
if chosenMonth is nextMonthLabel then set targetDate to nextMonth

-- Read all template rows upfront
set rowData to {}
tell application "Numbers"
	tell table "Spending Template" of sheet "Budget Template" of front document
		set rowCount to count of rows
		repeat with i from 2 to rowCount
			set r to row i
			set end of rowData to {¬
				value of cell 1 of r, ¬
				value of cell 2 of r, ¬
				value of cell 3 of r, ¬
				value of cell 4 of r, ¬
				value of cell 5 of r, ¬
				value of cell 6 of r}
		end repeat
	end tell
end tell

-- Process each template row
repeat with rowVals in rowData
	set budgetItem to item 1 of rowVals
	set category to item 2 of rowVals
	set paymentInterval to item 3 of rowVals
	set budgetedAmount to item 4 of rowVals
	set estimatedMonthlyPayment to item 5 of rowVals
	set firstPayment to item 6 of rowVals

	if budgetItem is missing value or budgetItem = "" then
		-- skip empty rows
	else
		set amountToUse to budgetedAmount
		set shouldAdd to false
		set showDialog to false

		if paymentInterval is "Monthly" then
			set shouldAdd to true
		else if paymentInterval is "Weekly" then
			set amountToUse to estimatedMonthlyPayment
			set shouldAdd to true
		else if paymentInterval is "Yearly" then
			if firstPayment is not missing value and (month of firstPayment) = (month of targetDate) then
				set showDialog to true
			end if
		else
			-- Discretionary and all other intervals: ask the user
			set showDialog to true
		end if

		if showDialog then
			set amountToUse to my promptForAmount(budgetItem, category, paymentInterval, firstPayment, budgetedAmount, estimatedMonthlyPayment)
			if amountToUse is not missing value then set shouldAdd to true
		end if

		if shouldAdd then
			my addBudgetRow(chosenMonth, budgetItem, category, paymentInterval, amountToUse)
		end if
	end if
end repeat

-- Handler: format a date as "Month YYYY"
on monthLabel(d)
	return (month of d as text) & " " & (year of d as text)
end monthLabel

-- Handler: format a number as £X.XX
on asCurrency(n)
	set rounded to (round (n * 100)) / 100
	return (character id 163) & (rounded as text)
end asCurrency

-- Handler: prompt user for an amount to budget
-- Returns the numeric amount, or missing value if skipped
-- Raises error -128 if user cancels the script
on promptForAmount(budgetItem, category, paymentInterval, firstPayment, budgetedAmount, estimatedMonthlyPayment)
	set firstPaymentText to "unknown"
	if firstPayment is not missing value then
		set firstPaymentText to (month of firstPayment as text) & " " & (year of firstPayment as text)
	end if

	if paymentInterval is "Discretionary" then
		set defaultAmount to estimatedMonthlyPayment
	else
		set defaultAmount to budgetedAmount
	end if

	set infoText to "Item: " & budgetItem & "
Category: " & category & "
Interval: " & paymentInterval & "
Payment Date: " & firstPaymentText & "

Enter the amount to budget this month:"

	set amountAnswer to display dialog infoText default answer my asCurrency(defaultAmount) ¬
		buttons {"Cancel Script", "Skip", "Add to Budget"} default button "Add to Budget"

	set clickedButton to button returned of amountAnswer
	if clickedButton is "Cancel Script" then error number -128
	if clickedButton is "Skip" then return missing value

	set enteredText to text returned of amountAnswer
	if enteredText starts with (character id 163) then set enteredText to text 2 thru -1 of enteredText
	if enteredText is "" then return missing value

	try
		return enteredText as number
	on error
		display alert "Invalid amount entered for \"" & budgetItem & "\" — skipping."
		return missing value
	end try
end promptForAmount

-- Handler: append one row to the Monthly Budget table
on addBudgetRow(currentMonth, budgetItem, category, paymentInterval, amountToUse)
	tell application "Numbers"
		try
			tell table "Monthly Budget" of sheet "Monthly Budget" of front document
				add row below (cell 1 of last row)
				set value of cell 1 of last row to currentMonth
				set value of cell 2 of last row to budgetItem
				set value of cell 3 of last row to category
				set value of cell 4 of last row to paymentInterval
				set value of cell 5 of last row to amountToUse
				set format of cell 5 of last row to currency
			end tell
		on error errMsg number errNum
			if errMsg contains "Unable to retrieve existing cell value" then
				display alert "Script cancelled: Don't forget to turn categories OFF in both Monthly Budget AND Budget Template"
				error number -128
			else
				error errMsg number errNum
			end if
		end try
	end tell
end addBudgetRow
