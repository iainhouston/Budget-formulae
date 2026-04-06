-- Ask user which month to use
set userChoice to button returned of (display dialog "Which budget month are you working on?
Dont forget to turn categories OFF in both Monthly Budget AND Budget Template" buttons {"Cancel", "Next Month", "This Month"} default button "This Month")

if userChoice is "Cancel" then error number -128

-- Get current date components
set currentDate to current date
set monthOffset to 0
if userChoice is "Next Month" then set monthOffset to 1

-- Approximate adjustedDate
set adjustedDate to currentDate + (monthOffset * 30 * days)

-- Derive currentMonth via AppleScript to avoid shell overhead
set currentMonth to (month of adjustedDate as text) & " " & (year of adjustedDate as text)

-- Numbers interaction
tell application "Numbers"
	tell front document
		tell sheet "Budget Template"
			set templateTable to table "Spending Template"
			set rowCount to count of rows of templateTable
		end tell

		repeat with i from 2 to rowCount
			-- Read all fields from template
			tell sheet "Budget Template"
				tell templateTable
					set budgetItem to the value of cell 1 of row i
					set category to the value of cell 2 of row i
					set paymentInterval to the value of cell 3 of row i
					set budgetedAmount to the value of cell 4 of row i
					set firstPayment to the value of cell 6 of row i
				end tell
			end tell

			-- Skip empty rows
			if budgetItem is missing value or budgetItem = "" then
				-- skip
			else
				-- Append to Monthly Budget
				try
					tell sheet "Monthly Budget" to tell table "Monthly Budget"
						add row below (cell 1 of last row)
						set the value of cell 1 of last row to currentMonth
						set the value of cell 2 of last row to budgetItem
						set the value of cell 3 of last row to category
						set the value of cell 4 of last row to paymentInterval
						set the value of cell 5 of last row to budgetedAmount
						set the value of cell 6 of last row to firstPayment
					end tell
				on error errMsg number errNum
					if errMsg contains "Unable to retrieve existing cell value" then
						display alert "Script cancelled: Don't forget to turn categories OFF in both Monthly Budget AND Budget Template"
						error number -128 -- user canceled
					else
						error errMsg number errNum -- propagate other errors
					end if
				end try
			end if
		end repeat
	end tell
end tell
