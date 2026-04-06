-- Pre-compute month names for the dialogue
set currentDate to current date
set thisMonthDate to currentDate
copy currentDate to nextMonthDate
set month of nextMonthDate to (month of nextMonthDate) + 1

set thisMonthName to (month of thisMonthDate as text) & " " & (year of thisMonthDate as text)
set nextMonthName to (month of nextMonthDate as text) & " " & (year of nextMonthDate as text)

-- Ask user which month to use
set userChoice to button returned of (display dialog "Which budget month are you working on?
Dont forget to turn categories OFF in both Monthly Budget AND Budget Template" buttons {"Cancel", nextMonthName, thisMonthName} default button thisMonthName)

if userChoice is "Cancel" then error number -128

-- Set target date and month string
if userChoice is thisMonthName then
	set targetDate to thisMonthDate
else
	set targetDate to nextMonthDate
end if
set currentMonth to userChoice

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
					set estimatedMonthlyPayment to the value of cell 5 of row i
					set firstPayment to the value of cell 6 of row i
				end tell
			end tell

			-- Skip empty rows
			if budgetItem is missing value or budgetItem = "" then
				-- skip
			else
				set amountToUse to budgetedAmount
				set shouldAdd to false
				set showDialog to false

				if paymentInterval is "Monthly" then
					-- Copy across using budgeted amount
					set shouldAdd to true
				else if paymentInterval is "Weekly" then
					-- Copy across using estimated monthly payment
					set amountToUse to estimatedMonthlyPayment
					set shouldAdd to true
				else if paymentInterval is "Yearly" then
					-- Only prompt if firstPayment falls in the target month; otherwise auto-skip
					if firstPayment is not missing value then
						if (month of firstPayment) = (month of targetDate) then
							set showDialog to true
						end if
					end if
				else
					-- All other intervals: ask the user
					set showDialog to true
				end if

				if showDialog then
					-- Format firstPayment as Month Year
					set firstPaymentText to "unknown"
					if firstPayment is not missing value then
						set firstPaymentText to (month of firstPayment as text) & " " & (year of firstPayment as text)
					end if

					-- Discretionary items default to estimated monthly payment; others use budgeted amount
					if paymentInterval is "Discretionary" then
						set dialogDefault to estimatedMonthlyPayment
					else
						set dialogDefault to budgetedAmount
					end if

					-- Format default as currency (use character id to avoid encoding issues)
					set amountNum to (round (dialogDefault * 100)) / 100
					set amountText to (character id 163) & (amountNum as text)

					set infoText to "Item: " & budgetItem & "
Category: " & category & "
Interval: " & paymentInterval & "
First Payment: " & firstPaymentText & "

Enter the amount to budget this month:"

					set amountAnswer to display dialog infoText default answer amountText buttons {"Cancel Script", "Skip", "Add to Budget"} default button "Add to Budget"

					set clickedButton to button returned of amountAnswer
					if clickedButton is "Cancel Script" then
						error number -128
					else if clickedButton is "Add to Budget" then
						set enteredText to text returned of amountAnswer
						-- Strip leading £ if present
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
								display alert "Invalid amount entered for \"" & budgetItem & "\" — skipping."
								set shouldAdd to false
							end try
						end if
					end if
					-- if Skip was pressed, shouldAdd remains false
				end if

				if shouldAdd then
					try
						tell sheet "Monthly Budget" to tell table "Monthly Budget"
							add row below (cell 1 of last row)
							set the value of cell 1 of last row to currentMonth
							set the value of cell 2 of last row to budgetItem
							set the value of cell 3 of last row to category
							set the value of cell 4 of last row to paymentInterval
							set the value of cell 5 of last row to amountToUse
							set format of cell 5 of last row to currency
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
			end if
		end repeat
	end tell
end tell
