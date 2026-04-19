use AppleScript version "2.4"
use scripting additions

-- Script: FromSetAside
-- Purpose: For rows with "From Set-aside?" = TRUE:
--          - Add value in column C to total
--          - Set column F ("Reimbursed already?") to TRUE
--          - Set column H ("Date reimbursed") to current timestamp

tell application "Numbers Creator Studio"
	activate
	try
		set doc to front document
		set sh to sheet "Actuals" of doc
		set tbl to table "Actuals" of sh
		set totalAmount to 0
		
		set currentTime to current date
		
		-- Loop through rows (skip header row)
		repeat with i from 2 to (count of rows of tbl)
			set thisRow to row i of tbl
			set fromSetAsideValue to value of cell 5 of thisRow -- column E
			if fromSetAsideValue is true then
				-- Add amount (column C)
				set amountValue to value of cell 3 of thisRow
				if amountValue is not missing value then
					set totalAmount to totalAmount + amountValue
				end if
				
				-- Set column F ("Reimbursed already?") to TRUE
				set value of cell 6 of thisRow to true
				
				-- Set column H ("Date reimbursed") to current timestamp
				set value of cell 8 of thisRow to currentTime
			end if
		end repeat
		
		-- Show result
		set formattedAmount to "£" & totalAmount
		display dialog "Total reimbursed from Set-aside: " & formattedAmount buttons {"OK"} default button "OK"
		
	on error errMsg number errNum
		display alert "Error in FromSetAside script:" message errMsg as warning
	end try
end tell
