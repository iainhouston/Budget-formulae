tell application "Numbers Creator Studio"
	activate
	tell document 1
		tell sheet "Sinking Funds"
			tell table "Sinking Fund Overview"
				
				-- discover how many rows to process
				set totalRows to row count
				set headerRows to header row count
				set footerRows to footer row count
				
				-- loop through each data row
				repeat with r from (headerRows + 1) to (totalRows - footerRows)
					-- get values from C and D
					set cVal to the value of cell 3 of row r
					set dVal to the value of cell 4 of row r
					
					-- add C into D
					set value of cell 4 of row r to (cVal + dVal)
					
					-- format column D as currency
					set format of cell 4 of row r to currency
					
					-- write today's date into column G (col 7)
					set value of cell 7 of row r to (current date)
				end repeat
				
			end tell
		end tell
	end tell
end tell
