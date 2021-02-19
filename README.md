Author: Nosorio Liberato

Summary:

	This excel sheet is a modified version of the previously made excel sheet. The difference is that I implemented VBA code
that doesn't require colon when inputting the military time for the "print time" column and the "delivery time" column.
For example, when inputting 1930 in the military time cell, the excell automatically converts this as 7:30 p.m.
This saves a lot of time when inputting the times and doesn't require to press SHIFT + ";".

	In addition, when delivering more than one tray, one can input the cart number only once on the same row as the 
first food tray of the entire delivery. The excel sheet automatically fills in the cart number below in the "leaving kitchen" column. 
For example, if one inputs 6 new food trays (i.e., 6 rows) for a single run/delivery, then one can just input the cart number 
on row 1 before leaving the kitchen and the excell sheet will automatically populate the cells below until the 6th row.
Same goes for the "return to kitchen" column.

Special thanks to mikerickson from: https://www.mrexcel.com/board/threads/accept-military-time-without-a-colon.1014107/
for providing the starter code and giving me a place to start for the military time without colon.


