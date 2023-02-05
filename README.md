# VBA-Loop
I modified the VBA code that Macro Recorder generated to generate a loop.
I recorded two Macros, the first one I inserted a line at the top of the list that will add the headers, category, months and total.
So I stop recording that Macro and start recording a second Macro that does the formatting.
After recording the two Macros, I'm going to loop these procedures inside the four worksheets
Within VBA I created a variable representing each Worksheet.
Using the For Each loop and the select method, then inside the Worksheet it will select the next one in line.
