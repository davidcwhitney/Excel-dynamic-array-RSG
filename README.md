# Description
Highly flexible random string generator written using Microsoft Excel 365 dynamic array formulae and no recursive lambdas. 

Very fast for most tasks.  Excel does not crash or barf or run out of memory or bog down the CPU if you want a lot of strings where some may be very long.  It just takes its own sweet time.  Tested 1 million random strings of 10 digits, run time between 2 and 3 mins.  The bottleneck is not count of strings or length of string, it is gap between min and max length of string.  For fun tested filling the cell to 32,267 character limit.  One row is instantaneous.  Run time for 20 rows was a few seconds but this significantly taxed the machine.

### Code
```
\\Inputs:  

min and max ANSI code, min and max string length, count of output strings, alpha case, ANSI range.  

No parameters are optional but all have default values, so the formula may be written =Random_String_Generator(, , , , , , ).

Alpha case only has effect when those characters are in the range selected by min/max or in the selected predefined range.  It is possible to submit parameters that result in an empty array error or no visible characters.  

Input boundaries:

ANSI code -- 1 - 255 (default 1, 255)

string length -- 1 to Excel cell content limit (default 3, 100)

count of output strings -- 1 to Excel worksheet row limit (default 1)

alpha case -- "Both", "Upper", "Lower" (default "Both")

ANSI range -- "Alpha", "Numeric", "Alphanumeric", "All" (Default "All")
