# Excel-dynamic-array-RSG
Highly flexible random string generator written using Microsoft 365 dynamic array formulae.  I wrote this as an exercise so it could be more efficient and/or elegant.  I am open to suggestions but not at the expense of making the formula unreadable.  As written, everything is obvious.

Acceptable performance tested up to 10,000 strings of random length between 3 and 200.  Excel does not crash or barf or run out of memory or bog down the CPU, it just takes several minutes if you want a lot of strings where some may be very long.  The bottleneck is not count of strings, it is max length of string.

Inputs:  min/max ANSI code, min/max string length, count of output strings, alpha case, ANSI range.  Alpha case only has effect when those characters are in the range selected by min/max or in the selected predefined range.  No parameters are optional but all have default values, so the formula may be written =Random_String_Generator(, , , , , , ).
Input boundaries:
min/max ANSI code -- 1 - 255
min/max string length -- 1 to Excel cell content limit at your own risk
count of output strings -- 1 to Excel worksheet row limit (tested 1 million random strings of 10 digits, run time between 2 and 3 mins)
alpha case -- "Both", "Upper", "Lower"
ANSI range -- "Alpha", "Numeric", "Alphanumeric", "All")
