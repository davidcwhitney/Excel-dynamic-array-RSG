# Description
Highly flexible random string generator with no recursion. 

Very fast for most tasks.  Passable if you want a lot of strings where some may be very short and others very long.  Tested 1 million random strings of 10 digits, run time between 2 and 3 mins.  Excel does not crash or barf or run out of memory or bog down the CPU.  It just takes its own sweet time.  

### Code
```
\\Inputs -- min and max ANSI code, min and max string length, count of output strings, alpha case, ANSI range.: 
\\
\\NOTE no parameter is optional but all have default values, so the formula may be written   
\\=Random_String_Generator(, , , , , , )
\\Alpha case only has effect when alpha characters are in the custom or predefined range selected.  
\\
\\Input boundaries:
\\ANSI code -- 1 - 255
\\string length -- 1 to Excel cell content limit
\\count of output strings -- 1 to Excel worksheet row limit
\\alpha case -- "Both", "Upper", "Lower"
\\ANSI range -- "Alpha", "Numeric", "Alphanumeric", "All"

Random_String_Generator = LET(
    DefaultMinANSI, 1,
    DefaultMaxANSI, 255,
    DefaultMinLength, 3,
    DefaultMaxLength, 100,
    DefaultCountOfStrings, 1,
    DefaultAlphaCase, "Both",
    DefaultANSIRange, "All",
    MinANSI, IF(ISOMITTED(intMinANSI), DefaultMinANSI, intMinANSI),
    MaxANSI, IF(ISOMITTED(intMaxANSI), DefaultMaxANSI, intMaxANSI),
    MinStringLength, IF(ISOMITTED(intMinLength), DefaultMinLength, intMinLength),
    MaxStringLength, IF(ISOMITTED(intMaxLength), DefaultMaxLength, intMaxLength),
    CountOfStrings, IF(ISOMITTED(intCountOfStrings), DefaultCountOfStrings, intCountOfStrings),
    AlphaCase, IF(ISOMITTED(strCaseSelection), DefaultAlphaCase, strCaseSelection),
    ANSIRange, IF(ISOMITTED(strANSIRangeSelection), DefaultANSIRange, strANSIRangeSelection),
    ASCIIDigits, SEQUENCE(10, , CODE("0")),
    ASCIILCaseLetters, SEQUENCE(26, , CODE("a")),
    ASCIIUCaseLetters, SEQUENCE(26, , CODE("A")),
    ASCIILetters, SWITCH(
        AlphaCase,
        "Both", VSTACK(ASCIILCaseLetters, ASCIIUCaseLetters),
        "Upper", ASCIIUCaseLetters,
        "Lower", ASCIILCaseLetters
    ),
    ANSI, SWITCH(
        ANSIRange,
        "Alpha", ASCIILetters,
        "Numeric", ASCIIDigits,
        "Alphanumeric", VSTACK(ASCIIDigits, ASCIILetters),
        "All", SEQUENCE(MaxANSI, , MinANSI)
    ),
    CountOfANSI, COUNT(ANSI),
    ANSIIndexed, HSTACK(SEQUENCE(CountOfANSI), ANSI),
    BaseArray, SWITCH(
        TRUE,
        LEFT(ANSIRange, 5) = "Alpha", CHAR(
            MAP(
                RANDARRAY(CountOfStrings, MaxStringLength, 1, CountOfANSI, TRUE),
                LAMBDA(x,
                    XLOOKUP(x, INDEX(ANSIIndexed, , 1), INDEX(ANSIIndexed, , 2), " absent", 0, 2)
                )
            )
        ),
        ANSIRange = "All", CHAR(RANDARRAY(CountOfStrings, MaxStringLength, MinANSI, MaxANSI, TRUE)),
        ANSIRange = "Numeric", CHAR(
            RANDARRAY(CountOfStrings, MaxStringLength, CODE("0"), CODE("9"), TRUE)
        )
    ),
    StringLengthArray, RANDARRAY(ROWS(BaseArray), , MinStringLength, MaxStringLength, TRUE),
    BaseArrayConcatenated, BYROW(BaseArray, LAMBDA(r, CONCAT(r))),
    output, LEFT(BaseArrayConcatenated, StringLengthArray),
    output
)
