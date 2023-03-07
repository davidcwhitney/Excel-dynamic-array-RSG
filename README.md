# Description
Highly flexible random string generator with no recursion.  
Includes options for min/max ANSI code, min/max string length, count of output strings, alpha case. 
Also permits restricting output to ANSI ranges Alpha, Numeric, Alphanumeric, All printing, All.

Very fast for most tasks.  Passable if you want a lot of strings where some may be very short and others very long.  Tested 1 million random strings of 10 digits, run time between 2 and 3 mins.  Excel does not crash or barf or run out of memory or bog down the CPU.  It just takes its own sweet time.  

### Code
```
\\Inputs -- min and max ANSI code, min and max string length, count of output strings, alpha case, ANSI range 
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
\\ANSI range -- "Alpha", "Numeric", "Alphanumeric", "All printing", "All"

=LET(
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
    ANSISequence, SEQUENCE(255),
    ASCIICanonicalNonPrinting, SEQUENCE(32),
    ASCIIDigits, SEQUENCE(10, , CODE("0")),
    ASCIIUCaseLetters, SEQUENCE(26, , CODE("A")),
    ASCIILCaseLetters, SEQUENCE(26, , CODE("a")),
    ASCIIDelete, 127,
    ANSINonPrinting, {129; 131; 136; 144; 152},
    ANSINonBreakingSpace, 160,
    comment, "ANSINonAlphanumericPrinting is inferred as 'absent from all enumerated sets'",
    ANSIAlphanumericNonPrinting, HSTACK(
        NOT(ISNA(XMATCH(ANSISequence, ASCIIDigits, 0, 2))),
        NOT(ISNA(XMATCH(ANSISequence, ASCIIUCaseLetters, 0, 2))),
        NOT(ISNA(XMATCH(ANSISequence, ASCIILCaseLetters, 0, 2))),
        NOT(ISNA(XMATCH(ANSISequence, ASCIICanonicalNonPrinting, 0, 2))),
        NOT(ISNA(XMATCH(ANSISequence, ASCIIDelete, 0, 2))),
        NOT(ISNA(XMATCH(ANSISequence, ANSINonPrinting, 0, 2))),
        NOT(ISNA(XMATCH(ANSISequence, ANSINonBreakingSpace, 0, 2)))
    ),
    ANSINonAlphanumericPrintingUnfiltered, SEQUENCE(255) *
        --BYROW(ANSIAlphanumericNonPrinting, LAMBDA(r, ISNA(MATCH(TRUE, r, 0)))),
    ANSINonAlphanumericPrinting, FILTER(
        ANSINonAlphanumericPrintingUnfiltered,
        ANSINonAlphanumericPrintingUnfiltered <> 0
    ),
    ASCIILetters, SWITCH(
        AlphaCase,
        "Both", VSTACK(ASCIIUCaseLetters, ASCIILCaseLetters),
        "Upper", ASCIIUCaseLetters,
        "Lower", ASCIILCaseLetters
    ),
    ANSI, SWITCH(
        ANSIRange,
        "Alpha", ASCIILetters,
        "Numeric", ASCIIDigits,
        "Alphanumeric", VSTACK(ASCIIDigits, ASCIILetters),
        "All printing", SORT(VSTACK(ASCIIDigits, ASCIILetters, ANSINonAlphanumericPrinting)),
        "All", SEQUENCE(MaxANSI, , MinANSI)
    ),
    CountOfANSI, COUNT(ANSI),
    ANSIIndexed, HSTACK(SEQUENCE(CountOfANSI, , 1), ANSI),
    BaseArray, SWITCH(
        TRUE,
        OR(ANSIRange = "All printing", LEFT(ANSIRange, 5) = "Alpha"), CHAR(
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
