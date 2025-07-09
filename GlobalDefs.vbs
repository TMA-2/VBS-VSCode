''' <summary>Returns the absolute value of a number.</summary>
''' <param name="expr">Any valid numeric expression.</param>
Function Abs(expr) ' As Integer
End Function

''' <summary>Returns a Variant containing an array.</summary>
''' <param name="arglist">arglist argument is a comma-delimited list of values that are assigned to the elements of an array</param>
Function Array(arglist)
End Function

''' <summary>Returns the unicode code of a character.</summary>
''' <param name="char">The character to get the code for. If a string is used, the code for the first character is given.</param>
Function Asc(char)
End Function

''' <summary>Returns the ANSI character code corresponding to the first letter in a string.</summary>
''' <param name="char">The character to get the ANSI code for.</param>
Function AscB(char)
End Function

''' <summary>Function that returns the Unicode (wide) character code that represents a specific Unicode character.</summary>
''' <param name="char">The Unicode character to get the code for.</param>
Function AscW(char)
End Function

''' <summary>Returns the arctangent of a number.</summary>
''' <param name="number">Any valid numeric expression.</param>
Function Atn(number)
End Function

''' <summary>Returns an expression that has been converted to a Variant of subtype Boolean.</summary>
''' <param name="expr">Any valid expression.</param>
Function CBool(expr) ' As Boolean
End Function

''' <summary>Returns an expression that has been converted to a Variant of subtype Byte.</summary>
''' <param name="expr">Any valid expression.</param>
Function CByte(expr) ' As Byte
End Function

''' <summary>Returns an expression that has been converted to a Variant of subtype Currency.</summary>
''' <param name="expr">Any valid string expression or numeric expression that can be converted to a currency value.</param>
Function CCur(expr) ' As Currency
End Function

''' <summary>Converts an expression to a Date subtype.</summary>
''' <param name="expr">Any valid expression that can be interpreted as a date.</param>
Function CDate(expr) ' As Date
End Function

''' <summary>Converts an expression to a Double subtype.</summary>
''' <param name="expr">Any valid numeric expression.</param>
Function CDbl(expr) ' As Double
End Function

''' <summary>Returns the character associated with the specified ANSI character code.</summary>
''' <param name="charcode">A number that identifies a character.</param>
Function Chr(charcode)
End Function

''' <summary>Returns a string containing the byte associated with the specified character code.</summary>
''' <param name="charcode">A number that identifies a character.</param>
Function ChrB(charcode)
End Function

''' <summary>Returns the Unicode character associated with the specified character code.</summary>
''' <param name="charcode">A number that identifies a Unicode character.</param>
Function ChrW(charcode)
End Function

''' <summary>Converts an expression to an Integer subtype.</summary>
''' <param name="expr">Any valid numeric expression.</param>
Function CInt(expr) ' As Integer
End Function

''' <summary>Converts an expression to a Long subtype.</summary>
''' <param name="expr">Any valid numeric expression.</param>
Function CLng(expr) ' As Long
End Function

''' <summary>Returns the cosine of an angle.</summary>
''' <param name="number">Any valid numeric expression that expresses an angle in radians.</param>
Function Cos(number)
End Function

''' <summary>Creates and returns a reference to an Automation object.</summary>
''' <param name="classname">The application name and class of the object to create.</param>
Function CreateObject(classname)
End Function

''' <summary>Creates and returns a reference to an Automation object on a remote server.</summary>
''' <param name="classname">The application name and class of the object to create.</param>
''' <param name="location">The name of the network server where the object is to be created.</param>
Function CreateObject(classname, location)
End Function

''' <summary>Converts an expression to a Single subtype.</summary>
''' <param name="expr">Any valid numeric expression.</param>
Function CSng(expr) ' As Single
End Function

''' <summary>Converts an expression to a String subtype.</summary>
''' <param name="expr">Any valid expression.</param>
Function CStr(expr) ' As String
End Function

''' <summary>Returns the current system date.</summary>
Function Date()
End Function

''' <summary>Returns a date to which a specified time interval has been added.</summary>
''' <param name="interval">String expression that is the interval you want to add</param>
''' <param name="number">Numeric expression that is the number of interval you want to add. The numeric expression can either be positive, for dates in the future, or negative, for dates in the past.</param>
''' <param name="date">Variant or literal representing the date to which interval is added.</param>
Function DateAdd(interval, number, date)
End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">String expression that is the interval you want to use to calculate the differences between date1 and date2</param>
''' <param name="date1">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Date expressions. Two dates you want to use in the calculation.</param>
Function DateDiff(interval, date1, date2)
End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">String expression that is the interval you want to use to calculate the differences between date1 and date2</param>
''' <param name="date1">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed</param>
Function DateDiff(interval, date1, date2, firstdayofweek)
End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">String expression that is the interval you want to use to calculate the differences between date1 and date2</param>
''' <param name="date1">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed</param>
''' <param name="firstweekofyear">Constant that specifies the first week of the year. If not specified, the first week is assumed to be the week in which January 1 occurs</param>
Function DateDiff(interval, date1, date2, firstdayofweek, firstweekofyear)
End Function

''' <summary>Returns the specified part of a given date.</summary>
''' <param name="interval">String expression that is the interval of time you want to return</param>
''' <param name="date">Date expression you want to evaluate</param>
Function DatePart(interval, date)
End Function

''' <summary>Returns the specified part of a given date.</summary>
''' <param name="interval">String expression that is the interval of time you want to return</param>
''' <param name="date">Date expression you want to evaluate</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed</param>
Function DatePart(interval, date, firstdayofweek)
End Function

''' <summary>Returns the specified part of a given date.</summary>
''' <param name="interval">String expression that is the interval of time you want to return</param>
''' <param name="date">Date expression you want to evaluate</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed</param>
''' <param name="firstweekofyear">Constant that specifies the first week of the year. If not specified, the first week is assumed to be the week in which January 1 occurs</param>
Function DatePart(interval, date, firstdayofweek, firstweekofyear)
End Function

''' <summary>Returns a Variant of subtype Date for a specified year, month, and day.</summary>
''' <param name="year">Number between 100 and 9999, inclusive, or a numeric expression.</param>
''' <param name="month">A number representing the month.</param>
''' <param name="day">A number representing the day.</param>
Function DateSerial(year, month, day)
End Function

''' <summary>Returns a Variant of subtype Date.</summary>
''' <param name="date">String expression representing a date.</param>
Function DateValue(date)
End Function

''' <summary>Returns a whole number between 1 and 31, inclusive, representing the day of the month.</summary>
''' <param name="date">Any expression that can represent a date.</param>
Function Day(date)
End Function

''' <summary>Returns a string with certain characters replaced with escape sequences.</summary>
''' <param name="str">String to be escaped.</param>
Function Escape(str) ' As String
End Function

''' <summary>Evaluates an expression and returns the result.</summary>
''' <param name="expr">String expression containing any valid VBScript expression.</param>
Function Eval(expr)
End Function

''' <summary>Returns e (the base of natural logarithms) raised to a power.</summary>
''' <param name="number">Any valid numeric expression.</param>
Function Exp(number)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
''' <param name="InputStrings">One-dimensional array of strings to be searched.</param>
''' <param name="Value">String to search for.</param>
Function Filter(InputStrings, Value)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
''' <param name="InputStrings">One-dimensional array of strings to be searched.</param>
''' <param name="Value">String to search for.</param>
''' <param name="Include">Boolean value indicating whether to return substrings that include or exclude Value. If Include is True, Filter returns the subset of the array that contains Value as a substring. If Include is False, Filter returns the subset of the array that does not contain Value as a substring.</param>
Function Filter(InputStrings, Value, Include)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
''' <param name="InputStrings">One-dimensional array of strings to be searched.</param>
''' <param name="Value">String to search for.</param>
''' <param name="Include">Boolean value indicating whether to return substrings that include or exclude Value. If Include is True, Filter returns the subset of the array that contains Value as a substring. If Include is False, Filter returns the subset of the array that does not contain Value as a substring.</param>
''' <param name="Compare">Numeric value indicating the kind of string comparison to use</param>
Function Filter(InputStrings, Value, Include, Compare)
End Function

''' <summary>Returns the integer portion of a number.</summary>
''' <param name="number">Any valid numeric expression.</param>
Function Fix(number)
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
Function FormatCurrency(Expression) ' As String
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal) ' As String
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit) ' As String
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values</param>
''' <param name="UseParensForNegativeNumbers">Tristate constant that indicates whether or not to place negative values within parentheses</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers) ' As String
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values</param>
''' <param name="UseParensForNegativeNumbers">Tristate constant that indicates whether or not to place negative values within parentheses</param>
''' <param name="GroupDigits">Tristate constant that indicates whether or not numbers are grouped using the group delimiter specified in the computer's regional settings</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) ' As String
End Function

Function FormatDateTime(Date) ' As String
End Function

Function FormatDateTime(Date, NamedFormat) ' As String
End Function

''' <summary>Returns an expression formatted as a number.</summary>
''' <param name="Expression">Expression to be formatted.</param>
Function FormatNumber(Expression) ' As String
End Function

''' <summary>Returns an expression formatted as a number.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed.</param>
Function FormatNumber(Expression, NumDigitsAfterDecimal) ' As String
End Function

''' <summary>Returns an expression formatted as a number.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed.</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values.</param>
Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit) ' As String
End Function

''' <summary>Returns an expression formatted as a number.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed.</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values.</param>
''' <param name="UseParensForNegativeNumbers">Tristate constant that indicates whether or not to place negative values within parentheses.</param>
Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers) ' As String
End Function

''' <summary>Returns an expression formatted as a number.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed.</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values.</param>
''' <param name="UseParensForNegativeNumbers">Tristate constant that indicates whether or not to place negative values within parentheses.</param>
''' <param name="GroupDigits">Tristate constant that indicates whether or not numbers are grouped using the group delimiter.</param>
Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) ' As String
End Function

''' <summary>Returns an expression formatted as a percentage (multiplied by 100) with a trailing % character.</summary>
''' <param name="Expression">Expression to be formatted.</param>
Function FormatPercent(Expression) ' As String
End Function

''' <summary>Returns an expression formatted as a percentage (multiplied by 100) with a trailing % character.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed.</param>
Function FormatPercent(Expression, NumDigitsAfterDecimal) ' As String
End Function

''' <summary>Returns an expression formatted as a percentage (multiplied by 100) with a trailing % character.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed.</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values.</param>
Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit) ' As String
End Function

''' <summary>Returns an expression formatted as a percentage (multiplied by 100) with a trailing % character.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed.</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values.</param>
''' <param name="UseParensForNegativeNumbers">Tristate constant that indicates whether or not to place negative values within parentheses.</param>
Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers) ' As String
End Function

''' <summary>Returns an expression formatted as a percentage (multiplied by 100) with a trailing % character.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed.</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values.</param>
''' <param name="UseParensForNegativeNumbers">Tristate constant that indicates whether or not to place negative values within parentheses.</param>
''' <param name="GroupDigits">Tristate constant that indicates whether or not numbers are grouped using the group delimiter.</param>
Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) ' As String
End Function

''' <summary>Returns the current locale identifier for the system.</summary>
Function GetLocale() ' As Long
End Function

''' <summary>Returns an Automation object from a file.</summary>
''' <param name="pathname">Full path and name of the file containing the object to retrieve.</param>
Function GetObject(pathname) ' As Object
End Function

''' <summary>Returns an Automation object from a file.</summary>
''' <param name="pathname">Full path and name of the file containing the object to retrieve.</param>
''' <param name="classname">String representing the class of the object.</param>
Function GetObject(pathname, classname) ' As Object
End Function

''' <summary>Returns a reference to a procedure.</summary>
''' <param name="procname">The name of the procedure.</param>
Function GetRef(procname)
End Function

''' <summary>Returns the locale identifier for the language used by the host application.</summary>
Function GetUILanguage() ' As Integer
End Function

''' <summary>Returns a string representing the hexadecimal value of a number.</summary>
''' <param name="number">Any valid numeric expression.</param>
Function Hex(number) ' As String
End Function

''' <summary>Returns a whole number between 0 and 23, inclusive, representing the hour of the day.</summary>
''' <param name="time">Any expression that can represent a time.</param>
Function Hour(time) ' As Integer
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box.</param>
Function InputBox(prompt)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box.</param>
Function InputBox(prompt, title)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box.</param>
''' <param name="default">String expression displayed in the text box as the default response.</param>
Function InputBox(prompt, title, default)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box.</param>
''' <param name="default">String expression displayed in the text box as the default response.</param>
''' <param name="xpos">Numeric expression that specifies the horizontal distance of the left edge of the dialog box from the left edge of the screen.</param>
Function InputBox(prompt, title, default, xpos)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box.</param>
''' <param name="default">String expression displayed in the text box as the default response.</param>
''' <param name="xpos">Numeric expression that specifies the horizontal distance of the left edge of the dialog box from the left edge of the screen.</param>
''' <param name="ypos">Numeric expression that specifies the vertical distance of the upper edge of the dialog box from the top of the screen.</param>
Function InputBox(prompt, title, default, xpos, ypos)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box.</param>
''' <param name="default">String expression displayed in the text box as the default response.</param>
''' <param name="xpos">Numeric expression that specifies the horizontal distance of the left edge of the dialog box from the left edge of the screen.</param>
''' <param name="ypos">Numeric expression that specifies the vertical distance of the upper edge of the dialog box from the top of the screen.</param>
''' <param name="helpfile">String expression that identifies the Help file to use to provide context-sensitive Help for the dialog box.</param>
''' <param name="context">Numeric expression that identifies the Help context number assigned by the Help author to the appropriate Help topic.</param>
Function InputBox(prompt, title, default, xpos, ypos, helpfile, context)
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
Function InStr(string1, string2) ' As Long
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="start">Specifies the starting position for search</param>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
Function InStr(start, string1, string2) ' As Long
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
''' <param name="compare">Specifies the string comparison to use</param>
Function InStr(string1, string2, compare) ' As Long
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="start">Specifies the starting position for search</param>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
''' <param name="compare">Specifies the string comparison to use</param>
Function InStr(start, string1, string2, compare) ' As Long
End Function

''' <summary>Returns the byte position of the first occurrence of one string within another.</summary>
''' <param name="string1">The string to be searched.</param>
''' <param name="string2">The string expression to search for.</param>
Function InStrB(string1, string2) ' As Long
End Function
''' <summary>Returns the byte position of the first occurrence of one string within another.</summary>
''' <param name="start">Specifies the starting position for search.</param>
''' <param name="string1">The string to be searched.</param>
''' <param name="string2">The string expression to search for.</param>
Function InStrB(start, string1, string2) ' As Long
End Function
''' <summary>Returns the byte position of the first occurrence of one string within another.</summary>
''' <param name="string1">The string to be searched.</param>
''' <param name="string2">The string expression to search for.</param>
''' <param name="compare">Specifies the string comparison to use.</param>
Function InStrB(string1, string2, compare) ' As Long
End Function
''' <summary>Returns the byte position of the first occurrence of one string within another.</summary>
''' <param name="start">Specifies the starting position for search.</param>
''' <param name="string1">The string to be searched.</param>
''' <param name="string2">The string expression to search for.</param>
''' <param name="compare">Specifies the string comparison to use.</param>
Function InStrB(start, string1, string2, compare) ' As Long
End Function

''' <summary>Returns the position of an occurrence of one string within another, from the end of string.</summary>
''' <param name="string1">The string to be searched.</param>
''' <param name="string2">The string expression to search for.</param>
Function InStrRev(string1, string2) ' As Long
End Function

''' <summary>Returns the position of an occurrence of one string within another, from the end of string.</summary>
''' <param name="string1">The string to be searched.</param>
''' <param name="string2">The string expression to search for.</param>
''' <param name="start">Specifies the starting position for search from the end.</param>
Function InStrRev(string1, string2, start) ' As Long
End Function

''' <summary>Returns the position of an occurrence of one string within another, from the end of string.</summary>
''' <param name="string1">The string to be searched.</param>
''' <param name="string2">The string expression to search for.</param>
''' <param name="start">Specifies the starting position for search from the end.</param>
''' <param name="compare">Specifies the string comparison to use.</param>
Function InStrRev(string1, string2, start, compare) ' As Long
End Function

''' <summary>Returns the integer portion of a number.</summary>
''' <param name="number">Any valid numeric expression.</param>
Function Int(number)
End Function

''' <summary>Returns a Boolean value indicating whether a variable is an array.</summary>
''' <param name="var">Variable to test.</param>
Function IsArray(var) ' As Boolean
End Function

''' <summary>Returns a Boolean value indicating whether an expression can be converted to a date.</summary>
''' <param name="expr">Any variable or expression.</param>
Function IsDate(expr) ' As Boolean
End Function

''' <summary>Returns a Boolean value indicating whether a variable has been initialized.</summary>
''' <param name="expr">A variable or expression.</param>
Function IsEmpty(expr) ' As Boolean
End Function

''' <summary>Returns a Boolean value indicating whether an expression contains no valid data (Null).</summary>
''' <param name="expr">Any variable or expression.</param>
Function IsNull(expr) ' As Boolean
End Function

''' <summary>Returns a Boolean value indicating whether an expression can be evaluated as a number.</summary>
''' <param name="expr">Any variable or expression.</param>
Function IsNumeric(expr) ' As Boolean
End Function

''' <summary>Returns a Boolean value indicating whether an expression references a valid Automation object.</summary>
''' <param name="expr">Any variable or expression.</param>
Function IsObject(expr) ' As Boolean
End Function

''' <summary>Returns a string created by joining a number of substrings contained in an array.</summary>
''' <param name="list">One-dimensional array containing substrings to be joined.</param>
Function Join(list) ' As String
End Function

''' <summary>Returns a string created by joining a number of substrings contained in an array.</summary>
''' <param name="list">One-dimensional array containing substrings to be joined.</param>
''' <param name="delimiter">String character used to separate the substrings in the returned string.</param>
Function Join(list, delimiter) ' As String
End Function

''' <summary>Returns the smallest available subscript for the indicated dimension of an array.</summary>
''' <param name="arrayname">Name of the array variable.</param>
Function LBound(arrayname)
End Function

''' <summary>Returns the smallest available subscript for the indicated dimension of an array.</summary>
''' <param name="arrayname">Name of the array variable.</param>
''' <param name="dimension">Whole number indicating which dimension's lower bound is returned.</param>
Function LBound(arrayname, dimension)
End Function

''' <summary>Returns a string that has been converted to lowercase.</summary>
''' <param name="str">String expression to be converted.</param>
Function LCase(str) ' As String
End Function

''' <summary>Returns a specified number of characters from the left side of a string.</summary>
''' <param name="str">String expression from which the leftmost characters are returned.</param>
''' <param name="length">Numeric expression indicating how many characters to return.</param>
Function Left(str, length) ' As String
End Function

''' <summary>Returns a specified number of bytes from the left side of a string.</summary>
''' <param name="str">String expression from which the leftmost bytes are returned.</param>
''' <param name="length">Numeric expression indicating how many bytes to return.</param>
Function LeftB(str, length) ' As String
End Function

''' <summary>Returns the number of characters in a string or the number of bytes required to store a variable.</summary>
''' <param name="str">Any valid string expression or variable name.</param>
Function Len(str) ' As Long
End Function

''' <summary>Returns the number of bytes required to store a string.</summary>
''' <param name="str">Any valid string expression.</param>
Function LenB(str) ' As Long
End Function

''' <summary>Returns a picture object.</summary>
''' <param name="picturename">String expression that indicates the name of the picture file to be loaded.</param>
Function LoadPicture(picturename)
End Function

''' <summary>Returns the natural logarithm of a number.</summary>
''' <param name="number">Any valid numeric expression greater than 0.</param>
Function Log(number)
End Function

''' <summary>Returns a string without leading spaces.</summary>
''' <param name="str">String expression from which leading spaces are removed.</param>
Function LTrim(str) ' As String
End Function

''' <summary>Returns a specified number of characters from a string.</summary>
''' <param name="str">String expression from which characters are returned.</param>
''' <param name="start">Character position in string at which the part to be taken begins.</param>
Function Mid(str, start) ' As String
End Function

''' <summary>Returns a specified number of characters from a string.</summary>
''' <param name="str">String expression from which characters are returned.</param>
''' <param name="start">Character position in string at which the part to be taken begins.</param>
''' <param name="length">Number of characters to return.</param>
Function Mid(str, start, length) ' As String
End Function

''' <summary>Returns a specified number of bytes from a string.</summary>
''' <param name="str">String expression from which bytes are returned.</param>
''' <param name="start">Byte position in string at which the part to be taken begins.</param>
Function MidB(str, start) ' As String
End Function

''' <summary>Returns a specified number of bytes from a string.</summary>
''' <param name="str">String expression from which bytes are returned.</param>
''' <param name="start">Byte position in string at which the part to be taken begins.</param>
''' <param name="length">Number of bytes to return.</param>
Function MidB(str, start, length) ' As String
End Function

''' <summary>Returns a whole number between 0 and 59, inclusive, representing the minute of the hour.</summary>
''' <param name="time">Any expression that can represent a time.</param>
Function Minute(time) ' As Integer
End Function

''' <summary>Returns a whole number between 1 and 12, inclusive, representing the month of the year.</summary>
''' <param name="date">Any expression that can represent a date.</param>
Function Month(date) ' As Integer
End Function

''' <summary>Returns a string indicating the specified month.</summary>
''' <param name="date">Numeric designation for a month.</param>
Function MonthName(date) ' As String
End Function

''' <summary>Returns a string indicating the specified month.</summary>
''' <param name="date">Numeric designation for a month.</param>
''' <param name="abbrevation">Boolean value that indicates if the month name is to be abbreviated.</param>
Function MonthName(date, abbrevation) ' As String
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
Function MsgBox(prompt)
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If omitted, the default value for buttons is 0.</param>
Function MsgBox(prompt, buttons)
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If omitted, the default value for buttons is 0.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.</param>
Function MsgBox(prompt, buttons, title)
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If omitted, the default value for buttons is 0.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.</param>
''' <param name="helpfile">String expression that identifies the Help file to use to provide context-sensitive Help for the dialog box. If helpfile is provided, context must also be provided. Not available on 16-bit platforms.</param>
''' <param name="context">Numeric expression that identifies the Help context number assigned by the Help author to the appropriate Help topic. If context is provided, helpfile must also be provided. Not available on 16-bit platforms.</param>
Function MsgBox(prompt, buttons, title, helpfile, context)
End Function

''' <summary>Returns the current date and time according to the setting of your computer's system date and time.</summary>
Function Now ' As Date
End Function

''' <summary>Returns a string representing the octal value of a number.</summary>
''' <param name="number">Any valid numeric expression.</param>
Function Oct(number) ' As String
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring.</summary>
''' <param name="str">String expression containing substring to replace.</param>
''' <param name="find">Substring being searched for.</param>
''' <param name="replacewith">Replacement substring.</param>
Function Replace(str, find, replacewith) ' As String
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring.</summary>
''' <param name="str">String expression containing substring to replace.</param>
''' <param name="find">Substring being searched for.</param>
''' <param name="replacewith">Replacement substring.</param>
''' <param name="start">Position within string where substring search is to begin.</param>
Function Replace(str, find, replacewith, start) ' As String
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring a specified number of times.</summary>
''' <param name="str">String expression containing substring to replace.</param>
''' <param name="find">Substring being searched for.</param>
''' <param name="replacewith">Replacement substring.</param>
''' <param name="start">Position within string where substring search is to begin.</param>
''' <param name="count">Number of substring substitutions to perform.</param>
Function Replace(str, find, replacewith, start, count) ' As String
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring a specified number of times.</summary>
''' <param name="str">String expression containing substring to replace.</param>
''' <param name="find">Substring being searched for.</param>
''' <param name="replacewith">Replacement substring.</param>
''' <param name="start">Position within string where substring search is to begin.</param>
''' <param name="count">Number of substring substitutions to perform.</param>
''' <param name="compare">Numeric value indicating the kind of comparison to use when evaluating substrings.</param>
Function Replace(str, find, replacewith, start, count, compare) ' As String
End Function

''' <summary>Returns a whole number representing an RGB color value.</summary>
''' <param name="red">Number in the range 0-255, inclusive, that represents the red component of the color.</param>
''' <param name="green">Number in the range 0-255, inclusive, that represents the green component of the color.</param>
''' <param name="blue">Number in the range 0-255, inclusive, that represents the blue component of the color.</param>
Function RGB(red, green, blue) ' As Long
End Function

''' <summary>Returns a specified number of characters from the right side of a string.</summary>
''' <param name="str">String expression from which the rightmost characters are returned.</param>
''' <param name="length">Numeric expression indicating how many characters to return.</param>
Function Right(str, length) ' As String
End Function

''' <summary>Returns a specified number of bytes from the right side of a string.</summary>
''' <param name="str">String expression from which the rightmost bytes are returned.</param>
''' <param name="length">Numeric expression indicating how many bytes to return.</param>
Function RightB(str, length) ' As String
End Function

''' <summary>Returns a random number.</summary>
Function Rnd()
End Function

''' <summary>Returns a random number.</summary>
''' <param name="number">Any valid numeric expression.</param>
Function Rnd(number)
End Function

''' <summary>Returns a number rounded to a specified number of decimal places.</summary>
''' <param name="number">Any valid numeric expression.</param>
''' <param name="digits">Number indicating how many places to the right of the decimal are included in the rounding.</param>
Function Round(number, digits)
End Function

''' <summary>Returns a string without trailing spaces.</summary>
''' <param name="str">String expression from which trailing spaces are removed.</param>
Function RTrim(str) ' As String
End Function

''' <summary>Returns a string identifying the script engine in use.</summary>
Function ScriptEngine ' As String
End Function

''' <summary>Returns the build version number of the script engine in use.</summary>
Function ScriptEngineBuildVersion ' As String
End Function

''' <summary>Returns the major version number of the script engine in use.</summary>
Function ScriptEngineMajorVersion ' As String
End Function

''' <summary>Returns the minor version number of the script engine in use.</summary>
Function ScriptEngineMinorVersion ' As String
End Function

''' <summary>Returns a whole number between 0 and 59, inclusive, representing the second of the minute.</summary>
''' <param name="time">Any expression that can represent a time.</param>
Function Second(time)
End Function

''' <summary>Sets the global locale and returns the previous locale.</summary>
''' <param name="int">Locale identifier.</param>
Function SetLocale(int)
End Function

''' <summary>Returns an integer indicating the sign of a number.</summary>
''' <param name="number">Any valid numeric expression.</param>
Function Sgn(number)
End Function

''' <summary>Returns the sine of an angle.</summary>
''' <param name="number">Any valid numeric expression that expresses an angle in radians.</param>
Function Sin(number)
End Function

''' <summary>Returns a string consisting of the specified number of spaces.</summary>
''' <param name="number">Number of spaces you want in the string.</param>
Function Space(number) ' As String
End Function

''' <summary>Returns a zero-based, one-dimensional array containing a specified number of substrings.</summary>
''' <param name="str">String expression containing substrings and delimiters.</param>
Function Split(str)
End Function

''' <summary>Returns a zero-based, one-dimensional array containing a specified number of substrings.</summary>
''' <param name="str">String expression containing substrings and delimiters.</param>
''' <param name="delimiter">String character used to identify substring limits.</param>
Function Split(str, delimiter)
End Function

''' <summary>Returns a zero-based, one-dimensional array containing a specified number of substrings.</summary>
''' <param name="str">String expression containing substrings and delimiters.</param>
''' <param name="delimiter">String character used to identify substring limits.</param>
''' <param name="count">Number of substrings to be returned.</param>
Function Split(str, delimiter, count)
End Function

''' <summary>Returns a zero-based, one-dimensional array containing a specified number of substrings.</summary>
''' <param name="str">String expression containing substrings and delimiters.</param>
''' <param name="delimiter">String character used to identify substring limits.</param>
''' <param name="count">Number of substrings to be returned.</param>
''' <param name="compare">Numeric value indicating the kind of comparison to use when evaluating substrings.</param>
Function Split(str, delimiter, count, compare)
End Function

''' <summary>Returns the square root of a number.</summary>
''' <param name="number">Any valid numeric expression greater than or equal to 0.</param>
Function Sqr(number)
End Function

''' <summary>Returns a value indicating the result of a string comparison.</summary>
''' <param name="string1">Any valid string expression.</param>
''' <param name="string2">Any valid string expression.</param>
Function StrComp(string1, string2)
End Function

''' <summary>Returns a value indicating the result of a string comparison.</summary>
''' <param name="string1">Any valid string expression.</param>
''' <param name="string2">Any valid string expression.</param>
''' <param name="compare">Numeric value indicating the kind of comparison to use when evaluating the strings.</param>
Function StrComp(string1, string2, compare)
End Function

''' <summary>Returns a string with characters in reverse order.</summary>
''' <param name="str">String whose characters are to be reversed.</param>
Function StrReverse(str)
End Function

''' <summary>Returns the tangent of an angle.</summary>
''' <param name="number">Any valid numeric expression that expresses an angle in radians.</param>
Function Tan(number)
End Function

''' <summary>Returns the current system time.</summary>
Function Time
End Function

''' <summary>Returns the number of seconds that have elapsed since 12:00 AM.</summary>
Function Timer
End Function

''' <summary>Returns a Variant of subtype Date for a specific hour, minute, and second.</summary>
''' <param name="hour">Number between 0 and 23, inclusive, or a numeric expression.</param>
''' <param name="minute">Number between 0 and 59, inclusive, or a numeric expression.</param>
''' <param name="second">Number between 0 and 59, inclusive, or a numeric expression.</param>
Function TimeSerial(hour, minute, second)
End Function

''' <summary>Returns a Variant of subtype Date.</summary>
''' <param name="time">String expression that represents a time.</param>
Function TimeValue(time)
End Function

''' <summary>Returns a string without leading and trailing spaces.</summary>
''' <param name="str">String expression from which leading and trailing spaces are removed.</param>
Function Trim(str) ' As String
End Function

''' <summary>Returns a string that provides subtype information about a variable.</summary>
''' <param name="var">Variable name.</param>
Function TypeName(var) ' As String
End Function

''' <summary>Returns the largest available subscript for the indicated dimension of an array.</summary>
''' <param name="arrayname">Name of the array variable.</param>
Function UBound(arrayname) ' As Long
End Function

''' <summary>Returns the largest available subscript for the indicated dimension of an array.</summary>
''' <param name="arrayname">Name of the array variable.</param>
''' <param name="dimension">Whole number indicating which dimension's upper bound is returned.</param>
Function UBound(arrayname, dimension) ' As Long
End Function

''' <summary>Returns a string that has been converted to uppercase.</summary>
''' <param name="str">String expression to be converted.</param>
Function UCase(str) ' As String
End Function

''' <summary>Returns a string with certain escape sequences converted back to their original characters.</summary>
''' <param name="str">String to be unescaped.</param>
Function Unescape(str) ' As String
End Function

''' <summary>Returns a value indicating the subtype of a variable.</summary>
''' <param name="var">Variable name.</param>
Function VarType(var) ' as Integer
End Function

''' <summary>Returns a whole number between 1 and 7, inclusive, representing the day of the week.</summary>
''' <param name="date">Any expression that can represent a date.</param>
Function Weekday(date) ' as Integer
End Function

''' <summary>Returns a whole number between 1 and 7, inclusive, representing the day of the week.</summary>
''' <param name="date">Any expression that can represent a date.</param>
''' <param name="firstdayofweek">Constant that specifies the first day of the week.</param>
Function Weekday(date, firstdayofweek) ' as Integer
End Function

''' <summary>Returns a string indicating the specified day of the week.</summary>
''' <param name="weekday">Numeric designation for a day of the week.</param>
Function WeekdayName(weekday) ' As String
End Function

''' <summary>Returns a string indicating the specified day of the week.</summary>
''' <param name="weekday">Numeric designation for a day of the week.</param>
''' <param name="abbreviate">Boolean value that indicates if the weekday name is to be abbreviated.</param>
Function WeekdayName(weekday, abbreviate) ' As String
End Function

''' <summary>Returns a string indicating the specified day of the week.</summary>
''' <param name="weekday">Numeric designation for a day of the week.</param>
''' <param name="abbreviate">Boolean value that indicates if the weekday name is to be abbreviated.</param>
''' <param name="firstdayofweek">Numeric value indicating the first day of the week.</param>
Function WeekdayName(weekday, abbreviate, firstdayofweek) ' As String
End Function

''' <summary>Returns a whole number representing the year.</summary>
''' <param name="date">Any expression that can represent a date</param>
Function Year(date)
End Function

''' Enum VbVarType
Const vbEmpty = 0 ' Uninitialized (default)
Const vbNull = 1 ' Contains no valid data
Const vbInteger = 2 ' Integer subtype
Const vbLong = 3 ' Long subtype
Const vbSingle = 4 ' Single subtype
Const vbDouble = 5 ' Double subtype
Const vbCurrency = 6 ' Currency subtype
Const vbDate = 7 ' Date subtype
Const vbString = 8 ' String subtype
Const vbObject = 9 ' Object subtype
Const vbError = 10 ' Error subtype
Const vbBoolean = 11 ' Boolean subtype
Const vbVariant = 12 ' Variant subtype (used only for arrays of variants)
Const vbDataObject = 13 ' Data object subtype
Const vbDecimal = 14 ' Decimal subtype
Const vbByte = 17 ' Byte subtype
Const vbArray = 8192 ' Array flag (OR'ed with other type constants)
''' End Enum ' VbVarType

Const Nothing = Nothing ' The Nothing keyword is used to indicate that an object variable does not refer to any object. It is not the same as Null or Empty. You can use the IsObject Function to determine whether a variable refers to an object.
Const Empty = Empty ' The Empty keyword is used to indicate an uninitialized variable value. This is not the same thing as Null. You can use the IsEmpty Function to determine whether a variable is initialized.
Const Null = Null ' The Null keyword is used to indicate that a variable contains no valid data. It is not the same as Empty or Nothing. You can use the IsNull Function to determine whether a variable contains no valid data.

Const False = False ' Boolean. Value is equal to 0.
Const True = True ' Boolean. Value is equal to -1.

''' Enum VbTriState
Const vbUseDefault = -2 ' Use system default setting
Const vbTrue = -1 ' Boolean true value
Const vbFalse = 0 ' Boolean false value
''' End Enum

''' Enum VbCompareMethod
Const vbBinaryCompare = 0 ' Perform a binary comparison
Const vbTextCompare = 1 ' Perform a textual comparison
Const vbDatabaseCompare = 2 ' Only in Access
''' End Enum

''' Enum StringConstants
Const vbBack = Chr(8) ' Backspace character
Const vbCr = Chr(13) ' Carriage return character
Const vbCrLf = Chr(13) & Chr(10) ' Carriage return and line feed characters
Const vbFormFeed = Chr(12) ' Form feed character
Const vbLf = Chr(10) ' Line feed character
Const vbNewLine = Chr(13) & Chr(10) ' Carriage return and line feed characters
Const vbNullChar = Chr(0) ' Null character
Const vbNullString = Empty ' A zero-length string, which is not the same as Null or Empty.
Const vbTab = Chr(9) ' Tab character
Const vbVerticalTab = Chr(11) ' Vertical tab character
''' End Enum

''' Enum VbDateTimeFormat
Const vbGeneralDate = 0 ' Display a date using the long date format specified in your computer's regional settings, and a time using the time format specified in your computer's regional settings.
Const vbLongDate = 1 ' Display a date using the long date format specified in your computer's regional settings.
Const vbShortDate = 2 ' Display a date using the short date format specified in your computer's regional settings.
Const vbLongTime = 3 ' Display a time using the time format specified in your computer's regional settings.
Const vbShortTime = 4 ' Display a time using the short time format specified in your computer's regional settings.
''' End Enum

''' Enum VbFirstWeekOfYear
Const vbUseSystemDayOfWeek = 0 ' Use the system setting for the first day of the week.
Const vbFirstJan1 = 1 ' The first week of the year is the week that contains January 1.
Const vbFirstFourDays = 2 ' The first week of the year is the week that contains at least four days in the new year.
Const vbFirstFullWeek = 3 ' The first week of the year is the first full week of the year, starting on the first day of the week.
''' End Enum

Const vbObjectError = &h80040000 ' User-defined error numbers should be greater than this value.

''' Enum VbDayOfWeek
Const vbMonday = 2
Const vbTuesday = 3
Const vbWednesday = 4
Const vbThursday = 5
Const vbFriday = 6
Const vbSaturday = 7
Const vbSunday = 1
Const vbUseSystem = 0 ' Use the system setting for the first day of the week.
''' End Enum

''' Enum VbMsgBoxStyle
Const vbOKOnly = 0 ' Display OK button only.
Const vbOKCancel = 1 ' Display OK and Cancel buttons
Const vbAbortRetryIgnore = 2 ' Display Abort, Retry, and Ignore buttons
Const vbYesNoCancel = 3 ' Display Yes, No, and Cancel buttons
Const vbYesNo = 4 ' Display Yes and No buttons
Const vbRetryCancel = 5 ' Display Retry and Cancel buttons.
Const vbCritical = 16 ' Display Critical Message icon
Const vbQuestion = 32 ' Display Warning Query icon.
Const vbExclamation = 48 ' Display Warning Message icon
Const vbInformation = 64 ' Display Information Message icon
Const vbDefaultButton1 = 0 ' First button is default
Const vbDefaultButton2 = 256 ' Second button is default
Const vbDefaultButton3 = 512 ' Third button is default
Const vbDefaultButton4 = 768 ' Fourth button is default
Const vbApplicationModal = 0 ' Application modal; the user must respond to the message box before continuing work in the current application
Const vbSystemModal         = &h01000 ' System modal; all applications are suspended until the user responds to the message box
Const vbMsgBoxHelpButton    = &h04000 ' Adds Help button to the message box
Const VbMsgBoxSetForeground = &h010000 ' Specifies the message box window as the foreground window
Const vbMsgBoxRight         = &h080000 ' Text is right aligned
Const vbMsgBoxRtlReading    = &h100000 ' Specifies text should appear as right-to-left reading on Hebrew and Arabic systems
''' End Enum

''' Enum VbMsgBoxResult
Const vbOK = 1 ' OK button was clicked
Const vbCancel = 2 ' Cancel button was clicked
Const vbAbort = 3 ' Abort button was clicked
Const vbRetry = 4 ' Retry button was clicked
Const vbIgnore = 5 ' Ignore button was clicked
Const vbYes = 6 ' Yes button was clicked
Const vbNo = 7 ' No button was clicked
''' End Enum

''' Enum ColorConstants
Const vbBlack   = &h000000
Const vbBlue    = &hFF0000
Const vbCyan    = &hFFFF00
Const vbGreen   = &h00FF00
Const vbMagenta = &hFF00FF
Const vbRed     = &h0000FF
Const vbWhite   = &hFFFFFF
Const vbYellow  = &h00FFFF
''' End Enum ' ColorConstants


' Const SystemFolder = 1
' Const TemporaryFolder = 2
' Const WindowsFolder = 0