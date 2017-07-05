<%
''''''''''''''''''''''''''''''''''
'' MSSQLSanitizer
''      Provides functions to sanitize and validate data before insertion into a MSSQL database.
''
'' [How To Use]
''      • To sanitize a value to its corresponding MSSQL value use the appropriate exposed function.
''      • To determine if any data resolution was lost check the public property 'bolLastSanitizationLossless'.
''      • If data resolution was lost check the public property 'strLastSanitizationMessage' for details.
''      • To allow/disallow NULL values to pass-through toggle the public property 'bolAllowNull'.
''
'' [Notes]
''      • Class does not support milliseconds - this resolution will be lost and not flagged
''
'' [Exposed Properties]
''      bolAllowNull                                            -- True/False value as to whether NULL values can pass-through
''      bolLastSanitizationLossless                             -- True/False value as to whether last sanitization was lossless
''      intRadixCount                                           -- Integer containing radix value
''      strClassName (default)                                  -- String identifier containing class name
''      strLastSanitizationMessage                              -- String containing brief description of last sanitization result
''
'' [Exposed Functions]
''      sanitizeBigInt(intValue)                                -- Sanitize a number into the specified TSQL 'bigint' data-type
''      sanitizeBit(vntValue)                                   -- Sanitize a variant into a bit
''      sanitizeChar(strValue, intLength)                       -- Sanitize a string to valid TSQL "char" or "varchar" data-type
''      sanitizeDate(strDate)                                   -- Sanitize a string into valid TSQL "date" data-type
''      sanitizeDateTime(strDate, strTime)                      -- Sanitize a string into valid TSQL "datetime" data-type
''      sanitizeDateTime2(strDate, strTime)                     -- Sanitize a string into valid TSQL "datetime2" data-type
''      sanitizeFloat(dblValue, intPrecision)                   -- Sanitize a floating point number into TSQL "float" data-type
''      sanitizeInt(intNumber)                                  -- Sanitize a number into the specified TSQL integer data-type
''      sanitizeIntegerRange(intValue, intMinValue, intMaxValue)
''                                                              -- Sanitize an integer to ensure it is in a specified range
''      sanitizeMediumInt(intValue)                             -- Sanitize a number into the specified TSQL 'MediumInt' data-type
''      sanitizeMoney(curMoney)                                 -- Sanitize a currency number into TSQL "money" data-type
''      sanitizeNumericDepth(dblValue, intMaxIntegerDigits, intMaxFractionalDigits)
''                                                              -- Sanitize a number's integer and fractional depths
''      Added sanitizeNumericRange(dblValue, dblMinValue, dblMaxValue)
''                                                              -- Sanitize a number to ensure it is in a specified range
''      sanitizePrecision(dblValue, dblAsymptote, dblLimit)     -- Sanitize a number into specified precision
''      sanitizeReal(dblValue)                                  -- Sanitize a floating point number into TSQL "real" data-type
''      sanitizeSmallDateTime(strDate, strTime)                 -- Sanitize a date and time into TSQL "smalldatetime" data-type
''      sanitizeSmallMoney(curSmallMoney)                       -- Sanitize a currency number into TSQL "smallmoney" data-type
''      sanitizeTime(strTime)                                   -- Sanitize a time value into TSQL "time" data-type
''      sanitizeTinyInt(intValue)
''
'' [ChangeLog]
''
''      1.5             Dudley      Added sanitizeIntegerRange()
''                                  Rewrote sanitizeBigInt(), sanitizeInt(), sanitizeMediumInt(), and sanitizeTinyInt()
''                                      to use sanitizeIntegerRange()
''      1.4             Dudley      Added sanitizeBigInt()
''                                      Added sanitizeMediumInt()
''                                      Added sanitizeTinyInt()
''                                      Finished bolAllowNull property results and updated UnitTests to check for lossless flag value
''                                      Fixed type-casting bugs in sanitizeChar, sanitizeFloat, sanitizeMoney, & sanitizeSmallMoney
''                                      Multiple re-factors to get things consistent
''                                      Rewrote sanitizeInt()
''                                      Rewrote sanitizeNumericRange()
''                                      Added missing UnitTests
''      1.3             Dudley      Added and partially integrated bolAllowNull property
''                                      Added sanitizeBit()
''                                      Added sanitizeNumericDepth(dblValue, intMaxIntegerDigits, intMaxFractionalDigits)
''                                      Updated Unit Tests
''      1.2             Dudley      Re-factored from collection of functions to a class
''                                      Added result tracking
''      1.1             Dudley      Added rounding to sanitizeInt()
''                                      Added sanitizeFloat()
''                                      Added sanitizePrecision()
''                                      Added sanitizeReal()
''                                      Wrote UnitTests - All Units Passing
''      1.0             Dudley      Added sanitizeChar()
''                                      Added sanitizeDate()
''                                      Added sanitizeDateTime()
''                                      Added sanitizeDateTime2()
''                                      Added sanitizeInt()
''                                      Added sanitizeMoney()
''                                      Added sanitizeNumericRange()
''                                      Added sanitizeSmallDateTime()
''                                      Added sanitizeSmallMoney()
''                                      Added sanitizeTime()
''''''''''''''''''''''''''''''''''

Class MSSQLSanitizer
    
    Private m_bolAllowNull
    Private m_bolLastSanitizationLossless
    Private m_intRadixCount
    Private m_strLastSanitizationMessage
    
    Public Property Get bolAllowNull
        bolAllowNull = m_bolAllowNull
    End Property
    
    Public Property Let bolAllowNull(bolValue)
        If sanitizeBit(bolValue) = 1 Then
            m_bolAllowNull = True
        Else
            m_bolAllowNull = False
        End If
    End Property
    
    Public Property Get bolLastSanitizationLossless()
        bolLastSanitizationLossless = m_bolLastSanitizationLossless
    End Property
    
    Private Property Let bolLastSanitizationLossless(bolValue)
        If isBoolean(bolValue) Then
            m_bolLastSanitizationLossless = bolValue
        ElseIf sanitizeBit(bolValue) = 1 Then
            m_bolLastSanitizationLossless = True
        Else
            m_bolLastSanitizationLossless = False
        End If
    End Property
    
    Public Property Get intRadixCount
        intRadixCount = m_intRadixCount
    End Property
    
    Public Default Property Get strClassName
        strClassName = "MSSQLSanitizer"
    End Property
    
    Public Property Get strLastSanitizationMessage()
        strLastSanitizationMessage = CStr(m_strLastSanitizationMessage)
    End Property
    
    Private Property Let strLastSanitizationMessage(strValue)
        m_strLastSanitizationMessage = Trim(m_strLastSanitizationMessage & " " & Left(Trim(CStr(strValue)), 384))
    End Property
    
    
    '' Subroutine: Class_Initialize
    ''      Class constructor
    ''''''''''''''''''''''''''''''''''
        Private Sub Class_Initialize()
            m_bolAllowNull = False
            m_intRadixCount = 10
            clearLastResult()
        End Sub
    
    
    '' Subroutine: clearLastResult
    ''      Clears the result information of the last sanitization (m_bolLastSanitizationLossless, m_strLastSanitizationMessage)
    ''''''''''''''''''''''''''''''''''
        Private Sub clearLastResult()
            m_bolLastSanitizationLossless = True
            m_strLastSanitizationMessage = ""
        End Sub
        
        
    '' Function: sanitizeBigInt
    ''      Sanitize a number into the specified TSQL 'bigint' data-type
    ''
    '' Params:
    ''      intValue - value to be sanitized
    ''
    '' Return: Integer
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeBigInt(ByVal intValue)
            sanitizeBigInt = sanitizeIntegerRange(intValue, -9223372036854755808, 9223372036854775807)
        End Function
    
        
    '' Function: sanitizeBit
    ''      Sanitizes a variant into a bit.  Positive values, objects, non-empty arrays, and non-empty strings become "1".
    ''      All other values become "0".
    ''
    '' Params:
    ''      vntValue - Value to be sanitized
    ''
    '' Return: Bit
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeBit(byVal vntValue)
            clearLastResult()
            
            If me.bolAllowNull And IsNull(vntValue) Then
                sanitizeBit = Null
            ElseIf (isBoolean(vntValue)) Then
                If vntValue = True Then
                    sanitizeBit = 1
                Else
                    sanitizeBit = 0
                End If
            ElseIf isNumeric(vntValue) Then
                If vntValue = 0 Then
                    sanitizeBit = 0
                Elseif vntValue = 1 Then
                    sanitizeBit = 1
                Elseif vntValue < 0 Then
                    strLastSanitizationMessage = "Received invalid negative value for bit type - assuming '0'."
                    sanitizeBit = 0
                Else
                    strLastSanitizationMessage = "Received invalid positive value for bit type - assuming '1'."
                    sanitizeBit = 1
                End If
            ElseIf isString(vntValue) Then
                If LEN(vntValue) = 0 Then
                    sanitizeBit = 0
                    strLastSanitizationMessage = "Received empty string value for bit - assuming '0'."
                Else
                    sanitizeBit = 1
                    strLastSanitizationMessage = "Received non-empty string value for bit - assuming '1'."
                End If
            Else
                sanitizeBit = 0
                
                If isObject(vntValue) Then
                    strLastSanitizationMessage = "Received unknown object for bit - assuming '0'."
                ElseIf isEmpty(vntValue) Then
                    strLastSanitizationMessage = "Received uninitiated value for bit - assuming '0'."
                ElseIf isNull(vntValue) Then
                    strLastSanitizationMessage = "Received NULL value for bit - assuming '0'."
                Else
                    strLastSanitizationMessage = "Received unknown data-type for bit - assuming '0'."
                End If
            End If
            
            If Len(strLastSanitizationMessage) > 0 Then
                bolLastSanitizationLossless = False
            End If
        End Function
        
    
    '' Function: sanitizeChar
    ''      Sanitizes a string to valid TSQL "char" or "varchar" data-type
    ''
    '' Params:
    ''      strValue - String containing value to be sanitized
    ''      intLength - Maximum length of string to be returned
    ''
    '' Return: String
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeChar(ByVal strValue, intLength)
            
            intLength = sanitizeIntegerRange(intLength, 0, 8000)
            clearLastResult()
            
            If isNull(strValue) Then
                If me.bolAllowNull Then
                    sanitizeChar = Null
                Else
                    sanitizeChar = ""
                    strLastSanitizationMessage = "Received NULL value for string - assuming ''."
                End If
            ElseIf isBoolean(strValue) Then
                If strValue Then
                    sanitizeChar = Left("Yes", intLength)
                    strLastSanitizationMessage = "Received TRUE Boolean assuming 'Yes'."
                Else
                    sanitizeChar = Left("No", intLength)
                    strLastSanitizationMessage = "Received False Boolean assuming 'No'."
                End If
            ElseIf isEmpty(strValue) OR isObject(strValue) Then
                    sanitizeChar = ""
                    strLastSanitizationMessage = "Received unknown data-type for string value - assuming ''."
            Else
                strValue = Server.HTMLEncode(Trim(strValue))
                
                If Len(strValue) > intLength Then
                    strLastSanitizationMessage = "String value was too long and was truncated by " & CStr(LEN(strValue) - intLength) & " characters. "
                End If
                
                sanitizeChar = Left(strValue, intLength)
            End If
            
            If Len(strLastSanitizationMessage) > 0 Then
                bolLastSanitizationLossless = False
            End If
        End Function
    
    
    '' Function: sanitizeDate
    ''      Sanitize a string into valid TSQL "date" data-type
    ''
    '' Params:
    ''      strDate - String containing date to be sanitized
    ''
    '' Return: String containing the date in IS08601 format (YYYY-MM-DD)
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeDate(strDate)
            clearLastResult()
        
            If me.bolAllowNull And IsNull(strDate) Then
                sanitizeDate = Null
                'sanitizeDate = "This will cause the function to fail on NULL"
            ElseIf IsDate(strDate) Then
                Dim dtmDate     : dtmDate = CDate(strDate)
                Dim strYear     : strYear = CStr(Right("0000" & CStr(DatePart("yyyy", dtmDate)), 4))
                Dim strMonth    : strMonth = CStr(Right("00" & CStr(DatePart("m", dtmDate)), 2))
                Dim strDay      : strDay = CStr(Right("00" & CStr(DatePart("d", dtmDate)), 2))
                sanitizeDate = CStr(strYear & "-" & strMonth & "-" & strDay)
            Else
                bolLastSanitizationLossless = False
                strLastSanitizationMessage = "Received invalid date - assuming '0000-00-00'. "
                sanitizeDate = "0000-00-00"
            End If
            'sanitizeDate = "this will cause the function to fail miserably"
            'error
        End Function
        
        
    '' Function: sanitizeDateTime
    ''      Sanitize a string into valid TSQL "datetime" data-type
    ''
    '' Params:
    ''      strDate - String containing date to be sanitized
    ''      strTime - String containing time to be sanitized
    ''
    '' Return: String of the date in the format "YYYY-MM-DDThh:mm:ss.mmm"
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeDateTime(ByVal strDate, ByVal strTime)
            DIM strResultMessage    : strResultMessage = ""
            
            strDate = sanitizeDate(strDate)
            strResultMessage = strResultMessage & strLastSanitizationMessage
            
            strTime = sanitizeTime(strTime)
            If LEN(strResultMessage) > 0 Then
                strResultMessage = strResultMessage & "  " & strLastSanitizationMessage
            End If
            
            If isNull(strDate) And isNull(strTime) Then
                sanitizeDateTime = Null
            ElseIf isNull(strDate) Then
                sanitizeDateTime = CStr(sanitizeDate("giveMeADate") & "T" & strTime)
                strResultMessage = strResultMessage & " Cannot combine NULL date with non-NULL time.  Assuming '0000-00-00'. "
            ElseIf isNull(strTime) Then
                sanitizeDateTime = CStr(strDate & "T" & sanitizeTime("giveMeATime"))
                strResultMessage = strResultMessage & " Cannot combine NULL time with non-NULL date.  Assuming '00:00:00:000'. "
            Else
                sanitizeDateTime = CStr(strDate & "T" & strTime)
            End If
            
            If LEN(strResultMessage) > 0 Then
                m_bolLastSanitizationLossless = False
                m_strLastSanitizationMessage = Trim(strResultMessage)
            Else
                clearLastResult()
            End If
            
        End Function
        
        
    '' Function: sanitizeDateTime2
    ''      Sanitize a string into valid TSQL "datetime2" data-type
    ''
    '' Params:
    ''      strDate - String containing date to be sanitized
    ''      strTime - String containing time to be sanitized
    ''
    '' Return: String of the date in the format "YYYY-MM-DDThh:mm:ss.nnnnnnn"
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeDateTime2(strDate, strTime)
            Dim strReturnValue      : strReturnValue = sanitizeDateTime(strDate, strTime)
            
            If IsNull(strReturnValue) Then
                sanitizeDateTime2 = Null
            Else
                sanitizeDateTime2 = CStr(strReturnValue) & "0000"
            End If
        End Function
        
        
    '' Function: sanitizeFloat
    ''      Sanitize a floating point number into TSQL "float" data-type
    ''
    '' Params:
    ''      dblValue - value to be sanitized
    ''      intPrecision - Determines number of bits that are used to store the mantissa
    ''
    '' Return: Floating Point Number
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeFloat(dblValue, intPrecision)
            If intPrecision < 25 Then
                sanitizeFloat = sanitizePrecision(dblValue, 1.18E-38, 3.40E+38)
            Else
                sanitizeFloat = sanitizePrecision(dblValue, 2.23E-308, 1.79E+308)
            End If
        End Function
        
        
    '' Function: sanitizeInt
    ''      Sanitize a number into the specified TSQL integer data-type
    ''
    '' Params:
    ''      intNumber - value to be sanitized
    ''
    '' Return: Integer
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeInt(ByVal intValue)
            sanitizeInt = sanitizeIntegerRange(intValue, -2147483648, 2147483647)
        End Function
        
        
    '' Function: sanitizeIntegerRange
    ''      Sanitize an integer to ensure it is in a specified range
    ''
    '' Params:
    ''      intValue - value to be sanitized
    ''      intMinValue - minimum value allowed
    ''      intMaxValue - maximum number allowed
    ''
    '' Return: Integer
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeIntegerRange(byVal intValue, byVal intMinValue, byVal intMaxValue)
            Dim intWholeValue       : intWholeValue = 0
            Dim strMessageQueue     : strMessageQueue = ""
            
            clearLastResult()
            
            If intMinValue <> Fix(intMinValue) Then
                intMinValue = Fix(intMinValue)
                strMessageQueue = Trim(strMessageQueue & "  Coded minimum value was not an integer and was rounded!")
            End If
            
            If intMaxValue <> Fix(intMaxValue) Then
                intMaxValue = Fix(intMaxValue)
                strMessageQueue = Trim(strMessageQueue & "  Coded maximum value was not an integer and was rounded!")
            End If
            
            If IsNumeric(intValue) Then
                intWholeValue = Round(intValue)
                If intValue <> intWholeValue Then
                    strMessageQueue = Trim(strMessageQueue & "  Provided value was not an integer and was rounded.")
                End If
                intValue = intWholeValue
            End If
            
            sanitizeIntegerRange = sanitizeNumericRange(intValue, intMinValue, intMaxValue)
            
            If Len(strMessageQueue) > 0 Then
                bolLastSanitizationLossless = False
                strLastSanitizationMessage = strMessageQueue
                strMessageQueue = ""
            End If
        End Function
        
        
    '' Function: sanitizeMediumInt
    ''      Sanitize a number into the specified TSQL 'MediumInt' data-type
    ''
    '' Params:
    ''      intValue - value to be sanitized
    ''
    '' Return: Integer
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeMediumInt(ByVal intValue)
            sanitizeMediumInt = sanitizeIntegerRange(intValue, -8388608, 8388607)
        End Function
        
        
    '' Function: sanitizeMoney
    ''      Sanitize a currency number into TSQL "money" data-type
    ''
    '' Params:
    ''      curMoney - value to be sanitized
    ''
    '' Return: Currency value
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeMoney(byVal curMoney)
            clearLastResult()
        
            If IsNull(curMoney) Then
                If me.bolAllowNull Then
                    sanitizeMoney = Null
                Else
                    strLastSanitizationMessage = "NULL value allowed for field - assuming '0.00'."
                    sanitizeMoney = 0.00
                End If
            Else
                If IsString(curMoney) Then
                    curMoney = Replace(curMoney, "$", "")
                    curMoney = Replace(curMoney, "¢", "")
                    curMoney = Replace(curMoney, "£", "")
                    curMoney = Replace(curMoney, "¤", "")
                    curMoney = Replace(curMoney, "¥", "")
                End If
                
                If IsNumeric(curMoney) Then
                    sanitizeMoney = CCur(sanitizeNumericRange(CDbl(curMoney), -9.22337203685477E14, 9.22337203685477E14))
                Else
                    strLastSanitizationMessage = "Received unknown value - assuming '0.00'."
                    sanitizeMoney = 0.00
                End If
            End If
            
            If Len(strLastSanitizationMessage) > 0 Then
                bolLastSanitizationLossless = False
            End If
        End Function
        
        
    '' Function: sanitizeNumericDepth
    ''      Sanitize a number to ensure that the number of digits to the left and right of the decimal point
    ''      do not exceed specified depths
    ''
    '' Params:
    ''      dblValue - value to be sanitized
    ''      intMaxIntegerDigits - maximum acceptable number of integer digits allowed
    ''      intMaxFractionalDigits - maximum acceptable number of fractional digits
    ''
    '' Return: Numeric
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeNumericDepth(byVal dblValue, byVal intMaxIntegerDigits, byVal intMaxFractionalDigits)
            clearLastResult()
        
            If IsNull(dblValue) Then
                If me.bolAllowNull Then
                    sanitizeNumericDepth = Null
                Else
                    sanitizeNumericDepth = 0
                    strLastSanitizationMessage = "Received NULL value which is not allowed - assuming '0'."
                End If
            ElseIf IsNumeric(dblValue) Then
                Dim intIntegerMax       : intIntegerMax = 0
                Dim intSign             : intSign = 1
                Dim intReturnValue      : intReturnValue = 0
                
                If Sgn(dblValue) = -1 Then
                    intSign = -1
                    dblValue = Abs(dblValue)
                End If
                
                intMaxIntegerDigits = sanitizeNumericRange(intMaxIntegerDigits, 0, 307)
                intMaxFractionalDigits = sanitizeNumericRange(intMaxFractionalDigits, 0, 323)
                
                clearLastResult()
                
                intReturnValue = Round(dblValue, intMaxFractionalDigits)
                If (intReturnValue <> dblValue) Then
                    strLastSanitizationMessage = "Value had too many fractional digits and was rounded."
                End If
                
                If intMaxIntegerDigits > 0 Then
                    intIntegerMax = (10^intMaxIntegerDigits) - 1
                End If
                
                If dblValue >= (intIntegerMax + 1) Then
                    intReturnValue = intIntegerMax + (intReturnValue - Fix(intReturnValue))
                    strLastSanitizationMessage = "Value had too many integer digits and was reduced by " & CStr(dblValue - intReturnValue) & "."
                End If
                
                intReturnValue = Round(intReturnValue, intMaxFractionalDigits)      'Good ol' vbscript
                
                sanitizeNumericDepth = (intSign * intReturnValue)
            Else
                sanitizeNumericDepth = 0
                strLastSanitizationMessage = "Received non-numeric value - assuming '0'."
            End If
            
            If Len(strLastSanitizationMessage) > 0 Then
                bolLastSanitizationLossless = False
            End If
        End Function
        
        
    '' Function: sanitizeNumericRange
    ''      Sanitize a number to ensure it is in a specified range
    ''
    '' Params:
    ''      dblValue - value to be sanitized
    ''      dblMinValue - minimum value allowed
    ''      dblMaxValue - maximum number allowed
    ''
    '' Return: Numeric
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeNumericRange(byVal dblValue, byVal dblMinValue, byVal dblMaxValue)
            Dim strMessageQueue     : strMessageQueue = ""
            
            clearLastResult()
        
            If me.bolAllowNull And IsNull(dblValue) Then
                sanitizeNumericRange = Null
            ElseIf IsNumeric(dblValue) Then
                If dblMaxValue < dblMinValue Then
                    dblMaxValue = dblMinValue
                    strMessageQueue = Trim(strMessageQueue & "  Coded minValue is larger than maxValue - assuming '" & dblMaxValue & "'.")
                End If
                
                If dblValue < dblMinValue Then
                    sanitizeNumericRange = dblMinValue
                    strMessageQueue = Trim(strMessageQueue & "  Provided value was too low and was set to '" & dblMinValue & "'.")
                ElseIf dblValue > dblMaxValue Then
                    sanitizeNumericRange = dblMaxValue
                    strMessageQueue = Trim(strMessageQueue & "  Provided value was too high and was set to '" & dblMaxValue & "'.")
                Else
                    sanitizeNumericRange = dblValue
                End If
            ElseIf IsBoolean(dblValue) Then
                If dblValue Then
                    sanitizeNumericRange = sanitizeNumericRange(1, dblMinValue, dblMaxValue)
                    clearLastResult()
                    strMessageQueue = Trim(strMessageQueue & "  Provided boolean value was 'TRUE' - assuming '" & sanitizeNumericRange & "'.")
                Else
                    sanitizeNumericRange = sanitizeNumericRange(0, dblMinValue, dblMaxValue)
                    clearLastResult()
                    strMessageQueue = Trim(strMessageQueue & "  Provided boolean value was 'FALSE' - assuming '" & sanitizeNumericRange & "'.")
                End If
            Else
                sanitizeNumericRange = sanitizeNumericRange(0, dblMinValue, dblMaxValue)
                strMessageQueue = Trim(strMessageQueue & "  Provided value was not numeric - assuming '" & sanitizeNumericRange & "'.")
            End If
            
            If Len(strMessageQueue) > 0 Then
                bolLastSanitizationLossless = False
                strLastSanitizationMessage = strMessageQueue
                strMessageQueue = ""
            End If
        End Function
        
        
    '' Function: sanitizePrecision
    ''      Sanitize a number into specified precision
    ''
    '' Params:
    ''      dblValue - value to be sanitized
    ''      dblAsymptote - Most precise asymptote allowed (approaching zero)
    ''      dblLimit - Most precise number allowed (approaching infinity)
    ''
    '' Return: Floating Point Number
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizePrecision(dblValue, dblAsymptote, dblLimit)
            clearLastResult()
            
            If bolAllowNull And (IsNull(dblValue) Or dblValue = "") Then
                sanitizePrecision = Null
            Else
                Dim intSign     : intSign = 1
                
                If Sgn(dblValue) = -1 Then
                    intSign = -1
                End If
                
                dblValue = CDbl(ABS(dblValue))
                dblAsymptote = CDbl(ABS(dblAsymptote))
                dblLimit = CDbl(ABS(dblLimit))
            
                If dblValue > dblLimit Then
                    dblValue = dblLimit
                    m_bolLastSanitizationLossless = False
                    m_strLastSanitizationMessage = "Value was too large to be stored in given memory space."
                ElseIf dblValue < dblAsymptote Then
                    dblValue = dblAsymptote
                    m_bolLastSanitizationLossless = False
                    m_strLastSanitizationMessage = "Value fractional was too large to be stored in given memory space."
                End If
            
                sanitizePrecision = (dblValue * intSign)
            End If
        End Function
        
        
    '' Function: sanitizeReal
    ''      Sanitize a floating point number into TSQL "real" data-type
    ''
    '' Params:
    ''      dblValue - value to be sanitized
    ''
    '' Return: Floating Point Number
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeReal(dblValue)
            sanitizeReal = sanitizePrecision(dblValue, 1.18E-38, 3.40E+38)
        End Function
        
        
    '' Function: sanitizeSmallDateTime
    ''      Sanitize a date and time into TSQL "smalldatetime" data-type
    ''
    '' Params:
    ''      curMoney - value to be sanitized
    ''
    '' Return: String containing date and time in the format of "YYYY-MM-DD HH:MM:00"
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeSmallDateTime(strDate, strTime)
            sanitizeSmallDateTime = sanitizeDate(strDate) & " " & LEFT(sanitizeTime(strTime), 6) & "00"
        End Function
        
        
    '' Function: sanitizeSmallInt
    ''      Sanitize a number into the specified TSQL 'bigint' data-type
    ''
    '' Params:
    ''      intValue - value to be sanitized
    ''
    '' Return: Integer
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeSmallInt(ByVal intValue)
            sanitizeSmallInt = sanitizeIntegerRange(intValue, -32768, 32767)
        End Function
        
        
    '' Function: sanitizeSmallMoney
    ''      Sanitize a currency number into TSQL "smallmoney" data-type
    ''
    '' Params:
    ''      curSmallMoney - value to be sanitized
    ''
    '' Return: Currency
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeSmallMoney(curSmallMoney)
            clearLastResult()
        
            If IsNull(curSmallMoney) Then
                If me.bolAllowNull Then
                    sanitizeSmallMoney = Null
                Else
                    strLastSanitizationMessage = "NULL value allowed for field - assuming '0.00'."
                    sanitizeSmallMoney = 0.00
                End If
            Else
                If IsString(curSmallMoney) Then
                    curSmallMoney = Replace(curSmallMoney, "$", "")
                    curSmallMoney = Replace(curSmallMoney, "¢", "")
                    curSmallMoney = Replace(curSmallMoney, "£", "")
                    curSmallMoney = Replace(curSmallMoney, "¤", "")
                    curSmallMoney = Replace(curSmallMoney, "¥", "")
                End If
                
                If IsNumeric(curSmallMoney) Then
                    sanitizeSmallMoney = CCur(sanitizeNumericRange(CDbl(curSmallMoney), -214748.3648, 214748.3647))
                Else
                    strLastSanitizationMessage = "Received unknown value - assuming '0.00'."
                    sanitizeSmallMoney = 0.00
                End If
            End If
            
            If Len(strLastSanitizationMessage) > 0 Then
                bolLastSanitizationLossless = False
            End If
        End Function

        
    '' Function: sanitizeTime
    ''      Sanitize a time value into TSQL "time" data-type
    ''
    '' Params:
    ''      strTime - time value to be sanitized
    ''
    '' Return: String in the format of "hh:mm:ss.mmm"
    ''''''''''''''''''''''''''''''''''
        Public Function sanitizeTime(ByVal strTime)
            clearLastResult()
            
            If me.bolAllowNull And IsNull(strTime) Then
                sanitizeTime = Null
            ElseIf IsDate(strTime) Then 
                Dim dtmTime         : dtmTime = CDate(strTime)              
                Dim strHour         : strHour = Right("00" & CStr(DatePart("h", dtmTime)), 2)
                Dim strMinute       : strMinute = Right("00" & CStr(DatePart("n", dtmTime)), 2)
                Dim strSecond       : strSecond = Right("00" & CStr(DatePart("s", dtmTime)), 2)
                Dim strMilliSecond  : strMilliSecond = "000"
                
                sanitizeTime = CStr(strHour & ":" & strMinute & ":" & strSecond & "." & strMilliSecond)
            Else
                sanitizeTime = "00:00:00.000"
                m_bolLastSanitizationLossless = False
                m_strLastSanitizationMessage = "Received invalid time format.  Defaulted to '00:00:00.000'."
            End If
        End Function
        
        
    '' Function: sanitizeTinyInt
    ''      Sanitize a number into the specified TSQL 'tinyint' data-type
    ''
    '' Params:
    ''      intValue - value to be sanitized
    ''
    '' Return: Integer
    ''''''''''''''''''''''''''''''''''
    Public Function sanitizeTinyInt(ByVal intValue)
        sanitizeTinyInt = sanitizeIntegerRange(intValue, 0, 255)
    End Function

End Class
%>
