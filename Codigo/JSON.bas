Attribute VB_Name = "JSON"
Option Explicit

Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()

Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_SGROUPING = &H10
Private Const A_CURLY_BRACKET_OPEN   As Integer = 123
Private Const A_CURLY_BRACKET_CLOSE  As Integer = 125
Private Const A_SQUARE_BRACKET_OPEN  As Integer = 91
Private Const A_SQUARE_BRACKET_CLOSE As Integer = 93
Private Const A_BRACKET_OPEN         As Integer = 40
Private Const A_BRACKET_CLOSE        As Integer = 41
Private Const A_COMMA                As Integer = 44
Private Const A_DOUBLE_QUOTE         As Integer = 34
Private Const A_SINGLE_QUOTE         As Integer = 39
Private Const A_BACKSLASH            As Integer = 92
Private Const A_FORWARDSLASH         As Integer = 47
Private Const A_COLON                As Integer = 58
Private Const A_SPACE                As Integer = 32
Private Const A_ASTERIX              As Integer = 42
Private Const A_VBCR                 As Integer = 13
Private Const A_VBLF                 As Integer = 10
Private Const A_VBTAB                As Integer = 9
Private Const A_VBCRLF               As Integer = 13
Private Const A_b                    As Integer = 98
Private Const A_f                    As Integer = 102
Private Const A_n                    As Integer = 110
Private Const A_r                    As Integer = 114
Private Const A_t                    As Integer = 116
Private Const A_u                    As Integer = 117
Private m_decSep                     As String
Private m_groupSep                   As String
Private m_parserrors                 As String
Private m_str()                      As Integer
Private m_length                     As Long

Public Function GetParserErrors() As String
    GetParserErrors = m_parserrors
End Function

Public Function parse(ByRef Str As String) As Object
    m_decSep = GetRegionalSettings(LOCALE_SDECIMAL)
    m_groupSep = GetRegionalSettings(LOCALE_SGROUPING)
    Dim index As Long
        index = 1
    Call GenerateStringArray(Str)
    m_parserrors = vbNullString
    On Error Resume Next
    Call skipChar(index)
    Select Case m_str(index)
        Case A_SQUARE_BRACKET_OPEN
            Set parse = parseArray(Str, index)

        Case A_CURLY_BRACKET_OPEN
            Set parse = parseObject(Str, index)

        Case Else
            m_parserrors = "JSON Invalido"
    End Select
    ReDim m_str(1)
End Function

Private Sub GenerateStringArray(ByRef Str As String)
    Dim i As Long
    m_length = Len(Str)
    ReDim m_str(1 To m_length)
    For i = 1 To m_length
        m_str(i) = AscW(mid$(Str, i, 1))
    Next i
End Sub

Private Function parseObject(ByRef Str As String, ByRef index As Long) As Dictionary
    Set parseObject = New Dictionary
    Dim sKey    As String
    Dim charint As Integer
    Call skipChar(index)
    If m_str(index) <> A_CURLY_BRACKET_OPEN Then
        m_parserrors = m_parserrors & "Objeto invalido en la posicion " & index & " : " & mid$(Str, index) & vbCrLf
        Exit Function
    End If
    index = index + 1
    Do
        Call skipChar(index)
        charint = m_str(index)
        Select Case charint
            Case A_COMMA
                index = index + 1
                Call skipChar(index)
            
            Case A_CURLY_BRACKET_CLOSE
                index = index + 1
                Exit Do
                
            Case index > m_length
                m_parserrors = m_parserrors & "Falta '}': " & Right$(Str, 20) & vbCrLf
                Exit Do
        End Select
        sKey = parseKey(index)
        On Error Resume Next
        Call parseObject.Add(sKey, parseValue(Str, index))
        If Err.Number <> 0 Then
            m_parserrors = m_parserrors & Err.description & ": " & sKey & vbCrLf
            Exit Do
        End If
    Loop
End Function

Private Function parseArray(ByRef Str As String, ByRef index As Long) As Collection
    Dim charint As Integer
    Set parseArray = New Collection
    Call skipChar(index)
    If mid$(Str, index, 1) <> "[" Then
        m_parserrors = m_parserrors & "Array invalido en la posicion " & index & " : " + mid$(Str, index, 20) & vbCrLf
        Exit Function
    End If
    index = index + 1
    Do
        Call skipChar(index)
        charint = m_str(index)
        If charint = A_SQUARE_BRACKET_CLOSE Then
            index = index + 1
            Exit Do
        ElseIf charint = A_COMMA Then
            index = index + 1
            Call skipChar(index)
        ElseIf index > m_length Then
            m_parserrors = m_parserrors & "Falta ']': " & Right$(Str, 20) & vbCrLf
            Exit Do
        End If
        On Error Resume Next
        Call parseArray.Add(parseValue(Str, index))
        If Err.Number <> 0 Then
            m_parserrors = m_parserrors & Err.description & ": " & mid$(Str, index, 20) & vbCrLf
            Exit Do
        End If
    Loop
End Function

Private Function parseValue(ByRef Str As String, ByRef index As Long)
    Call skipChar(index)
    Select Case m_str(index)
        Case A_DOUBLE_QUOTE, A_SINGLE_QUOTE
            parseValue = parseString(Str, index)
            Exit Function

        Case A_SQUARE_BRACKET_OPEN
            Set parseValue = parseArray(Str, index)
            Exit Function

        Case A_t, A_f
            parseValue = parseBoolean(Str, index)
            Exit Function

        Case A_n
            parseValue = parseNull(Str, index)
            Exit Function

        Case A_CURLY_BRACKET_OPEN
            Set parseValue = parseObject(Str, index)
            Exit Function

        Case Else
            parseValue = parseNumber(Str, index)
            Exit Function
    End Select
End Function

Private Function parseString(ByRef Str As String, ByRef index As Long) As String
    Dim quoteint As Integer
    Dim charint  As Integer
    Dim code     As String
    Call skipChar(index)
    quoteint = m_str(index)
    index = index + 1
    Do While index > 0 And index <= m_length
        charint = m_str(index)
        Select Case charint
            Case A_BACKSLASH
                index = index + 1
                charint = m_str(index)
                Select Case charint
                    Case A_DOUBLE_QUOTE, A_BACKSLASH, A_FORWARDSLASH, A_SINGLE_QUOTE
                        parseString = parseString & ChrW$(charint)
                        index = index + 1

                    Case A_b
                        parseString = parseString & vbBack
                        index = index + 1

                    Case A_f
                        parseString = parseString & vbFormFeed
                        index = index + 1

                    Case A_n
                        parseString = parseString & vbLf
                        index = index + 1

                    Case A_r
                        parseString = parseString & vbCr
                        index = index + 1

                    Case A_t
                        parseString = parseString & vbTab
                        index = index + 1

                    Case A_u
                        index = index + 1
                        code = mid$(Str, index, 4)

                        parseString = parseString & ChrW$(val("&h" + code))
                        index = index + 4
                End Select

            Case quoteint
                index = index + 1
                Exit Function

            Case Else
                parseString = parseString & ChrW$(charint)
                index = index + 1
        End Select
    Loop
End Function

Private Function parseNumber(ByRef Str As String, ByRef index As Long)
    Dim Value As String
    Dim Char  As String
    Call skipChar(index)
    Do While index > 0 And index <= m_length
        Char = mid$(Str, index, 1)
        If InStr("+-0123456789.eE", Char) Then
            Value = Value & Char
            index = index + 1
        Else
            If Not m_decSep = "." Then
                Value = Replace(Value, ".", m_decSep)
            End If
            If m_groupSep = "." Then
                Value = Replace(Value, ".", m_decSep)
            End If
            parseNumber = CDec(Value)
            Exit Function
        End If
    Loop
End Function

Private Function parseBoolean(ByRef Str As String, ByRef index As Long) As Boolean
    Call skipChar(index)
    If mid$(Str, index, 4) = "true" Then
        parseBoolean = True
        index = index + 4
    ElseIf mid$(Str, index, 5) = "false" Then
        parseBoolean = False
        index = index + 5
    Else
        m_parserrors = m_parserrors & "Boolean invalido en la posicion " & index & " : " & mid$(Str, index) & vbCrLf
    End If
End Function

Private Function parseNull(ByRef Str As String, ByRef index As Long)
    Call skipChar(index)
    If mid$(Str, index, 4) = "null" Then
        parseNull = Null
        index = index + 4
    Else
        m_parserrors = m_parserrors & "Valor nulo invalido en la posicion " & index & " : " & mid$(Str, index) & vbCrLf
    End If
End Function

Private Function parseKey(ByRef index As Long) As String
    Dim dquote  As Boolean
    Dim squote  As Boolean
    Dim charint As Integer
    Call skipChar(index)
    Do While index > 0 And index <= m_length
        charint = m_str(index)
        Select Case charint
            Case A_DOUBLE_QUOTE
                dquote = Not dquote
                index = index + 1
                If Not dquote Then
                    Call skipChar(index)
                    If m_str(index) <> A_COLON Then
                        m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & index & " : " & parseKey & vbCrLf
                        Exit Do
                    End If
                End If
                
            Case A_SINGLE_QUOTE
                squote = Not squote
                index = index + 1
                If Not squote Then
                    Call skipChar(index)
                    If m_str(index) <> A_COLON Then
                        m_parserrors = m_parserrors & "Valor clave invalido en la posicion " & index & " : " & parseKey & vbCrLf
                        Exit Do
                    End If
                End If
        
            Case A_COLON
                index = index + 1
                If Not dquote And Not squote Then
                    Exit Do
                Else
                    parseKey = parseKey & ChrW$(charint)
                End If

            Case Else
                If A_VBCRLF = charint Then
                ElseIf A_VBCR = charint Then
                ElseIf A_VBLF = charint Then
                ElseIf A_VBTAB = charint Then
                ElseIf A_SPACE = charint Then
                Else
                    parseKey = parseKey & ChrW$(charint)
                End If
                index = index + 1
        End Select
    Loop
End Function

Private Sub skipChar(ByRef index As Long)
    Dim bComment      As Boolean
    Dim bStartComment As Boolean
    Dim bLongComment  As Boolean
    Do While index > 0 And index <= m_length
        Select Case m_str(index)
            Case A_VBCR, A_VBLF
                If Not bLongComment Then
                    bStartComment = False
                    bComment = False
                End If
    
            Case A_VBTAB, A_SPACE, A_BRACKET_OPEN, A_BRACKET_CLOSE
        
            Case A_FORWARDSLASH
                If Not bLongComment Then
                    If bStartComment Then
                        bStartComment = False
                        bComment = True
                    Else
                        bStartComment = True
                        bComment = False
                        bLongComment = False
                    End If
                Else
                    If bStartComment Then
                        bLongComment = False
                        bStartComment = False
                        bComment = False
                    End If
                End If

            Case A_ASTERIX
                If bStartComment Then
                    bStartComment = False
                    bComment = True
                    bLongComment = True
                Else
                    bStartComment = True
                End If

            Case Else
                If Not bComment Then
                    Exit Do
                End If
        End Select
        index = index + 1
    Loop
End Sub

Public Function GetRegionalSettings(ByVal regionalsetting As Long) As String
    On Error GoTo ErrorHandler
    Dim Locale      As Long
    Dim Symbol      As String
    Dim iRet1       As Long
    Dim iRet2       As Long
    Dim lpLCDataVar As String
    Dim Pos         As Integer
    Locale = GetUserDefaultLCID()
    iRet1 = GetLocaleInfo(Locale, regionalsetting, lpLCDataVar, 0)
    Symbol = String$(iRet1, 0)
    iRet2 = GetLocaleInfo(Locale, regionalsetting, Symbol, iRet1)
    Pos = InStr(Symbol, Chr$(0))
    If Pos > 0 Then
        Symbol = Left$(Symbol, Pos - 1)
    End If
ErrorHandler:
    GetRegionalSettings = Symbol
    Select Case Err.Number
        Case 0

        Case Else
            Call Err.Raise(123, "GetRegionalSetting", "GetRegionalSetting: " & regionalsetting)
    End Select
End Function

Private Function Encode(Str) As String
    Dim SB  As New cStringBuilder
    Dim i   As Long
    Dim j   As Long
    Dim aL1 As Variant
    Dim aL2 As Variant
    Dim c   As String
    Dim p   As Boolean
    aL1 = Array(&H22, &H5C, &H2F, &H8, &HC, &HA, &HD, &H9)
    aL2 = Array(&H22, &H5C, &H2F, &H62, &H66, &H6E, &H72, &H74)
    For i = 1 To LenB(Str)
        p = True
        c = mid$(Str, i, 1)
        For j = 0 To 7
            If c = Chr$(aL1(j)) Then
                Call SB.Append("\" & Chr$(aL2(j)))
                p = False
                Exit For
            End If
        Next
        If p Then
            Dim a As Integer
                a = AscW(c)
            If a > 31 And a < 127 Then
                Call SB.Append(c)
            ElseIf a > -1 Or a < 65535 Then
                Call SB.Append("\u" & String$(4 - LenB(Hex$(a)), "0") & Hex$(a))
            End If
        End If
    Next
    Encode = SB.toString
    Set SB = Nothing
End Function

Public Function StringToJSON(st As String) As String
    Const FIELD_SEP = "~"
    Const RECORD_SEP = "|"
    Dim sFlds   As String
    Dim sRecs   As New cStringBuilder
    Dim lRecCnt As Long
    Dim lFld    As Long
    Dim fld     As Variant
    Dim rows    As Variant
    Dim Lower_fld As Long, Upper_fld As Long
    lRecCnt = 0
    If LenB(st) = 0 Then
        StringToJSON = "null"
    Else
        rows = Split(st, RECORD_SEP)
        For lRecCnt = LBound(rows) To UBound(rows)
            sFlds = vbNullString
            fld = Split(rows(lRecCnt), FIELD_SEP)
            For lFld = LBound(fld) To UBound(fld) Step 2
                sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld(lFld) & """:""" & toUnicode(fld(lFld + 1) & "") & """")
            Next
            Call sRecs.Append(IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}")
        Next
        StringToJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")
    End If
End Function

Public Function RStoJSON(rs As ADODB.Recordset) As String
    On Error GoTo ErrorHandler
    Dim sFlds   As String
    Dim sRecs   As New cStringBuilder
    Dim lRecCnt As Long
    Dim fld     As ADODB.Field
    lRecCnt = 0
    If rs.State = adStateClosed Then
        RStoJSON = "null"
    Else
        If rs.EOF Or rs.BOF Then
            RStoJSON = "null"
        Else
            Do While Not rs.EOF And Not rs.BOF
                lRecCnt = lRecCnt + 1
                sFlds = vbNullString
                For Each fld In rs.Fields
                    sFlds = (sFlds & IIf(sFlds <> "", ",", "") & """" & fld.Name & """:""" & toUnicode(fld.Value & "") & """")
                Next
                Call sRecs.Append(IIf((Trim$(sRecs.toString) <> ""), "," & vbNewLine, "") & "{" & sFlds & "}")
                Call rs.MoveNext
            Loop
            RStoJSON = ("( {""Records"": [" & vbNewLine & sRecs.toString & vbNewLine & "], " & """RecordCount"":""" & lRecCnt & """ } )")
        End If
    End If
    Exit Function
ErrorHandler:

End Function

Public Function toUnicode(Str As String) As String
    Dim X        As Long
    Dim uStr     As New cStringBuilder
    Dim uChrCode As Integer
    For X = 1 To LenB(Str)
        uChrCode = Asc(mid$(Str, X, 1))
        Select Case uChrCode
            Case 8:
                Call uStr.Append("\b")

            Case 9:
                Call uStr.Append("\t")

            Case 10:
                Call uStr.Append("\n")

            Case 12:
                Call uStr.Append("\f")

            Case 13:
                Call uStr.Append("\r")

            Case 34:
                Call uStr.Append("\""")

            Case 39:
                Call uStr.Append("\'")

            Case 92:
                Call uStr.Append("\\")

            Case 123, 125:
                Call uStr.Append("\u" & Right$("0000" & Hex$(uChrCode), 4))

            Case Is < 32, Is > 127:
                Call uStr.Append("\u" & Right$("0000" & Hex$(uChrCode), 4))

            Case Else
                Call uStr.Append(Chr$(uChrCode))
        End Select
    Next
    toUnicode = uStr.toString
    Exit Function
End Function
