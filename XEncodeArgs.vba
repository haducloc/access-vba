Option Explicit

' Configuration
Private Const sep As String = "|"
Private Const esc As String = "\"

'========================
' Public API
'========================

' Encode: accepts ParamArray of values, returns a single string.
Public Function EncodeArgs(ParamArray values() As Variant) As String
    Dim i As Long
    Dim parts() As String

    ' Handle empty ParamArray safely (LBound/UBound can error depending on host/call)
    On Error GoTo NoArgs
    ReDim parts(LBound(values) To UBound(values))
    On Error GoTo 0

    For i = LBound(values) To UBound(values)
        parts(i) = EscapePart(NormalizeToString(values(i)))
    Next i

    EncodeArgs = Join(parts, sep)
    Exit Function

NoArgs:
    EncodeArgs = vbNullString
End Function

' Decode: returns String() array.
' - If encoded is "", returns a TRUE empty array (0 elements).
Public Function DecodeArgs(ByVal encoded As String) As String()
    Dim emptyArr() As String

    If Len(encoded) = 0 Then
        DecodeArgs = emptyArr
        Exit Function
    End If

    DecodeArgs = SplitEscaped(encoded, sep, esc)
End Function

'========================
' Internal Helpers
'========================

Private Function NormalizeToString(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        NormalizeToString = vbNullString
    Else
        NormalizeToString = CStr(v)
    End If
End Function

Private Function EscapePart(ByVal s As String) As String
    If Len(s) = 0 Then
        EscapePart = vbNullString
        Exit Function
    End If

    Dim result As String
    result = Replace$(s, esc, esc & esc)          ' Escape the escape char first
    result = Replace$(result, sep, esc & sep)     ' Then escape the separator
    EscapePart = result
End Function

Private Function UnescapePart(ByVal s As String) As String
    Dim out As String
    Dim i As Long
    Dim length As Long

    length = Len(s)
    i = 1

    Do While i <= length
        If Mid$(s, i, 1) = esc Then
            ' If there is a character after the escape, take it literally
            If i < length Then
                out = out & Mid$(s, i + 1, 1)
                i = i + 2
            Else
                ' Trailing escape char: treat as literal
                out = out & esc
                i = i + 1
            End If
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop

    UnescapePart = out
End Function

Private Function SplitEscaped(ByVal s As String, ByVal sep As String, ByVal esc As String) As String()
    Dim parts() As String
    Dim cur As String
    Dim i As Long
    Dim ch As String
    Dim length As Long

    ReDim parts(0 To 0)
    length = Len(s)
    i = 1

    Do While i <= length
        ch = Mid$(s, i, 1)

        ' If we see an escape, keep it and the next char in the buffer
        ' so UnescapePart can handle it later.
        If ch = esc And i < length Then
            cur = cur & ch & Mid$(s, i + 1, 1)
            i = i + 2

        ElseIf ch = sep Then
            parts(UBound(parts)) = UnescapePart(cur)
            ReDim Preserve parts(0 To UBound(parts) + 1)
            cur = vbNullString
            i = i + 1

        Else
            cur = cur & ch
            i = i + 1
        End If
    Loop

    parts(UBound(parts)) = UnescapePart(cur)
    SplitEscaped = parts
End Function
