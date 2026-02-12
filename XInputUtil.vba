Option Compare Database
Option Explicit

'Build an XInputState and optionally set the control BorderColor based on validity.
Private Function MakeState(ctrl As Control, fieldName As String, val As Variant, valType As XValueType, valid As Boolean, Optional errMsg As String = "") As XInputState
    Dim st As New XInputState
    st.FieldName = fieldName
    st.Value = val
    st.ValueType = valType
    st.IsValid = valid
    st.ErrorMessage = errMsg
    Set MakeState = st

    On Error Resume Next
    CallByName ctrl, "BorderColor", VbLet, IIf(valid, vbWindowText, vbRed)
    On Error GoTo 0
End Function

'Read a trimmed string value (Null when empty unless required).
Public Function GetString(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim v As Variant
    v = TrimToNull(ctrl.Value)

    If IsNull(v) Then
        If required Then
            Set GetString = MakeState(ctrl, fieldName, Null, Type_String, False, ctrl.Name & " is required.")
        Else
            Set GetString = MakeState(ctrl, fieldName, Null, Type_String, True)
        End If
        Exit Function
    End If

    Set GetString = MakeState(ctrl, fieldName, CStr(v), Type_String, True)
End Function

'Read an uppercased code string (Null when empty unless required).
Public Function GetCode(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim v As Variant
    v = TrimToNull(ctrl.Value)

    If IsNull(v) Then
        If required Then
            Set GetCode = MakeState(ctrl, fieldName, Null, Type_String, False, ctrl.Name & " is required.")
        Else
            Set GetCode = MakeState(ctrl, fieldName, Null, Type_String, True)
        End If
        Exit Function
    End If

    Set GetCode = MakeState(ctrl, fieldName, UCase$(CStr(v)), Type_String, True)
End Function

'Read a signed Byte-sized whole number (Null when empty unless required).
Public Function GetByte(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseByte(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetByte = MakeState(ctrl, fieldName, Null, Type_Byte, False, ctrl.Name & " is required.")
            Else
                Set GetByte = MakeState(ctrl, fieldName, Null, Type_Byte, True)
            End If
        Else
            Set GetByte = MakeState(ctrl, fieldName, parsed, Type_Byte, True)
        End If
    Else
        Set GetByte = MakeState(ctrl, fieldName, Null, Type_Byte, False, ctrl.Name & " must be a whole number. -128 to 127")
    End If
End Function

'Read an unsigned Byte-sized whole number (Null when empty unless required).
Public Function GetUByte(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseUByte(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetUByte = MakeState(ctrl, fieldName, Null, Type_UByte, False, ctrl.Name & " is required.")
            Else
                Set GetUByte = MakeState(ctrl, fieldName, Null, Type_UByte, True)
            End If
        Else
            Set GetUByte = MakeState(ctrl, fieldName, parsed, Type_UByte, True)
        End If
    Else
        Set GetUByte = MakeState(ctrl, fieldName, Null, Type_UByte, False, ctrl.Name & " must be a whole number. 0 to 255")
    End If
End Function

'Read a SmallInt-sized whole number (Null when empty unless required).
Public Function GetInt2(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseInt2(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetInt2 = MakeState(ctrl, fieldName, Null, Type_Int2, False, ctrl.Name & " is required (smallint).")
            Else
                Set GetInt2 = MakeState(ctrl, fieldName, Null, Type_Int2, True)
            End If
        Else
            Set GetInt2 = MakeState(ctrl, fieldName, parsed, Type_Int2, True)
        End If
    Else
        Set GetInt2 = MakeState(ctrl, fieldName, Null, Type_Int2, False, ctrl.Name & " must be a whole number.")
    End If
End Function

'Read an Int-sized whole number (Null when empty unless required).
Public Function GetInt4(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseInt4(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetInt4 = MakeState(ctrl, fieldName, Null, Type_Int4, False, ctrl.Name & " is required (int).")
            Else
                Set GetInt4 = MakeState(ctrl, fieldName, Null, Type_Int4, True)
            End If
        Else
            Set GetInt4 = MakeState(ctrl, fieldName, parsed, Type_Int4, True)
        End If
    Else
        Set GetInt4 = MakeState(ctrl, fieldName, Null, Type_Int4, False, ctrl.Name & " must be a whole number.")
    End If
End Function

'Read a BigInt-sized whole number (Null when empty unless required).
Public Function GetInt8(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseInt8(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetInt8 = MakeState(ctrl, fieldName, Null, Type_Int8, False, ctrl.Name & " is required (bigint).")
            Else
                Set GetInt8 = MakeState(ctrl, fieldName, Null, Type_Int8, True)
            End If
        Else
            Set GetInt8 = MakeState(ctrl, fieldName, parsed, Type_Int8, True)
        End If
    Else
        Set GetInt8 = MakeState(ctrl, fieldName, Null, Type_Int8, False, ctrl.Name & " must be a whole number.")
    End If
End Function

'Read a Single-precision number (Null when empty unless required).
Public Function GetFloat(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseFloat(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetFloat = MakeState(ctrl, fieldName, Null, Type_Float, False, ctrl.Name & " is required (float).")
            Else
                Set GetFloat = MakeState(ctrl, fieldName, Null, Type_Float, True)
            End If
        Else
            Set GetFloat = MakeState(ctrl, fieldName, parsed, Type_Float, True)
        End If
    Else
        Set GetFloat = MakeState(ctrl, fieldName, Null, Type_Float, False, ctrl.Name & " must be a valid number.")
    End If
End Function

'Read a Double-precision number (Null when empty unless required).
Public Function GetDouble(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseDouble(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetDouble = MakeState(ctrl, fieldName, Null, Type_Double, False, ctrl.Name & " is required (double).")
            Else
                Set GetDouble = MakeState(ctrl, fieldName, Null, Type_Double, True)
            End If
        Else
            Set GetDouble = MakeState(ctrl, fieldName, parsed, Type_Double, True)
        End If
    Else
        Set GetDouble = MakeState(ctrl, fieldName, Null, Type_Double, False, ctrl.Name & " must be a valid number.")
    End If
End Function

'Read a Decimal value (stored as Variant) (Null when empty unless required).
Public Function GetDecimal(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseDecimal(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetDecimal = MakeState(ctrl, fieldName, Null, Type_Decimal, False, ctrl.Name & " is required (decimal).")
            Else
                Set GetDecimal = MakeState(ctrl, fieldName, Null, Type_Decimal, True)
            End If
        Else
            Set GetDecimal = MakeState(ctrl, fieldName, parsed, Type_Decimal, True)
        End If
    Else
        Set GetDecimal = MakeState(ctrl, fieldName, Null, Type_Decimal, False, ctrl.Name & " must be a valid number.")
    End If
End Function

'Read a Date-only value (DateValue) (Null when empty unless required).
Public Function GetDate(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseDate(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetDate = MakeState(ctrl, fieldName, Null, Type_Date, False, ctrl.Name & " is required (date).")
            Else
                Set GetDate = MakeState(ctrl, fieldName, Null, Type_Date, True)
            End If
        Else
            Set GetDate = MakeState(ctrl, fieldName, parsed, Type_Date, True)
        End If
    Else
        Set GetDate = MakeState(ctrl, fieldName, Null, Type_Date, False, ctrl.Name & " must be a valid date.")
    End If
End Function

'Read a Time-only value (TimeValue) (Null when empty unless required).
Public Function GetTime(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseTime(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetTime = MakeState(ctrl, fieldName, Null, Type_Time, False, ctrl.Name & " is required (time).")
            Else
                Set GetTime = MakeState(ctrl, fieldName, Null, Type_Time, True)
            End If
        Else
            Set GetTime = MakeState(ctrl, fieldName, parsed, Type_Time, True)
        End If
    Else
        Set GetTime = MakeState(ctrl, fieldName, Null, Type_Time, False, ctrl.Name & " must be a valid time.")
    End If
End Function

'Read a Date+Time value (CDate) (Null when empty unless required).
Public Function GetDateTime(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseDateTime(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetDateTime = MakeState(ctrl, fieldName, Null, Type_DateTime, False, ctrl.Name & " is required (date/time).")
            Else
                Set GetDateTime = MakeState(ctrl, fieldName, Null, Type_DateTime, True)
            End If
        Else
            Set GetDateTime = MakeState(ctrl, fieldName, parsed, Type_DateTime, True)
        End If
    Else
        Set GetDateTime = MakeState(ctrl, fieldName, Null, Type_DateTime, False, ctrl.Name & " must be a valid date/time.")
    End If
End Function

'Read a Boolean from common text tokens (Null when empty unless required).
Public Function GetBool(ctrl As Control, Optional fieldName As String = "", Optional required As Boolean = False) As XInputState
    Dim parsed As Variant, ok As Boolean
    ok = ParseBool(ctrl.Value, parsed)

    If ok Then
        If IsNull(parsed) Then
            If required Then
                Set GetBool = MakeState(ctrl, fieldName, Null, Type_Bool, False, ctrl.Name & " is required.")
            Else
                Set GetBool = MakeState(ctrl, fieldName, Null, Type_Bool, True)
            End If
        Else
            Set GetBool = MakeState(ctrl, fieldName, parsed, Type_Bool, True)
        End If
    Else
        Set GetBool = MakeState(ctrl, fieldName, Null, Type_Bool, False, ctrl.Name & " must be a valid Boolean value.")
    End If
End Function
