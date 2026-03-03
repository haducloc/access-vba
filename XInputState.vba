' ===== Class Module: XInputState =====
Option Compare Database
Option Explicit

Private mInput As Control
Private mFieldName As String
Private mValueType As XValueType
Private mValue As Variant
Private mIsValid As Boolean
Private mErrorMessage As String

Public Property Get FormInput() As Control
    FormInput = mInput
End Property

Public Property Let FormInput(ByVal v As Control)
    Set mInput = v
End Property

Public Property Get FieldName() As String
    FieldName = mFieldName
End Property

Public Property Let FieldName(ByVal v As String)
    mFieldName = v
End Property

Public Property Get ValueType() As XValueType
    ValueType = mValueType
End Property

Public Property Let ValueType(ByVal v As XValueType)
    mValueType = v
End Property

Public Property Get Value() As Variant
    Value = mValue
End Property

Public Property Let Value(ByVal v As Variant)
    mValue = v
End Property

Public Property Get IsValid() As Boolean
    IsValid = mIsValid
End Property

Public Property Let IsValid(ByVal v As Boolean)
    mIsValid = v
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = mErrorMessage
End Property

Public Property Let ErrorMessage(ByVal v As String)
    mErrorMessage = v
End Property

' =========================
' STRING / NUMBER / DATE VALIDATION
' (MarkInvalidInput added everywhere an invalid state is set)
' =========================

Private Sub MarkInvalidInput()
    On Error Resume Next
    CallByName mInput, "BorderColor", VbLet, vbRed
    On Error GoTo 0
End Sub

' =========================
' STRING VALIDATION
' =========================

Public Sub ValidateMaxLength(ByVal maxLen As Integer)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub
    
    Dim str As String: str = CStr(mValue)
    If Len(str) > maxLen Then
        mIsValid = False
        mErrorMessage = mInput.name & " is too long (max length=" & maxLen & ")"
        MarkInvalidInput
    End If
End Sub

Public Sub ValidateMinLength(ByVal minLen As Integer)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub

    Dim str As String: str = CStr(mValue)
    If Len(str) < minLen Then
        mIsValid = False
        mErrorMessage = mInput.name & " is too short (min length=" & minLen & ")"
        MarkInvalidInput
    End If
End Sub

Public Sub ValidateLength(ByVal exactLen As Integer)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub

    Dim str As String: str = CStr(mValue)
    If Len(str) <> exactLen Then
        mIsValid = False
        mErrorMessage = mInput.name & " must be exactly " & exactLen & " characters"
        MarkInvalidInput
    End If
End Sub

Public Sub ValidateLengthRange(ByVal minLen As Integer, ByVal maxLen As Integer)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub

    Dim str As String: str = CStr(mValue)
    Dim n As Long: n = Len(str)

    If n < minLen Or n > maxLen Then
        mIsValid = False
        mErrorMessage = mInput.name & " length must be between " & minLen & " and " & maxLen
        MarkInvalidInput
    End If
End Sub

' =========================
' INTEGER VALIDATION (Long)
' =========================

Public Sub ValidateIntegerMin(ByVal minVal As Long)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub
    If IsNumeric(mValue) = False Then Exit Sub

    If CLng(mValue) < minVal Then
        mIsValid = False
        mErrorMessage = mInput.name & " must be >= " & minVal
        MarkInvalidInput
    End If
End Sub

Public Sub ValidateIntegerMax(ByVal maxVal As Long)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub
    If IsNumeric(mValue) = False Then Exit Sub

    If CLng(mValue) > maxVal Then
        mIsValid = False
        mErrorMessage = mInput.name & " must be <= " & maxVal
        MarkInvalidInput
    End If
End Sub

Public Sub ValidateIntegerRange(ByVal minVal As Long, ByVal maxVal As Long)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub
    If IsNumeric(mValue) = False Then Exit Sub

    Dim n As Long: n = CLng(mValue)
    If n < minVal Or n > maxVal Then
        mIsValid = False
        mErrorMessage = mInput.name & " must be between " & minVal & " and " & maxVal
        MarkInvalidInput
    End If
End Sub

' =========================
' DOUBLE VALIDATION
' =========================

Public Sub ValidateDoubleMin(ByVal minVal As Double)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub
    If IsNumeric(mValue) = False Then Exit Sub

    If CDbl(mValue) < minVal Then
        mIsValid = False
        mErrorMessage = mInput.name & " must be >= " & CStr(minVal)
        MarkInvalidInput
    End If
End Sub

Public Sub ValidateDoubleMax(ByVal maxVal As Double)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub
    If IsNumeric(mValue) = False Then Exit Sub

    If CDbl(mValue) > maxVal Then
        mIsValid = False
        mErrorMessage = mInput.name & " must be <= " & CStr(maxVal)
        MarkInvalidInput
    End If
End Sub

Public Sub ValidateDoubleRange(ByVal minVal As Double, ByVal maxVal As Double)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub
    If IsNumeric(mValue) = False Then Exit Sub

    Dim d As Double: d = CDbl(mValue)
    If d < minVal Or d > maxVal Then
        mIsValid = False
        mErrorMessage = mInput.name & " must be between " & CStr(minVal) & " and " & CStr(maxVal)
        MarkInvalidInput
    End If
End Sub

' =========================
' DATE VALIDATION
' =========================

Public Sub ValidateDateMin(ByVal minDate As Date)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub
    If IsDate(mValue) = False Then Exit Sub

    If CDate(mValue) < minDate Then
        mIsValid = False
        mErrorMessage = mInput.name & " must be on/after " & Format$(minDate, "mm/dd/yyyy")
        MarkInvalidInput
    End If
End Sub

Public Sub ValidateDateMax(ByVal maxDate As Date)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub
    If IsDate(mValue) = False Then Exit Sub

    If CDate(mValue) > maxDate Then
        mIsValid = False
        mErrorMessage = mInput.name & " must be on/before " & Format$(maxDate, "mm/dd/yyyy")
        MarkInvalidInput
    End If
End Sub

Public Sub ValidateDateRange(ByVal minDate As Date, ByVal maxDate As Date)
    If mIsValid = False Then Exit Sub
    If IsNull(mValue) Then Exit Sub
    If IsDate(mValue) = False Then Exit Sub

    Dim d As Date: d = CDate(mValue)
    If d < minDate Or d > maxDate Then
        mIsValid = False
        mErrorMessage = mInput.name & " must be between " & _
                        Format$(minDate, "mm/dd/yyyy") & " and " & _
                        Format$(maxDate, "mm/dd/yyyy")
        MarkInvalidInput
    End If
End Sub
