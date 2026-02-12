' ===== Class Module: XInputState =====
Option Compare Database
Option Explicit

Private mFieldName As String
Private mValueType As XValueType
Private mValue As Variant
Private mIsValid As Boolean
Private mErrorMessage As String

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
