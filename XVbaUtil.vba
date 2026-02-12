Option Compare Database
Option Explicit

Public Const X_ERR_NUMBER As Long = vbObjectError + 1000

'Raise a standardized X error
Public Sub XRaise(ByVal source As String, ByVal message As String)
    Err.Raise X_ERR_NUMBER, source, message
End Sub

'Close and release late-bound object safely.
Public Sub CloseObj(ByRef obj As Object, Optional ByVal closeMethod As String = "Close")
    On Error Resume Next
    If Not obj Is Nothing Then
        CallByName obj, closeMethod, VbMethod
        Set obj = Nothing
    End If
    On Error GoTo 0
End Sub

'Convert a VBA ErrObject to an XError instance.
Public Function ToXError(ByVal err As VBA.ErrObject) As XError
    Dim xe As XError
    Set xe = New XError

    xe.ErrNum = err.Number
    xe.ErrDesc = err.Description
    xe.ErrSrc = err.source

    Set ToXError = xe
End Function

'Create and return a new Scripting.Dictionary instance.
'By default, keys are case-insensitive.
Public Function NewDict(Optional ByVal caseSensitive As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    If caseSensitive Then
        d.CompareMode = vbBinaryCompare
    Else
        d.CompareMode = vbTextCompare
    End If

    Set NewDict = d
End Function

'Check whether a dictionary contains the specified key (case-insensitive).
Public Function HasField(ByVal values As Object, ByVal fieldName As String) As Boolean
    On Error GoTo TCError

    Dim v As Variant
    v = values(fieldName)

    HasField = True
    Exit Function

TCError:
    HasField = False
End Function

' Create and return a new VBScript.RegExp instance.
' Params: rePattern, isIgnoreCase, isGlobal, isMultiLine
Public Function NewRegEx( _
    ByVal rePattern As String, _
    Optional ByVal isIgnoreCase As Boolean = True, _
    Optional ByVal isGlobal As Boolean = False, _
    Optional ByVal isMultiLine As Boolean = False _
) As Object

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    With re
        .Pattern = rePattern
        .IgnoreCase = isIgnoreCase
        .Global = isGlobal
        .MultiLine = isMultiLine
    End With

    Set NewRegEx = re
End Function
