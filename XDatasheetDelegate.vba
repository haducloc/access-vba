' ===== Class Module: XDatasheetDelegate =====
Option Compare Database
Option Explicit

' Delegate state
Private hostForm As Access.Form

' One-shot flags so we can defer work until Access finishes its internal sort/filter processing
Private pendingReload As Boolean
Private pendingSortAdo As String

' Configuration
Private parentFormName As String
Private reloadMethodName As String
Private editFormName As String
Private pkFieldName As String
Private timerMs As Long

' Initializes the delegate with the host form and configuration values
Public Sub Init( _
    ByVal datasheetForm As Access.Form, _
    ByVal parentForm As String, _
    ByVal reloadMethod As String, _
    ByVal editForm As String, _
    ByVal pkField As String, _
    Optional ByVal timerIntervalMs As Long = 100 _
)
    Set hostForm = AssertNotNothing(datasheetForm, "XDatasheetDelegate.Init", "datasheetForm is required.")

    parentFormName = AssertHasValue(parentForm, "XDatasheetDelegate.Init", "parentForm is required.")
    reloadMethodName = AssertHasValue(reloadMethod, "XDatasheetDelegate.Init", "reloadMethod is required.")
    editFormName = AssertHasValue(editForm, "XDatasheetDelegate.Init", "editForm is required.")
    pkFieldName = AssertHasValue(pkField, "XDatasheetDelegate.Init", "pkField is required.")
    timerMs = timerIntervalMs

    pendingReload = False
    pendingSortAdo = vbEmptyString
End Sub

' Runs the one-shot timer handler to perform any deferred reload
Public Sub Form_Timer()
    ' One-shot timer
    hostForm.TimerInterval = 0

    ' Nothing pending
    If Not pendingReload Then Exit Sub
    pendingReload = False

    ' Invoke <ParentFormName>.<ReloadMethodName>(pendingSortAdo)
    InvokeFormMethod1 parentFormName, reloadMethodName, pendingSortAdo
End Sub

' Opens the configured edit form for the double-clicked record
Public Sub Form_DblClick(ByRef Cancel As Integer)
    If hostForm.Recordset Is Nothing Then
        Exit Sub
    End If
    
    Dim pkFieldValue As Variant
    
    ' Try to get pkFieldValue from the recordset
    On Error Resume Next
    pkFieldValue = hostForm.Recordset.fields(pkFieldName).value
    On Error GoTo 0

    If IsEmpty(pkFieldValue) Then
        Exit Sub
    End If

    ' Open Edit Form
    If FormExists(editFormName) Then
        DoCmd.OpenForm editFormName, , , , , , pkFieldValue
    End If
End Sub

' Cancels datasheet sort/filter and defers reload to avoid re-entrancy issues
Public Sub Form_ApplyFilter(ByRef Cancel As Integer, ByVal ApplyType As Integer)
    ' ApplyType: 1 - datasheet sorting
    If ApplyType = 1 Then
        ' Stop Access from applying its own sort/filter (can crash / throw 3075 with ADO recordset sources)
        Cancel = True

        ' Capture the current sort, convert Access OrderBy to ADO Sort format
        pendingReload = True
        pendingSortAdo = ToRsOrderByAdo(CStr(Nz(hostForm.OrderBy, vbEmptyString)))

        ' Defer reload to avoid re-entrancy while Access is still in ApplyFilter call stack
        hostForm.TimerInterval = timerMs
    End If
End Sub

' Handles dataset form errors
Public Sub Form_Error(ByVal DataErr As Integer, ByRef Response As Integer)
    If hostForm.CurrentView = acCurViewDatasheet Then
        If DataErr = 3075 Then
            Response = acDataErrContinue
            Exit Sub
        End If
    End If

    Response = acDataErrDisplay
End Sub
