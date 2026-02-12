Option Compare Database
Option Explicit

' Returns True if a form object exists in the project (regardless of whether it's open).
Public Function FormExists(ByVal formName As String) As Boolean
    On Error Resume Next
    err.Clear
    Dim ao As AccessObject: Set ao = CurrentProject.AllForms(formName)
    FormExists = (err.Number = 0)
    err.Clear: On Error GoTo 0
End Function

' Returns the Form object if the form is loaded/open; otherwise returns Nothing.
Public Function FormLoaded(ByVal formName As String) As Form
    On Error Resume Next
    Set FormLoaded = Forms(formName)
    On Error GoTo 0
End Function

' Attempts to open a form if it exists and is not already loaded.
Public Sub TryOpenForm(ByVal formName As String, Optional ByVal formArgs As Variant = Null)
    If Not FormExists(formName) Then Exit Sub

    Dim frm As Form
    Set frm = FormLoaded(formName)

    If Not (frm Is Nothing) Then
        Dim displayName As String
        displayName = Nz(frm.Caption, "")
        
        If Len(Trim$(displayName)) = 0 Then displayName = frm.name
            MsgBox displayName & " is in use." & vbCrLf & _
            "Finish your work and close it before opening another instance.", _
            vbExclamation, "Form In Use"

        Exit Sub
    End If

    If IsNull(formArgs) Then
        DoCmd.OpenForm formName
    Else
        DoCmd.OpenForm formName, , , , , , formArgs
    End If
End Sub

' Returns True if the given form is loaded as a subform of the specified parent form.
Public Function HasLoadedParent(ByVal frm As Form, ByVal parentFormName As String) As Boolean
    Dim p As Object

    On Error Resume Next
    Set p = frm.Parent
    On Error GoTo 0

    If p Is Nothing Then Exit Function
    If StrComp(p.name, parentFormName, vbTextCompare) <> 0 Then Exit Function

    HasLoadedParent = True
End Function

' Returns True if a control exists on an OPEN form.
Public Function ControlExists(ByVal formName As String, ByVal ControlName As String) As Boolean
    On Error Resume Next
    Dim ctl As Control
    Set ctl = Forms(formName).Controls(ControlName)
    ControlExists = (err.Number = 0)
    err.Clear: On Error GoTo 0
End Function

' Returns the Control object if it exists on the given (open) form; otherwise returns Nothing.
Public Function GetControl(ByVal frm As Form, ByVal ControlName As String) As Control
    On Error Resume Next
    Set GetControl = frm.Controls(ControlName)
    On Error GoTo 0
End Function

' Initializes a datasheet form and enforces Datasheet DefaultView
Public Sub ConfigCustomDatasheet(ByVal frm As Form)
    AssertTrue frm.DefaultView = acDefViewDatasheet, _
           "XFormUtil.ConfigCustomDatasheet", "The DefaultView of " & frm.name & " must be Datasheet Form."
    frm.Modal = False
    frm.PopUp = False
    
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = False
    frm.AllowFilters = False
    
    frm.NavigationButtons = True
    frm.RecordSelectors = False
End Sub

' Initializes an custom form and enforces Single Form DefaultView
Public Sub ConfigCustomForm(ByVal frm As Form)
    AssertTrue frm.DefaultView = acDefViewSingle, _
           "XFormUtil.ConfigCustomForm", "The DefaultView of " & frm.name & " must be Single Form."

    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.AllowEdits = True
    frm.AllowFilters = False
    
    frm.NavigationButtons = False
    frm.RecordSelectors = False
End Sub

' Invokes a Public method on a loaded form by name.
Public Sub TryCallForm(ByVal formName As String, ByVal methodName As String, ParamArray args() As Variant)
    Dim frm As Form
    Set frm = FormLoaded(formName)
    If frm Is Nothing Then Exit Sub

    ' ParamArray always exists as an array.
    ' If no args were passed, UBound(args) = -1
    Select Case UBound(args)
        Case -1
            CallByName frm, methodName, VbMethod

        Case 0
            CallByName frm, methodName, VbMethod, args(0)

        Case 1
            CallByName frm, methodName, VbMethod, args(0), args(1)

        Case Else
            XRaise "XFormUtil.TryCallForm", "Too many args."
    End Select
End Sub

' Assumes DatasheetForm.OrderBy looks like:
'   [Datasheet_FormName].[TicketName] ASC
'   [Datasheet.Form.Name].[Ticket.Name] DESC
'
' Returns:
'   [TicketName] ASC/DESC  (or [Ticket.Name] ASC/DESC)
Public Function ToRsSortByAdo(ByVal formOrderBy As String) As String
    Dim s As String
    Dim dir As String
    Dim pos As Long

    s = Trim$(formOrderBy)
    If Len(s) = 0 Then
        ToRsSortByAdo = vbNullString
        Exit Function
    End If

    ' Direction (default ASC)
    dir = " ASC"
    If UCase$(Right$(s, 5)) = " DESC" Then
        dir = " DESC"
        s = Trim$(Left$(s, Len(s) - 5))
    ElseIf UCase$(Right$(s, 4)) = " ASC" Then
        dir = " ASC"
        s = Trim$(Left$(s, Len(s) - 4))
    End If

    ' Keep only the field part after "].["
    pos = InStrRev(s, "].[")
    If pos > 0 Then
        s = Mid$(s, pos + 2)
    End If

    ' Ensure bracketed (in case it came as unqualified)
    s = Trim$(s)
    If Len(s) > 0 Then
        If Left$(s, 1) <> "[" Then s = "[" & s
        If Right$(s, 1) <> "]" Then s = s & "]"
    End If

    ToRsSortByAdo = s & dir
End Function

Public Function BuildRsSortByAdo( _
    ByVal sortByAdo As String, _
    ByVal datasheetForm As Form, _
    ByVal defaultVal As String _
) As String

    ' Explicit override
    If Len(sortByAdo) > 0 Then
        BuildRsSortByAdo = sortByAdo
        Exit Function
    End If

    ' Form's current recordset sort (ADO only; may error or recordset may be Nothing)
    Dim sb As String
    sb = ""

    If Not (datasheetForm.Recordset Is Nothing) Then
        On Error Resume Next
        sb = datasheetForm.Recordset.Sort
        On Error GoTo 0
    End If

    If Len(sb) > 0 Then
        BuildRsSortByAdo = sb
        Exit Function
    End If

    ' Fallback
    BuildRsSortByAdo = defaultVal
End Function
