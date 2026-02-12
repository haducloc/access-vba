Option Compare Database
Option Explicit

' Convert a 2-column ADO Recordset (Value, DisplayName) into XDropdownOptions
Private Function ToDropdownOptionsAdo(ByVal rs As Object) As XDropdownOptions
    Dim opts As XDropdownOptions
    Set opts = New XDropdownOptions

    If (rs Is Nothing) Then
        Set ToDropdownOptionsAdo = opts
        Exit Function
    End If

    If rs.Fields.Count <> 2 Then
        XRaise "XAdoDropdown.ToDropdownOptionsAdo", "Recordset must have only 2 columns (Value, DisplayName)."
    End If

    ' Read rows
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            opts.Add rs.Fields(0).Value, Nz(rs.Fields(1).Value, "")
            rs.MoveNext
        Loop
    End If

    Set ToDropdownOptionsAdo = opts
End Function

' Execute an ADO Command and convert the 2-column result into XDropdownOptions
Public Function ExecuteDropdownOptionsAdo(ByVal cmd As Object) As XDropdownOptions
    Dim rs As Object
    Dim xe As XError
    On Error GoTo TCError

    Set rs = cmd.Execute
    Set ExecuteDropdownOptionsAdo = ToDropdownOptionsAdo(rs)

    CloseObj rs
    Exit Function

TCError:
    ' Preserve error, cleanup, then rethrow
    Set xe = ToXError(Err)

    CloseObj rs
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Execute a SQL query and convert the 2-column result into XDropdownOptions
Public Function ExecuteDropdownOptionsSqlAdo(ByVal connAdo As Object, ByVal sql As String) As XDropdownOptions
    Dim cmd As Object
    Set cmd = CreateCommandAdo(connAdo, sql)
    Set ExecuteDropdownOptionsSqlAdo = ExecuteDropdownOptionsAdo(cmd)
End Function
