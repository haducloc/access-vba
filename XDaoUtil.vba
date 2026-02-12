Option Compare Database
Option Explicit

' Begin a DAO transaction
Public Sub BeginTransDao(ByVal ws As DAO.Workspace)
    If Not ws Is Nothing Then ws.BeginTrans
End Sub

' Commit a DAO transaction
Public Sub CommitTransDao(ByVal ws As DAO.Workspace)
    On Error Resume Next
    If Not ws Is Nothing Then ws.CommitTrans
    On Error GoTo 0
End Sub

' Roll back a DAO transaction
Public Sub RollbackTransDao(ByVal ws As DAO.Workspace)
    On Error Resume Next
    If Not ws Is Nothing Then ws.Rollback
    On Error GoTo 0
End Sub

' Create a temporary DAO QueryDef from SQL
Public Function CreateCommandDao(ByVal db As DAO.Database, ByVal sql As String) As DAO.QueryDef
    Dim qd As DAO.QueryDef
    Set qd = db.CreateQueryDef("", sql)
    Set CreateCommandDao = qd
End Function

' Execute an action query and return records affected
Public Function ExecuteUpdateDao(ByVal qd As DAO.QueryDef) As Long
    qd.Execute dbFailOnError
    ExecuteUpdateDao = qd.RecordsAffected
End Function

' Execute a scalar query and return the first column of the first row
Public Function ExecuteScalarDao(ByVal qd As DAO.QueryDef) As Variant
    Dim rs As DAO.Recordset
    Dim xe As XError
    On Error GoTo TCError

    Set rs = qd.OpenRecordset(dbOpenSnapshot)

    If (rs.BOF And rs.EOF) Then
        ExecuteScalarDao = Null
    Else
        ExecuteScalarDao = rs.Fields(0).Value
    End If

    CloseObj rs
    Exit Function

TCError:
    Set xe = ToXError(Err)
    CloseObj rs
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Execute a query and return a snapshot recordset
Public Function ExecuteQueryDao(ByVal qd As DAO.QueryDef) As DAO.Recordset
    Set ExecuteQueryDao = qd.OpenRecordset(dbOpenSnapshot)
End Function

' Assign a value to a positional DAO parameter
Private Sub SetParam(ByVal qd As DAO.QueryDef, ByVal index As Long, ByVal value As Variant)
    If IsNull(value) Or IsEmpty(value) Then
        qd.Parameters(index).Value = Null
    Else
        qd.Parameters(index).Value = value
    End If
End Sub

' Set a LIKE parameter with normalization and truncation
Public Sub ParamLikeDao( _
    ByVal qd As DAO.QueryDef, _
    ByVal index As Long, _
    ByVal value As Variant, _
    Optional ByVal maxLikeSize As Long = 255, _
    Optional ByVal dbType As XDbType = XDbType.Db_Access _
)
    If IsNull(value) Or IsEmpty(value) Then
        SetParam qd, index, Null
        Exit Sub
    End If

    Dim str As String
    str = CStr(value)

    If maxLikeSize > 0 And Len(str) > maxLikeSize Then
        str = Left$(str, maxLikeSize)
    End If

    SetParam qd, index, ToLikeParamValue(str, dbType)
End Sub

' Convert the current DAO record into a Dictionary
Public Function RecordToDictDao(ByVal rs As DAO.Recordset) As Object
    Dim row As Object
    Set row = NewDict()

    Dim f As DAO.Field
    For Each f In rs.Fields
        row(f.Name) = f.Value
    Next

    Set RecordToDictDao = row
End Function

' Escape/quote an identifier (column/table/parameter name) for Access/DAO.
' Allows dotted names like dbo.Table or schema.table by escaping each segment.
Public Function EscapeIdentDao(ByVal ident As String) As String
    Dim parts() As String
    Dim i As Long
    Dim p As String
    Dim outSql As String

    AssertHasValue ident, "XDaoUtil.EscapeIdentDao", "ident is blank."

    parts = Split(ident, ".")

    For i = LBound(parts) To UBound(parts)
        p = Trim$(parts(i))
        AssertHasValue p, "XDaoUtil.EscapeIdentDao", "ident contains a blank segment."
        p = Replace(p, "]", "]]")
        If Len(outSql) > 0 Then outSql = outSql & "."
        outSql = outSql & "[" & p & "]"
    Next i

    EscapeIdentDao = outSql
End Function

' Resolve the Workspace that owns the provided Database.
' Falls back to the default workspace if not found.
Public Function GetWorkspace(ByVal db As DAO.Database) As DAO.Workspace
    Dim ws As DAO.Workspace
    Dim wsDb As DAO.Database

    On Error Resume Next

    For Each ws In DBEngine.Workspaces
        For Each wsDb In ws.Databases
            ' Matching Name?
            If StrComp(wsDb.Name, db.Name, vbTextCompare) = 0 Then
                
                ' Matching .Connect?
                If StrComp(Nz(wsDb.Connect, vbNullString), Nz(db.Connect, vbNullString), vbTextCompare) = 0 Then
                    Set GetWorkspace = ws
                    Exit Function
                End If
            End If
        Next wsDb
    Next ws

    Err.Clear
    Set GetWorkspace = DBEngine.Workspaces(0)
End Function
