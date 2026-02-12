Option Compare Database
Option Explicit

' Insert one row using parameter names = "p" + field name.
' fieldsCsv example: "Id,CtName,ReturnDate1"
' typesCsv example: "LONG,TEXT(100),DATETIME"
' values: dictionary-like object
Public Sub InsertRowDao(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldsCsv As String, ByVal typesCsv As String, ByVal values As Object)
    AssertHasValue tableName, "XDaoCrud.InsertRowDao", "tableName is blank."
    AssertHasValue fieldsCsv, "XDaoCrud.InsertRowDao", "fieldsCsv is blank."
    AssertHasValue typesCsv, "XDaoCrud.InsertRowDao", "typesCsv is blank."
    AssertNotNothing values, "XDaoCrud.InsertRowDao", "values is Nothing."
    AssertNotNothing db, "XDaoCrud.InsertRowDao", "db is Nothing."

    Dim qd As DAO.QueryDef
    Dim xe As XError

    On Error GoTo TCError
    Set qd = BuildInsertQd(db, tableName, fieldsCsv, typesCsv)

    BindParamsFromValues qd, values
    qd.Execute dbFailOnError

    CloseObj qd
    Exit Sub

TCError:
    Set xe = ToXError(Err)
    CloseObj qd
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Sub

' Update rows using parameter names = "p" + field name.
' setFieldsCsv example:   "CtName,ReturnDate1"
' setTypesCsv example:    "TEXT(100),DATETIME"
' whereFieldsCsv example: "Id"
' whereTypesCsv example:  "LONG"
' values: dictionary-like object (should include set + where fields; missing -> Null via BindParamsFromValues)
Public Sub UpdateRowDao( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal setFieldsCsv As String, _
    ByVal setTypesCsv As String, _
    ByVal whereFieldsCsv As String, _
    ByVal whereTypesCsv As String, _
    ByVal values As Object _
)
    AssertHasValue tableName, "XDaoCrud.UpdateRowDao", "tableName is blank."
    AssertHasValue setFieldsCsv, "XDaoCrud.UpdateRowDao", "setFieldsCsv is blank."
    AssertHasValue setTypesCsv, "XDaoCrud.UpdateRowDao", "setTypesCsv is blank."
    AssertHasValue whereFieldsCsv, "XDaoCrud.UpdateRowDao", "whereFieldsCsv is blank."
    AssertHasValue whereTypesCsv, "XDaoCrud.UpdateRowDao", "whereTypesCsv is blank."
    AssertNotNothing values, "XDaoCrud.UpdateRowDao", "values is Nothing."
    AssertNotNothing db, "XDaoCrud.UpdateRowDao", "db is Nothing."

    Dim qd As DAO.QueryDef
    Dim xe As XError

    On Error GoTo TCError
    Set qd = BuildUpdateQd(db, tableName, setFieldsCsv, setTypesCsv, whereFieldsCsv, whereTypesCsv)

    BindParamsFromValues qd, values
    qd.Execute dbFailOnError

    CloseObj qd
    Exit Sub

TCError:
    Set xe = ToXError(Err)
    CloseObj qd
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Sub

' Delete rows using parameter names = "p" + field name.
' whereFieldsCsv example: "Id"
' whereTypesCsv example:  "LONG"
' values: dictionary-like object (missing -> Null via BindParamsFromValues)
Public Sub DeleteRowDao( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal whereFieldsCsv As String, _
    ByVal whereTypesCsv As String, _
    ByVal values As Object _
)
    AssertHasValue tableName, "XDaoCrud.DeleteRowDao", "tableName is blank."
    AssertHasValue whereFieldsCsv, "XDaoCrud.DeleteRowDao", "whereFieldsCsv is blank."
    AssertHasValue whereTypesCsv, "XDaoCrud.DeleteRowDao", "whereTypesCsv is blank."
    AssertNotNothing values, "XDaoCrud.DeleteRowDao", "values is Nothing."
    AssertNotNothing db, "XDaoCrud.DeleteRowDao", "db is Nothing."

    Dim qd As DAO.QueryDef
    Dim xe As XError

    On Error GoTo TCError
    Set qd = BuildDeleteQd(db, tableName, whereFieldsCsv, whereTypesCsv)

    BindParamsFromValues qd, values
    qd.Execute dbFailOnError

    CloseObj qd
    Exit Sub

TCError:
    Set xe = ToXError(Err)
    CloseObj qd
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Sub

' Get single row by PK as Dictionary (param name = "p" + pkField).
' pkFieldsCsv example: "Id"
' pkTypesCsv example:  "LONG"
' values: dictionary-like object containing pk fields (missing -> Null via BindParamsFromValues)
Public Function GetRowByPkDao( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal pkFieldsCsv As String, _
    ByVal pkTypesCsv As String, _
    ByVal values As Object _
) As Object

    AssertHasValue tableName, "XDaoCrud.GetRowByPkDao", "tableName is blank."
    AssertHasValue pkFieldsCsv, "XDaoCrud.GetRowByPkDao", "pkFieldsCsv is blank."
    AssertHasValue pkTypesCsv, "XDaoCrud.GetRowByPkDao", "pkTypesCsv is blank."
    AssertNotNothing values, "XDaoCrud.GetRowByPkDao", "values is Nothing."
    AssertNotNothing db, "XDaoCrud.GetRowByPkDao", "db is Nothing."

    Dim qd As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim result As Object
    Dim xe As XError

    Dim paramsSql As String
    Dim whereSql As String

    On Error GoTo TCError

    AppendWhereParts pkFieldsCsv, pkTypesCsv, "XDaoCrud.GetRowByPkDao", paramsSql, whereSql

    Set qd = db.CreateQueryDef("", _
        "PARAMETERS " & paramsSql & ";" & vbCrLf & _
        "SELECT * FROM " & EscapeIdentDao(tableName) & vbCrLf & _
        "WHERE " & whereSql & ";")

    BindParamsFromValues qd, values
    Set rs = qd.OpenRecordset(dbOpenSnapshot)

    If (rs Is Nothing) Or (rs.BOF And rs.EOF) Then
        Set result = Nothing
    Else
        Set result = RecordToDictDao(rs)
    End If

    CloseObj rs
    CloseObj qd

    Set GetRowByPkDao = result
    Exit Function

TCError:
    Set xe = ToXError(Err)
    CloseObj rs
    CloseObj qd
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Returns True if a row exists for the given PK definition + values (param name = "p" + pkField).
' pkFieldsCsv example: "Id"
' pkTypesCsv example:  "LONG"
' values: dictionary-like object containing pk fields (missing -> Null via BindParamsFromValues)
Public Function ExistsByPkDao( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal pkFieldsCsv As String, _
    ByVal pkTypesCsv As String, _
    ByVal values As Object _
) As Boolean

    AssertHasValue tableName, "XDaoCrud.ExistsByPkDao", "tableName is blank."
    AssertHasValue pkFieldsCsv, "XDaoCrud.ExistsByPkDao", "pkFieldsCsv is blank."
    AssertHasValue pkTypesCsv, "XDaoCrud.ExistsByPkDao", "pkTypesCsv is blank."
    AssertNotNothing values, "XDaoCrud.ExistsByPkDao", "values is Nothing."
    AssertNotNothing db, "XDaoCrud.ExistsByPkDao", "db is Nothing."

    Dim qd As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim xe As XError

    Dim paramsSql As String
    Dim whereSql As String

    On Error GoTo TCError

    AppendWhereParts pkFieldsCsv, pkTypesCsv, "XDaoCrud.ExistsByPkDao", paramsSql, whereSql

    Set qd = db.CreateQueryDef("", _
        "PARAMETERS " & paramsSql & ";" & vbCrLf & _
        "SELECT TOP 1 1" & vbCrLf & _
        "FROM " & EscapeIdentDao(tableName) & vbCrLf & _
        "WHERE " & whereSql & ";")

    BindParamsFromValues qd, values
    Set rs = qd.OpenRecordset(dbOpenSnapshot)

    ' Exists if at least one row returned
    ExistsByPkDao = Not ((rs Is Nothing) Or (rs.BOF And rs.EOF))

    CloseObj rs
    CloseObj qd
    Exit Function

TCError:
    Set xe = ToXError(Err)
    CloseObj rs
    CloseObj qd
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Build a parameterized INSERT QueryDef (param names = "p" + field name).
Private Function BuildInsertQd(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldsCsv As String, ByVal typesCsv As String) As DAO.QueryDef
    Dim fields() As String
    Dim types() As String

    AssertNotNothing db, "XDaoCrud.BuildInsertQd", "db is Nothing."
    AssertHasValue tableName, "XDaoCrud.BuildInsertQd", "tableName is blank."
    AssertHasValue fieldsCsv, "XDaoCrud.BuildInsertQd", "fieldsCsv is blank."
    AssertHasValue typesCsv, "XDaoCrud.BuildInsertQd", "typesCsv is blank."

    fields = Split(fieldsCsv, ",")
    types = Split(typesCsv, ",")

    If UBound(fields) <> UBound(types) Then
        XRaise "XDaoCrud.BuildInsertQd", "fieldsCsv and typesCsv must have the same number of items."
    End If

    Dim paramsSql As String
    Dim colsSql As String
    Dim valsSql As String

    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        Dim f As String
        Dim t As String
        Dim p As String

        f = Trim$(fields(i))
        t = Trim$(types(i))
        p = "p" & f

        If Len(paramsSql) > 0 Then paramsSql = paramsSql & ", "
        paramsSql = paramsSql & EscapeIdentDao(p) & " " & t

        If Len(colsSql) > 0 Then colsSql = colsSql & ", "
        colsSql = colsSql & EscapeIdentDao(f)

        If Len(valsSql) > 0 Then valsSql = valsSql & ", "
        valsSql = valsSql & EscapeIdentDao(p)
    Next

    Dim sqlText As String
    sqlText = "PARAMETERS " & paramsSql & ";" & vbCrLf & _
              "INSERT INTO " & EscapeIdentDao(tableName) & " (" & colsSql & ")" & vbCrLf & _
              "VALUES (" & valsSql & ");"

    Set BuildInsertQd = db.CreateQueryDef("", sqlText)
End Function

' Build a parameterized UPDATE QueryDef (param names = "p" + field name).
Private Function BuildUpdateQd( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal setFieldsCsv As String, _
    ByVal setTypesCsv As String, _
    ByVal whereFieldsCsv As String, _
    ByVal whereTypesCsv As String _
) As DAO.QueryDef

    Dim setFields() As String
    Dim setTypes() As String

    AssertNotNothing db, "XDaoCrud.BuildUpdateQd", "db is Nothing."
    AssertHasValue tableName, "XDaoCrud.BuildUpdateQd", "tableName is blank."
    AssertHasValue setFieldsCsv, "XDaoCrud.BuildUpdateQd", "setFieldsCsv is blank."
    AssertHasValue setTypesCsv, "XDaoCrud.BuildUpdateQd", "setTypesCsv is blank."
    AssertHasValue whereFieldsCsv, "XDaoCrud.BuildUpdateQd", "whereFieldsCsv is blank."
    AssertHasValue whereTypesCsv, "XDaoCrud.BuildUpdateQd", "whereTypesCsv is blank."

    setFields = Split(setFieldsCsv, ",")
    setTypes = Split(setTypesCsv, ",")

    AssertTrue (UBound(setFields) = UBound(setTypes)), "XDaoCrud.BuildUpdateQd", _
        "setFieldsCsv and setTypesCsv must have the same number of items."

    Dim paramsSql As String
    Dim setSql As String
    Dim whereSql As String

    Dim i As Long

    ' PARAMETERS + SET clause
    For i = LBound(setFields) To UBound(setFields)
        Dim sf As String
        Dim st As String
        Dim sp As String

        sf = Trim$(setFields(i))
        st = Trim$(setTypes(i))

        AssertHasValue sf, "XDaoCrud.BuildUpdateQd", "setFieldsCsv contains a blank field name."
        AssertHasValue st, "XDaoCrud.BuildUpdateQd", "setTypesCsv contains a blank type."

        sp = "p" & sf

        If Len(paramsSql) > 0 Then paramsSql = paramsSql & ", "
        paramsSql = paramsSql & EscapeIdentDao(sp) & " " & st

        If Len(setSql) > 0 Then setSql = setSql & ", "
        setSql = setSql & EscapeIdentDao(sf) & " = " & EscapeIdentDao(sp)
    Next

    ' PARAMETERS + WHERE clause
    AppendWhereParts whereFieldsCsv, whereTypesCsv, "XDaoCrud.BuildUpdateQd", paramsSql, whereSql

    Dim sqlText As String
    sqlText = "PARAMETERS " & paramsSql & ";" & vbCrLf & _
              "UPDATE " & EscapeIdentDao(tableName) & vbCrLf & _
              "SET " & setSql & vbCrLf & _
              "WHERE " & whereSql & ";"

    Set BuildUpdateQd = db.CreateQueryDef("", sqlText)
End Function

' Build a parameterized DELETE QueryDef (param names = "p" + field name).
Private Function BuildDeleteQd( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal whereFieldsCsv As String, _
    ByVal whereTypesCsv As String _
) As DAO.QueryDef

    Dim paramsSql As String
    Dim whereSql As String

    AssertNotNothing db, "XDaoCrud.BuildDeleteQd", "db is Nothing."
    AssertHasValue tableName, "XDaoCrud.BuildDeleteQd", "tableName is blank."
    AssertHasValue whereFieldsCsv, "XDaoCrud.BuildDeleteQd", "whereFieldsCsv is blank."
    AssertHasValue whereTypesCsv, "XDaoCrud.BuildDeleteQd", "whereTypesCsv is blank."

    AppendWhereParts whereFieldsCsv, whereTypesCsv, "XDaoCrud.BuildDeleteQd", paramsSql, whereSql

    Dim sqlText As String
    sqlText = "PARAMETERS " & paramsSql & ";" & vbCrLf & _
              "DELETE FROM " & EscapeIdentDao(tableName) & vbCrLf & _
              "WHERE " & whereSql & ";"

    Set BuildDeleteQd = db.CreateQueryDef("", sqlText)
End Function

' Build/append WHERE clause + PARAMETERS entries from CSV field/type lists.
' - Appends to paramsSql using ", " when needed
' - Appends to whereSql using " AND " when needed
' - Uses param name convention: p + field name
Private Sub AppendWhereParts( _
    ByVal whereFieldsCsv As String, _
    ByVal whereTypesCsv As String, _
    ByVal source As String, _
    ByRef paramsSql As String, _
    ByRef whereSql As String _
)
    Dim whereFields() As String
    Dim whereTypes() As String
    Dim i As Long

    AssertHasValue whereFieldsCsv, source, "whereFieldsCsv is blank."
    AssertHasValue whereTypesCsv, source, "whereTypesCsv is blank."

    whereFields = Split(whereFieldsCsv, ",")
    whereTypes = Split(whereTypesCsv, ",")

    AssertTrue (UBound(whereFields) = UBound(whereTypes)), source, _
        "whereFieldsCsv and whereTypesCsv must have the same number of items."

    For i = LBound(whereFields) To UBound(whereFields)
        Dim wf As String
        Dim wt As String
        Dim wp As String

        wf = Trim$(whereFields(i))
        wt = Trim$(whereTypes(i))

        AssertHasValue wf, source, "whereFieldsCsv contains a blank field name."
        AssertHasValue wt, source, "whereTypesCsv contains a blank type."

        wp = "p" & wf

        If Len(paramsSql) > 0 Then paramsSql = paramsSql & ", "
        paramsSql = paramsSql & EscapeIdentDao(wp) & " " & wt

        If Len(whereSql) > 0 Then whereSql = whereSql & " AND "
        whereSql = whereSql & EscapeIdentDao(wf) & " = " & EscapeIdentDao(wp)
    Next
End Sub

' Bind QueryDef parameters from the values object (missing -> Null).
Private Sub BindParamsFromValues(ByVal qd As DAO.QueryDef, ByVal values As Object)
    Dim p As DAO.Parameter

    AssertNotNothing qd, "XDaoCrud.BindParamsFromValues", "qd is Nothing."
    AssertNotNothing values, "XDaoCrud.BindParamsFromValues", "values is Nothing."

    For Each p In qd.Parameters
        Dim fieldName As String
        fieldName = NormalizeParamName(p.Name)

        ' expected parameter naming: p + FieldName
        If LCase$(Left$(fieldName, 1)) = "p" And Len(fieldName) > 1 Then
            fieldName = Mid$(fieldName, 2)
        End If

        If HasField(values, fieldName) Then
            p.Value = values(fieldName)
        Else
            p.Value = Null
        End If
    Next
End Sub

' Normalize a parameter name by stripping surrounding brackets.
Private Function NormalizeParamName(ByVal paramName As String) As String
    AssertHasValue paramName, "XDaoCrud.NormalizeParamName", "paramName is blank."

    If Left$(paramName, 1) = "[" And Right$(paramName, 1) = "]" Then
        NormalizeParamName = Mid$(paramName, 2, Len(paramName) - 2)
        Exit Function
    End If
    NormalizeParamName = paramName
End Function
