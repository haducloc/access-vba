Option Compare Database
Option Explicit

' Module-level cached RegExp objects for validating type spec
Private mReVarchar  As Object
Private mReNvarchar As Object
Private mReChar     As Object
Private mReNChar    As Object
Private mReDecimal  As Object
Private mReNumeric  As Object
Private mReBinary   As Object
Private mReVarbinary As Object

' Insert one row using parameter names = "p" + field name.
' fieldsCsv example: "Id,CtName,ReturnDate1"
' typesCsv example:  "INT4,VARCHAR(100),DATETIME"
' values: dictionary-like object
Public Sub InsertRowAdo( _
    ByVal cn As Object, _
    ByVal tableName As String, _
    ByVal fieldsCsv As String, _
    ByVal typesCsv As String, _
    ByVal values As Object _
)
    Dim cmd As Object
    Dim dbType As XDbType

    AssertNotNothing cn, "XAdoCrud.InsertRowAdo", "cn is Nothing."
    AssertHasValue tableName, "XAdoCrud.InsertRowAdo", "tableName is blank."
    AssertHasValue fieldsCsv, "XAdoCrud.InsertRowAdo", "fieldsCsv is blank."
    AssertHasValue typesCsv, "XAdoCrud.InsertRowAdo", "typesCsv is blank."
    AssertNotNothing values, "XAdoCrud.InsertRowAdo", "values is Nothing."

    dbType = GetDbType(cn)
    Set cmd = CreateCommandAdo(cn, BuildInsertSqlAdo(dbType, tableName, fieldsCsv))

    ' Params must be appended in the same order as fieldsCsv
    AppendParamsFromValuesAdo cmd, fieldsCsv, typesCsv, values

    ExecuteUpdateAdo cmd
End Sub

' Update rows using parameter names = "p" + field name.
' setFieldsCsv example:   "CtName,ReturnDate1"
' setTypesCsv example:    "VARCHAR(100),DATETIME"
' whereFieldsCsv example: "Id"
' whereTypesCsv example:  "INT4"
' values: dictionary-like object (should include set + where fields; missing -> Null)
Public Sub UpdateRowAdo( _
    ByVal cn As Object, _
    ByVal tableName As String, _
    ByVal setFieldsCsv As String, _
    ByVal setTypesCsv As String, _
    ByVal whereFieldsCsv As String, _
    ByVal whereTypesCsv As String, _
    ByVal values As Object _
)
    Dim cmd As Object
    Dim dbType As XDbType

    AssertNotNothing cn, "XAdoCrud.UpdateRowAdo", "cn is Nothing."
    AssertHasValue tableName, "XAdoCrud.UpdateRowAdo", "tableName is blank."
    AssertHasValue setFieldsCsv, "XAdoCrud.UpdateRowAdo", "setFieldsCsv is blank."
    AssertHasValue setTypesCsv, "XAdoCrud.UpdateRowAdo", "setTypesCsv is blank."
    AssertHasValue whereFieldsCsv, "XAdoCrud.UpdateRowAdo", "whereFieldsCsv is blank."
    AssertHasValue whereTypesCsv, "XAdoCrud.UpdateRowAdo", "whereTypesCsv is blank."
    AssertNotNothing values, "XAdoCrud.UpdateRowAdo", "values is Nothing."

    dbType = GetDbType(cn)
    Set cmd = CreateCommandAdo(cn, BuildUpdateSqlAdo(dbType, tableName, setFieldsCsv, whereFieldsCsv))

    ' SET params first, WHERE params second (order matters)
    AppendParamsFromValuesAdo cmd, setFieldsCsv, setTypesCsv, values
    AppendParamsFromValuesAdo cmd, whereFieldsCsv, whereTypesCsv, values

    ExecuteUpdateAdo cmd
End Sub

' Delete rows using parameter names = "p" + field name.
' whereFieldsCsv example: "Id"
' whereTypesCsv example:  "INT4"
' values: dictionary-like object (missing -> Null)
Public Sub DeleteRowAdo( _
    ByVal cn As Object, _
    ByVal tableName As String, _
    ByVal whereFieldsCsv As String, _
    ByVal whereTypesCsv As String, _
    ByVal values As Object _
)
    Dim cmd As Object
    Dim dbType As XDbType

    AssertNotNothing cn, "XAdoCrud.DeleteRowAdo", "cn is Nothing."
    AssertHasValue tableName, "XAdoCrud.DeleteRowAdo", "tableName is blank."
    AssertHasValue whereFieldsCsv, "XAdoCrud.DeleteRowAdo", "whereFieldsCsv is blank."
    AssertHasValue whereTypesCsv, "XAdoCrud.DeleteRowAdo", "whereTypesCsv is blank."
    AssertNotNothing values, "XAdoCrud.DeleteRowAdo", "values is Nothing."

    dbType = GetDbType(cn)
    Set cmd = CreateCommandAdo(cn, BuildDeleteSqlAdo(dbType, tableName, whereFieldsCsv))

    ' WHERE params only (order matters)
    AppendParamsFromValuesAdo cmd, whereFieldsCsv, whereTypesCsv, values

    ExecuteUpdateAdo cmd
End Sub

' Get single row by PK as Dictionary (param name = "p" + field name).
' pkFieldsCsv example: "Id"
' pkTypesCsv example:  "INT4"
' values: dictionary-like object containing pk fields (missing -> Null)
Public Function GetRowByPkAdo( _
    ByVal cn As Object, _
    ByVal tableName As String, _
    ByVal pkFieldsCsv As String, _
    ByVal pkTypesCsv As String, _
    ByVal values As Object _
) As Object

    AssertNotNothing cn, "XAdoCrud.GetRowByPkAdo", "cn is Nothing."
    AssertHasValue tableName, "XAdoCrud.GetRowByPkAdo", "tableName is blank."
    AssertHasValue pkFieldsCsv, "XAdoCrud.GetRowByPkAdo", "pkFieldsCsv is blank."
    AssertHasValue pkTypesCsv, "XAdoCrud.GetRowByPkAdo", "pkTypesCsv is blank."
    AssertNotNothing values, "XAdoCrud.GetRowByPkAdo", "values is Nothing."

    Dim cmd As Object
    Dim rs As Object
    Dim result As Object
    Dim xe As XError
    Dim dbType As XDbType

    On Error GoTo TCError
    dbType = GetDbType(cn)
    Set cmd = CreateCommandAdo(cn, BuildSelectByPkSqlAdo(dbType, tableName, pkFieldsCsv))

    ' WHERE params in PK order (order matters)
    AppendParamsFromValuesAdo cmd, pkFieldsCsv, pkTypesCsv, values

    Set rs = ExecuteQueryAdo(cmd, True)

    If (rs Is Nothing) Or (rs.BOF And rs.EOF) Then
        Set result = Nothing
    Else
        Set result = RecordToDictAdo(rs)
    End If

    CloseObj rs
    Set GetRowByPkAdo = result
    Exit Function

TCError:
    Set xe = ToXError(Err)
    CloseObj rs
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Returns True if a row exists for the given PK definition + values (param name = "p" + field name).
' pkFieldsCsv example: "Id"
' pkTypesCsv example:  "INT4"
' values: dictionary-like object containing pk fields (missing -> Null)
Public Function ExistsByPkAdo( _
    ByVal cn As Object, _
    ByVal tableName As String, _
    ByVal pkFieldsCsv As String, _
    ByVal pkTypesCsv As String, _
    ByVal values As Object _
) As Boolean

    AssertNotNothing cn, "XAdoCrud.ExistsByPkAdo", "cn is Nothing."
    AssertHasValue tableName, "XAdoCrud.ExistsByPkAdo", "tableName is blank."
    AssertHasValue pkFieldsCsv, "XAdoCrud.ExistsByPkAdo", "pkFieldsCsv is blank."
    AssertHasValue pkTypesCsv, "XAdoCrud.ExistsByPkAdo", "pkTypesCsv is blank."
    AssertNotNothing values, "XAdoCrud.ExistsByPkAdo", "values is Nothing."

    Dim cmd As Object
    Dim rs As Object
    Dim xe As XError
    Dim dbType As XDbType

    On Error GoTo TCError
    dbType = GetDbType(cn)
    Set cmd = CreateCommandAdo(cn, BuildExistsByPkSqlAdo(dbType, tableName, pkFieldsCsv))

    ' WHERE params in PK order (order matters)
    AppendParamsFromValuesAdo cmd, pkFieldsCsv, pkTypesCsv, values

    Set rs = ExecuteQueryAdo(cmd, True)

    ' Exists if at least one row returned
    ExistsByPkAdo = Not ((rs Is Nothing) Or (rs.BOF And rs.EOF))

    CloseObj rs
    Exit Function

TCError:
    Set xe = ToXError(Err)
    CloseObj rs
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function

' Build INSERT SQL with positional ? placeholders matching fieldsCsv order.
Private Function BuildInsertSqlAdo(ByVal dbType As XDbType, ByVal tableName As String, ByVal fieldsCsv As String) As String
    Dim fields() As String
    Dim colsSql As String
    Dim valsSql As String
    Dim i As Long
    Dim f As String

    AssertHasValue tableName, "XAdoCrud.BuildInsertSqlAdo", "tableName is blank."
    AssertHasValue fieldsCsv, "XAdoCrud.BuildInsertSqlAdo", "fieldsCsv is blank."

    fields = Split(fieldsCsv, ",")

    For i = LBound(fields) To UBound(fields)
        f = Trim$(fields(i))
        AssertHasValue f, "XAdoCrud.BuildInsertSqlAdo", "fieldsCsv contains a blank field name."

        If Len(colsSql) > 0 Then colsSql = colsSql & ", "
        colsSql = colsSql & EscapeIdentAdo(dbType, f)

        If Len(valsSql) > 0 Then valsSql = valsSql & ", "
        valsSql = valsSql & "?"
    Next i

    BuildInsertSqlAdo = "INSERT INTO " & EscapeIdentAdo(dbType, tableName) & " (" & colsSql & ") VALUES (" & valsSql & ");"
End Function

' Build UPDATE SQL with positional ? placeholders: SET fields first, then WHERE fields.
Private Function BuildUpdateSqlAdo(ByVal dbType As XDbType, ByVal tableName As String, ByVal setFieldsCsv As String, ByVal whereFieldsCsv As String) As String
    Dim setFields() As String
    Dim whereFields() As String
    Dim setSql As String
    Dim whereSql As String
    Dim i As Long
    Dim f As String

    AssertHasValue tableName, "XAdoCrud.BuildUpdateSqlAdo", "tableName is blank."
    AssertHasValue setFieldsCsv, "XAdoCrud.BuildUpdateSqlAdo", "setFieldsCsv is blank."
    AssertHasValue whereFieldsCsv, "XAdoCrud.BuildUpdateSqlAdo", "whereFieldsCsv is blank."

    setFields = Split(setFieldsCsv, ",")
    whereFields = Split(whereFieldsCsv, ",")

    For i = LBound(setFields) To UBound(setFields)
        f = Trim$(setFields(i))
        AssertHasValue f, "XAdoCrud.BuildUpdateSqlAdo", "setFieldsCsv contains a blank field name."

        If Len(setSql) > 0 Then setSql = setSql & ", "
        setSql = setSql & EscapeIdentAdo(dbType, f) & " = ?"
    Next i

    For i = LBound(whereFields) To UBound(whereFields)
        f = Trim$(whereFields(i))
        AssertHasValue f, "XAdoCrud.BuildUpdateSqlAdo", "whereFieldsCsv contains a blank field name."

        If Len(whereSql) > 0 Then whereSql = whereSql & " AND "
        whereSql = whereSql & EscapeIdentAdo(dbType, f) & " = ?"
    Next i

    BuildUpdateSqlAdo = "UPDATE " & EscapeIdentAdo(dbType, tableName) & " SET " & setSql & " WHERE " & whereSql & ";"
End Function

' Build DELETE SQL with positional ? placeholders matching whereFieldsCsv order.
Private Function BuildDeleteSqlAdo(ByVal dbType As XDbType, ByVal tableName As String, ByVal whereFieldsCsv As String) As String
    Dim whereFields() As String
    Dim whereSql As String
    Dim i As Long
    Dim f As String

    AssertHasValue tableName, "XAdoCrud.BuildDeleteSqlAdo", "tableName is blank."
    AssertHasValue whereFieldsCsv, "XAdoCrud.BuildDeleteSqlAdo", "whereFieldsCsv is blank."

    whereFields = Split(whereFieldsCsv, ",")

    For i = LBound(whereFields) To UBound(whereFields)
        f = Trim$(whereFields(i))
        AssertHasValue f, "XAdoCrud.BuildDeleteSqlAdo", "whereFieldsCsv contains a blank field name."

        If Len(whereSql) > 0 Then whereSql = whereSql & " AND "
        whereSql = whereSql & EscapeIdentAdo(dbType, f) & " = ?"
    Next i

    BuildDeleteSqlAdo = "DELETE FROM " & EscapeIdentAdo(dbType, tableName) & " WHERE " & whereSql & ";"
End Function

' Build SELECT * by PK SQL with positional ? placeholders matching pkFieldsCsv order.
Private Function BuildSelectByPkSqlAdo(ByVal dbType As XDbType, ByVal tableName As String, ByVal pkFieldsCsv As String) As String
    Dim pkFields() As String
    Dim whereSql As String
    Dim i As Long
    Dim f As String

    AssertHasValue tableName, "XAdoCrud.BuildSelectByPkSqlAdo", "tableName is blank."
    AssertHasValue pkFieldsCsv, "XAdoCrud.BuildSelectByPkSqlAdo", "pkFieldsCsv is blank."

    pkFields = Split(pkFieldsCsv, ",")

    For i = LBound(pkFields) To UBound(pkFields)
        f = Trim$(pkFields(i))
        AssertHasValue f, "XAdoCrud.BuildSelectByPkSqlAdo", "pkFieldsCsv contains a blank field name."

        If Len(whereSql) > 0 Then whereSql = whereSql & " AND "
        whereSql = whereSql & EscapeIdentAdo(dbType, f) & " = ?"
    Next i

    BuildSelectByPkSqlAdo = "SELECT * FROM " & EscapeIdentAdo(dbType, tableName) & " WHERE " & whereSql & ";"
End Function

' Build SELECT 1 by PK SQL with positional ? placeholders matching pkFieldsCsv order.
Private Function BuildExistsByPkSqlAdo(ByVal dbType As XDbType, ByVal tableName As String, ByVal pkFieldsCsv As String) As String
    Dim pkFields() As String
    Dim whereSql As String
    Dim i As Long
    Dim f As String

    AssertHasValue tableName, "XAdoCrud.BuildExistsByPkSqlAdo", "tableName is blank."
    AssertHasValue pkFieldsCsv, "XAdoCrud.BuildExistsByPkSqlAdo", "pkFieldsCsv is blank."

    pkFields = Split(pkFieldsCsv, ",")

    For i = LBound(pkFields) To UBound(pkFields)
        f = Trim$(pkFields(i))
        AssertHasValue f, "XAdoCrud.BuildExistsByPkSqlAdo", "pkFieldsCsv contains a blank field name."

        If Len(whereSql) > 0 Then whereSql = whereSql & " AND "
        whereSql = whereSql & EscapeIdentAdo(dbType, f) & " = ?"
    Next i

    Select Case dbType
        Case Db_SQLServer
            BuildExistsByPkSqlAdo = _
                "SELECT TOP 1 1 FROM " & EscapeIdentAdo(dbType, tableName) & " WHERE " & whereSql

        Case Db_Access
            BuildExistsByPkSqlAdo = _
                "SELECT TOP 1 1 FROM " & EscapeIdentAdo(dbType, tableName) & " WHERE " & whereSql

        Case Db_Oracle
            BuildExistsByPkSqlAdo = _
                "SELECT 1 FROM " & EscapeIdentAdo(dbType, tableName) & " WHERE " & whereSql & " AND ROWNUM = 1"

        Case Db_Postgres
            BuildExistsByPkSqlAdo = _
                "SELECT 1 FROM " & EscapeIdentAdo(dbType, tableName) & " WHERE " & whereSql & " LIMIT 1"

        Case Db_MySQL
            BuildExistsByPkSqlAdo = _
                "SELECT 1 FROM " & EscapeIdentAdo(dbType, tableName) & " WHERE " & whereSql & " LIMIT 1"

        Case Else
            XRaise "XAdoCrud.BuildExistsByPkSqlAdo", "Unsupported dbType: " & CStr(dbType)
    End Select
End Function

' Append parameters in fieldsCsv order, pairing each field with its typesCsv item; missing values -> Null.
Private Sub AppendParamsFromValuesAdo(ByVal cmd As Object, ByVal fieldsCsv As String, ByVal typesCsv As String, ByVal values As Object)
    Dim fields() As String
    Dim types() As String
    Dim i As Long

    AssertNotNothing cmd, "XAdoCrud.AppendParamsFromValuesAdo", "cmd is Nothing."
    AssertHasValue fieldsCsv, "XAdoCrud.AppendParamsFromValuesAdo", "fieldsCsv is blank."
    AssertHasValue typesCsv, "XAdoCrud.AppendParamsFromValuesAdo", "typesCsv is blank."
    AssertNotNothing values, "XAdoCrud.AppendParamsFromValuesAdo", "values is Nothing."

    fields = Split(fieldsCsv, ",")
    types = Split(typesCsv, ",")

    AssertTrue (UBound(fields) = UBound(types)), "XAdoCrud.AppendParamsFromValuesAdo", _
        "fieldsCsv and typesCsv must have the same number of items."

    For i = LBound(fields) To UBound(fields)
        Dim f As String: f = Trim$(fields(i))
        Dim t As String: t = UCase$(Trim$(types(i)))
        Dim v As Variant

        AssertHasValue f, "XAdoCrud.AppendParamsFromValuesAdo", "fieldsCsv contains a blank field name."
        AssertHasValue t, "XAdoCrud.AppendParamsFromValuesAdo", "typesCsv contains a blank type."

        If HasField(values, f) Then
            v = values(f)
        Else
            v = Null
        End If

        AppendParamAdo cmd, "p" & f, t, v
    Next i
End Sub

' Append one ADO parameter based on typeSpec (INT2/INT4/INT8, VARCHAR/NVARCHAR/CHAR/NCHAR, DATE/TIME/DATETIME/TIMESTAMP, BOOL/BIT, numeric, decimal).
Private Sub AppendParamAdo(ByVal cmd As Object, ByVal paramName As String, ByVal typeSpec As String, ByVal value As Variant)
    Dim sz As Long

    AssertNotNothing cmd, "XAdoCrud.AppendParamAdo", "cmd is Nothing."
    AssertHasValue paramName, "XAdoCrud.AppendParamAdo", "paramName is blank."
    AssertHasValue typeSpec, "XAdoCrud.AppendParamAdo", "typeSpec is blank."

    ' Disallow Access-style aliases here
    If typeSpec = "MEMO" Or typeSpec = "LONGTEXT" Then
        XRaise "XAdoCrud.AppendParamAdo", typeSpec & " is invalid type spec."
    End If

    ' VARCHAR(n|MAX)
    If InitReVarchar().Test(typeSpec) Then
        sz = ParseTypeSize(typeSpec)
        AssertTrue sz = -1 Or sz > 0, "XAdoCrud.AppendParamAdo", typeSpec & " is invalid type spec."
        If sz = -1 Then
            ParamVarcharMaxAdo cmd, paramName, value
        Else
            ParamVarcharAdo cmd, paramName, value, sz
        End If
        Exit Sub
    End If

    ' NVARCHAR(n|MAX)
    If InitReNVarchar().Test(typeSpec) Then
        sz = ParseTypeSize(typeSpec)
        AssertTrue sz = -1 Or sz > 0, "XAdoCrud.AppendParamAdo", typeSpec & " is invalid type spec."
        If sz = -1 Then
            ParamNVarcharMaxAdo cmd, paramName, value
        Else
            ParamNVarcharAdo cmd, paramName, value, sz
        End If
        Exit Sub
    End If

    ' CHAR(n)  (MAX not allowed)
    If InitReChar().Test(typeSpec) Then
        sz = ParseTypeSize(typeSpec)
        AssertTrue sz > 0, "XAdoCrud.AppendParamAdo", typeSpec & " is invalid type spec."
        ParamCharAdo cmd, paramName, value, sz
        Exit Sub
    End If

    ' NCHAR(n) (MAX not allowed)
    If InitReNChar().Test(typeSpec) Then
        sz = ParseTypeSize(typeSpec)
        AssertTrue sz > 0, "XAdoCrud.AppendParamAdo", typeSpec & " is invalid type spec."
        ParamNCharAdo cmd, paramName, value, sz
        Exit Sub
    End If

    If typeSpec = "DATETIME" Or typeSpec = "TIMESTAMP" Then
        ParamDateTimeAdo cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "DATE" Then
        ParamDateAdo cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "TIME" Then
        ParamTimeAdo cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "BOOL" Or typeSpec = "BOOLEAN" Or typeSpec = "BIT" Then
        ParamBoolAdo cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "BYTE" Then
        ParamByteAdo cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "UBYTE" Then
        ParamUByteAdo cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "INT2" Or typeSpec = "SHORT" Then
        ParamInt2Ado cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "INT4" Or typeSpec = "INTEGER" Then
        ParamInt4Ado cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "INT8" Or typeSpec = "BIGINT" Then
        ParamInt8Ado cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "FLOAT" Or typeSpec = "SINGLE" Then
        ParamFloatAdo cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "DOUBLE" Then
        ParamDoubleAdo cmd, paramName, value
        Exit Sub
    End If

    If typeSpec = "CURRENCY" Or typeSpec = "MONEY" Then
        ParamCurrencyAdo cmd, paramName, value
        Exit Sub
    End If

    ' DECIMAL or NUMERIC
    Dim ps As Variant

    If InitReDecimal().Test(typeSpec) Then
        ps = ParseDecimalPS(typeSpec)
        AssertHasValue ps, "XAdoCrud.AppendParamAdo", typeSpec & " is invalid type spec."

        ParamDecimalAdo cmd, paramName, value, ps(0), ps(1)
        Exit Sub
    End If

    If InitReNumeric().Test(typeSpec) Then
        ps = ParseDecimalPS(typeSpec)
        AssertHasValue ps, "XAdoCrud.AppendParamAdo", typeSpec & " is invalid type spec."

        ParamDecimalAdo cmd, paramName, value, ps(0), ps(1)
        Exit Sub
    End If

    ' BINARY(n) (MAX not allowed)
    If InitReBinary().Test(typeSpec) Then
        sz = ParseTypeSize(typeSpec)
        AssertTrue sz > 0, "XAdoCrud.AppendParamAdo", typeSpec & " is invalid type spec."
        ParamBinaryAdo cmd, paramName, value, sz
        Exit Sub
    End If

    ' VARBINARY(n)
    If InitReVarbinary().Test(typeSpec) Then
        sz = ParseTypeSize(typeSpec)
        AssertTrue sz = -1 Or sz > 0, "XAdoCrud.AppendParamAdo", typeSpec & " is invalid type spec."
        If sz = -1 Then
            ParamVarBinaryMaxAdo cmd, paramName, value
        Else
            ParamVarBinaryAdo cmd, paramName, value, sz
        End If
        Exit Sub
    End If

    XRaise "XAdoCrud.AppendParamAdo", typeSpec & " is invalid type spec."
End Sub

' Assumes typeSpec is validated and in valid format.
' Returns: -1 for MAX, >0 for numeric size, -2 for invalid.
Private Function ParseTypeSize(ByVal typeSpec As String) As Long
    Dim p1 As Long, p2 As Long
    Dim inner As String
    Dim n As Long

    Dim size As Long
    size = -2

    p1 = InStr(1, typeSpec, "(", vbTextCompare)
    p2 = InStrRev(typeSpec, ")")

    inner = Trim$(Mid$(typeSpec, p1 + 1, p2 - p1 - 1))

    If inner = "MAX" Then
        size = -1
    Else
        size = CLng(inner)
        If size < 1 Then size = -2
    End If

    ParseTypeSize = size
End Function

' Assumes typeSpec is validated and in valid format.
' Returns Array(CByte(p), CByte(s)) or Null if invalid per your rules.
Private Function ParseDecimalPS(ByVal typeSpec As String) As Variant
    Dim p1 As Long, p2 As Long, commaPos As Long
    Dim pStr As String, sStr As String
    Dim pVal As Long, sVal As Long

    p1 = InStr(1, typeSpec, "(", vbTextCompare)
    p2 = InStrRev(typeSpec, ")")

    commaPos = InStr(p1 + 1, typeSpec, ",", vbTextCompare)

    pStr = Trim$(Mid$(typeSpec, p1 + 1, commaPos - p1 - 1))
    sStr = Trim$(Mid$(typeSpec, commaPos + 1, p2 - commaPos - 1))

    pVal = CLng(pStr)
    sVal = CLng(sStr)

    If pVal < 1 Or pVal > 38 Then
        ParseDecimalPS = Null
        Exit Function
    End If

    If sVal < 0 Or sVal > pVal Then
        ParseDecimalPS = Null
        Exit Function
    End If

    ParseDecimalPS = Array(CByte(pVal), CByte(sVal))
End Function

Private Function InitReVarchar() As Object
    If mReVarchar Is Nothing Then
        Set mReVarchar = NewRegEx("^VARCHAR\s*\(\s*(\d+|MAX)\s*\)$", True, False, False)
    End If
    Set InitReVarchar = mReVarchar
End Function

Private Function InitReNVarchar() As Object
    If mReNvarchar Is Nothing Then
        Set mReNvarchar = NewRegEx("^NVARCHAR\s*\(\s*(\d+|MAX)\s*\)$", True, False, False)
    End If
    Set InitReNVarchar = mReNvarchar
End Function

Private Function InitReChar() As Object
    If mReChar Is Nothing Then
        ' MAX not allowed
        Set mReChar = NewRegEx("^CHAR\s*\(\s*(\d+)\s*\)$", True, False, False)
    End If
    Set InitReChar = mReChar
End Function

Private Function InitReNChar() As Object
    If mReNChar Is Nothing Then
        ' MAX not allowed
        Set mReNChar = NewRegEx("^NCHAR\s*\(\s*(\d+)\s*\)$", True, False, False)
    End If
    Set InitReNChar = mReNChar
End Function

Private Function InitReDecimal() As Object
    If mReDecimal Is Nothing Then
        Set mReDecimal = NewRegEx("^DECIMAL\s*\(\s*(\d+)\s*,\s*(\d+)\s*\)$", True, False, False)
    End If
    Set InitReDecimal = mReDecimal
End Function

Private Function InitReNumeric() As Object
    If mReNumeric Is Nothing Then
        Set mReNumeric = NewRegEx("^NUMERIC\s*\(\s*(\d+)\s*,\s*(\d+)\s*\)$", True, False, False)
    End If
    Set InitReNumeric = mReNumeric
End Function

Private Function InitReBinary() As Object
    If mReBinary Is Nothing Then
        ' MAX not allowed
        Set mReBinary = NewRegEx("^BINARY\s*\(\s*(\d+)\s*\)$", True, False, False)
    End If
    Set InitReBinary = mReBinary
End Function

Private Function InitReVarbinary() As Object
    If mReVarbinary Is Nothing Then
        Set mReVarbinary = NewRegEx("^VARBINARY\s*\(\s*(\d+|MAX)\s*\)$", True, False, False)
    End If
    Set InitReVarbinary = mReVarbinary
End Function
