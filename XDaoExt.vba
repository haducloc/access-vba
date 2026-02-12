Option Compare Database
Option Explicit

' Create temp table if it doesn't exist.
' schemaSql example: "ReturnTrackingID LONG NOT NULL, CtName TEXT(100) NOT NULL, ReturnDate1 DATETIME"
' pkFields example: "ReturnTrackingID" OR "PK1,PK2"
Public Function EnsureTempTableAdo( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal schemaSql As String, _
    Optional ByVal pkFields As String = "" _
) As Boolean

    AssertNotNothing db, "XDaoExt.EnsureTempTableAdo", "db is Nothing."
    AssertHasValue tableName, "XDaoExt.EnsureTempTableAdo", "tableName is blank."
    AssertHasValue schemaSql, "XDaoExt.EnsureTempTableAdo", "schemaSql is blank."

    If TableExists(db, tableName) Then
        EnsureTempTableAdo = False
        Exit Function
    End If

    db.Execute "CREATE TABLE " & EscapeIdentDao(tableName) & " (" & schemaSql & ");", dbFailOnError
    db.TableDefs.Refresh

    If Len(pkFields) > 0 Then
        Dim tdf As DAO.TableDef
        Set tdf = db.TableDefs(tableName)

        AddPrimaryKeyCsv tdf, pkFields, "XDaoExt.EnsureTempTableAdo"
    End If

    EnsureTempTableAdo = True
End Function

' Create temp table from ADO RS if missing; otherwise clear it; then load data.
Public Sub EnsureTempTableFromAdoRs( _
    ByVal db As DAO.Database, _
    ByVal rs As Object, _
    ByVal tableName As String, _
    Optional ByVal pkFields As String = "" _
)
    Dim tdf As DAO.TableDef
    Dim i As Long
    Dim xe As XError

    On Error GoTo TCError

    AssertNotNothing db, "XDaoExt.EnsureTempTableFromAdoRs", "db is Nothing."
    AssertHasValue tableName, "XDaoExt.EnsureTempTableFromAdoRs", "tableName is blank."

    If rs Is Nothing Then XRaise "XDaoExt.EnsureTempTableFromAdoRs", "Recordset is Nothing."
    If rs.State = 0 Then XRaise "XDaoExt.EnsureTempTableFromAdoRs", "Recordset is closed."

    If TableExists(db, tableName) Then
        If Not TableMatchesAdoRecordset(db, tableName, rs) Then
            db.Execute "DROP TABLE " & EscapeIdentDao(tableName) & ";", dbFailOnError
            db.TableDefs.Refresh
        End If
    End If

    If TableExists(db, tableName) Then
        db.Execute "DELETE FROM " & EscapeIdentDao(tableName) & ";", dbFailOnError
    Else
        Set tdf = db.CreateTableDef(tableName)

        Dim fld As DAO.Field
        For i = 0 To rs.Fields.Count - 1
            Set fld = MapAdoFieldToDaoField(db, tdf, rs.Fields(i))
            tdf.Fields.Append fld

            ' If Decimal
            If fld.Type = dbDecimal Then
                tdf.Fields.Refresh
                InitDecimalPS fld, rs.Fields(i)
            End If
        Next i

        db.TableDefs.Append tdf
        db.TableDefs.Refresh

        If Len(pkFields) > 0 Then
            ' Re-acquire persisted TableDef before adding indexes
            Set tdf = db.TableDefs(tableName)

            AddPrimaryKeyCsv tdf, pkFields, "XDaoExt.EnsureTempTableFromAdoRs"
        End If
    End If

    ' Always load data after ensuring table exists and is empty
    InsertAdoRecordsetRows db, rs, tableName
    Exit Sub

TCError:
    Set xe = ToXError(Err)
    Err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Sub

' Clear all rows from the temp table.
Public Sub ClearTempTableDao(ByVal db As DAO.Database, ByVal tableName As String)
    AssertNotNothing db, "XDaoExt.ClearTempTableDao", "db is Nothing."
    AssertHasValue tableName, "XDaoExt.ClearTempTableDao", "tableName is blank."
    db.Execute "DELETE FROM " & EscapeIdentDao(tableName) & ";", dbFailOnError
End Sub

' Check whether a table exists in the database.
Public Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef

    AssertNotNothing db, "XDaoExt.TableExists", "db is Nothing."
    AssertHasValue tableName, "XDaoExt.TableExists", "tableName is blank."

    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next
    TableExists = False
End Function

' Create a DAO Field based on an ADO Field definition.
Private Function MapAdoFieldToDaoField( _
    ByVal db As DAO.Database, _
    ByVal tdf As DAO.TableDef, _
    ByVal adoFld As Object _
) As DAO.Field

    AssertNotNothing tdf, "XDaoExt.MapAdoFieldToDaoField", "tdf is Nothing."
    AssertNotNothing adoFld, "XDaoExt.MapAdoFieldToDaoField", "adoFld is Nothing."

    Dim daoType As Long
    daoType = MapAdoTypeToDaoType(db, adoFld.Type, adoFld.DefinedSize)

    Dim fld As DAO.Field

    Select Case daoType
        Case dbText
            Set fld = tdf.CreateField(adoFld.Name, dbText, adoFld.DefinedSize)

        Case dbBinary, dbVarBinary
            Set fld = tdf.CreateField(adoFld.Name, daoType, adoFld.DefinedSize)

        Case Else
            Set fld = tdf.CreateField(adoFld.Name, daoType)
    End Select

    Set MapAdoFieldToDaoField = fld
End Function

' Apply Precision & Scale (field must be appended first).
Private Sub InitDecimalPS(ByVal fld As DAO.Field, ByVal adoFld As Object)
    Dim p As Integer, s As Integer

    On Error Resume Next
    p = adoFld.Precision
    s = adoFld.NumericScale
    On Error GoTo 0

    ' Access supports Precision 1..28; Scale 0..28 and <= Precision

    If p > 28 Then p = 28

    ' Only apply when ADO actually provided something usable
    If p > 0 And s >= 0 And s <= p Then
        On Error Resume Next
        fld.Properties("Precision").Value = p
        fld.Properties("Scale").Value = s
        On Error GoTo 0
    End If
End Sub

' Map an ADO data type to the closest DAO data type.
' NOTE: Uses Long to avoid requiring an ADODB reference; ADO enum values are Longs.
Private Function MapAdoTypeToDaoType( _
    ByVal db As DAO.Database, _
    ByVal adoType As Long, _
    ByVal definedSize As Long _
) As Long

    Select Case adoType
        ' Numbers
        Case adUnsignedTinyInt
            MapAdoTypeToDaoType = dbByte

        Case adTinyInt
            ' DAO has no signed 1-byte integer; use Integer to preserve negatives
            MapAdoTypeToDaoType = dbInteger

        Case adSmallInt
            MapAdoTypeToDaoType = dbInteger

        Case adInteger
            MapAdoTypeToDaoType = dbLong

        Case adBigInt
            If SupportsDaoBigInt(db) Then
                MapAdoTypeToDaoType = dbBigInt
            Else
                MapAdoTypeToDaoType = dbDecimal
            End If

        Case adSingle
            MapAdoTypeToDaoType = dbSingle

        Case adDouble
            MapAdoTypeToDaoType = dbDouble

        Case adCurrency
            MapAdoTypeToDaoType = dbCurrency

        Case adDecimal, adNumeric
            MapAdoTypeToDaoType = dbDecimal

        ' Boolean
        Case adBoolean
            MapAdoTypeToDaoType = dbBoolean

        ' Date / time
        Case adDBDate, adDBTime, adDBTimeStamp
            MapAdoTypeToDaoType = dbDate

        ' Identifiers
        Case adGUID
            MapAdoTypeToDaoType = dbGUID

        ' Text
        Case adChar, adVarChar, adWChar, adVarWChar
            If definedSize <= 0 Or definedSize > 255 Then
                MapAdoTypeToDaoType = dbMemo
            Else
                MapAdoTypeToDaoType = dbText
            End If

        Case adLongVarChar, adLongVarWChar
            MapAdoTypeToDaoType = dbMemo

        ' Binary
        Case adBinary, adVarBinary
            If definedSize <= 0 Or definedSize > 255 Then
                MapAdoTypeToDaoType = dbLongBinary
            Else
                If adoType = adVarBinary Then
                    MapAdoTypeToDaoType = dbVarBinary
                Else
                    MapAdoTypeToDaoType = dbBinary
                End If
            End If

        Case adLongVarBinary
            MapAdoTypeToDaoType = dbLongBinary

        Case Else
            MapAdoTypeToDaoType = dbText

    End Select
End Function

' Detect whether the current Access version supports DAO BigInt.
Private Function SupportsDaoBigInt(ByVal db As DAO.Database) As Boolean
    On Error GoTo TCError
    Dim tdf As DAO.TableDef
    Dim f As DAO.Field

    Set tdf = db.CreateTableDef("")
    Set f = tdf.CreateField("x", dbBigInt)
    SupportsDaoBigInt = True
    Exit Function
TCError:
    SupportsDaoBigInt = False
End Function

' Add a (single or composite) primary key to a TableDef.
' pkFieldsCsv example: "ReturnTrackingID" OR "PK1,PK2"
Private Sub AddPrimaryKeyCsv(ByRef tdf As DAO.TableDef, ByVal pkFieldsCsv As String, ByVal source As String)
    Dim idx As DAO.Index

    AssertNotNothing tdf, "XDaoExt.AddPrimaryKeyCsv", "tdf is Nothing."
    AssertHasValue pkFieldsCsv, "XDaoExt.AddPrimaryKeyCsv", "pkFieldsCsv is blank."
    AssertHasValue source, "XDaoExt.AddPrimaryKeyCsv", "source is blank."

    Dim parts() As String
    Dim i As Long
    Dim f As String

    parts = Split(pkFieldsCsv, ",")

    ' Validate all PK fields exist before creating the index
    For i = LBound(parts) To UBound(parts)
        f = Trim$(parts(i))
        AssertHasValue f, source, "pkFields contains a blank field name."

        If Not FieldExistsInTdf(tdf, f) Then
            XRaise source, "PK field not found: " & f
        End If
    Next i

    Set idx = tdf.CreateIndex("PK_" & tdf.Name)
    With idx
        .Primary = True
        .Unique = True

        For i = LBound(parts) To UBound(parts)
            f = Trim$(parts(i))
            .Fields.Append .CreateField(f)
        Next i
    End With

    tdf.Indexes.Append idx
End Sub

' Check whether a field exists in a TableDef.
Private Function FieldExistsInTdf(ByVal tdf As DAO.TableDef, ByVal fieldName As String) As Boolean
    On Error GoTo TCError

    AssertNotNothing tdf, "XDaoExt.FieldExistsInTdf", "tdf is Nothing."
    AssertHasValue fieldName, "XDaoExt.FieldExistsInTdf", "fieldName is blank."

    Dim f As DAO.Field
    Set f = tdf.Fields(fieldName)
    FieldExistsInTdf = True
    Exit Function
TCError:
    FieldExistsInTdf = False
End Function

' Final Production Version for LOCAL Tables
Private Sub InsertAdoRecordsetRows(ByVal db As DAO.Database, ByVal rs As Object, ByVal tableName As String)
    Dim rsDest As DAO.Recordset
    Dim i As Long, rowNum As Long
    Dim inTrans As Boolean
    Dim xe As XError

    Dim destFlds() As DAO.Field
    Dim srcFlds() As Object
    Dim validColCount As Long

    Dim ws As DAO.Workspace

    On Error GoTo TCError

    AssertNotNothing db, "XDaoExt.InsertAdoRecordsetRows", "db is Nothing."
    AssertHasValue tableName, "XDaoExt.InsertAdoRecordsetRows", "tableName is blank."
    If rs Is Nothing Then XRaise "XDaoExt.InsertAdoRecordsetRows", "Recordset is Nothing."
    If rs.State = 0 Then XRaise "XDaoExt.InsertAdoRecordsetRows", "Recordset is closed."

    Set ws = GetWorkspace(db)
    Set rsDest = db.OpenRecordset(tableName, dbOpenTable, dbAppendOnly)

    rowNum = 0

    ' PRE-MAPPING PHASE
    Dim fieldLimit As Long: fieldLimit = rs.Fields.Count - 1
    ReDim destFlds(fieldLimit)
    ReDim srcFlds(fieldLimit)
    validColCount = 0

    Dim tdf As DAO.TableDef: Set tdf = db.TableDefs(tableName)

    For i = 0 To fieldLimit
        Dim fn As String: fn = rs.Fields(i).Name
        If FieldExistsInTdf(tdf, fn) Then
            Dim dFld As DAO.Field: Set dFld = rsDest.Fields(fn)
            ' Skip Autonumber
            If (dFld.Attributes And dbAutoIncrField) = 0 Then
                Set destFlds(validColCount) = dFld
                Set srcFlds(validColCount) = rs.Fields(i)
                validColCount = validColCount + 1
            End If
        End If
    Next i

    If validColCount = 0 Then
        CloseObj rsDest
        Exit Sub
    End If

    ' EXECUTION PHASE
    ws.BeginTrans
    inTrans = True

    If Not (rs.BOF And rs.EOF) Then
        ' MoveFirst can fail on some ADO forward-only cursors.
        ' Only call if supported.
        If rs.Supports(adMovePrevious) Then rs.MoveFirst

        Do While Not rs.EOF
            rowNum = rowNum + 1

            rsDest.AddNew
            For i = 0 To validColCount - 1
                ' Direct value assignment from cached objects
                destFlds(i).Value = srcFlds(i).Value
            Next i
            rsDest.Update

            ' Batch Commit (Every 5k rows)
            If rowNum Mod 5000 = 0 Then
                ws.CommitTrans
                ws.BeginTrans
            End If

            rs.MoveNext
        Loop
    End If

    If inTrans Then
        ws.CommitTrans
        inTrans = False
    End If

    CloseObj rsDest
    Exit Sub

TCError:
    Set xe = ToXError(Err)
    On Error Resume Next
    If inTrans Then
        If Not ws Is Nothing Then
            ws.Rollback
        End If
    End If
    On Error GoTo 0
    CloseObj rsDest
    Err.Raise xe.ErrNum, xe.ErrSrc, "Transfer failed at row " & rowNum & ": " & xe.ErrDesc
End Sub

' FIX: Schema drift check (minimal). Table must contain all ADO recordset field names.
Private Function TableMatchesAdoRecordset(ByVal db As DAO.Database, ByVal tableName As String, ByVal rs As Object) As Boolean
    On Error GoTo TCError

    Dim tdf As DAO.TableDef
    Dim i As Long

    AssertNotNothing db, "XDaoExt.TableMatchesAdoRecordset", "db is Nothing."
    AssertHasValue tableName, "XDaoExt.TableMatchesAdoRecordset", "tableName is blank."
    If rs Is Nothing Then XRaise "XDaoExt.TableMatchesAdoRecordset", "Recordset is Nothing."
    If rs.State = 0 Then XRaise "XDaoExt.TableMatchesAdoRecordset", "Recordset is closed."

    Set tdf = db.TableDefs(tableName)

    For i = 0 To rs.Fields.Count - 1
        If Not FieldExistsInTdf(tdf, rs.Fields(i).Name) Then
            TableMatchesAdoRecordset = False
            Exit Function
        End If
    Next i

    TableMatchesAdoRecordset = True
    Exit Function

TCError:
    TableMatchesAdoRecordset = False
End Function
