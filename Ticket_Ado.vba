Option Compare Database
Option Explicit

' Search Ticket records using optional TicketID and Name filters (NULL-aware predicates)
Public Function SearchTicketAdo( _
    ByVal connAdo As Object, ByVal ticketID As Variant, ByVal name As Variant) As Object

    Dim rs As Object
    Dim cmd As Object
    Dim xe As XError

    ' Error handling (On Error GoTo)
    On Error GoTo TCError

    ' Build a parameterized SQL query with NULL-aware predicates
    Dim sql As String
    sql = "SELECT TK.*, TP.Name AS TicketTypeName " & _
          "FROM Ticket AS TK " & _
            "INNER JOIN TicketType AS TP ON TP.TicketTypeID = TK.TicketTypeID " & _
            "WHERE (? IS NULL OR TK.TicketID = ?) " & _
            "AND (? IS NULL OR TK.Name LIKE ?)"

    Set cmd = CreateCommandAdo(connAdo, sql)

    ' TicketID filter: (? IS NULL OR TK.TicketID = ?)
    ParamInt4Ado cmd, "@p1", ticketID
    ParamInt4Ado cmd, "@p2", ticketID

    ' Name filter: (? IS NULL OR TK.Name LIKE ?)
    ' Name column is VARCHAR(100); size 255 is sufficient for the LIKE parameter
    ParamLikeAdo cmd, "@p3", name, 255, Db_SQLServer
    ParamLikeAdo cmd, "@p4", name, 255, Db_SQLServer

    ' Execute and return the recordset (client-side cursor / disconnected if True is supported by ExecuteQueryAdo)
    Set rs = ExecuteQueryAdo(cmd, True)

    Set SearchTicketAdo = rs
    Exit Function

TCError:
    ' Preserve original error, cleanup, then rethrow
    Set xe = ToXError(err)
    CloseObj rs
    err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function
