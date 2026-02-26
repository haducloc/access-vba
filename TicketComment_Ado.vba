Option Compare Database
Option Explicit

' Load Ticket comments
Public Function LoadCommentsAdo( _
    ByVal connAdo As Object, ByVal ticketID As Long) As Object

    Dim rs As Object
    Dim cmd As Object
    Dim xe As XError

    ' Error handling (On Error GoTo)
    On Error GoTo TCError

    ' Build a parameterized SQL query with NULL-aware predicates
    Dim sql As String
    sql = "SELECT C.* " & _
          "FROM TicketComment AS C " & _
          "WHERE C.TicketID = ?"

    Set cmd = CreateCommandAdo(connAdo, sql)

    ' TicketID filter: C.TicketID = ?
    ParamInt4Ado cmd, "@p1", ticketID

    ' Execute and return the recordset
    Set rs = ExecuteQueryAdo(cmd, True)

    Set LoadCommentsAdo = rs
    Exit Function

TCError:
    ' Preserve original error, cleanup, then rethrow
    Set xe = ToXError(err)
    CloseObj rs
    err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function
