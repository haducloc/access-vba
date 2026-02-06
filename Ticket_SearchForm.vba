Option Compare Database
Option Explicit

' ADO connection (cached)
Private connAdo As Object

' Returns an ADO connection (creates/caches it via GetConnection if needed)
Private Function GetConn() As Object
    Set GetConn = GetConnection(connAdo)
End Function

' Initialize form and load initial results
Private Sub Form_Load()
    ' Initialize custom form properties
    ConfigCustomForm Me

    ' Load records into the datasheet using default criteria/sort
    DoSearch
End Sub

' Validate inputs, run the search, and bind results to the datasheet subform
Public Sub DoSearch(Optional ByVal orderByAdo As String = "")
    ' Input validation states
    Dim stTicketID As XInputState: Set stTicketID = GetInt4(Me.txtTicketID)
    Dim stName As XInputState: Set stName = GetString(Me.txtName)

    ' Collect states so we can validate all inputs together
    Dim states As XStateCollection: Set states = New XStateCollection
    states.AddStates stTicketID, stName

    Dim rsAdo As Object

    ' Error handling (On Error GoTo)
    On Error GoTo TCError

    If Not states.AllValid Then
        ' If inputs are invalid, bind an empty recordset (keeps the datasheet stable)
        Set rsAdo = CreateEmptyRsAdo(Me.Ticket_Datasheet.Form.Recordset)
    Else
        ' Execute search
        Set rsAdo = SearchTicketAdo(GetConn(), stTicketID.value, stName.value)

        ' Apply sorting (default to DateCreated DESC)
        rsAdo.Sort = IIf(orderByAdo <> "", orderByAdo, "DateCreated DESC")
    End If

    ' Rebind datasheet recordset: close existing subform recordset first
    CloseObj Me.Ticket_Datasheet.Form.Recordset
    Set Me.Ticket_Datasheet.Form.Recordset = rsAdo

    Exit Sub

TCError:
    CloseObj rsAdo
    MsgBox "Search failed: " & err.Description, vbCritical
End Sub

' Open the edit form in Add mode
Private Sub btnAddNew_Click()
    DoCmd.OpenForm "Ticket_EditForm", , , , acFormAdd
End Sub

' Run the search using the current criteria
Private Sub btnSearch_Click()
    DoSearch
End Sub

' Cleanup: release recordsets/connections when the form closes
Private Sub Form_Close()
    CloseObj Me.Ticket_Datasheet.Form.Recordset
    CloseObj connAdo
End Sub
