Option Compare Database
Option Explicit

' ADO connection
Private connAdo As Object

' Returns an ADO connection
Private Function GetConn() As Object
    Set GetConn = GetConnection(connAdo)
End Function

' Handle Form Load
Private Sub Form_Load()
    ' Initialize custom form
    ConfigCustomForm Me

    ' Load all tickets at form load
    SearchTickets
End Sub

' Search tickets
Public Sub SearchTickets(Optional ByVal sortByAdo As String = "")
    ' Search Inputs
    Dim stTicketID As XInputState: Set stTicketID = GetInt4(Me.txtTicketID)
    Dim stName As XInputState: Set stName = GetString(Me.txtName)

    ' State Collection
    Dim states As XStateCollection: Set states = New XStateCollection
    states.AddStates stTicketID, stName

    Dim rsAdo As Object

    ' Try/Cache Error
    On Error GoTo TCError

    If Not states.AllValid Then
        ' Empty result
        Set rsAdo = CreateEmptyRsAdo(Me.Ticket_Datasheet.Form.Recordset)
    Else
        ' Execute search
        Set rsAdo = SearchTicketAdo(GetConn(), stTicketID.Value, stName.Value)
        
        ' Apply sorting
        rsAdo.Sort = BuildRsSortByAdo(sortByAdo, Me.Ticket_Datasheet.Form, "DateCreated DESC")
    End If
    
    ' Bind rsAdo to Ticket_Datasheet
    CloseObj Me.Ticket_Datasheet.Form.Recordset
    Set Me.Ticket_Datasheet.Form.Recordset = rsAdo
    
    Exit Sub
TCError:
    CloseObj rsAdo
    MsgBox "Search failed: " & err.Description, vbCritical
End Sub

' Handle btnAddNew clicked
Private Sub btnAddNew_Click()
    TryOpenForm "Ticket_EditForm"
End Sub

' Handle btnSearch clicked
Private Sub btnSearch_Click()
    SearchTickets
End Sub

' Handle Form Close: Release resources
Private Sub Form_Close()
    CloseObj Me.Ticket_Datasheet.Form.Recordset
    CloseObj connAdo
End Sub

' Public callback functions used by the Ticket_Datasheet and Ticket_EditForm
' Purpose: To refresh tickets

Public Sub RefreshTickets(Optional ByVal sortByAdo As String = "")
   SearchTickets sortByAdo
End Sub
