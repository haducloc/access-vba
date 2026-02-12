Option Compare Database
Option Explicit

Private delegate As XDatasheetDelegate

' NOTES: You ONLY implement RefreshFromParent and OpenEditForm

Public Sub RefreshFromParent(ByVal sortByAdo As String)
    If HasLoadedParent(Me, "Ticket_MainForm") Then
        Me.Parent.RefreshTickets sortByAdo
    End If
End Sub

Public Sub OpenEditForm(ByVal selectedRow As Object)
    Dim ticketID As Long
    ticketID = selectedRow("TicketID")

    TryOpenForm "Ticket_EditForm", CStr(ticketID)
End Sub

' DON'T touch the Private functions below

Private Sub Form_Load()
    Set delegate = New XDatasheetDelegate
    delegate.Attach Me
    delegate.OnLoad
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set delegate = Nothing
End Sub

Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    delegate.OnApplyFilter Cancel, ApplyType
End Sub

Private Sub Form_Timer()
    If delegate Is Nothing Then Exit Sub
    delegate.OnTimer
End Sub

Private Sub Form_DblClick(Cancel As Integer)
    delegate.OnDblClick
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    delegate.OnError DataErr, Response
End Sub
