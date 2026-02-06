Option Compare Database
Option Explicit

' Datasheet delegate used to handle common datasheet behaviors (timer, double-click, filter, errors)
Private datasheetDelegate As XDatasheetDelegate

' Form startup: initialize datasheet form styling and configure the delegate
Private Sub Form_Load()
    ' Initialize custom datasheet form properties
    ConfigCustomDatasheet Me

    ' Create and initialize the datasheet delegate
    Set datasheetDelegate = New XDatasheetDelegate
    datasheetDelegate.Init _
        datasheetForm:=Me, _
        parentForm:="Ticket_SearchForm", _
        reloadMethod:="DoSearch", _
        editForm:="Ticket_EditForm", _
        pkField:="TicketID", _
        timerIntervalMs:=100
End Sub

' Forward the Timer event to the datasheet delegate
Private Sub Form_Timer()
    datasheetDelegate.Form_Timer
End Sub

' Forward the Double-Click event to the datasheet delegate
Private Sub Form_DblClick(Cancel As Integer)
    datasheetDelegate.Form_DblClick Cancel
End Sub

' Forward the ApplyFilter event to the datasheet delegate
Private Sub Form_ApplyFilter(Cancel As Integer, ApplyType As Integer)
    datasheetDelegate.Form_ApplyFilter Cancel, ApplyType
End Sub

' Forward the form Error event to the datasheet delegate
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    datasheetDelegate.Form_Error DataErr, Response
End Sub