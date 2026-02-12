Option Compare Database
Option Explicit

' ADO connection
Private connAdo As Object

' Current TicketID
Private ticketID As Variant

' Returns an ADO connection
Private Function GetConn() As Object
    Set GetConn = GetConnection(connAdo)
End Function

' Form startup
Private Sub Form_Load()
    ' Initialize custom form
    ConfigCustomForm Me

    ' Initialize dropdown controls
    InitDropdowns

    ' Try to Parse TicketID from OpenArgs
    ParseInt4 Me.OpenArgs, ticketID

    If IsNull(ticketID) Then
        ' Add New Case
        Me.btnDelete.Enabled = False
        
        Me.txtTicketID.Enabled = False
        Me.txtDateCreated.Enabled = False
    Else
        ' Update Case
        
        ' Make readonly
        Me.txtTicketID.Locked = True
        Me.txtDateCreated.Locked = True
    End If
    
    ' Load ticket
    LoadTicket
End Sub

' Initialize dropdown controls
Private Sub InitDropdowns()
    ' cboTicketTypeID
    InitDropdown Me.cboTicketTypeID, True

    Dim opts As XDropdownOptions
    Set opts = ExecuteDropdownOptionsSqlAdo(GetConn(), "SELECT TicketTypeID, Name FROM TicketType")
    opts.ToValueList Me.cboTicketTypeID

    ' Other dropdowns...
End Sub

' Try to load ticket
Private Sub LoadTicket()
    If IsNull(ticketID) Then
        Exit Sub
    End If
    
    Dim dict As Object

    ' Try/Catch Error
    On Error GoTo TCError

    ' Build PK
    Dim pk As Object: Set pk = NewDict
    pk("TicketID") = ticketID

    ' Load Ticket
    Set dict = GetRowByPkAdo(GetConn(), "Ticket", "TicketID", "INT4", pk)

    If dict Is Nothing Then
        MsgBox "Record not found for TicketID " & ticketID, vbInformation
        DoCmd.Close acForm, Me.name, acSaveNo
    Else
        ' Bind record fields to controls
        Me.txtTicketID = dict("TicketID")
        Me.txtName = dict("Name")
        Me.txtDescription = dict("Description")

        Me.chkIsDone = dict("IsDone")
        Me.cboTicketTypeID = dict("TicketTypeID")
        Me.txtDateCreated = dict("DateCreated")

        ' Change form caption
        Me.Caption = "EDIT TICKET: " & dict("TicketID")
    End If

    Exit Sub

TCError:
    MsgBox "Failed to load record: " & err.Description, vbCritical
End Sub

' Handle btnSave Clicked
Private Sub btnSave_Click()
    ' Form Inputs
    Dim stTicketID As XInputState: Set stTicketID = GetInt4(Me.txtTicketID, "TicketID", False)
    Dim stName As XInputState: Set stName = GetString(Me.txtName, "Name", True)
    Dim stDescription As XInputState: Set stDescription = GetString(Me.txtDescription, "Description", False)

    Dim stIsDone As XInputState: Set stIsDone = GetBool(Me.chkIsDone, "IsDone", True)
    Dim stTicketTypeID As XInputState: Set stTicketTypeID = GetInt4(Me.cboTicketTypeID, "TicketTypeID", True)
    Dim stDateCreated As XInputState: Set stDateCreated = GetDate(Me.txtDateCreated, "DateCreated", False)

    ' State Collection
    Dim states As XStateCollection: Set states = New XStateCollection
    states.AddStates stTicketID, stName, stDescription, _
                     stIsDone, stTicketTypeID, stDateCreated

    ' Show Errors
    If Not states.AllValid Then
        MsgBox "Please fix errors:" & vbCrLf & vbCrLf & states.ToErrorString, vbExclamation
        Exit Sub
    End If

    ' Try/Catch Error
    On Error GoTo TCError

    ' Build values dictionary
    Dim values As Object: Set values = states.ToValuesDict

    If IsNull(ticketID) Then
        ' Add New

        ' Set DateCreated = current date
        values("DateCreated") = Date

        InsertRowAdo GetConn(), "Ticket", _
        "Name, Description, IsDone, TicketTypeID, DateCreated", _
        "VARCHAR(100), VARCHAR(4000), BOOL, INt4, DATE", _
        values
    Else
        ' Update Case

        UpdateRowAdo GetConn(), "Ticket", _
        "Name, Description, IsDone, TicketTypeID", _
        "VARCHAR(100), VARCHAR(4000), BOOL, INt4", _
        "TicketID", "INT4", _
        values
    End If

    MsgBox "Record saved successfully.", vbInformation
    DoCmd.Close acForm, Me.name, acSaveNo
    Exit Sub

TCError:
    MsgBox "Failed to save record: " & err.Description, vbCritical
End Sub

' Handle btnDelete clicked
Private Sub btnDelete_Click()

    ' Confirm delete
    If MsgBox("Are you sure you want to delete this record: " & ticketID & "?", vbYesNo + vbQuestion) <> vbYes Then
        Exit Sub
    End If

    ' Try/Catch Error
    On Error GoTo TCError

    ' Build PK
    Dim pk As Object: Set pk = NewDict
    pk("TicketID") = ticketID

    ' Delete the ticket
    DeleteRowAdo GetConn(), "Ticket", "TicketID", "INT4", pk

    MsgBox "Record deleted.", vbInformation
    DoCmd.Close acForm, Me.name, acSaveNo
    Exit Sub

TCError:
    MsgBox "Failed to delete record: " & err.Description, vbCritical
End Sub

' Handle Form Close: Release resources
Private Sub Form_Close()
    CloseObj connAdo

    ' Call Ticket_MainForm.RefreshTickets
    TryCallForm "Ticket_MainForm", "RefreshTickets"
End Sub
