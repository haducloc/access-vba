Option Compare Database
Option Explicit

' Cached ADO connection
Private connAdo As Object

' Current TicketID (Null when adding a new record)
Private ticketID As Variant

' Returns an ADO connection (creates/caches it via GetConnection if needed)
Private Function GetConn() As Object
    Set GetConn = GetConnection(connAdo)
End Function

' Form startup: init UI, parse OpenArgs, and load record when editing
Private Sub Form_Load()
    ' Initialize custom form properties and set this form to modal
    ConfigCustomForm Me, True

    ' Initialize dropdown controls and load their option lists
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
    
    ' If we have a TicketID, load the existing record into the form controls
    If Not IsNull(ticketID) Then
        LoadTicket ticketID
        Exit Sub
    End If
End Sub

' Initialize dropdown controls and populate their options
Private Sub InitDropdowns()
    ' Initialize cboTicketTypeID as a value-list dropdown (no blank option)
    InitDropdown Me.cboTicketTypeID, True

    ' Load dropdown options from the TicketType table
    Dim opts As XDropdownOptions
    Set opts = ExecuteDropdownOptionsSqlAdo(GetConn(), "SELECT TicketTypeID, Name FROM TicketType")
    opts.ToValueList Me.cboTicketTypeID

    ' Other dropdowns...
End Sub

' Load a Ticket by TicketID and bind its fields to form controls
Public Sub LoadTicket(ByVal ticketID As Long)
    Dim dict As Object

    ' Error handling (On Error GoTo)
    On Error GoTo TCError

    ' Build PK dictionary for lookup
    Dim pk As Object: Set pk = NewDict
    pk("TicketID") = ticketID

    ' Fetch row as a dictionary (column name -> value)
    Set dict = GetRowByPkAdo(connAdo, "Ticket", "TicketID", "INT4", pk)

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
    End If

    Exit Sub

TCError:
    MsgBox "Failed to load record: " & err.Description, vbCritical
End Sub

' Validate inputs and save: INSERT when adding, UPDATE when editing
Private Sub btnSave_Click()
    ' Input states
    ' NOTES: The third argument is True/False means required or optional

    Dim stTicketID As XInputState: Set stTicketID = GetInt4(Me.txtTicketID, "TicketID", False)
    Dim stName As XInputState: Set stName = GetString(Me.txtName, "Name", True)
    Dim stDescription As XInputState: Set stDescription = GetString(Me.txtDescription, "Description", False)

    Dim stIsDone As XInputState: Set stIsDone = GetBool(Me.chkIsDone, "IsDone", True)
    Dim stTicketTypeID As XInputState: Set stTicketTypeID = GetInt4(Me.cboTicketTypeID, "TicketTypeID", True)
    Dim stDateCreated As XInputState: Set stDateCreated = GetDate(Me.txtDateCreated, "DateCreated", False)

    ' Collect states so we can validate all inputs together
    Dim states As XStateCollection: Set states = New XStateCollection
    states.AddStates stTicketID, stName, stDescription, _
                     stIsDone, stTicketTypeID, stDateCreated

    ' Stop if validation failed and show aggregated errors
    If Not states.AllValid Then
        MsgBox "Please fix errors:" & vbCrLf & vbCrLf & states.ToErrorString, vbExclamation
        Exit Sub
    End If

    ' Error handling (On Error GoTo)
    On Error GoTo TCError

    ' Values dictionary derived from validated inputs
    Dim values As Object: Set values = states.ToValuesDict

    If IsNull(ticketID) Then
        ' Add New
        ' Set DateCreated = current date
        values("DateCreated") = Date

        ' Insert new Ticket row
        InsertRowAdo connAdo, _
        "Ticket", _
        "Name, Description, IsDone, TicketTypeID, DateCreated", _
        "VARCHAR(100), VARCHAR(4000), BOOL, INt4, DATE", _
        values
    Else
        ' Update Case
        UpdateRowAdo connAdo, _
        "Ticket", _
        "Name, Description, IsDone, TicketTypeID", _
        "VARCHAR(100), VARCHAR(4000), BOOL, INt4", _
        "TicketID", _
        "INT4", _
        values
    End If

    MsgBox "Record saved successfully.", vbInformation
    DoCmd.Close acForm, Me.name, acSaveNo
    Exit Sub

TCError:
    MsgBox "Failed to save record: " & err.Description, vbCritical
End Sub

' Confirm and delete the current Ticket (by TicketID)
Private Sub btnDelete_Click()
    ' TicketID is required for delete
    Dim stTicketID As XInputState: Set stTicketID = GetInt4(Me.txtTicketID, "TicketID", True)

    ' Confirm delete
    If MsgBox("Are you sure you want to delete this record: " & stTicketID.value & "?", vbYesNo + vbQuestion) <> vbYes Then
        Exit Sub
    End If

    ' Error handling (On Error GoTo)
    On Error GoTo TCError

    ' Build PK dictionary for delete
    Dim pk As Object: Set pk = NewDict
    pk("TicketID") = stTicketID.value

    ' Delete the row
    DeleteRowAdo connAdo, "Ticket", "TicketID", "INT4", pk

    MsgBox "Record deleted.", vbInformation
    DoCmd.Close acForm, Me.name, acSaveNo
    Exit Sub

TCError:
    MsgBox "Failed to delete record: " & err.Description, vbCritical
End Sub

' Cleanup and refresh the search form when this form closes
Private Sub Form_Close()
    CloseObj connAdo

    ' Refresh results on the search form
    InvokeFormMethod "Ticket_SearchForm", "DoSearch"
End Sub
