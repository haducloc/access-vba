Option Compare Database
Option Explicit

' ADO connection
Private connAdo As Object

' Current ticketCommentID
Private ticketCommentID As Variant

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
    
    ' OpenArgs must be in this format: ticketCommentID|ticketID
    If IsNull(Me.OpenArgs) Then
        Exit Sub
    End If
    
    ' Parse ticketCommentID, ticketID
    Dim args() As String
    args = DecodeArgs(CStr(Me.OpenArgs))
    
    ParseInt4 args(0), ticketCommentID
    ParseInt4 args(1), ticketID

    ' Init ticketCommentID & txtTicketID
    Me.txtTicketCommentID = ticketCommentID
    Me.txtTicketID = ticketID
        
    ' TicketCommentID & TicketID are readonly
    Me.txtTicketCommentID.Enabled = False
    Me.txtTicketID.Enabled = False

    If IsNull(ticketCommentID) Then
        ' Add New
        Me.btnDelete.Enabled = False
    Else
        ' Update
    End If
    
    ' Try to load ticketComment
    LoadTicketComment
End Sub

' Try to load ticket comment
Private Sub LoadTicketComment()
    ' Add New -> Skip
    If IsNull(ticketCommentID) Then
        Exit Sub
    End If
    
    Dim dict As Object

    ' Try/Catch Error
    On Error GoTo TCError

    ' Build PK
    Dim pk As Object: Set pk = NewDict
    pk("ticketCommentID") = ticketCommentID

    ' Load the record
    Set dict = GetRowByPkAdo(GetConn(), "TicketComment", "TicketCommentID", "INT4", pk)

    If dict Is Nothing Then
        MsgBox "Record not found for TicketCommentID " & ticketCommentID, vbInformation
        DoCmd.Close acForm, Me.name, acSaveNo
    Else
        ' Bind record fields to controls
        
        ' Me.txtTicketCommentID = dict("TicketCommentID")
        ' Me.txtTicketID = dict("TicketID")
        Me.txtComment = dict("Comment")

        ' Change form caption
        Me.Caption = "EDIT TICKET COMMENT: " & dict("ticketCommentID")
    End If

    Exit Sub

TCError:
    MsgBox "Failed to load record: " & err.Description, vbCritical
End Sub

' Handle btnSave Clicked
Private Sub btnSave_Click()
    ' Form Inputs
    
    ' TicketCommentID is auto generated, so it is optional
    Dim stTicketCommentID As XInputState: Set stTicketCommentID = GetInt4(Me.txtTicketCommentID, "TicketCommentID", False)
    
    ' TicketID always has a value, so it is required
    Dim stTicketID As XInputState: Set stTicketID = GetInt4(Me.txtTicketID, "TicketID", True)
    
    Dim stComment As XInputState: Set stComment = GetString(Me.txtComment, "Comment", True)

    ' State Collection
    Dim states As XStateCollection: Set states = New XStateCollection
    states.AddStates stTicketCommentID, stTicketID, stComment

    ' Show Errors
    If Not states.AllValid Then
        MsgBox "Please fix errors:" & vbCrLf & vbCrLf & states.ToErrorString, vbExclamation
        Exit Sub
    End If

    ' Try/Catch Error
    On Error GoTo TCError

    ' Build values dictionary
    Dim values As Object: Set values = states.ToValuesDict
    
    If IsNull(ticketCommentID) Then
        ' Add New

        InsertRowAdo GetConn(), "TicketComment", _
        "TicketID, Comment", _
        "INT4, VARCHAR(255)", _
        values
    Else
        ' Update

        UpdateRowAdo GetConn(), "TicketComment", _
        "Comment", "VARCHAR(255)", _
        "TicketCommentID", "INT4", _
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
    If MsgBox("Are you sure you want to delete this record: " & ticketCommentID & "?", vbYesNo + vbQuestion) <> vbYes Then
        Exit Sub
    End If

    ' Try/Catch Error
    On Error GoTo TCError

    ' Build PK
    Dim pk As Object: Set pk = NewDict
    pk("ticketCommentID") = ticketCommentID

    ' Delete the record
    DeleteRowAdo GetConn(), "TicketComment", "ticketCommentID", "INT4", pk

    MsgBox "Record deleted.", vbInformation
    DoCmd.Close acForm, Me.name, acSaveNo
    Exit Sub

TCError:
    MsgBox "Failed to delete record: " & err.Description, vbCritical
End Sub

' Handle Form Close: Release resources
Private Sub Form_Close()
    CloseObj connAdo

    ' Call Ticket_EditForm.LoadComments to reload comments
    TryCallForm "Ticket_EditForm", "LoadComments"
End Sub
