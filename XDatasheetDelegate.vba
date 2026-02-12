Option Compare Database
Option Explicit

Private mHost As Form

Private mPendingReload As Boolean
Private mSortByAdo As String

Public Sub Attach(ByVal hostForm As Form)
    Set mHost = hostForm
End Sub

Public Sub OnLoad()
    ConfigCustomDatasheet mHost
End Sub

Public Sub OnApplyFilter(ByRef Cancel As Integer, ByVal ApplyType As Integer)
    If ApplyType <> 1 Then Exit Sub

    ' prevent Access native sort on ADO datasheet
    Cancel = True

    mPendingReload = True
    mSortByAdo = ToRsSortByAdo(mHost.OrderBy)

    ' defer to timer to avoid re-entrancy
    mHost.TimerInterval = 100
End Sub

Public Sub OnTimer()
    mHost.TimerInterval = 0

    If Not mPendingReload Then Exit Sub
    mPendingReload = False

    ' host implements RefreshFromParent(sortByAdo)
    CallByName mHost, "RefreshFromParent", VbMethod, mSortByAdo
End Sub

Public Sub OnDblClick()
    If (mHost.Recordset Is Nothing) Then Exit Sub
    If mHost.Recordset.BOF Or mHost.Recordset.EOF Then Exit Sub

    ' host implements OpenEditForm(selectedRow)
    CallByName mHost, "OpenEditForm", VbMethod, mHost.Recordset.fields
End Sub

Public Sub OnError(ByVal DataErr As Integer, ByRef Response As Integer)
    If mHost.CurrentView = acCurViewDatasheet And DataErr = 3075 Then
        Response = acDataErrContinue
    Else
        Response = acDataErrDisplay
    End If
End Sub
