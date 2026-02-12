Option Compare Database
Option Explicit

' Build and return the ADO connection string used by this app.
' NOTES: Update the connection properties below as needed for your environment.

Public Function GetDbConString(Optional ByVal timeoutSeconds As Long = 5) As String
    Dim cs As XConnStr
    Set cs = New XConnStr

    ' MS-SQL ADO Provider
    cs.Add "Provider", "MSOLEDBSQL"
    
    ' Database Server IP/Domain/Name
    cs.Add "Data Source", "localhost"
    
    ' Database Name
    cs.Add "Initial Catalog", "TestDB"
    
    ' Windows Authentication
    cs.Add "Integrated Security", "SSPI"
    
    cs.Add "TrustServerCertificate", "True"
    cs.Add "Encrypt", "False"
    
    cs.Add "Connect Timeout", CStr(timeoutSeconds)

    ' NOTES: SQL Server Authentication

    ' cs.Add "User ID", "dbuser"
    ' cs.Add "Password", "dbpassword"
    ' Delete the line: "Integrated Security", "SSPI"

    GetDbConString = cs.Build()
End Function

' Return an open ADODB.Connection, creating/opening it if needed
Public Function GetConnection(ByRef cn As Object, Optional ByVal timeoutSeconds As Long = 5) As Object
    Dim xe As XError
    On Error GoTo TCError

    If cn Is Nothing Then
        Set cn = CreateObject("ADODB.Connection")
        cn.ConnectionString = GetDbConString(timeoutSeconds)
        ' cn.ConnectionTimeout = timeoutSeconds
    End If

    If cn.state = 0 Then cn.Open

    Set GetConnection = cn
    Exit Function

TCError:
    Set xe = ToXError(err)

    CloseObj cn
    err.Raise xe.ErrNum, xe.ErrSrc, xe.ErrDesc
End Function
