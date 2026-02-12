Option Compare Database
Option Explicit

Public Enum XDbType
    Db_SQLServer = 1
    Db_Access = 2
    Db_Oracle = 3
    Db_Postgres = 4
    Db_MySQL = 5
End Enum

Public Function GetDbType(ByVal cn As Object) As XDbType
    Dim prov As String
    Dim cs As String
    
    ' Read Provider safely (cn might be closed, uninitialized, or not an ADO connection)
    On Error Resume Next
    prov = LCase$(Trim$(CStr(cn.Provider)))
    On Error GoTo 0
    
    If Len(prov) = 0 Then
        XRaise "XDbTypes.GetDbType", _
               "Connection object does not expose Provider (or is not initialized)."
        Exit Function
    End If
    
    Select Case True
        ' SQL Server (MSOLEDBSQL only)
        Case prov Like "msoledbsql*"
            GetDbType = Db_SQLServer
            Exit Function
        
        ' Access (ACE only)
        Case InStr(prov, "microsoft.ace.oledb") > 0
            GetDbType = Db_Access
            Exit Function
        
        ' Oracle (OraOLEDB only)
        Case InStr(prov, "oraoledb.oracle") > 0
            GetDbType = Db_Oracle
            Exit Function
        
        ' Postgres/MySQL via ODBC (MSDASQL only)
        Case InStr(prov, "msdasql") > 0
            ' Read ConnectionString safely
            On Error Resume Next
            cs = LCase$(CStr(cn.ConnectionString))
            On Error GoTo 0
            
            If Len(cs) = 0 Then
                XRaise "XDbTypes.GetDbType", _
                       "Provider is MSDASQL but ConnectionString is unavailable."
                Exit Function
            End If
            
            ' Prefer DRIVER={...} checks first (more reliable than generic substring matches)
            
            ' PostgreSQL ODBC (common driver names)
            If InStr(cs, "driver={postgresql") > 0 _
               Or InStr(cs, "driver={psqlodbc") > 0 _
               Or InStr(cs, "driver={postgresql ansi") > 0 _
               Or InStr(cs, "driver={postgresql unicode") > 0 _
               Or InStr(cs, "psqlodbc") > 0 Then
                GetDbType = Db_Postgres
                Exit Function
            End If
            
            ' MySQL ODBC (common driver names)
            If InStr(cs, "driver={mysql") > 0 _
               Or InStr(cs, "driver={myodbc") > 0 _
               Or InStr(cs, "myodbc") > 0 _
               Or InStr(cs, "connector/odbc") > 0 Then
                GetDbType = Db_MySQL
                Exit Function
            End If
            
            ' Fallback (less reliable): generic keywords in the connection string
            ' (kept last to reduce false positives)
            If InStr(cs, "postgresql") > 0 Then
                GetDbType = Db_Postgres
                Exit Function
            End If
            
            If InStr(cs, "mysql") > 0 Then
                GetDbType = Db_MySQL
                Exit Function
            End If
    End Select
    
    XRaise "XDbTypes.GetDbType", _
           "Could not determine database type for the provider: " & CStr(cn.Provider)
End Function