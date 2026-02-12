Option Compare Database
Option Explicit

' Converts a value into a SQL LIKE parameter, treating quoted input as a literal pattern
Public Function ToLikeParamValue(ByVal s As Variant, ByVal dbType As XDbType) As Variant
    If IsNull(s) Then
        ToLikeParamValue = Null
        Exit Function    
    End If

    Dim val As String: val = Trim$(CStr(s))

    If val = "" Then
        ToLikeParamValue = Null
        Exit Function
    End If

    If (Left$(val, 1) = """" And Right$(val, 1) = """") Then
        ToLikeParamValue = Trim$(Mid$(val, 2, Len(val) - 2))
    Else
        ToLikeParamValue = ToLikePattern(val, dbType)
    End If
End Function

' Escapes special characters for a LIKE pattern and wraps in wildcards.
Private Function ToLikePattern(ByVal val As String, ByVal dbType As XDbType) As Variant
    If val = "" Then
        ToLikePattern = Null
        Exit Function
    End If

    Select Case dbType
        Case Db_Access
            ' Access uses [] escaping and * / ? wildcards.
            val = Replace(val, "[", "[[]")
            val = Replace(val, "]", "[]]")
            val = Replace(val, "#", "[#]")
            val = Replace(val, "*", "[*]")
            val = Replace(val, "?", "[?]")
            ToLikePattern = "*" & val & "*"

        Case Db_SQLServer
            ' SQL Server LIKE escaping: [%] [_] [[] (no ESCAPE clause needed)
            val = Replace(val, "[", "[[]")
            val = Replace(val, "%", "[%]")
            val = Replace(val, "_", "[_]")
            ToLikePattern = "%" & val & "%"

        Case Db_Oracle, Db_Postgres, Db_MySQL
            ' Safest with: ... ESCAPE '\'
            val = Replace(val, "\", "\\")
            val = Replace(val, "%", "\%")
            val = Replace(val, "_", "\_")
            ToLikePattern = "%" & val & "%"

        Case Else
            ToLikePattern = val
    End Select
End Function
