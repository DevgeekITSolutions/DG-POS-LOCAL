Imports System.Management
Imports System.Management.Instrumentation
Imports System
Imports System.IO
Imports MySql.Data.MySqlClient
Module DeleteModule
    Public Sub GLOBAL_DELETE_ALL_FUNCTION(ByVal tablename As String, ByVal where As String)
        Try
            sql = "DELETE FROM " & tablename & " WHERE " & where
            With cmd
                .Connection = LocalhostConn()
                .CommandText = sql
            End With
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    'Public Sub truncatetable(ByVal tablename As String)
    '    Try
    '        sql = "TRUNCATE TABLE " & tablename & ";"
    '        cmd = New MySqlCommand
    '        With cmd
    '            .Connection = LocalhostConn()
    '            .CommandText = sql
    '        End With
    '        cmd.ExecuteNonQuery()
    '    Catch ex As Exception
    '        SendErrorReport(ex.ToString)
    '    End Try
    'End Sub
End Module
