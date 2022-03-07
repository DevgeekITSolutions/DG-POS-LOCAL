Imports MySql.Data.MySqlClient
Module Updatemodule
    Public Sub GLOBAL_FUNCTION_UPDATE(ByVal table, ByVal fields, ByVal where)
        Try
            If LocalhostConn.State <> ConnectionState.Open Then
                LocalhostConn.Open()
            End If
            Dim sql = "UPDATE " + table + " SET " + fields + " WHERE " & where
            Dim cmd As MySqlCommand = New MySqlCommand(sql, LocalhostConn)
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        Finally
            LocalhostConn.Close()
            cmd.Dispose()
        End Try
    End Sub
End Module
