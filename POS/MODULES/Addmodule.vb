Imports MySql.Data.MySqlClient
Module Addmodule
    Public Sub GLOBAL_SYSTEM_LOGS(ByVal logtype As String, ByVal logdesc As String)
        Try
            If ClientCrewID = "" Then
                ClientCrewID = 0
            End If
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim Query As String = "INSERT INTO `loc_system_logs`(`log_type`,`crew_id`,`log_description`, `log_store`, `guid`, `loc_systemlog_id`, `synced`, `zreading`, `log_date_time`) VALUES (@1,@2,@3,@4,@5,@6,@7,@8,@9)"
            Dim Command As MySqlCommand = New MySqlCommand(Query, ConnectionLocal)
            Command.Parameters.Add("@1", MySqlDbType.VarChar).Value = logtype
            Command.Parameters.Add("@2", MySqlDbType.VarChar).Value = ClientCrewID
            Command.Parameters.Add("@3", MySqlDbType.Text).Value = logdesc
            Command.Parameters.Add("@4", MySqlDbType.VarChar).Value = ClientStoreID
            Command.Parameters.Add("@5", MySqlDbType.VarChar).Value = ClientGuid
            Command.Parameters.Add("@6", MySqlDbType.VarChar).Value = Format(Now, ("yyyyMMdd-HHmmss"))
            Command.Parameters.Add("@7", MySqlDbType.VarChar).Value = "Unsynced"
            If IsNothing(S_Zreading) Then
                Command.Parameters.Add("@8", MySqlDbType.VarChar).Value = ""
            Else
                Command.Parameters.Add("@8", MySqlDbType.VarChar).Value = S_Zreading
            End If
            Command.Parameters.Add("@9", MySqlDbType.Text).Value = FullDate24HR()
            Command.ExecuteNonQuery()
            ConnectionLocal.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Dim result As Integer
    Public Sub GLOBAL_INSERT_FUNCTION(ByVal table As String, ByVal fields As String, ByVal values As String)
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim Query As String = "INSERT INTO " + table + fields + " VALUES " + values
            Dim cmd As MySqlCommand = New MySqlCommand(Query, ConnectionLocal)
            cmd.ExecuteNonQuery()
            ConnectionLocal.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub SendErrorReport(MSG)

        Dim ConnectionLocal As MySqlConnection = LocalhostConn()
        'If ConnectionLocal.State = ConnectionState.Open Then
        Try
                Dim Query As String = "INSERT INTO `loc_send_bug_report`(`bug_desc`, `crew_id`, `guid`, `store_id`, `date_created`, `synced`) VALUES (@1,@2,@3,@4,@5,@6)"
                Dim Command As MySqlCommand = New MySqlCommand(Query, ConnectionLocal)
                Command.Parameters.Add("@1", MySqlDbType.Text).Value = MSG
                Command.Parameters.Add("@2", MySqlDbType.Text).Value = ClientCrewID
                Command.Parameters.Add("@3", MySqlDbType.Text).Value = ClientGuid
                Command.Parameters.Add("@4", MySqlDbType.Text).Value = ClientStoreID
                Command.Parameters.Add("@5", MySqlDbType.Text).Value = FullDate24HR()
                Command.Parameters.Add("@6", MySqlDbType.Text).Value = "Unsynced"
                Command.ExecuteNonQuery()
                ConnectionLocal.Close()

                MessageBox.Show("An error occured. Please contact the System Administrator", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch ex As Exception
            Console.WriteLine("Error: " & ex.ToString)
        End Try
        'Else
        '    MsgBox("Localhost connection is not valid.")
        'End If
    End Sub
End Module
