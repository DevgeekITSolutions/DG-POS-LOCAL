Imports MySql.Data.MySqlClient

Public Class SeniorDetails
    Private Sub ButtonCANCEL_Click(sender As Object, e As EventArgs) Handles ButtonCANCEL.Click
        SENIORDETAILSBOOL = False
        Close()
    End Sub
    Private Sub SeniorDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        CouponCode.Enabled = True
    End Sub
    Private Sub ButtonSubmit_Click(sender As Object, e As EventArgs) Handles ButtonSubmit.Click
        Try

            If TextboxIsEmpty(Me) Then
                Dim LimitToOne As Boolean = False
                Dim ConnectionLocal As MySqlConnection = LocalhostConn()
                Dim SQL = "SELECT senior_id FROM `loc_senior_details` WHERE senior_id = '" & Trim(TextBoxSENIORID.Text) & "'"
                Console.WriteLine(SQL)
                Dim Cmd As MySqlCommand = New MySqlCommand(SQL, ConnectionLocal)
                Using reader As MySqlDataReader = Cmd.ExecuteReader
                    If reader.HasRows Then
                        LimitToOne = True
                    Else
                        LimitToOne = False
                    End If
                End Using
                If LimitToOne Then
                    MessageBox.Show("Benefit limit reached for the day. Please use another ID. Thank you", "NOTICE", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    SeniorDetailsID = Trim(TextBoxSENIORID.Text)
                    SeniorDetailsName = Trim(TextBoxSENIORNAME.Text)
                    SENIORDETAILSBOOL = True
                    CouponCode.couponpercentage()
                    Close()
                End If
            Else
                MessageBox.Show("Fill up all the blanks.", "NOTICE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Private Sub TextBoxSENIORID_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxSENIORID.KeyPress
        Try
            'Numeric(sender, e)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class