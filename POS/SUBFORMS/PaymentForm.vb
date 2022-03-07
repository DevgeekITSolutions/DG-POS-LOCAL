Public Class PaymentForm
    Private Sub PaymentForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        POS.Enabled = True
        POS.BringToFront()
    End Sub
    Private Sub PaymentForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TopMost = True
        TextBoxMONEY.Text = "0"
        TextBoxDiscType.Text = POS.DISCOUNTTYPE
        TextBoxTransactionType.Text = TRANSACTIONMODE
        TextBoxMONEY.Focus()
        TextBoxMONEY.Focus()
    End Sub
    Private Sub TextBoxMONEY_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxMONEY.KeyDown
        If e.KeyCode = Keys.F9 Then
            ButtonESC.PerformClick()
        ElseIf e.KeyCode = Keys.Enter Then
            ButtonSubmitPayment.PerformClick()
        End If
    End Sub
    Private Sub ButtonSubmitPayment_Click(sender As Object, e As EventArgs) Handles ButtonSubmitPayment.Click
        Try
            With POS
                If Double.Parse(TextBoxTOTALPAY.Text) > Val(TextBoxMONEY.Text) Then
                    MsgBox("Insufficient money")
                    TextBoxMONEY.Clear()
                    TextBoxCHANGE.Clear()
                    messageboxappearance = False
                Else
                    If TextBoxMONEY.Text = "" Then
                        TextBoxMONEY.Text = 0
                    End If
                    TEXTBOXMONEYVALUE = TextBoxMONEY.Text
                    TEXTBOXCHANGEVALUE = TextBoxCHANGE.Text
                    .BackgroundWorker1.WorkerSupportsCancellation = True
                    .BackgroundWorker1.WorkerReportsProgress = True
                    .BackgroundWorker1.RunWorkerAsync()
                    .Label9.Text = NUMBERFORMAT(Val(TextBoxMONEY.Text))
                    .Label13.Text = NUMBERFORMAT(Val(TextBoxCHANGE.Text))
                    If My.Settings.LedDisplayTrue Then
                        LedDisplay(TextBoxCHANGE.Text, False)
                    End If
                    Close()
                    WaitFrm.Show()
                End If
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub TextBoxMONEY_TextChanged(sender As Object, e As EventArgs) Handles TextBoxMONEY.TextChanged
        Try
            If Val(TextBoxMONEY.Text) = 0 Then
                TextBoxCHANGE.Text = "0"
            End If
            TextBoxCHANGE.Text = NUMBERFORMAT(Double.Parse(TextBoxMONEY.Text) - Double.Parse(TextBoxTOTALPAY.Text))
        Catch ex As Exception
        End Try
    End Sub
    Public Sub buttonpressedenter(ByVal btntext As String)
        If Val(TextBoxMONEY.Text) <> 0 Then
            TextBoxMONEY.Text += btntext
        Else
            TextBoxMONEY.Text = btntext
        End If
    End Sub
    Private Sub ButtonESC_Click(sender As Object, e As EventArgs) Handles ButtonESC.Click
        Me.Close()
    End Sub
    Private Sub TextBoxMONEY_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxMONEY.KeyPress
        If (e.KeyChar.ToString = ".") And (TextBoxMONEY.Text.Contains(e.KeyChar.ToString)) Then
            e.Handled = True
            Exit Sub
        End If
        Numeric(sender:=sender, e:=e)
    End Sub
    Private Sub ButtonNo9_Click(sender As Object, e As EventArgs) Handles ButtonNo9.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo9.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo9.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo8_Click(sender As Object, e As EventArgs) Handles ButtonNo8.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo8.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo8.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo7_Click(sender As Object, e As EventArgs) Handles ButtonNo7.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo7.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo7.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo6_Click(sender As Object, e As EventArgs) Handles ButtonNo6.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo6.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo6.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo5_Click(sender As Object, e As EventArgs) Handles ButtonNo5.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo5.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo5.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo4_Click(sender As Object, e As EventArgs) Handles ButtonNo4.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo4.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo4.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo3_Click(sender As Object, e As EventArgs) Handles ButtonNo3.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo3.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo3.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo2_Click(sender As Object, e As EventArgs) Handles ButtonNo2.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo2.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo2.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo1_Click(sender As Object, e As EventArgs) Handles ButtonNo1.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo1.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo1.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo0_Click(sender As Object, e As EventArgs) Handles ButtonNo0.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo0.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo0.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo00_Click(sender As Object, e As EventArgs) Handles ButtonNo00.Click
        If payment = False Then
            If TextBoxMONEY.Text.Length > 5 Then
            Else
                buttonpressedenter(btntext:=ButtonNo00.Text)
            End If
        Else
            If TextBoxMONEY.Text.Length > 5 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo00.Text)
            End If
        End If
    End Sub
    Private Sub Buttondot_Click(sender As Object, e As EventArgs) Handles Buttondot.Click
        If payment = False Then
            If Not TextBoxMONEY.Text.Contains(".") Then
                TextBoxMONEY.Text += "."
            End If
        End If
    End Sub
    Private Sub ButtonClear_Click(sender As Object, e As EventArgs) Handles ButtonClear.Click
        If payment = False Then
            TextBoxMONEY.Text = 0
        End If
    End Sub
    Private Sub PaymentForm_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.NumPad0 Then
            ButtonNo0.PerformClick()
        ElseIf e.KeyCode = Keys.NumPad1 Then
            ButtonNo1.PerformClick()
        ElseIf e.KeyCode = Keys.NumPad2 Then
            ButtonNo2.PerformClick()
        ElseIf e.KeyCode = Keys.NumPad3 Then
            ButtonNo3.PerformClick()
        ElseIf e.KeyCode = Keys.NumPad4 Then
            ButtonNo4.PerformClick()
        ElseIf e.KeyCode = Keys.NumPad5 Then
            ButtonNo5.PerformClick()
        ElseIf e.KeyCode = Keys.NumPad6 Then
            ButtonNo6.PerformClick()
        ElseIf e.KeyCode = Keys.NumPad7 Then
            ButtonNo7.PerformClick()
        ElseIf e.KeyCode = Keys.NumPad8 Then
            ButtonNo8.PerformClick()
        ElseIf e.KeyCode = Keys.NumPad9 Then
            ButtonNo9.PerformClick()
        ElseIf e.KeyCode = Keys.Back Then
            ButtonClear.PerformClick()
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBoxMONEY.Text.Count > 0 Then
            TextBoxMONEY.Text = TextBoxMONEY.Text.Remove(TextBoxMONEY.Text.Count - 1)
        End If
    End Sub
End Class