Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Net
Imports MySql.Data.MySqlClient
Imports System.Drawing.Printing
Imports System.Text.RegularExpressions
Imports System.Globalization

Module publicfunctions
    Public drasd
    Dim dr2
    Dim hashable As String
    Dim dateformat
    Dim timeformat
    Declare Function Wow64DisableWow64FsRedirection Lib "kernel32" (ByRef oldvalue As Long) As Boolean
    Private osk As String = "C:\Windows\System32\osk.exe"

    Dim sp As New System.IO.Ports.SerialPort("COM2", 2400, IO.Ports.Parity.None And IO.Ports.StopBits.One)
    Public Declare Function com_init Lib "api_com.dll" (ByVal com As Integer, ByVal baud As Integer) As Boolean
    Public Declare Function com_send Lib "api_com.dll" (ByVal buf As String, ByVal lens As Long) As Boolean
    Public Declare Function com_rest Lib "api_com.dll" () As Boolean

    Public Sub ShowKeyboard()
        Try
            Wow64DisableWow64FsRedirection(0)
            Process.Start(osk)
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub ButtonEnableability(ByVal root As Control, ENB As Boolean)
        For Each ctrl As Control In root.Controls
            ButtonEnableability(ctrl, ENB)
            If TypeOf ctrl Is Button Then
                CType(ctrl, Button).Enabled = ENB
            End If
        Next ctrl
    End Sub
    Public Sub CheckBoxEnabled(ByVal root As Control, ENB As Boolean)
        For Each ctrl As Control In root.Controls
            CheckBoxEnabled(ctrl, ENB)
            If TypeOf ctrl Is CheckBox Then
                CType(ctrl, CheckBox).Checked = ENB
            End If
        Next ctrl
    End Sub
    Public Function TextboxIsEmpty(ByVal root As Control)
        Dim ReturnThisThing As Boolean
        For Each tb As TextBox In root.Controls.OfType(Of TextBox)()
            If tb.Text = String.Empty Then

                ReturnThisThing = False
                Exit For
            Else
                ReturnThisThing = True
            End If
        Next
        Return ReturnThisThing
    End Function
    Public Sub TextboxEnableability(ByVal root As Control, ENB As Boolean)
        Try
            For Each ctrl As Control In root.Controls
                TextboxEnableability(ctrl, ENB)
                If TypeOf ctrl Is TextBox Then
                    CType(ctrl, TextBox).Enabled = ENB
                End If
            Next ctrl
        Catch ex As Exception
            MsgBox(ex.ToString)
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Public Sub ReturnZero(ByVal root As Control)

        '        If String.IsNullOrEmpty(textBox2.Text)) Then


        '    textBox2.Text = "0";                
        '}

        Try
            For Each ctrl As Control In root.Controls
                ReturnZero(ctrl)
                If TypeOf ctrl Is TextBox Then
                    If String.IsNullOrEmpty(CType(ctrl, TextBox).Text) Then
                        CType(ctrl, TextBox).Text = "0"
                        CType(ctrl, TextBox).SelectAll()
                    End If
                End If
            Next ctrl

        Catch ex As Exception
            MsgBox(ex.ToString)
            SendErrorReport(ex.ToString)
        End Try

    End Sub
    'Public Sub TextboxReadOnly(ByVal root As Control, ENB As Boolean)
    '    Try
    '        For Each ctrl As Control In root.Controls
    '            TextboxReadOnly(ctrl, ENB)
    '            If TypeOf ctrl Is TextBox Then
    '                CType(ctrl, TextBox).ReadOnly = ENB
    '            End If
    '        Next ctrl
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '        SendErrorReport(ex.ToString)
    '    End Try
    'End Sub
    Public Sub ClearTextBox(ByVal root As Control)
        Try
            For Each ctrl As Control In root.Controls
                ClearTextBox(ctrl)
                If TypeOf ctrl Is TextBox Then
                    CType(ctrl, TextBox).Text = String.Empty
                End If
            Next ctrl
        Catch ex As Exception
            MsgBox(ex.ToString)
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub ClearDataGridViewRows(ByVal root As Control)
        Try
            For Each ctrl As Control In root.Controls
                ClearDataGridViewRows(ctrl)
                If TypeOf ctrl Is DataGridView Then
                    CType(ctrl, DataGridView).DataSource = Nothing
                    CType(ctrl, DataGridView).Rows.Clear()
                End If
            Next ctrl
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub SpecialCharRestriction(ByVal root As Control, ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Try
            For Each ctrl As Control In root.Controls
                SpecialCharRestriction(ctrl, sender, e)
                If TypeOf ctrl Is TextBox Then
                    Dim allowedChars As String = "[`~!@#\$%\^&\*\(\)_\-\+=\{\}\[\]\\\|:;""'<>,\.\?/"
                    If Not allowedChars.IndexOf(e.KeyChar) = -1 Then
                        e.Handled = True
                    End If
                End If
            Next ctrl
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub buttonpressedenter(ByVal btntext As String)
        If Val(POS.TextBoxQTY.Text) <> 0 Then
            POS.TextBoxQTY.Text += btntext
        Else
            POS.TextBoxQTY.Text = btntext
        End If
    End Sub
    Public Sub buttonpressedenterpayment(ByVal btntext As String)
        If Val(PaymentForm.TextBoxMONEY.Text) <> 0 Then
            PaymentForm.TextBoxMONEY.Text += btntext
        Else
            PaymentForm.TextBoxMONEY.Text = btntext
        End If
    End Sub
    Public Sub btnformcolor(ByVal changecolor As Button)
        changecolor.BackColor = Color.FromArgb(23, 162, 184)
    End Sub
    Public Sub btndefaut(ByVal defaultcolor As Button, ByVal form As Form)
        For Each P As Control In form.Controls
            If TypeOf P Is Panel Then
                For Each ctrl As Control In P.Controls
                    If TypeOf ctrl Is Button Then
                        If ctrl.Name <> defaultcolor.Name Then
                            CType(ctrl, Button).BackColor = Color.FromArgb(41, 39, 40)
                        End If
                    End If
                Next
            End If
        Next
    End Sub
    Public Function ConvertToBase64(str As String)
        Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(str)
        Dim byt2 = Convert.ToBase64String(byt)
        Return byt2
    End Function
    Public Function ConvertB64ToString(str As String)
        Dim b As Byte() = Convert.FromBase64String(str)
        Dim byt2 = System.Text.Encoding.UTF8.GetString(b)
        Return byt2
    End Function
    'Dim MyPublicIpAddress As String
    'Public Function PublicIpAddress(myform As Form)
    '    Dim client As New WebClient
    '    '// Add a user agent header in case the requested URI contains a query.
    '    client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR1.0.3705;)")
    '    Dim baseurl As String = "http://localhost/ipadd/"
    '    ' with proxy server only:
    '    Dim proxy As IWebProxy = WebRequest.GetSystemWebProxy()
    '    proxy.Credentials = CredentialCache.DefaultNetworkCredentials
    '    client.Proxy = proxy
    '    Dim data As Stream
    '    Try
    '        data = client.OpenRead(baseurl)
    '    Catch ex As Exception
    '        MsgBox("open url " & ex.Message)
    '        Exit Function
    '    End Try
    '    Dim reader As StreamReader = New StreamReader(data)
    '    MyPublicIpAddress = reader.ReadToEnd()
    '    data.Close()
    '    reader.Close()
    '    MyPublicIpAddress = MyPublicIpAddress.Replace("<html><head><title>Current IP Check</title></head><body>", "").Replace("</body></html>", "").ToString()
    '    myform.Text = MyPublicIpAddress
    '    Return MyPublicIpAddress
    'End Function

    Public Function RemoveCharacter(ByVal stringToCleanUp, ByVal characterToRemove)
        ' replace the target with nothing
        ' Replace() returns a new String and does not modify the current one
        Return stringToCleanUp.Replace(characterToRemove, "")
    End Function
    '_________________________________________________________________________________________________________________
    'IMAGE TO TEXT
    Public Function ImageToBase64(ByVal image As Image, ByVal format As System.Drawing.Imaging.ImageFormat) As String
        Dim base64String As String = ""
        Try
            Using ms As New MemoryStream()
                ' Convert Image to byte[]
                image.Save(ms, format)
                Dim imageBytes As Byte() = ms.ToArray()
                ' Convert byte[] to Base64 String
                base64String = Convert.ToBase64String(imageBytes)
            End Using
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return base64String
    End Function
    'TEXT TO IMAGE
    Public Function Base64ToImage(ByVal base64String As String) As Image
        ' Convert Base64 String to byte[]
        Dim imageBytes As Byte() = Convert.FromBase64String(base64String)
        Dim ms As New MemoryStream(imageBytes, 0, imageBytes.Length)
        ' Convert byte[] to Image
        ms.Write(imageBytes, 0, imageBytes.Length)
        Dim ConvertedBase64Image As Image = Image.FromStream(ms, True)
        Return ConvertedBase64Image
    End Function
    Private ImagePath As String = ""
    '_________________________________________________________________________________________________________________
    'POS FUNCTIONS
    Public Function GetHash(theInput As String) As String
        Using hasher As MD5 = MD5.Create()    ' create hash object
            ' Convert to byte array and get hash
            Dim dbytes As Byte() =
             hasher.ComputeHash(Encoding.UTF8.GetBytes(theInput))
            ' sb to create string from bytes
            Dim sBuilder As New StringBuilder()
            ' convert byte data to hex string
            For n As Integer = 0 To dbytes.Length - 1
                sBuilder.Append(dbytes(n).ToString("x2"))
            Next n
            Return sBuilder.ToString()
        End Using
    End Function
    Public Function ConvertPassword(ByVal SourceString As String)
        Dim ConvertedString As String
        Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(SourceString)
        ConvertedString = Convert.ToBase64String(byt)
        Using md5Hash As MD5 = MD5.Create()
            hashable = GetHash(ConvertedString)
        End Using
        Return hashable
    End Function
    Public Function CheckForInternetConnection() As Boolean
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("https://www.google.com/")
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function
    Public Function GetMonthName(dat As Date) As String
        Dim iMonth As Integer = Month(dat)
        GetMonthName = MonthName(iMonth)
    End Function
    Public resetinventory As Boolean
    Public Function FirstDayOfMonth(ByVal sourceDate As DateTime)
        Dim displaythis = ""
        Try
            Dim FirstDay As DateTime = New DateTime(sourceDate.Year, sourceDate.Month, 1)
            Dim FormatDay As String = "yyyy-MM-dd"
            displaythis = FirstDay.ToString(FormatDay)
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return displaythis
    End Function
    Dim dtRESET As DataTable
    Public Function CheckIfNeedToReset() As Boolean
        Try
            Dim cmd As MySqlCommand
            Dim da As MySqlDataAdapter
            Dim firstday = FirstDayOfMonth(Date.Now)
            Try
                Dim sql = "SELECT * FROM loc_inv_temp_data WHERE created_at = '" & firstday & "'"
                cmd = New MySqlCommand(sql, LocalhostConn)
                da = New MySqlDataAdapter(cmd)
                dtRESET = New DataTable
                da.Fill(dtRESET)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        Catch ex As Exception
            MsgBox(ex.ToString)
            SendErrorReport(ex.ToString)
        End Try
        If dtRESET.Rows.Count = 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Dim DateNow
    Public Function FullDate24HR()
        Try
            DateNow = Format(Now(), "yyyy-MM-dd HH:mm:ss")
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return DateNow
    End Function
    Dim DateSave
    Public Function FullDateFormatForSaving()
        Try
            DateSave = Format(Now(), "yyyy-MM-dd HH-mm-ss")
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return DateSave
    End Function
    Public Sub EndBalance()
        Try
            If Shift = "First Shift" Then
                SystemLogType = "END-1"
                Dim DailySales = sum(table:="loc_daily_transaction_details WHERE created_at = '" & Format(Now(), "yyyy-MM-dd") & "' AND active = 1 AND store_id = '" & ClientStoreID & "' AND guid = '" & ClientGuid & "' ", tototal:="total")
                EndingBalance = BeginningBalance + Val(DailySales)
            ElseIf Shift = "Second Shift" Then
                SystemLogType = "END-2"
                Dim DailySales = sum(table:="loc_daily_transaction_details WHERE created_at = '" & Format(Now(), "yyyy-MM-dd") & "' AND active = 1 AND store_id = '" & ClientStoreID & "' AND guid = '" & ClientGuid & "' ", tototal:="total")
                EndingBalance = BeginningBalance + Val(DailySales)
            ElseIf Shift = "Third Shift" Then
                SystemLogType = "END-3"
                Dim DailySales = sum(table:="loc_daily_transaction_details WHERE created_at = '" & Format(Now(), "yyyy-MM-dd") & "' AND active = 1 AND store_id = '" & ClientStoreID & "' AND guid = '" & ClientGuid & "' ", tototal:="total")
                EndingBalance = BeginningBalance + Val(DailySales)
            Else
                SystemLogType = "END-4"
                Dim DailySales = sum(table:="loc_daily_transaction_details WHERE created_at = '" & Format(Now(), "yyyy-MM-dd") & "' AND active = 1 AND store_id = '" & ClientStoreID & "' AND guid = '" & ClientGuid & "' ", tototal:="total")
                EndingBalance = BeginningBalance + Val(DailySales)
            End If
            SystemLogDesc = EndingBalance
            GLOBAL_SYSTEM_LOGS(SystemLogType, SystemLogDesc)
            Shift = ""
            BeginningBalance = 0
            EndingBalance = 0
            BegBalanceBool = False
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Dim total
    Public Function SumOfColumnsToDecimal(ByVal datagrid As DataGridView, ByVal celltocompute As Integer)
        Dim SumTotal As Decimal
        Try
            With datagrid
                For i As Integer = 0 To .Rows.Count() - 1 Step +1
                    SumTotal = SumTotal + .Rows(i).Cells(celltocompute).Value
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return NUMBERFORMAT(SumTotal)
    End Function
    Public Function SumOfColumnsToInt(ByVal datagrid As DataGridView, ByVal celltocompute As Integer)
        Dim SumTotal As Integer
        Try
            With datagrid
                For i As Integer = 0 To .Rows.Count() - 1 Step +1
                    SumTotal += .Rows(i).Cells(celltocompute).Value
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return SumTotal
    End Function
    Public Sub Numeric(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        If e.KeyChar <> ControlChars.Back Then
            e.Handled = Not (Char.IsDigit(e.KeyChar) Or e.KeyChar = ".")
        End If
    End Sub
    Public Function NUMBERFORMAT(formatthis)
        Return Format(formatthis, "###,###,##0.00")
    End Function
    Dim ReturnRowIndex
    Public Function getCurrentCellButton_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        With POS
            If .DataGridViewOrders.Rows.Count > 0 Then
                ReturnRowIndex = .DataGridViewOrders.CurrentCell.RowIndex
            End If
        End With
        Return ReturnRowIndex
    End Function
    Public Sub RightToLeftDisplay(sender As Object, e As PrintPageEventArgs, position As Integer, lefttext As String, righttext As String, myfont As Font, wth As Single, frompoint As Single)
        Dim format As StringFormat = New StringFormat(StringFormatFlags.DirectionRightToLeft)
        Dim rect3 As RectangleF
        If My.Settings.PrintSize = "57mm" Then
            rect3 = New RectangleF(10.0F + frompoint, position, 173.0F + wth, 100.0F)
        Else
            rect3 = New RectangleF(10.0F + frompoint, position, 203.0F + wth, 100.0F)
        End If
        e.Graphics.DrawString(lefttext, myfont, Brushes.Black, rect3)
        e.Graphics.DrawString(righttext, myfont, Brushes.Black, rect3, format)
    End Sub
    Public Sub PrintStars(sender As Object, e As PrintPageEventArgs, FontDefault As Font, RowA As Integer)
        If My.Settings.PrintSize = "57mm" Then
            SimpleTextDisplay(sender, e, "*************************************", FontDefault, 0, RowA)
        Else
            SimpleTextDisplay(sender, e, "*********************************************", FontDefault, 0, RowA)
        End If
    End Sub
    Public Sub PrintLine(sender As Object, e As PrintPageEventArgs, FontDefault As Font, RowA As Integer)

        If My.Settings.PrintSize = "57mm" Then
            SimpleTextDisplay(sender, e, "----------------------------------------------------------------", FontDefault, 0, RowA)
        Else
            SimpleTextDisplay(sender, e, "----------------------------------------------------------------------", FontDefault, 0, RowA)
        End If
    End Sub
    Public Sub PrintSmallLine(sender As Object, e As PrintPageEventArgs, FontDefault As Font, RowA As Integer)

        If My.Settings.PrintSize = "57mm" Then
            SimpleTextDisplay(sender, e, "-------------------------------------------------------------", FontDefault, 0, RowA)
        Else
            SimpleTextDisplay(sender, e, "-------------------------------------------------------------------", FontDefault, 0, RowA)
        End If
    End Sub


    Public Sub RightDisplay1(sender As Object, e As PrintPageEventArgs, position As Integer, lefttext As String, righttext As String, myfont As Font, wth As Single, frompoint As Single)
        Dim format As StringFormat = New StringFormat(StringFormatFlags.DirectionRightToLeft)
        Dim rect3 As RectangleF = New RectangleF(10.0F + frompoint, position, 0 + wth, 0)
        e.Graphics.DrawString(lefttext, myfont, Brushes.Black, rect3)
        e.Graphics.DrawString(righttext, myfont, Brushes.Black, rect3, format)
    End Sub
    Public Sub RightDisplay(sender As Object, e As PrintPageEventArgs, position As Integer, righttext As String, myfont As Font, wth As Single, frompoint As Single)
        Dim format As StringFormat = New StringFormat(StringFormatFlags.DirectionRightToLeft)
        Dim rect3 As RectangleF = New RectangleF(10.0F + frompoint, position, 120.0F + wth, 100.0F)
        e.Graphics.DrawString(righttext, myfont, Brushes.Black, rect3, format)
    End Sub
    Public Sub CenterTextDisplay(sender As Object, e As PrintPageEventArgs, myText As String, myFont As Font, myPosition As Integer)
        Dim sngCenterPagebrand As Single
        sngCenterPagebrand = Convert.ToSingle(e.PageBounds.Width / 2 - e.Graphics.MeasureString(myText, myFont).Width / 2)
        e.Graphics.DrawString(myText, myFont, Brushes.Black, sngCenterPagebrand, myPosition)
    End Sub
    Public Sub SimpleTextDisplay(sender As Object, e As PrintPageEventArgs, myText As String, myFont As Font, ShopX As Integer, ShopY As Integer)
        Dim shopnameX As Integer = 10, shopnameY As Integer = 20
        e.Graphics.DrawString(myText, myFont, Brushes.Black, New PointF(shopnameX + ShopX, shopnameY + ShopY))
    End Sub
    Public Sub FormIsOpen()
        If Application.OpenForms().OfType(Of SynctoCloud).Any Then
            SynctoCloud.Close()
        End If
    End Sub
    Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Enum ProgressBarColor
        Green = &H1
        Red = &H2
        Yellow = &H3
    End Enum
    Public Sub ChangeProgBarColor(ByVal ProgressBar_Name As System.Windows.Forms.ProgressBar, ByVal ProgressBar_Color As ProgressBarColor)
        SendMessage(ProgressBar_Name.Handle, &H410, ProgressBar_Color, 0)
    End Sub
    Public Sub ReceiptHeader(sender As Object, e As PrintPageEventArgs, VoidReturn As Boolean)
        Try
            Dim brandfont As New Font("Tahoma", 7, FontStyle.Bold)
            Dim font As New Font("Tahoma", 6)
            Dim brand = ClientBrand.ToUpper
            Dim AddLine As Integer = 0
            CenterTextDisplay(sender, e, brand, brandfont, 10)

            If VoidReturn Then
                CenterTextDisplay(sender, e, "VOID/RETURN", font, 21)
                AddLine += 10
            End If

            CenterTextDisplay(sender, e, "VAT REG TIN " & ClientTin, font, 21 + AddLine)
            CenterTextDisplay(sender, e, "MSN : " & ClientMSN, font, 31 + AddLine)
            '============================================================================================================================
            CenterTextDisplay(sender, e, "MIN : " & ClientMIN, font, 41 + AddLine)
            '============================================================================================================================
            CenterTextDisplay(sender, e, "PTUN : " & ClientPTUN, font, 51 + AddLine)
            CenterTextDisplay(sender, e, ClientAddress, font, 61 + AddLine)
            CenterTextDisplay(sender, e, ClientBrgy & ", ", font, 71 + AddLine)
            CenterTextDisplay(sender, e, getmunicipality & ", " & getprovince, font, 81 + AddLine)
            CenterTextDisplay(sender, e, "TEL. NO.: " & ClientTel, font, 91 + AddLine)



            If SENIORDETAILSBOOL = True Then
                SimpleTextDisplay(sender, e, SeniorDetailsName & " - " & SeniorDetailsID, font, 30, 82 + AddLine)
            End If


            If My.Settings.PrintSize = "57mm" Then
                e.Graphics.DrawLine(Pens.Black, 40, 112 + AddLine, 180, 112 + AddLine)
                e.Graphics.DrawLine(Pens.Black, 28, 122 + AddLine, 180, 122 + AddLine)
                e.Graphics.DrawLine(Pens.Black, 49, 132 + AddLine, 180, 132 + AddLine)
                e.Graphics.DrawLine(Pens.Black, 75, 142 + AddLine, 180, 142 + AddLine)
            Else
                e.Graphics.DrawLine(Pens.Black, 40, 112 + AddLine, 210, 112 + AddLine)
                e.Graphics.DrawLine(Pens.Black, 28, 122 + AddLine, 210, 122 + AddLine)
                e.Graphics.DrawLine(Pens.Black, 49, 132 + AddLine, 210, 132 + AddLine)
                e.Graphics.DrawLine(Pens.Black, 75, 142 + AddLine, 210, 142 + AddLine)
            End If

            SimpleTextDisplay(sender, e, "Name:", font, 0, 85 + AddLine)
            SimpleTextDisplay(sender, e, "Tin:", font, 0, 95 + AddLine)
            SimpleTextDisplay(sender, e, "Address:", font, 0, 105 + AddLine)
            SimpleTextDisplay(sender, e, "Business Style:", font, 0, 115 + AddLine)

        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub ReceiptFooter(sender As Object, e As PrintPageEventArgs, a As Integer, ItemReturn As Boolean)
        Try
            Dim sql As String = "SELECT `Dev_Company_Name`, `Dev_Address`, `Dev_Tin`, `Dev_Accr_No`, `Dev_Accr_Date_Issued`, `Dev_Accr_Valid_Until`, `Dev_PTU_No`, `Dev_PTU_Date_Issued`, `Dev_PTU_Valid_Until` FROM loc_settings WHERE settings_id = 1"
            Dim cmd As MySqlCommand = New MySqlCommand(sql, LocalhostConn())
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            Dim brandfont As New Font("Tahoma", 7, FontStyle.Bold)
            Dim font As New Font("Tahoma", 6)
            Dim font1 As New Font("Tahoma", 5, FontStyle.Bold)
            CenterTextDisplay(sender, e, dt(0)(0).ToUpper, brandfont, a + 200)
            CenterTextDisplay(sender, e, "VAT REG TIN : " & dt(0)(2).ToString, font, a + 210)
            CenterTextDisplay(sender, e, dt(0)(1), font, a + 220)
            CenterTextDisplay(sender, e, "ACCR # : " & dt(0)(3), font, a + 230)
            CenterTextDisplay(sender, e, "DATE ISSUED : " & dt(0)(4), font, a + 240)
            CenterTextDisplay(sender, e, "VALID UNTIL : " & dt(0)(5), font, a + 250)
            CenterTextDisplay(sender, e, "PERMIT TO OPERATE : " & dt(0)(6), font, a + 260)
            CenterTextDisplay(sender, e, "DATE ISSUED : " & dt(0)(7), font, a + 270)
            CenterTextDisplay(sender, e, "VALID UNTIL : " & dt(0)(8), font, a + 280)
            CenterTextDisplay(sender, e, "THIS INVOICE/RECEIPT SHALL BE ", font1, a + 290)
            CenterTextDisplay(sender, e, "VALID FOR FIVE(5) YEARS FROM THE DATE ", font1, a + 300)
            CenterTextDisplay(sender, e, "OF THE PERMIT TO USE", font1, a + 310)

            If ItemReturn Then
                CenterTextDisplay(sender, e, "THIS DOCUMENT SHALL BE ", font1, a + 330)
                CenterTextDisplay(sender, e, "VALID FOR FIVE(5) YEARS FROM THE DATE ", font1, a + 340)
                CenterTextDisplay(sender, e, "OF THE PERMIT TO USE", font1, a + 350)
                CenterTextDisplay(sender, e, "THIS DOCUMENT IS NOT ", font1, a + 360)
                CenterTextDisplay(sender, e, "VALID FOR CLAIM OF INPUT TAX", font1, a + 370)
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Public Function TwoDecimalPlaces(ToRound)
        Try
            ToRound = Math.Round(ToRound, 2, MidpointRounding.AwayFromZero)
        Catch ex As Exception
            MsgBox(ex.ToString)
            SendErrorReport(ex.ToString)
        End Try
        Return ToRound
    End Function
    Public Sub Compute()
        With POS
            .Label76.Text = SumOfColumnsToDecimal(.DataGridViewOrders, 3)
            .TextBoxSUBTOTAL.Text = .Label76.Text
            .TextBoxGRANDTOTAL.Text = Double.Parse(.Label76.Text) - Double.Parse(.TextBoxDISCOUNT.Text)
            .TextBoxGRANDTOTAL.Text = NUMBERFORMAT(Double.Parse(.TextBoxGRANDTOTAL.Text))
            .TextBoxDISCOUNT.Text = NUMBERFORMAT(0)
        End With
    End Sub

    Public Sub LedDisplay(TextToDisplay As String, TotalOrChange As Boolean)
        Try
            Dim ComPort As String = My.Settings.SpPort
            Dim BaudRate As Integer = My.Settings.SpBaudrate

            If ComPort <> "" And BaudRate <> 0 Then
                Dim sp As New System.IO.Ports.SerialPort(ComPort, BaudRate, IO.Ports.Parity.None And IO.Ports.StopBits.One)
                sp.Open()
                sp.Write(Convert.ToString(ChrW(12)))
                If TotalOrChange Then
                    'Displays Price Amount AND word TOTAL in the pole display
                    sp.WriteLine(Chr(27) + Chr(81) + Chr(65) + TextToDisplay + Chr(13) + Chr(27) + Chr(115) + ”2”)
                Else
                    'Displays Price Amount AND word CHANGE in the pole display
                    sp.WriteLine(Chr(27) + Chr(81) + Chr(65) + TextToDisplay + Chr(13) + Chr(27) + Chr(115) + ”4”)
                End If
                ' sp.WriteLine(Chr(27) + Chr(115) + ”2”)
                sp.Close()
                sp.Dispose()
                sp = Nothing
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Public Sub LedConfig(TextToDisplay As String, ComPort As String, BaudRate As Integer)
        Try
            'Displays Price Amount in the pole display
            Dim sp As New System.IO.Ports.SerialPort(ComPort, BaudRate, IO.Ports.Parity.None And IO.Ports.StopBits.One)
            sp.Open()
            sp.Write(Convert.ToString(ChrW(12)))
            sp.WriteLine(Chr(27) + Chr(81) + Chr(65) + TextToDisplay + Chr(13))
            sp.Close()
            sp.Dispose()
            sp = Nothing
            Dim msg = MessageBox.Show("Does sample text displays on LED panel?", "LED DISPLAY CONFIGURATION", MessageBoxButtons.YesNo)
            If msg = DialogResult.Yes Then
                My.Settings.LedDisplayTrue = True
                My.Settings.Save()
            Else
                My.Settings.LedDisplayTrue = False
                My.Settings.Save()
            End If
        Catch ex As Exception
            My.Settings.LedDisplayTrue = False
            My.Settings.Save()
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub GetPorts(ToFill)
        Try
            ToFill.items.clear
            For Each sp As String In My.Computer.Ports.SerialPortNames
                ToFill.Items.Add(sp)
            Next
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Public Function ReturnPrintSize() As Integer
        Dim PrintSize = 0
        If My.Settings.PrintSize = "57mm" Then
            PrintSize = 205
        Else
            PrintSize = 235
        End If
        Return PrintSize
    End Function

    Public Function ReturnStringToDate(StringToDate) As DateTime
        Dim iDate As String = StringToDate
        Dim oDate As DateTime = Convert.ToDateTime(iDate)
        Return oDate.Year & "-" & oDate.Month & "-" & oDate.Day
    End Function
End Module
