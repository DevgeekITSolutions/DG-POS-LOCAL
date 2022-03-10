Imports MySql.Data.MySqlClient
Imports System.Drawing.Printing
Imports System.Threading
Imports System.Data
Imports System.Linq
Public Class POS
    Private WithEvents printdoc As PrintDocument = New PrintDocument
    Private PrintPreviewDialog1 As New PrintPreviewDialog

    Private Location_control As New Point(0, 0)

    Public ButtonClickCount As Integer = 0

    Dim RowA = 0
    Dim RowB = 0


    Public SUPERAMOUNTDUE
    Public WaffleUpgrade As Boolean = False

    Private Shared _instance As POS

    Dim Font1Bold As New Font("Tahoma", 6, FontStyle.Bold)
    Dim Font2Bold As New Font("Tahoma", 7, FontStyle.Bold)
    Dim FontDefault As New Font("Tahoma", 6)
    Dim FontAddOn As New Font("Tahoma", 5)
    Public ReadOnly Property Instance As POS
        Get
            Return _instance
        End Get
    End Property
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _instance = Me
        ChangeProgBarColor(ProgressBar1, ProgressBarColor.Yellow)
        LabelFOOTER.Text = My.Settings.Footer
        Try
            If Application.OpenForms().OfType(Of SynctoCloud).Any Then
                SynctoCloud.BringToFront()
            End If

            LabelStorename.Text = ClientStorename
            Label11.Focus()
            Timer1.Start()

            selectmax(whatform:=1)
            DataGridViewOrders.Font = New Font("Tahoma", 11.25)
            LoadCategory()
            For Each btn As Button In Panel3.Controls.OfType(Of Button)()
                If btn.Text = "Simply Perfect" Then
                    btn.PerformClick()
                End If
            Next
            DataGridViewOrders.CellBorderStyle = DataGridViewCellBorderStyle.None


            If CheckForInternetConnection() Then
                Enabled = False
                CheckingForUpdates.Show()
                CheckingForUpdates.TopMost = True

                If ValidCloudConnection = True Then
                    GetRowCount()
                    BackgroundWorker2.WorkerReportsProgress = True
                    BackgroundWorker2.WorkerSupportsCancellation = True
                    BackgroundWorker2.RunWorkerAsync()
                Else
                    Enabled = False
                    BegBalance.Show()
                    BegBalance.TopMost = True
                    'CheckingForUpdates.LabelCheckingUpdates.Text = "Invalid cloud server connection."
                End If
            Else
                Enabled = False
                BegBalance.Show()
                BegBalance.TopMost = True
            End If

        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Public Sub LoadCategory()
        Try
            Panel3.Controls.Clear()
            Location_control = New Point(0, 0)
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql = "SELECT category_name FROM loc_admin_category WHERE status = 1"
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable()
            da.Fill(dt)
            With cmd
                For Each row As DataRow In dt.Rows
                    Dim buttonname As String = row("category_name")
                    Dim new_Button As New Button
                    Dim panellocation As New Panel
                    With new_Button
                        .Name = buttonname
                        .Text = buttonname
                        .TextImageRelation = TextImageRelation.ImageBeforeText
                        .TextAlign = ContentAlignment.MiddleCenter
                        .ForeColor = Color.White
                        .Font = New Font("Tahoma", 9, FontStyle.Bold)
                        .FlatStyle = FlatStyle.Flat
                        .FlatAppearance.BorderSize = 0
                        .Location = New Point(Location_control.X, Location_control.Y)
                        .Width = 120
                        .Height = 53
                        .Cursor = Cursors.Hand
                        Location_control.X += .Height + 65
                        AddHandler .Click, AddressOf new_Button_click_category
                    End With
                    Panel3.Controls.Add(new_Button)
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub new_Button_click_category(ByVal sender As Object, ByVal e As EventArgs)
        Try
            'NEW BUTTON ON CLICK EVENT 
            If TypeOf sender Is Button Then
                Dim btn = sender
                Dim name = btn.name
                btnformcolor(changecolor:=sender)
                btndefaut(defaultcolor:=sender, form:=Me)

                If name = "Others" Or name = "Famous Blends" Or name = "Add-Ons" Then
                    WaffleUpgrade = False
                    ButtonWaffleUpgrade.Text = "Brownie Upgrade"
                    ButtonWaffleUpgrade.BackColor = Color.FromArgb(221, 114, 46)
                    ButtonWaffleUpgrade.Enabled = False
                Else
                    ButtonWaffleUpgrade.Enabled = True
                End If
                listviewproductsshow(where:=name)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub ButtonLogout_Click(sender As Object, e As EventArgs) Handles ButtonLogout.Click
        'LOGOUT
        If SyncIsOnProcess = True Then
            MessageBox.Show("Sync is on process please wait.", "Syncing", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            If MessageBox.Show("Are you sure you really want to Logout ?", "Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
                Enabled = False
                LOGOUTFROMPOS = True
                CashBreakdown.Show()
            End If
        End If
    End Sub
    Private Sub ButtonSettings_Click(sender As Object, e As EventArgs) Handles ButtonSettings.Click
        SettingsForm.Show()
        Enabled = False
    End Sub
    Private Sub ButtonExpense_Click(sender As Object, e As EventArgs) Handles ButtonExpense.Click
        Enabled = False
        Dim newMDIchild As New Addexpense()
        If Application.OpenForms().OfType(Of Addexpense).Any Then
            Addexpense.BringToFront()
        Else
            Addexpense.Show()
            Addexpense.Focus()
        End If
        'VIEW EXPENSE FORM
    End Sub
    Private Sub ButtonMenu_Click(sender As Object, e As EventArgs) Handles ButtonMenu.Click
        'VIEW MENU FORM
        messageboxappearance = False
        SystemLogType = "MENU FORM"
        SystemLogDesc = "Accessed by :" & returnfullname(ClientCrewID) & " : " & ClientRole
        GLOBAL_SYSTEM_LOGS(SystemLogType, SystemLogDesc)
        Enabled = False
        MDIFORM.Show()
    End Sub
    Private Sub ButtonPromo_Click(sender As Object, e As EventArgs) Handles ButtonPromo.Click
        'VIEW PROMO FORM
        Me.Enabled = False
        If Application.OpenForms().OfType(Of CouponCode).Any Then
            CouponCode.BringToFront()
        Else
            CouponCode.Show()
            CouponCode.ButtonSubmit.Enabled = False
        End If
    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles ButtonEnter.Click
        Try
            If payment = False Then
                If Val(TextBoxQTY.Text) > 0 Then
                    Dim Tax = 1 + Val(S_Tax)
                    Dim TotalProductPrice As Double = 0
                    Dim productprice = DataGridViewOrders.SelectedRows(0).Cells(2).Value
                    'Procedure: 1 Product Qty
                    If DataGridViewOrders.Rows.Count > 0 Then
                        If S_ZeroRated = "0" Then
                            'Price not / by 1.12
                            If WaffleUpgrade Then
                                'Price plus waffle upgrade price
                                Dim TotalPrice As Integer = 0
                                TotalPrice = Val(TextBoxQTY.Text) * Val(productprice)
                                Dim TotalUpgrade As Integer = 0
                                TotalUpgrade = Val(TextBoxQTY.Text) * Val(S_Upgrade_Price)
                                TotalProductPrice = TwoDecimalPlaces(TotalPrice + TotalUpgrade)
                                DataGridViewOrders.SelectedRows(0).Cells(11).Value = TextBoxQTY.Text
                            Else
                                Dim TotalPrice As Integer = 0
                                TotalPrice = Val(TextBoxQTY.Text) * Val(productprice)
                                TotalProductPrice = TwoDecimalPlaces(TotalPrice)
                            End If
                            DataGridViewOrders.SelectedRows(0).Cells(1).Value = TextBoxQTY.Text
                            DataGridViewOrders.SelectedRows(0).Cells(3).Value = TotalProductPrice

                            Compute()

                            For i As Integer = 0 To DataGridViewInv.Rows.Count - 1 Step +1
                                If DataGridViewOrders.SelectedRows(0).Cells(0).Value = DataGridViewInv.Rows(i).Cells(4).Value Then
                                    DataGridViewInv.Rows(i).Cells(2).Value = Val(TextBoxQTY.Text)
                                    DataGridViewInv.Rows(i).Cells(0).Value = DataGridViewInv.Rows(i).Cells(2).Value * DataGridViewInv.Rows(i).Cells(5).Value
                                    DataGridViewInv.Rows(i).Cells(6).Value = DataGridViewInv.Rows(i).Cells(2).Value * DataGridViewInv.Rows(i).Cells(7).Value
                                End If
                            Next

                        Else
                            If WaffleUpgrade Then
                                Dim TotalPrice As Integer = 0
                                TotalPrice = Val(TextBoxQTY.Text) * Val(productprice)
                                Dim TotalUgrade As Integer = 0
                                TotalUgrade = Val(TextBoxQTY.Text) * Val(S_Upgrade_Price)
                                Dim WaffleAddPriceTotal = TotalPrice + TotalUgrade
                                TotalProductPrice = TwoDecimalPlaces(WaffleAddPriceTotal / Tax)
                                DataGridViewOrders.SelectedRows(0).Cells(11).Value = TextBoxQTY.Text
                            Else
                                Dim TotalPrice As Integer = 0
                                TotalPrice = Val(TextBoxQTY.Text) * Val(productprice)
                                TotalProductPrice = TwoDecimalPlaces(TotalPrice / Tax)
                            End If
                            DataGridViewOrders.SelectedRows(0).Cells(1).Value = TextBoxQTY.Text
                            DataGridViewOrders.SelectedRows(0).Cells(3).Value = TotalProductPrice

                            Compute()

                            For i As Integer = 0 To DataGridViewInv.Rows.Count - 1 Step +1
                                If DataGridViewOrders.Rows(0).Cells(0).Value = DataGridViewInv.Rows(i).Cells(4).Value Then
                                    DataGridViewInv.Rows(i).Cells(2).Value = Val(TextBoxQTY.Text)
                                    DataGridViewInv.Rows(i).Cells(0).Value = DataGridViewInv.Rows(i).Cells(2).Value * DataGridViewInv.Rows(i).Cells(5).Value
                                    DataGridViewInv.Rows(i).Cells(6).Value = DataGridViewInv.Rows(i).Cells(2).Value * DataGridViewInv.Rows(i).Cells(7).Value
                                End If
                            Next
                        End If

                    End If
                End If
                TextBoxQTY.Text = 0
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles ButtonPendingOrders.Click
        Dim newMDIchild As New PendingOrders()
        If Application.OpenForms().OfType(Of PendingOrders).Any Then
            PendingOrders.BringToFront()
        Else
            PendingOrders.Show()
            posandpendingenter = True
            Me.Enabled = False
        End If
    End Sub
    Private Sub Buttonholdoder_Click(sender As Object, e As EventArgs) Handles Buttonholdoder.Click
        If Application.OpenForms().OfType(Of HoldOrder).Any Then
            HoldOrder.BringToFront()
        Else
            HoldOrder.Show()
            Me.Enabled = False
        End If
    End Sub
    Private Sub ButtonPay_Click(sender As Object, e As EventArgs) Handles ButtonPayMent.Click
        Try
            If Double.Parse(TextBoxGRANDTOTAL.Text) <= 999999999.99 Then
                If ButtonPayMent.Text = "Checkout" Then
                    'If Shift = "" Then
                    '    MessageBox.Show("Input cashier balance first", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    'Else
                    If S_Zreading = Format(Now(), "yyyy-MM-dd") Or S_Zreading = Format(Now().AddDays(1), "yyyy-MM-dd") Then
                            Enabled = False
                            PaymentForm.Show()
                            Application.DoEvents()
                            PaymentForm.TextBoxMONEY.Focus()
                            PaymentForm.TextBoxTOTALPAY.Text = TextBoxGRANDTOTAL.Text
                            PaymentForm.Focus()
                        Else
                            MessageBox.Show("Z-read first", "Z-Reading", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    'If S_Zreading <> Format(Now(), "yyyy-MM-dd") Or Format(Now(), "yyyy-MM-dd") > Format(ReturnStringToDate(S_Zreading).AddDays(1), "yyyy-MM-dd") Then
                    'Else

                    'End If
                    'If S_Zreading <> Format(Now(), "yyyy-MM-dd") Then
                    '    
                    'Else

                    'End If
                    'End If
                Else
                    BackgroundWorker3.WorkerReportsProgress = True
                    BackgroundWorker3.WorkerSupportsCancellation = True
                    BackgroundWorker3.RunWorkerAsync()
                End If
            Else
                MsgBox("Maximum sales capacity already reached. Please contact your administrator for immediate solution.")
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub ButtonWaffleUpgrade_Click(sender As Object, e As EventArgs) Handles ButtonWaffleUpgrade.Click
        Try
            If WaffleUpgrade = False Then
                WaffleUpgrade = True
                ButtonWaffleUpgrade.Text = "Classic Waffle"
                ButtonWaffleUpgrade.BackColor = Color.Brown
            Else
                WaffleUpgrade = False
                ButtonWaffleUpgrade.Text = "Brownie Upgrade"
                ButtonWaffleUpgrade.BackColor = Color.FromArgb(221, 114, 46)

            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Mix()
        Try
            Dim inventory_id As Integer = 0
            Dim totalQuantity As Integer = 0
            Dim Ingredient As String = ""
            Dim Query As String = ""
            Dim SqlCommand As MySqlCommand
            Dim SqlAdapter As MySqlDataAdapter
            Dim SqlDt As DataTable = New DataTable

            Dim FORMPrimaryval As Double = 0
            Dim FORMSecondaryval As Double = 0
            Dim FORMServingval As Double = 0
            Dim FORMNoofservings As Double = 0

            Dim TotalPrimaryVal As Double = 0
            Dim TotalSecondaryVal As Double = 0
            Dim TotalNoOfServings As Double = 0

            Dim RetStockSec As Double = 0
            Dim RetStockPrim As Double = 0
            Dim RetNoServ As Double = 0
            Dim ProductID As Integer = 0
            With DataGridViewOrders
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    ProductID = .Rows(i).Cells(5).Value
                    inventory_id = .Rows(i).Cells(10).Value
                    totalQuantity = .Rows(i).Cells(1).Value
                    Ingredient = .Rows(i).Cells(0).Value
                    'Console.WriteLine("INV ID - " & inventory_id)
                    If .Rows(i).Cells(14).Value > 0 Then
                        Query = "SELECT `primary_value`, `secondary_value`, `serving_value`, `no_servings` FROM loc_product_formula WHERE server_formula_id = " & inventory_id
                        ' Console.WriteLine("HALF BATCH" & Query)
                        SqlCommand = New MySqlCommand(Query, LocalhostConn)
                        SqlAdapter = New MySqlDataAdapter(SqlCommand)
                        SqlDt = New DataTable
                        SqlAdapter.Fill(SqlDt)
                        For Each row As DataRow In SqlDt.Rows
                            FORMPrimaryval = row("primary_value") / 2
                            FORMSecondaryval = row("secondary_value") / 2
                            FORMServingval = row("serving_value") / 2
                            FORMNoofservings = row("no_servings") / 2

                            '  Console.WriteLine("INVENTORY PV - " & FORMPrimaryval & ", SV - " & FORMSecondaryval & ", SERVAL - " & FORMServingval & ", NO. SERV - " & FORMNoofservings)
                        Next
                    Else
                        Query = "SELECT `primary_value`, `secondary_value`, `serving_value`, `no_servings` FROM loc_product_formula WHERE server_formula_id = " & inventory_id
                        SqlCommand = New MySqlCommand(Query, LocalhostConn)
                        SqlAdapter = New MySqlDataAdapter(SqlCommand)
                        SqlDt = New DataTable
                        SqlAdapter.Fill(SqlDt)
                        For Each row As DataRow In SqlDt.Rows
                            FORMPrimaryval = row("primary_value")
                            FORMSecondaryval = row("secondary_value")
                            FORMServingval = row("serving_value")
                            FORMNoofservings = row("no_servings")

                            '  Console.WriteLine("INVENTORY PV - " & FORMPrimaryval & ", SV - " & FORMSecondaryval & ", SERVAL - " & FORMServingval & ", NO. SERV - " & FORMNoofservings)
                        Next
                    End If



                    TotalPrimaryVal = totalQuantity * FORMPrimaryval
                    TotalSecondaryVal = FORMSecondaryval * totalQuantity
                    TotalNoOfServings = FORMNoofservings * totalQuantity

                    '  Console.WriteLine("TOTAL PV - " & TotalPrimaryVal & ", TOTAL SEC - " & TotalSecondaryVal & ", TOTAL NO. SERV - " & TotalNoOfServings)

                    If .Rows(i).Cells(14).Value > 0 Then
                        Query = "SELECT `stock_primary`,`stock_secondary`,`stock_no_of_servings` FROM `loc_pos_inventory` WHERE server_inventory_id = " & inventory_id
                    Else
                        Query = "SELECT `stock_primary`,`stock_secondary`,`stock_no_of_servings` FROM `loc_pos_inventory` WHERE server_inventory_id = " & inventory_id
                    End If

                    SqlCommand = New MySqlCommand(Query, LocalhostConn)
                    SqlAdapter = New MySqlDataAdapter(SqlCommand)
                    SqlDt = New DataTable
                    SqlAdapter.Fill(SqlDt)
                    For Each row As DataRow In SqlDt.Rows
                        RetStockPrim = row("stock_primary")
                        RetStockSec = row("stock_secondary")
                        RetNoServ = row("stock_no_of_servings")

                        ' Console.WriteLine("RetStockPrim - " & RetStockPrim & ", RetStockSec - " & RetStockSec & ", RetNoServ - " & RetNoServ)
                    Next

                    Dim TotalPrimary As Double = RetStockPrim + TotalPrimaryVal
                    Dim Secondary As Double = RetStockSec + TotalSecondaryVal
                    Dim ServingValue As Double = RetNoServ + TotalNoOfServings


                    ' Console.WriteLine("TOTAL Primary - " & TotalPrimary & ", TOTAL Secondary - " & Secondary & ", TOTAL ServingValue - " & ServingValue)

                    If .Rows(i).Cells(14).Value > 0 Then
                        Query = "UPDATE loc_pos_inventory SET `stock_secondary` = " & Secondary & " , `stock_no_of_servings` = " & ServingValue & " , `stock_primary` = " & TotalPrimary & ", `date_modified` = '" & FullDate24HR() & "' WHERE `server_inventory_id` = " & inventory_id
                    Else
                        Query = "UPDATE loc_pos_inventory SET `stock_secondary` = " & Secondary & " , `stock_no_of_servings` = " & ServingValue & " , `stock_primary` = " & TotalPrimary & ", `date_modified` = '" & FullDate24HR() & "' WHERE `server_inventory_id` = " & inventory_id
                    End If
                    ' Console.WriteLine(Query)
                    SqlCommand = New MySqlCommand(Query, LocalhostConn())
                    SqlCommand.ExecuteNonQuery()
                    GLOBAL_SYSTEM_LOGS("MIX", "MIXED : " & Ingredient & ", Crew : " & ClientCrewID)
                Next


            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Dim secondary_value As Double = 0
    Dim stock_secondary As Double = 0
    Private Sub UpdateInventory()
        Dim SqlCommand As MySqlCommand
        Dim SqlAdapter As MySqlDataAdapter
        Dim SqlDt As DataTable
        Dim UpdateInventoryCon As MySqlConnection = LocalhostConn()
        Try
            Dim Query As String = ""
            With DataGridViewInv
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    Dim TotalQuantity As Double = 0
                    Dim TotalServingValue As Double = 0
                    Dim Secondary As Double = 0
                    Dim ServingValue As Double = 0
                    Dim TotalPrimary As Double = 0
                    TotalQuantity = .Rows(i).Cells(2).Value
                    TotalServingValue = Double.Parse(.Rows(i).Cells(0).Value.ToString)
                    If .Rows(i).Cells(9).Value.ToString = "Server" Then
                        Query = "SELECT `secondary_value` FROM `loc_product_formula` WHERE server_formula_id = " & .Rows(i).Cells(1).Value
                    Else
                        Query = "SELECT `secondary_value` FROM `loc_product_formula` WHERE formula_id = " & .Rows(i).Cells(1).Value

                    End If
                    SqlCommand = New MySqlCommand(Query, UpdateInventoryCon)
                    SqlAdapter = New MySqlDataAdapter(SqlCommand)
                    SqlDt = New DataTable
                    SqlAdapter.Fill(SqlDt)
                    For Each row As DataRow In SqlDt.Rows
                        secondary_value = row("secondary_value")
                    Next
                    If .Rows(i).Cells(9).Value.ToString = "Server" Then
                        Query = "SELECT `stock_secondary` FROM `loc_pos_inventory` WHERE server_inventory_id = " & .Rows(i).Cells(1).Value
                    Else
                        Query = "SELECT `stock_secondary` FROM `loc_pos_inventory` WHERE inventory_id = " & .Rows(i).Cells(1).Value
                    End If
                    SqlCommand = New MySqlCommand(Query, UpdateInventoryCon)
                    SqlAdapter = New MySqlDataAdapter(SqlCommand)
                    SqlDt = New DataTable
                    SqlAdapter.Fill(SqlDt)
                    For Each row As DataRow In SqlDt.Rows
                        stock_secondary = row("stock_secondary")
                    Next

                    'Console.WriteLine(Double.Parse(.Rows(i).Cells(5).Value.ToString) - secondary_value)
                    Secondary = stock_secondary - TotalServingValue
                    ServingValue = Secondary / Double.Parse(.Rows(i).Cells(5).Value.ToString)
                    TotalPrimary = Secondary / secondary_value

                    'Console.WriteLine("TOTAL SECONDARY " & Secondary)
                    'Console.WriteLine("TOTAL SERVING " & ServingValue)
                    ' Console.WriteLine("TOTAL PRIMARY  " & TotalPrimary)

                    If .Rows(i).Cells(9).Value.ToString = "Server" Then
                        Query = "UPDATE loc_pos_inventory SET stock_secondary = " & Secondary & " , stock_no_of_servings = " & ServingValue & " , stock_primary = " & TotalPrimary & ", date_modified = '" & FullDate24HR() & "' WHERE server_inventory_id = " & .Rows(i).Cells(1).Value
                    Else
                        Query = "UPDATE loc_pos_inventory SET stock_secondary = " & Secondary & " , stock_no_of_servings = " & ServingValue & " , stock_primary = " & TotalPrimary & ", date_modified = '" & FullDate24HR() & "' WHERE inventory_id = " & .Rows(i).Cells(1).Value
                    End If
                    SqlCommand = New MySqlCommand(Query, UpdateInventoryCon)
                    SqlCommand.ExecuteNonQuery()
                Next
                UpdateInventoryCon.Close()
            End With
        Catch ex As Exception
            MsgBox(ex.ToString)
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    'Private Sub UpdateInventory()
    '    Dim SqlCommand As MySqlCommand
    '    Dim SqlAdapter As MySqlDataAdapter
    '    Dim SqlDt As DataTable
    '    Dim UpdateInventoryCon As MySqlConnection = LocalhostConn()
    '    Try
    '        Dim Query As String = ""
    '        With DataGridViewInv
    '            For i As Integer = 0 To .Rows.Count - 1 Step +1

    '                Dim GetStock_Primary As Double = 0
    '                Dim GetStock_Secondary As Double = 0
    '                Dim GetStock_NoOfServings As Double = 0

    '                Dim GetFormPrimaryValue As Double = 0
    '                Dim GetFormSecondaryValue As Double = 0
    '                Dim GetFormServingValue As Double = 0

    '                Dim TotalStockPrimary As Double = 0
    '                Dim TotalStockSecondary As Double = 0
    '                Dim TotaStockNoOfServings As Double = 0

    '                Dim FormulaID = .Rows(i).Cells(1).Value

    '                'Dim TotalServingValue As Double = Double.Parse(.Rows(i).Cells(0).Value.ToString)
    '                'Dim ServingValue As Double = 0
    '                'TotalQuantity = .Rows(i).Cells(2).Value

    '                If .Rows(i).Cells(9).Value.ToString = "Server" Then
    '                    Query = "SELECT `primary_value`,`secondary_value`,`serving_value` FROM `loc_product_formula` WHERE server_formula_id = " & .Rows(i).Cells(1).Value
    '                Else
    '                    Query = "SELECT `primary_value`,`secondary_value`,`serving_value` FROM `loc_product_formula` WHERE formula_id = " & .Rows(i).Cells(1).Value
    '                End If

    '                SqlCommand = New MySqlCommand(Query, UpdateInventoryCon)
    '                SqlAdapter = New MySqlDataAdapter(SqlCommand)
    '                SqlDt = New DataTable
    '                SqlAdapter.Fill(SqlDt)

    '                For Each row As DataRow In SqlDt.Rows
    '                    GetFormPrimaryValue = row("primary_value")
    '                    GetFormSecondaryValue = row("secondary_value")
    '                    GetFormServingValue = row("serving_value")
    '                Next
    '                Console.WriteLine("FORMULA PRIMARY " & GetFormPrimaryValue)

    '                If .Rows(i).Cells(9).Value.ToString = "Server" Then
    '                    Query = "SELECT `stock_primary`,`stock_secondary`,`stock_no_of_servings` FROM `loc_pos_inventory` WHERE server_inventory_id = " & FormulaID
    '                Else
    '                    Query = "SELECT `stock_primary`,`stock_secondary`,`stock_no_of_servings` FROM `loc_pos_inventory` WHERE inventory_id = " & FormulaID
    '                End If
    '                SqlCommand = New MySqlCommand(Query, UpdateInventoryCon)
    '                SqlAdapter = New MySqlDataAdapter(SqlCommand)
    '                SqlDt = New DataTable
    '                SqlAdapter.Fill(SqlDt)
    '                For Each row As DataRow In SqlDt.Rows
    '                    GetStock_Primary = row("stock_primary")
    '                    GetStock_Secondary = row("stock_secondary")
    '                    GetStock_NoOfServings = row("stock_no_of_servings")
    '                Next
    '                Console.WriteLine("GET STOCK SEC: " & GetStock_Secondary)
    '                'TotalStockPrimary = GetStock_Primary -
    '                TotalStockSecondary = GetStock_Secondary - Double.Parse(.Rows(i).Cells(0).Value)
    '                'TotaStockNoOfServings =  
    '                TotalStockPrimary = TotalStockSecondary / Double.Parse(.Rows(i).Cells(5).Value.ToString)
    '                'TotaStockNoOfServings
    '                ''
    '                ''    secondary_value = secondary_value / 2
    '                ''    secondary_value = secondary_value / 2
    '                ''End If
    '                'If .Rows(i).Cells(10).Value = 1 Then
    '                '    secondary_value = secondary_value / 2
    '                '    stock_secondary = stock_secondary / 2
    '                'End If

    '                'Secondary = stock_secondary - TotalServingValue
    '                'ServingValue = Secondary / Double.Parse(.Rows(i).Cells(5).Value.ToString)
    '                'TotalPrimary = Secondary / secondary_value



    '                Console.WriteLine("TOTAL SECONDARY " & TotalStockSecondary)
    '                Console.WriteLine("TOTAL SERVING " & TotaStockNoOfServings)
    '                Console.WriteLine("TOTAL PRIMARY  " & TotalStockPrimary)
    '                'If .Rows(i).Cells(9).Value.ToString = "Server" Then
    '                '    Query = "UPDATE loc_pos_inventory SET stock_secondary = " & Secondary & " , stock_no_of_servings = " & ServingValue & " , stock_primary = " & TotalPrimary & ", date_modified = '" & FullDate24HR() & "' WHERE server_inventory_id = " & .Rows(i).Cells(1).Value
    '                'Else
    '                '    Query = "UPDATE loc_pos_inventory SET stock_secondary = " & Secondary & " , stock_no_of_servings = " & ServingValue & " , stock_primary = " & TotalPrimary & ", date_modified = '" & FullDate24HR() & "' WHERE inventory_id = " & .Rows(i).Cells(1).Value
    '                'End If
    '                '' Console.WriteLine(Query)
    '                'SqlCommand = New MySqlCommand(Query, UpdateInventoryCon)
    '                'SqlCommand.ExecuteNonQuery()
    '            Next
    '            UpdateInventoryCon.Close()
    '        End With
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '        SendErrorReport(ex.ToString)
    '    End Try
    'End Sub
    Dim ThreadlistMIX As List(Of Thread) = New List(Of Thread)
    Dim ThreadMix As Thread
    Private Sub BackgroundWorker3_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker3.DoWork
        Try
            ThreadMix = New Thread(AddressOf Mix)
            ThreadMix.Start()
            ThreadlistMIX.Add(ThreadMix)
            For Each t In ThreadlistMIX
                t.Join()
            Next
            ThreadMix = New Thread(AddressOf UpdateInventory)
            ThreadMix.Start()
            ThreadlistMIX.Add(ThreadMix)
            For Each t In ThreadlistMIX
                t.Join()
            Next
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Private Sub BackgroundWorker3_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker3.RunWorkerCompleted
        MessageBox.Show("Ingredient Mixed", "Mix Products", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Panel3.Enabled = True
        ButtonPayMent.Enabled = False
        ButtonWaffleUpgrade.Enabled = True
        ButtonPayMent.Text = "Checkout"
        ButtonTransactionMode.Text = "Transaction Type"
        DataGridViewOrders.Rows.Clear()
        DataGridViewInv.Rows.Clear()
        DISABLESERVEROTHERSPRODUCT = False
    End Sub
    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles ButtonCancel.Click
        Try

            'ButtonCDISC.PerformClick()
            If DataGridViewOrders.Rows.Count > 0 Then
                Dim datas = DataGridViewOrders.SelectedRows(0).Cells(0).Value.ToString()
                For x As Integer = DataGridViewInv.Rows.Count - 1 To 0 Step -1
                    If DataGridViewInv.Rows(x).Cells("Column10").Value = datas Then
                        DataGridViewInv.Rows.Remove(DataGridViewInv.Rows(x))
                    End If
                Next
                datas = ""
                Deleteitem = True
                Dim dr As DataGridViewRow
                For Each dr In DataGridViewOrders.SelectedRows
                    Dim sum As String = DataGridViewOrders.SelectedRows(0).Cells(3).Value.ToString
                    DataGridViewOrders.Rows.Remove(dr)
                Next

            Else
                TextBoxQTY.Text = 0
                MessageBox.Show("Add item first", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            If DataGridViewOrders.Rows.Count > 0 Then
                ButtonPayMent.Enabled = True
                ButtonPendingOrders.Enabled = False
                Buttonholdoder.Enabled = True
            Else
                HASOTHERSLOCALPRODUCT = False
                HASOTHERSSERVERPRODUCT = False
                DISABLESERVEROTHERSPRODUCT = False
                Panel3.Enabled = True

                ButtonPayMent.Text = "Checkout"
                ButtonTransactionMode.Text = "Transaction Type"
                ButtonClickCount = 0
                ButtonPayMent.Enabled = False
                Buttonholdoder.Enabled = False
                ButtonPendingOrders.Enabled = True
            End If
            Compute()
            CouponDefault()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub Label76_TextChanged(sender As Object, e As EventArgs) Handles Label76.TextChanged
        Try
            If DataGridViewOrders.RowCount > 0 Then
                ButtonApplyCoupon.Enabled = True
            Else
                ButtonApplyCoupon.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub POS_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Expenses.Dispose()
        Couponisavailable = False
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Label11.Text = Date.Now.ToString("hh:mm:ss tt").ToUpper
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Private Sub CouponDefault()
        Try
            LESSVAT = 0
            VATEXEMPTSALES = 0
            SeniorDetailsID = ""
            SeniorDetailsName = ""
            SENIORDETAILSBOOL = False
            TextBoxDISCOUNT.Text = "0.00"
            CouponApplied = False
            CouponDesc = ""
            CouponTotal = 0
            CouponName = ""
            TOTALDISCOUNT = 0
            LESSVAT = 0
            DISCOUNTTYPE = "N/A"
            TOTALDISCOUNT = 0
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles ButtonTransactionMode.Click
        If ButtonTransactionMode.Text = "Transaction Type" Then
            Enabled = False
            TransactionType.Show()
        Else
            ButtonCancel.PerformClick()
            ButtonPayMent.Text = "Checkout"
            ButtonTransactionMode.Text = "Transaction Type"
            Panel3.Enabled = True
            ButtonWaffleUpgrade.Enabled = True
        End If
    End Sub
    Private Sub Button1_Click_3(sender As Object, e As EventArgs) Handles Button1.Click
        Enabled = False
        TakeOut.Show()
    End Sub
#Region "Button Functions"
    Private Sub POS_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
        If e.KeyCode = Keys.F9 Then
            ButtonPayMent.PerformClick()
        ElseIf e.KeyCode = Keys.Enter Then
            ButtonEnter.PerformClick()
        ElseIf e.KeyCode = Keys.F10 Then
            ButtonTransactionMode.PerformClick()
        ElseIf e.KeyCode = Keys.F11 Then
            Buttonholdoder.PerformClick()
        ElseIf e.KeyCode = Keys.F12 Then
            ButtonPendingOrders.PerformClick()
        ElseIf e.KeyCode = Keys.Delete Then
            ButtonCancel.PerformClick()
            '=================================
        End If
        If payment = False Then
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
        End If
    End Sub
    Private Sub ButtonNo9_Click(sender As Object, e As EventArgs) Handles ButtonNo9.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo9.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo9.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo8_Click(sender As Object, e As EventArgs) Handles ButtonNo8.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo8.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo8.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo7_Click(sender As Object, e As EventArgs) Handles ButtonNo7.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo7.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo7.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo6_Click(sender As Object, e As EventArgs) Handles ButtonNo6.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo6.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo6.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo5_Click(sender As Object, e As EventArgs) Handles ButtonNo5.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo5.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo5.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo4_Click(sender As Object, e As EventArgs) Handles ButtonNo4.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo4.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo4.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo3_Click(sender As Object, e As EventArgs) Handles ButtonNo3.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo3.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo3.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo2_Click(sender As Object, e As EventArgs) Handles ButtonNo2.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo2.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo2.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo1_Click(sender As Object, e As EventArgs) Handles ButtonNo1.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo1.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo1.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo0_Click(sender As Object, e As EventArgs) Handles ButtonNo0.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenter(btntext:=ButtonNo0.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 6 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo0.Text)
            End If
        End If
    End Sub
    Private Sub ButtonNo00_Click(sender As Object, e As EventArgs) Handles ButtonNo00.Click
        If payment = False Then
            If TextBoxQTY.Text.Length > 5 Then
            Else
                buttonpressedenter(btntext:=ButtonNo00.Text)
            End If
        Else
            If TextBoxQTY.Text.Length > 5 Then
            Else
                buttonpressedenterpayment(btntext:=ButtonNo00.Text)
            End If
        End If
    End Sub
    Private Sub Buttondot_Click(sender As Object, e As EventArgs) Handles Buttondot.Click
        If payment = False Then
            If Not TextBoxQTY.Text.Contains(".") Then
                TextBoxQTY.Text += "."
            End If
        End If
    End Sub
    Private Sub ButtonClear_Click(sender As Object, e As EventArgs) Handles ButtonClear.Click
        If payment = False Then
            TextBoxQTY.Text = 0
        End If
    End Sub
    Private Sub TextBoxQTY_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxQTY.KeyPress
        Try
            If InStr(DisallowedCharacters, e.KeyChar) > 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
#End Region
#Region "POS Coupon Application/ Print/ Transaction"
    Public DISCOUNTTYPE As String = "N/A"
    Public TOTALDISCOUNT As Double = 0
    Public GROSSSALE As Double = 0
    Public VATEXEMPTSALES As Double = 0
    Public LESSVAT As Double = 0
    Public TOTALDISCOUNTEDAMOUNT As Double = 0
    Public TOTALAMOUNTDUE As Double = 0
    Public VATABLESALES As Double = 0
    Public VAT12PERCENT As Double = 0
    Public ZERORATEDSALES As Double = 0
    Public ZERORATEDNETSALES As Double = 0
    Dim THREADLIST As List(Of Thread) = New List(Of Thread)
    Dim THREADLISTUPDATE As List(Of Thread) = New List(Of Thread)
    Dim TIMETOINSERT As String
    Dim ACTIVE As Integer = 1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonApplyCoupon.Click
        Try
            MessageBox.Show("Apply coupon after taking all customer orders", "NOTICE", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ButtonCDISC.PerformClick()
            Enabled = False
            GetHighest()
            CouponCode.Show()
            CouponCode.ButtonSubmit.Enabled = True
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub GetHighest()
        Try
            Dim HighestWafflesPrice As Double = 0
            Dim HighestDrinksPrice As Double = 0

            With DataGridViewOrders
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    If .Rows(i).Cells(9).Value.ToString = "WAFFLE" Then
                        If .Rows(i).Cells(11).Value > 0 Then
                            If HighestWafflesPrice < .Rows(i).Cells(2).Value + S_Upgrade_Price Then
                                HighestWafflesPrice = .Rows(i).Cells(2).Value + S_Upgrade_Price
                            End If
                        Else
                            If HighestWafflesPrice < .Rows(i).Cells(2).Value Then
                                HighestWafflesPrice = .Rows(i).Cells(2).Value
                            End If
                        End If
                    Else
                        'Console.WriteLine(.Rows(i).Cells(6).Value.ToString)
                        If .Rows(i).Cells(6).Value.ToString <> "BW20" And .Rows(i).Cells(6).Value.ToString <> "BW25" Then
                            'MsgBox(True)
                            If HighestDrinksPrice < .Rows(i).Cells(2).Value Then
                                HighestDrinksPrice = .Rows(i).Cells(2).Value
                            End If
                        Else
                            'MsgBox(False)
                        End If
                    End If
                Next
                'Console.WriteLine("HIGHEST: " & HighestWafflesPrice & " DRINKS " & SeniorPWdDrinks)
                SeniorPWd = HighestWafflesPrice
                SeniorPWdDrinks = HighestDrinksPrice
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

#Region "Transaction Process"

    Private Sub InsertFMStock()
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim cmd As MySqlCommand
            Dim sql = "INSERT INTO loc_fm_stock (`formula_id`, `stock_primary`, `stock_secondary`,`crew_id`, `store_id`, `guid`, `created_at`, `status`)
                       VALUES (@1,@2,@3,@4,@5,@6,@7,@8)"
            For i As Integer = 0 To DataGridViewInv.Rows.Count - 1 Step +1
                cmd = New MySqlCommand(sql, ConnectionLocal)
                cmd.Parameters.Add("@1", MySqlDbType.VarChar).Value = DataGridViewInv.Rows(i).Cells(1).Value
                cmd.Parameters.Add("@2", MySqlDbType.Decimal).Value = DataGridViewInv.Rows(i).Cells(2).Value
                cmd.Parameters.Add("@3", MySqlDbType.Decimal).Value = DataGridViewInv.Rows(i).Cells(0).Value
                cmd.Parameters.Add("@4", MySqlDbType.VarChar).Value = ClientCrewID
                cmd.Parameters.Add("@5", MySqlDbType.VarChar).Value = ClientStoreID
                cmd.Parameters.Add("@6", MySqlDbType.VarChar).Value = ClientGuid
                cmd.Parameters.Add("@7", MySqlDbType.Text).Value = FullDate24HR()
                cmd.Parameters.Add("@8", MySqlDbType.Int64).Value = 1
                cmd.ExecuteNonQuery()
            Next
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InsertDailyTransaction()
        Try

            Dim table As String = "loc_daily_transaction"
            Dim fields As String = " (`transaction_number`, `amounttendered`, `totaldiscount`, `change`, `amountdue`, `vatablesales`, `vatexemptsales`, `zeroratedsales`
                     , `lessvat`, `si_number`, `crew_id`, `guid`, `active`, `store_id`, `created_at`, `transaction_type`, `shift`, `zreading`, `synced`
                     , `discount_type`, `vatpercentage`, `grosssales`, `totaldiscountedamount`) "
            Dim NetSales As Double = 0
            If S_ZeroRated = "0" Then
                NetSales = SUPERAMOUNTDUE
            Else
                NetSales = ZERORATEDNETSALES
            End If

            If CouponName = "" Then
                CouponName = "N/A"
            End If

            Dim value As String = "('" & TextBoxMAXID.Text & "'," & TEXTBOXMONEYVALUE & "," & TOTALDISCOUNT & "," & TEXTBOXCHANGEVALUE & "," & NetSales & "," & VATABLESALES & "
                     ," & VATEXEMPTSALES & "," & ZERORATEDSALES & "," & LESSVAT & "," & SINumber & ",'" & ClientCrewID & "','" & ClientGuid & "','" & ACTIVE & "','" & ClientStoreID & "'
                     ,'" & INSERTTHISDATE & "','" & TRANSACTIONMODE & "','" & Shift & "','" & S_Zreading & "','Unsynced','" & CouponName & "'," & VAT12PERCENT & "," & GROSSSALE & "," & TOTALDISCOUNTEDAMOUNT & ")"
            GLOBAL_INSERT_FUNCTION(table, fields, value)
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InsertDailyDetails()
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim cmd As MySqlCommand
            Dim sql = "INSERT INTO loc_daily_transaction_details (`product_id`,`product_sku`,`product_name`,`quantity`,`price`,`total`,`crew_id`,`transaction_number`,`active`,`created_at`,`guid`,`store_id`,`synced`,`total_cost_of_goods`,`product_category`,`zreading`,`transaction_type`,`upgraded`,`addontype`) 
                       VALUES (@1,@2,@3,@4,@5,@6,@7,@8,@9,@10,@11,@12,@13,@14,@15,@16,@17,@18,@19)"
            For i As Integer = 0 To DataGridViewOrders.Rows.Count - 1 Step +1
                Dim totalcostofgoods As Decimal
                For a As Integer = 0 To DataGridViewInv.Rows.Count - 1 Step +1
                    If DataGridViewInv.Rows(a).Cells(4).Value = DataGridViewOrders.Rows(i).Cells(0).Value Then
                        totalcostofgoods += DataGridViewInv.Rows(a).Cells(6).Value
                    End If
                Next
                cmd = New MySqlCommand(sql, ConnectionLocal)
                cmd.Parameters.Add("@1", MySqlDbType.Int64).Value = DataGridViewOrders.Rows(i).Cells(5).Value
                cmd.Parameters.Add("@2", MySqlDbType.VarChar).Value = DataGridViewOrders.Rows(i).Cells(6).Value
                cmd.Parameters.Add("@3", MySqlDbType.VarChar).Value = DataGridViewOrders.Rows(i).Cells(0).Value
                cmd.Parameters.Add("@4", MySqlDbType.Int64).Value = DataGridViewOrders.Rows(i).Cells(1).Value
                cmd.Parameters.Add("@5", MySqlDbType.Decimal).Value = DataGridViewOrders.Rows(i).Cells(2).Value
                cmd.Parameters.Add("@6", MySqlDbType.Decimal).Value = DataGridViewOrders.Rows(i).Cells(3).Value
                cmd.Parameters.Add("@7", MySqlDbType.VarChar).Value = ClientCrewID
                cmd.Parameters.Add("@8", MySqlDbType.VarChar).Value = TextBoxMAXID.Text
                cmd.Parameters.Add("@9", MySqlDbType.Int64).Value = ACTIVE
                cmd.Parameters.Add("@10", MySqlDbType.Text).Value = INSERTTHISDATE
                cmd.Parameters.Add("@11", MySqlDbType.VarChar).Value = ClientGuid
                cmd.Parameters.Add("@12", MySqlDbType.VarChar).Value = ClientStoreID
                cmd.Parameters.Add("@13", MySqlDbType.VarChar).Value = "Unsynced"
                cmd.Parameters.Add("@14", MySqlDbType.Decimal).Value = totalcostofgoods
                cmd.Parameters.Add("@15", MySqlDbType.VarChar).Value = DataGridViewOrders.Rows(i).Cells(7).Value
                cmd.Parameters.Add("@16", MySqlDbType.Text).Value = S_Zreading
                cmd.Parameters.Add("@17", MySqlDbType.Text).Value = TRANSACTIONMODE
                cmd.Parameters.Add("@18", MySqlDbType.Int64).Value = DataGridViewOrders.Rows(i).Cells(11).Value
                cmd.Parameters.Add("@19", MySqlDbType.Text).Value = DataGridViewOrders.Rows(i).Cells(13).Value
                cmd.ExecuteNonQuery()
                totalcostofgoods = 0
            Next
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InsertModeofTransaction()
        Try

            Dim table As String = "loc_transaction_mode_details"
            Dim fields As String = "(`transaction_type`, `transaction_number`, `fullname`, `reference`, `status`, `synced`, `store_id`, `guid`, `created_at`)"
            Dim value As String = "( '" & TRANSACTIONMODE & "'
                            ,'" & TextBoxMAXID.Text & "'
                            , '" & TEXTBOXFULLNAMEVALUE & "'
                            , '" & TEXTBOXREFERENCEVALUE & "'
                            , " & 1 & "
                            , 'Unsynced'
                            , '" & ClientStoreID & "'
                            , '" & ClientGuid & "'
                            , '" & FullDate24HR() & "')"
            GLOBAL_INSERT_FUNCTION(table:=table, fields:=fields, values:=value)
            ButtonClickCount = 0
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InsertCouponData()
        Try
            Dim table As String = "loc_coupon_data"
            Dim fields As String = "(`transaction_number`, `coupon_name`, `coupon_type`, `coupon_desc`, `coupon_line`, `coupon_total`, `zreading`, `status`, `synced`)"
            Dim value As String = "( '" & TextBoxMAXID.Text & "'
                      ,'" & CouponName & "'
                      , '" & DISCOUNTTYPE & "'
                      , '" & CouponDesc & "'
                      , '" & CouponLine & "'
                      , '" & CouponTotal & "', '" & S_Zreading & "', '1', 'Unsynced')"
            GLOBAL_INSERT_FUNCTION(table, fields, value)


        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InsertSeniorDetails()
        Try
            Dim table As String = "loc_senior_details"
            Dim fields As String = "(`transaction_number`, `senior_id`, `senior_name`, `active`, `crew_id`, `store_id`, `guid`, `date_created`, `synced`)"
            Dim value As String = "( '" & TextBoxMAXID.Text & "'
                      , '" & SeniorDetailsID & "'
                      , '" & SeniorDetailsName & "'
                      , '" & 1 & "'
                      , '" & ClientCrewID & "'
                      , '" & ClientStoreID & "'
                      , '" & ClientGuid & "'
                      , '" & S_Zreading & "'
                      , 'Unsynced')"
            GLOBAL_INSERT_FUNCTION(table, fields, value)
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
#End Region
    Dim INSERTTHISDATE
    Dim ThreadOrder As Thread
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            With WaitFrm
                INSERTTHISDATE = S_Zreading & " " & Format(Now(), "HH:mm:ss")
                SUPERAMOUNTDUE = Convert.ToDecimal(Double.Parse(TextBoxGRANDTOTAL.Text))
                If TRANSACTIONMODE = "Representation Expenses" Then
                    ACTIVE = 3
                End If
                GROSSSALE = NUMBERFORMAT(Double.Parse(Label76.Text))

                If S_ZeroRated = "0" Then
                    If CouponApplied = False Then
                        VATABLESALES = Math.Round(SUPERAMOUNTDUE / Val(1 + S_Tax), 2, MidpointRounding.AwayFromZero)
                        VAT12PERCENT = Math.Round(SUPERAMOUNTDUE - VATABLESALES, 2, MidpointRounding.AwayFromZero)
                    End If
                Else

                    Dim TotalPrice As Double = 0
                    Dim GrandTotal As Double = 0
                    With DataGridViewOrders
                        For i As Integer = 0 To .Rows.Count - 1 Step +1
                            TotalPrice = .Rows(i).Cells(1).Value * .Rows(i).Cells(2).Value
                            GrandTotal += TotalPrice
                        Next
                    End With

                    If CouponApplied = False Then
                        LESSVAT = Math.Round(GrandTotal - SUPERAMOUNTDUE, 2, MidpointRounding.AwayFromZero)
                        ZERORATEDNETSALES = GROSSSALE
                        ZERORATEDSALES = GROSSSALE
                    Else
                        Dim SubTotal = Convert.ToDecimal(Double.Parse(TextBoxSUBTOTAL.Text))
                        LESSVAT = Math.Round(GrandTotal - SubTotal, 2, MidpointRounding.AwayFromZero)
                    End If
                End If


                sql = "SELECT si_number FROM loc_daily_transaction ORDER BY transaction_id DESC limit 1"
                cmd = New MySqlCommand(sql, LocalhostConn)
                da = New MySqlDataAdapter(cmd)
                dt = New DataTable
                da.Fill(dt)
                If dt.Rows.Count > 0 Then
                    SINumber = dt(0)(0)
                Else
                    SINumber = S_SIBeg
                End If
                If SINumber = 0 Then
                    SINumber = 1
                    SiNumberToString = SINumber.ToString(S_SIFormat)
                Else
                    SINumber += 1
                    SiNumberToString = SINumber.ToString(S_SIFormat)
                End If
                For i = 0 To 100
                    BackgroundWorker1.ReportProgress(i)
                    If i = 0 Then
                        .Label1.Text = "Transaction is processing. Please wait."
                        If S_TrainingMode = False Then
                            ThreadOrder = New Thread(AddressOf InsertFMStock)
                            ThreadOrder.Start()
                            THREADLIST.Add(ThreadOrder)
                            For Each t In THREADLIST
                                t.Join()
                            Next
                            ThreadOrder = New Thread(Sub() UpdateInventory())
                            ThreadOrder.Start()
                            THREADLIST.Add(ThreadOrder)
                            For Each t In THREADLIST
                                t.Join()
                            Next
                            ThreadOrder = New Thread(AddressOf InsertDailyTransaction)
                            ThreadOrder.Start()
                            THREADLIST.Add(ThreadOrder)
                            For Each t In THREADLIST
                                t.Join()
                            Next
                            ThreadOrder = New Thread(AddressOf InsertDailyDetails)
                            ThreadOrder.Start()
                            THREADLIST.Add(ThreadOrder)
                            For Each t In THREADLIST
                                t.Join()
                            Next

                            If modeoftransaction = True Then
                                ThreadOrder = New Thread(AddressOf InsertModeofTransaction)
                                ThreadOrder.Start()
                                THREADLIST.Add(ThreadOrder)
                                For Each t In THREADLIST
                                    t.Join()
                                Next
                            End If

                            If SENIORDETAILSBOOL = True Then
                                ThreadOrder = New Thread(AddressOf InsertSeniorDetails)
                                ThreadOrder.Start()
                                THREADLIST.Add(ThreadOrder)
                                For Each t In THREADLIST
                                    t.Join()
                                Next
                            End If
                            If CouponApplied = True Then
                                ThreadOrder = New Thread(AddressOf InsertCouponData)
                                ThreadOrder.Start()
                                THREADLIST.Add(ThreadOrder)
                                For Each t In THREADLIST
                                    t.Join()
                                Next
                            End If
                        End If
                    End If
                    Thread.Sleep(10)
                Next
                For Each t In THREADLIST
                    t.Join()
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        With WaitFrm
            .ProgressBar1.Value = e.ProgressPercentage
            If e.ProgressPercentage = 20 Then
                .Label1.Text = "Transaction is processing. Please wait.."
            End If
            If e.ProgressPercentage = 40 Then
                .Label1.Text = "Transaction is processing. Please wait..."
            End If
            If e.ProgressPercentage = 60 Then
                .Label1.Text = "Transaction is processing. Please wait."
            End If
            If e.ProgressPercentage = 80 Then
                .Label1.Text = "Transaction is processing. Please wait.."
            End If
            If e.ProgressPercentage = 100 Then
                .Label1.Text = "Transaction is processing. Please wait..."
            End If
        End With
    End Sub
    Dim Reprint As Integer = 1
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Enabled = True
        WaitFrm.Close()
        PaymentForm.Close()
        If DataGridViewOrders.Rows.Count > 0 Then
            Try
                For i As Integer = 0 To DataGridViewOrders.Rows.Count - 1 Step +1
                    RowB += 10
                    If DataGridViewOrders.Rows(i).Cells(11).Value > 0 Then
                        RowB += 10
                    End If
                Next

                If CouponApplied Then
                    printdoc.DefaultPageSettings.PaperSize = New PaperSize("Custom", ReturnPrintSize(), 550 + RowB)
                Else
                    printdoc.DefaultPageSettings.PaperSize = New PaperSize("Custom", ReturnPrintSize(), 510 + RowB)
                End If

                If S_Print = "YES" Then
                    For i = 1 To S_PrintCount
                        printdoc.Print()
                        Reprint += 1
                    Next
                    Reprint = 1
                Else
                    For i = 1 To S_PrintCount
                        PrintPreviewDialog1.Document = printdoc
                        PrintPreviewDialog1.ShowDialog()
                        Reprint += 1
                    Next
                    Reprint = 1
                End If
            Catch ex As Exception
                SendErrorReport(ex.ToString)
                MessageBox.Show("An error occurred while trying to load the " &
                    "document for Print Preview. Make sure you currently have " &
                    "access to a printer. A printer must be localconnected and " &
                    "accessible for Print Preview to work.", Text,
                     MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            '=================================================================================================
            selectmax(whatform:=1)
            '=================================================================================================
            messageboxappearance = False
            SystemLogType = "TRANSACTION"
            SystemLogDesc = "Transaction of :" & returnfullname(ClientCrewID) & " Item(s): " & DataGridViewOrders.Rows.Count
            GLOBAL_SYSTEM_LOGS(SystemLogType, SystemLogDesc)
            '=================================================================================================
            DataGridViewOrders.Rows.Clear()
            DataGridViewInv.Rows.Clear()
            modeoftransaction = False
            ButtonApplyCoupon.Enabled = False
            ButtonPayMent.Enabled = False
            Buttonholdoder.Enabled = False
            ButtonPendingOrders.Enabled = True
            payment = False


            TRANSACTIONMODE = "Walk-In"
            CouponApplied = False
            CouponName = ""
            CouponDesc = ""
            ACTIVE = 1
            DISCOUNTTYPE = "N/A"
            TOTALDISCOUNT = 0
            GROSSSALE = 0
            VATEXEMPTSALES = 0
            LESSVAT = 0
            TOTALDISCOUNTEDAMOUNT = 0
            TOTALAMOUNTDUE = 0
            VATABLESALES = 0
            VAT12PERCENT = 0
            CouponLine = 10
            ZERORATEDNETSALES = 0
            ZERORATEDSALES = 0
            DISABLESERVEROTHERSPRODUCT = False


            SeniorGCDiscount = False
            SENIORDETAILSBOOL = False
            SeniorDetailsID = ""
            SeniorDetailsName = ""
            Label76.Text = "0.00"
            TextBoxDISCOUNT.Text = "0.00"
            TextBoxSUBTOTAL.Text = "0.00"
            TextBoxGRANDTOTAL.Text = "0.00"
            LabelTransactionType.Text = "Walk-In"
            WaffleUpgrade = False
            ButtonWaffleUpgrade.Text = "Brownie Upgrade"
            ButtonWaffleUpgrade.BackColor = Color.FromArgb(221, 114, 46)
        Else
            MsgBox("Select Transaction First!")
        End If
        RowA = 0
        RowB = 0
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles printdoc.PrintPage
        Try
            RowA = 40
            ReceiptHeader(sender, e, False)
            Dim abc As Integer = 40
            With DataGridViewOrders
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    Dim price = NUMBERFORMAT(.Rows(i).Cells(3).Value)
                    If .Rows(i).Cells(7).Value.ToString = "Add-Ons" Then
                        If .Rows(i).Cells(13).Value.ToString = "Classic" Then
                            RightToLeftDisplay(sender, e, abc + 115, "     @" & .Rows(i).Cells(0).Value, price, FontAddOn, 0, 0)
                        Else
                            RightToLeftDisplay(sender, e, abc + 115, .Rows(i).Cells(1).Value & " " & .Rows(i).Cells(0).Value, price, FontDefault, 0, 0)
                        End If
                    Else
                        RightToLeftDisplay(sender, e, abc + 115, .Rows(i).Cells(1).Value & " " & .Rows(i).Cells(0).Value, price, FontDefault, 0, 0)
                        If .Rows(i).Cells(11).Value > 0 Then
                            abc += 10
                            RowA += 10
                            RightToLeftDisplay(sender, e, abc + 115, "     + UPGRADE BRWN " & .Rows(i).Cells(11).Value, "", FontAddOn, 0, 0)
                        End If
                    End If
                    RowA += 10
                    abc += 10
                Next
            End With

            If CouponApplied = True Then
                RowA += 100
                If SeniorGCDiscount Then
                    SimpleTextDisplay(sender, e, CouponName & "(" & DISCOUNTTYPE & ")", FontDefault, 0, RowA)
                Else
                    SimpleTextDisplay(sender, e, CouponName & "(" & DISCOUNTTYPE & ")", FontDefault, 0, RowA)
                End If
                SimpleTextDisplay(sender, e, CouponDesc, FontDefault, 0, RowA + 10)
                RowA += 40 + CouponLine
                RightToLeftDisplay(sender, e, RowA - 18, "Total Discount:", "P" & NUMBERFORMAT(CouponTotal), FontDefault, 0, 0)
            Else
                RowA += 120
            End If

            Dim Qty = SumOfColumnsToInt(DataGridViewOrders, 1)

            Dim NETSALES As Double = 0
            If S_ZeroRated = "0" Then
                NETSALES = NUMBERFORMAT(SUPERAMOUNTDUE)
            Else
                NETSALES = ZERORATEDNETSALES
            End If
            If Val(TextBoxDISCOUNT.Text) < 1 Then
                Dim aNumber As Double = TEXTBOXMONEYVALUE
                Dim cash = String.Format("{0:n2}", aNumber)
                Dim aNumber1 As Double = TEXTBOXCHANGEVALUE
                Dim change = String.Format("{0:n2}", aNumber1)

                RightToLeftDisplay(sender, e, RowA, "AMOUNT DUE:", "P" & String.Format("{0:n2}", NETSALES), Font2Bold, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 15, "CASH:", "P" & cash, Font1Bold, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 25, "CHANGE:", "P" & change, Font1Bold, 0, 0)

                PrintStars(sender, e, FontDefault, RowA + 23)

                RightToLeftDisplay(sender, e, RowA + 52, "     VATable Sales", "    " & NUMBERFORMAT(VATABLESALES), FontDefault, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 62, "     VAT Exempt Sales", "    " & NUMBERFORMAT(VATEXEMPTSALES), FontDefault, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 72, "     Zero-Rated Sales", "    " & NUMBERFORMAT(ZERORATEDSALES), FontDefault, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 82, "     VAT Amount" & "(" & Val(S_Tax) * 100 & "%)", "    " & NUMBERFORMAT(VAT12PERCENT), FontDefault, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 92, "     Less Vat", "    " & NUMBERFORMAT(LESSVAT), FontDefault, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 102, "     Total", "    " & NUMBERFORMAT(NETSALES), FontDefault, 0, 0)
                RowA += 4

                PrintStars(sender, e, FontDefault, RowA + 92)

                RowA += 1
                SimpleTextDisplay(sender, e, "Transaction Type: " & Trim(TRANSACTIONMODE), FontDefault, 0, RowA + 100)
                SimpleTextDisplay(sender, e, "Total Item(s): " & Qty, FontDefault, 0, RowA + 110)
                SimpleTextDisplay(sender, e, "Cashier: " & ClientCrewID & " " & returnfullname(where:=ClientCrewID), FontDefault, 0, RowA + 120)
                SimpleTextDisplay(sender, e, "Str No: " & ClientStoreID, FontDefault, 110, RowA + 110)
                SimpleTextDisplay(sender, e, "Date & Time: " & INSERTTHISDATE, FontDefault, 0, RowA + 130)
                SimpleTextDisplay(sender, e, "Terminal No: " & S_Terminal_No, FontDefault, 110, RowA + 140)
                SimpleTextDisplay(sender, e, "Ref. #: " & TextBoxMAXID.Text, FontDefault, 0, RowA + 140)
                SimpleTextDisplay(sender, e, "SI No: " & SiNumberToString, FontDefault, 0, RowA + 150)
                If Reprint = 1 Then
                    SimpleTextDisplay(sender, e, "Customers Copy", FontDefault, 0, RowA + 160)
                Else
                    SimpleTextDisplay(sender, e, "Reprint Copy", FontDefault, 0, RowA + 160)
                End If
                If S_TrainingMode Then
                    SimpleTextDisplay(sender, e, "THIS IS NOT AN OFFICIAL RECEIPT", FontDefault, 0, RowA + 170)
                Else
                    SimpleTextDisplay(sender, e, "THIS SERVES AS AN OFFICIAL RECEIPT", FontDefault, 0, RowA + 170)
                End If

                PrintStars(sender, e, FontDefault, RowA + 184)

                ReceiptFooter(sender, e, RowA + 12, False)
            Else

                Dim aNumber1 As Double = TEXTBOXCHANGEVALUE
                Dim change = String.Format("{0:n2}", aNumber1)
                Dim aNumber As Double = TEXTBOXMONEYVALUE
                Dim cash = String.Format("{0:n2}", aNumber)
                RightToLeftDisplay(sender, e, RowA, "SUB TOTAL:", "P" & NUMBERFORMAT(Double.Parse(Label76.Text)), Font1Bold, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 10, "DISCOUNT:", NUMBERFORMAT(Double.Parse(TextBoxDISCOUNT.Text)) & "-", Font1Bold, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 20, "AMOUNT DUE:", "P" & NUMBERFORMAT(Double.Parse(TextBoxGRANDTOTAL.Text)), Font1Bold, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 30, "CASH:", "P" & cash, Font1Bold, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 40, "CHANGE:", "P" & change, Font1Bold, 0, 0)

                PrintStars(sender, e, FontDefault, RowA + 37)

                RightToLeftDisplay(sender, e, RowA + 65, "     VATable Sales", "    " & NUMBERFORMAT(VATABLESALES), FontDefault, 0, 0)
                If DISCOUNTTYPE = "Percentage(w/o vat)" Then
                    RightToLeftDisplay(sender, e, RowA + 75, "     VAT Exempt Sales", "    " & Format(VATEXEMPTSALES, "0.00"), FontDefault, 0, 0)
                Else
                    If SeniorGCDiscount = False Then
                        RightToLeftDisplay(sender, e, RowA + 75, "     Vat Exempt Sales", "    " & "0.00", FontDefault, 0, 0)
                    Else
                        RightToLeftDisplay(sender, e, RowA + 75, "     Vat Exempt Sales", "    " & Format(VATEXEMPTSALES, "0.00"), FontDefault, 0, 0)
                    End If
                End If
                RightToLeftDisplay(sender, e, RowA + 85, "     Zero-Rated Sales", "    " & Format(ZERORATEDSALES, "0.00"), FontDefault, 0, 0)
                RightToLeftDisplay(sender, e, RowA + 95, "     VAT Amount" & "(" & Val(S_Tax) * 100 & "%)", "    " & Format(VAT12PERCENT, "0.00"), FontDefault, 0, 0)
                If DISCOUNTTYPE = "Percentage(w/o vat)" Then
                    RightToLeftDisplay(sender, e, RowA + 105, "     Less Vat", "    " & Format(LESSVAT, "0.00"), FontDefault, 0, 0)
                Else
                    If SeniorGCDiscount = False Then
                        If S_ZeroRated = "0" Then
                            RightToLeftDisplay(sender, e, RowA + 105, "     Less Vat", "    " & "0.00", FontDefault, 0, 0)
                        Else
                            RightToLeftDisplay(sender, e, RowA + 105, "     Less Vat", "    " & Format(LESSVAT, "0.00"), FontDefault, 0, 0)
                        End If
                    Else
                        RightToLeftDisplay(sender, e, RowA + 105, "     Less Vat", "    " & Format(LESSVAT, "0.00"), FontDefault, 0, 0)
                    End If
                End If
                RightToLeftDisplay(sender, e, RowA + 115, "     Total", "    " & NUMBERFORMAT(NETSALES), FontDefault, 0, 0)
                RowA += 5

                PrintStars(sender, e, FontDefault, RowA + 101)

                RowA += 4
                SimpleTextDisplay(sender, e, "Transaction Type: " & Trim(TRANSACTIONMODE), FontDefault, 0, RowA + 110)
                SimpleTextDisplay(sender, e, "Total Item(s): " & Qty, FontDefault, 0, RowA + 120)
                SimpleTextDisplay(sender, e, "Cashier: " & ClientCrewID & " " & returnfullname(where:=ClientCrewID), FontDefault, 0, RowA + 130)
                SimpleTextDisplay(sender, e, "Str No: " & ClientStoreID, FontDefault, 120, RowA + 120)
                SimpleTextDisplay(sender, e, "Date & Time: " & INSERTTHISDATE, FontDefault, 0, RowA + 140)
                SimpleTextDisplay(sender, e, "Terminal No: " & S_Terminal_No, FontDefault, 120, RowA + 150)
                SimpleTextDisplay(sender, e, "Ref. #: " & TextBoxMAXID.Text, FontDefault, 0, RowA + 150)
                SimpleTextDisplay(sender, e, "SI No: " & SiNumberToString, FontDefault, 0, RowA + 160)
                If Reprint = 1 Then
                    SimpleTextDisplay(sender, e, "Customers Copy", FontDefault, 0, RowA + 170)
                Else
                    SimpleTextDisplay(sender, e, "Reprint Copy", FontDefault, 0, RowA + 170)
                End If
                If S_TrainingMode Then
                    SimpleTextDisplay(sender, e, "THIS IS NOT AN OFFICIAL RECEIPT", FontDefault, 0, RowA + 180)
                Else
                    SimpleTextDisplay(sender, e, "THIS SERVES AS AN OFFICIAL RECEIPT", FontDefault, 0, RowA + 180)
                End If
                RowA += 6

                PrintStars(sender, e, FontDefault, RowA + 190)

                ReceiptFooter(sender, e, RowA + 20, False)
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
#End Region
#Region "Updates"
#Region "Categories Update"
    Public PriceChangeDatatabe As DataTable
    Private Sub CheckPriceChanges()
        Try
            Dim ConnectionServer As MySqlConnection = ServerCloudCon()
            Dim Query = "SELECT * FROM admin_price_request WHERE store_id = '" & ClientStoreID & "' AND guid = '" & ClientGuid & "' AND synced = 'Unsynced' AND active = 2"
            Dim CmdCheck As MySqlCommand = New MySqlCommand(Query, ConnectionServer)
            Dim DaCheck As MySqlDataAdapter = New MySqlDataAdapter(CmdCheck)
            PriceChangeDatatabe = New DataTable
            DaCheck.Fill(PriceChangeDatatabe)
            If PriceChangeDatatabe.Rows.Count > 0 Then
                PRICECHANGE = True
            Else
                PRICECHANGE = False
            End If
        Catch ex As Exception
            BackgroundWorker2.CancelAsync()
            'SendErrorReport(ex.ToString)
            Exit Sub
        End Try
    End Sub
    Public CouponDatatable As DataTable
    Dim CouponApp As Boolean = False
    Private Sub CouponApproval()
        Try
            Dim ConnectionServer As MySqlConnection = ServerCloudCon()
            Dim Query = "SELECT ID FROM admin_custom_coupon WHERE store_id = '" & ClientStoreID & "' AND guid = '" & ClientGuid & "' AND active = 1 AND synced = 'Unsynced'"
            Dim CmdCheck As MySqlCommand = New MySqlCommand(Query, ConnectionServer)
            Dim DaCheck As MySqlDataAdapter = New MySqlDataAdapter(CmdCheck)
            CouponDatatable = New DataTable
            DaCheck.Fill(CouponDatatable)
            If CouponDatatable.Rows.Count > 0 Then
                CouponApp = True
            Else
                CouponApp = False
            End If
        Catch ex As Exception
            BackgroundWorker2.CancelAsync()
            'SendErrorReport(ex.ToString)
            Exit Sub
        End Try
    End Sub
    Public CustomProductsApproval As DataTable
    Dim CustomProdctsAppBool As Boolean = False
    Private Sub CustomProductApproval()
        Try
            Dim ConnectionServer As MySqlConnection = ServerCloudCon()
            Dim Query = "SELECT loc_product_id, product_formula_id FROM loc_product_list WHERE store_id = '" & ClientStoreID & "' AND guid = '" & ClientGuid & "' AND product_status = 1 AND synced = 'Unsynced'"
            Dim CmdCheck As MySqlCommand = New MySqlCommand(Query, ConnectionServer)
            Dim DaCheck As MySqlDataAdapter = New MySqlDataAdapter(CmdCheck)
            CustomProductsApproval = New DataTable
            DaCheck.Fill(CustomProductsApproval)
            If CustomProductsApproval.Rows.Count > 0 Then
                CustomProdctsAppBool = True
            Else
                CustomProdctsAppBool = False
            End If
        Catch ex As Exception
            BackgroundWorker2.CancelAsync()
            'SendErrorReport(ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Function LoadCouponsLocal() As DataTable
        Dim cmdlocal As MySqlCommand
        Dim dalocal As MySqlDataAdapter
        Dim dtlocal As DataTable = New DataTable
        dtlocal.Columns.Add("date_created")
        dtlocal.Columns.Add("ID")
        Dim dtlocal1 As DataTable = New DataTable
        Try
            Dim sql = "SELECT date_created, ID FROM tbcoupon"
            cmdlocal = New MySqlCommand(sql, LocalhostConn())
            dalocal = New MySqlDataAdapter(cmdlocal)
            dalocal.Fill(dtlocal1)
            For i As Integer = 0 To dtlocal1.Rows.Count - 1 Step +1
                Dim Coup As DataRow = dtlocal.NewRow
                Coup("date_created") = dtlocal1(i)(0).ToString
                Coup("ID") = dtlocal1(i)(1)
                dtlocal.Rows.Add(Coup)
            Next
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return dtlocal
    End Function
    Private Sub Function5()
        Try
            Dim Query = "SELECT * FROM tbcoupon"
            Dim CmdCheck As MySqlCommand = New MySqlCommand(Query, LocalhostConn)
            Dim DaCheck As MySqlDataAdapter = New MySqlDataAdapter(CmdCheck)
            Dim DtCheck As DataTable = New DataTable
            DaCheck.Fill(DtCheck)
            Dim cmdserver As MySqlCommand
            Dim daserver As MySqlDataAdapter
            Dim dtserver As DataTable
            If DtCheck.Rows.Count < 1 Then
                Dim sql = "SELECT `ID`,`Couponname_`,`Desc_`,`Discountvalue_`,`Referencevalue_`,`Type`,`Bundlebase_`,`BBValue_`,`Bundlepromo_`,`BPValue_`,`Effectivedate`,`Expirydate`,`date_created` FROM admin_coupon"
                cmdserver = New MySqlCommand(sql, ServerCloudCon())
                daserver = New MySqlDataAdapter(cmdserver)
                dtserver = New DataTable
                daserver.Fill(dtserver)
                For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                    DataGridView5.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3).ToString, dtserver(i)(4), dtserver(i)(5), dtserver(i)(6), dtserver(i)(7), dtserver(i)(8), dtserver(i)(9), dtserver(i)(10), dtserver(i)(11), dtserver(i)(12))
                Next
            Else
                Dim Ids As String = ""
                If ValidCloudConnection = True Then
                    For i As Integer = 0 To LoadCouponsLocal.Rows.Count - 1 Step +1
                        If Ids = "" Then
                            Ids = "" & LoadCouponsLocal(i)(1) & ""
                        Else
                            Ids += "," & LoadCouponsLocal(i)(1) & ""
                        End If
                    Next
                    Dim sql = "SELECT `ID`,`Couponname_`,`Desc_`,`Discountvalue_`,`Referencevalue_`,`Type`,`Bundlebase_`,`BBValue_`,`Bundlepromo_`,`BPValue_`,`Effectivedate`,`Expirydate`,`date_created` FROM admin_coupon WHERE ID IN (" & Ids & ")"
                    cmdserver = New MySqlCommand(sql, ServerCloudCon())
                    daserver = New MySqlDataAdapter(cmdserver)
                    dtserver = New DataTable
                    daserver.Fill(dtserver)
                    For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                        If LoadCouponsLocal(i)(0).ToString <> dtserver(i)(12).ToString Then
                            DataGridView5.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3).ToString, dtserver(i)(4), dtserver(i)(5), dtserver(i)(6), dtserver(i)(7), dtserver(i)(8), dtserver(i)(9), dtserver(i)(10), dtserver(i)(11), dtserver(i)(12))
                        End If
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   If DataGridViewUpdate.Rows.Count > 0 Then
                                                                       CheckingForUpdates.ProgressBar1.Value += 1
                                                                       CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                                   End If
                                                               End Sub)
                        End If

                    Next
                    Dim sql2 = "SELECT `ID`,`Couponname_`,`Desc_`,`Discountvalue_`,`Referencevalue_`,`Type`,`Bundlebase_`,`BBValue_`,`Bundlepromo_`,`BPValue_`,`Effectivedate`,`Expirydate`,`date_created` FROM admin_coupon WHERE ID NOT IN (" & Ids & ")"
                    cmdserver = New MySqlCommand(sql2, ServerCloudCon())
                    daserver = New MySqlDataAdapter(cmdserver)
                    dtserver = New DataTable
                    daserver.Fill(dtserver)
                    For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                        If LoadCouponsLocal(i)(0) <> dtserver(i)(12) Then
                            DataGridView5.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3).ToString, dtserver(i)(4), dtserver(i)(5), dtserver(i)(6), dtserver(i)(7), dtserver(i)(8), dtserver(i)(9), dtserver(i)(10), dtserver(i)(11), dtserver(i)(12))
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            BackgroundWorker2.CancelAsync()
            'SendErrorReport(ex.ToString)
            Exit Sub
            'If table doesnt have data
        End Try
    End Sub


    Private Function LoadCategoryLocal() As DataTable
        Dim cmdlocal As MySqlCommand
        Dim dalocal As MySqlDataAdapter
        Dim dtlocal As DataTable = New DataTable
        dtlocal.Columns.Add("updated_at")
        dtlocal.Columns.Add("category_id")
        Dim dtlocal1 As DataTable = New DataTable
        Try
            Dim sql = "SELECT updated_at, category_id FROM loc_admin_category"
            cmdlocal = New MySqlCommand(sql, LocalhostConn())
            dalocal = New MySqlDataAdapter(cmdlocal)
            dalocal.Fill(dtlocal1)
            For i As Integer = 0 To dtlocal1.Rows.Count - 1 Step +1
                Dim Cat As DataRow = dtlocal.NewRow
                Cat("updated_at") = dtlocal1(i)(0).ToString
                Cat("category_id") = dtlocal1(i)(1)
                dtlocal.Rows.Add(Cat)
            Next
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return dtlocal
    End Function

    Private Sub Function1()
        Try
            Dim Query = "SELECT * FROM loc_admin_category"
            Dim CmdCheck As MySqlCommand = New MySqlCommand(Query, LocalhostConn)
            Dim DaCheck As MySqlDataAdapter = New MySqlDataAdapter(CmdCheck)
            Dim DtCheck As DataTable = New DataTable
            DaCheck.Fill(DtCheck)
            Dim cmdserver As MySqlCommand
            Dim daserver As MySqlDataAdapter
            Dim dtserver As DataTable
            If DtCheck.Rows.Count < 1 Then
                Dim sql = "SELECT `category_id`, `category_name`, `brand_name`, `updated_at`, `origin`, `status` FROM admin_category"
                cmdserver = New MySqlCommand(sql, ServerCloudCon())
                daserver = New MySqlDataAdapter(cmdserver)
                dtserver = New DataTable
                daserver.Fill(dtserver)
                If WorkerUpdateCancel = False Then
                    CheckingForUpdates.Instance.Invoke(Sub()
                                                           CheckingForUpdates.ProgressBar1.Maximum += dtserver.Rows.Count
                                                           'Console.WriteLine("CATEGORY : " & CheckingForUpdates.ProgressBar1.Maximum)
                                                       End Sub)
                End If

                For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                    DataGridView1.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3).ToString, dtserver(i)(4), dtserver(i)(5))
                    If WorkerUpdateCancel = False Then
                        CheckingForUpdates.Instance.Invoke(Sub()
                                                               If DataGridViewUpdate.Rows.Count > 0 Then
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End If
                                                           End Sub)
                    End If

                Next
            Else
                Dim Ids As String = ""
                If ValidCloudConnection = True Then
                    For i As Integer = 0 To LoadCategoryLocal.Rows.Count - 1 Step +1
                        If Ids = "" Then
                            Ids = "" & LoadCategoryLocal(i)(1) & ""
                        Else
                            Ids += "," & LoadCategoryLocal(i)(1) & ""
                        End If
                    Next
                    Dim sql = "SELECT `category_id`, `category_name`, `brand_name`, `updated_at`, `origin`, `status` FROM admin_category WHERE category_id IN (" & Ids & ")"
                    cmdserver = New MySqlCommand(sql, ServerCloudCon())
                    daserver = New MySqlDataAdapter(cmdserver)
                    dtserver = New DataTable
                    daserver.Fill(dtserver)
                    For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                        If LoadCategoryLocal(i)(0).ToString <> dtserver(i)(3).ToString Then
                            DataGridView1.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3).ToString, dtserver(i)(4), dtserver(i)(5))
                        End If
                        If WorkerUpdateCancel = False Then

                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   If DataGridViewUpdate.Rows.Count > 0 Then
                                                                       CheckingForUpdates.ProgressBar1.Value += 1
                                                                       CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                                   End If
                                                               End Sub)
                        End If

                    Next
                    Dim sql2 = "SELECT `category_id`, `category_name`, `brand_name`, `updated_at`, `origin`, `status` FROM admin_category WHERE category_id NOT IN (" & Ids & ")"
                    cmdserver = New MySqlCommand(sql2, ServerCloudCon())
                    daserver = New MySqlDataAdapter(cmdserver)
                    dtserver = New DataTable
                    daserver.Fill(dtserver)
                    For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                        If LoadCategoryLocal(i)(0) <> dtserver(i)(3) Then
                            DataGridView1.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3).ToString, dtserver(i)(4), dtserver(i)(5))

                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            BackgroundWorker2.CancelAsync()
            SendErrorReport(ex.ToString)
            Exit Sub
            'If table doesnt have data
        End Try
    End Sub
#End Region
#Region "Products Update"

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            WorkerUpdateCancel = False
            'UPDATEPRODUCTONLY = True
            HASUPDATE = False
            DataGridView1.Rows.Clear()
            DataGridView2.DataSource = Nothing
            DataGridView3.Rows.Clear()
            DataGridView4.Rows.Clear()
            DataGridView5.Rows.Clear()
            DataGridViewPartners.Rows.Clear()
            DataGridViewUpdate.Rows.Clear()


            If CheckForInternetConnection() Then
                Enabled = False
                CheckingForUpdates.Show()
                CheckingForUpdates.TopMost = True

                If ValidCloudConnection = True Then
                    GetRowCount()
                    BackgroundWorker2.WorkerReportsProgress = True
                    BackgroundWorker2.WorkerSupportsCancellation = True
                    BackgroundWorker2.RunWorkerAsync()
                Else
                    If BegBalanceBool = False Then
                        Enabled = False
                        BegBalance.Show()
                        BegBalance.TopMost = True
                    End If

                    'CheckingForUpdates.LabelCheckingUpdates.Text = "Invalid cloud server connection."
                End If
            Else
                MsgBox("No internet connection.")
                If BegBalanceBool = False Then
                    BegBalance.Show()
                    BegBalance.TopMost = True
                End If
            End If

            'Enabled = False
            'CheckingForUpdates.Show()
            'CheckingForUpdates.TopMost = True
            'CheckingForUpdates.LabelCheckingUpdates.Text = "Checking for updates."
            'GetRowCount()

            'BackgroundWorker2.WorkerReportsProgress = True
            'BackgroundWorker2.WorkerSupportsCancellation = True
            'BackgroundWorker2.RunWorkerAsync()

        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public POSISUPDATING As Boolean = False
    Dim PRICECHANGE As Boolean = False
    Dim TestInternetCon As Boolean = False
    Public WorkerUpdateCancel As Boolean = False
    Dim ThreadUpdate As Thread
    Private Sub BackgroundWorker2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Try
            If ValidLocalConnection Then
                ThreadUpdate = New Thread(Sub() TestInternetCon = CheckForInternetConnection())
                ThreadUpdate.Start()
                THREADLISTUPDATE.Add(ThreadUpdate)
                For Each t In THREADLISTUPDATE
                    t.Join()
                    If (BackgroundWorker2.CancellationPending) Then
                        ' Indicate that the task was canceled.
                        e.Cancel = True
                        WorkerUpdateCancel = True
                        If WorkerUpdateCancel Then
                            Exit For
                        End If
                        Exit Sub
                    End If
                Next
                If TestInternetCon Then
                    ThreadUpdate = New Thread(AddressOf ServerCloudCon)
                    ThreadUpdate.Start()
                    THREADLISTUPDATE.Add(ThreadUpdate)
                    For Each t In THREADLISTUPDATE
                        t.Join()

                        If (BackgroundWorker2.CancellationPending) Then
                            ' Indicate that the task was canceled.
                            e.Cancel = True
                            WorkerUpdateCancel = True
                            If WorkerUpdateCancel Then
                                Exit For
                            End If
                            Exit Sub
                        End If
                    Next
                    If ServerCloudCon.State = ConnectionState.Open Then
                        'If UPDATEPRODUCTONLY = False Then
                        'POSISUPDATING = True
                        ThreadUpdate = New Thread(AddressOf CheckPriceChanges)
                        ThreadUpdate.Start()
                        THREADLISTUPDATE.Add(ThreadUpdate)

                        For Each t In THREADLISTUPDATE
                            t.Join()

                            If (BackgroundWorker2.CancellationPending) Then
                                ' Indicate that the task was canceled.
                                e.Cancel = True
                                WorkerUpdateCancel = True
                                If WorkerUpdateCancel Then
                                    Exit For
                                End If
                                Exit Sub
                            End If
                        Next

                        ThreadUpdate = New Thread(AddressOf GetProducts)
                        ThreadUpdate.Start()
                        THREADLISTUPDATE.Add(ThreadUpdate)

                        ThreadUpdate = New Thread(AddressOf PromptMessage)
                        ThreadUpdate.Start()
                        THREADLISTUPDATE.Add(ThreadUpdate)
                        ThreadUpdate = New Thread(AddressOf Function1)
                        ThreadUpdate.Start()
                        THREADLISTUPDATE.Add(ThreadUpdate)

                        ThreadUpdate = New Thread(AddressOf Function3)
                        ThreadUpdate.Start()
                        THREADLISTUPDATE.Add(ThreadUpdate)
                        ThreadUpdate = New Thread(AddressOf Function4)
                        ThreadUpdate.Start()
                        THREADLISTUPDATE.Add(ThreadUpdate)
                        ThreadUpdate = New Thread(AddressOf Function5)
                        ThreadUpdate.Start()
                        THREADLISTUPDATE.Add(ThreadUpdate)
                        ThreadUpdate = New Thread(AddressOf Function6)
                        ThreadUpdate.Start()
                        THREADLISTUPDATE.Add(ThreadUpdate)
                        ThreadUpdate = New Thread(AddressOf CouponApproval)
                        ThreadUpdate.Start()
                        THREADLISTUPDATE.Add(ThreadUpdate)
                        ThreadUpdate = New Thread(AddressOf CustomProductApproval)
                        ThreadUpdate.Start()
                        THREADLISTUPDATE.Add(ThreadUpdate)

                        For Each t In THREADLISTUPDATE
                            t.Join()
                            If (BackgroundWorker2.CancellationPending) Then
                                ' Indicate that the task was canceled.
                                e.Cancel = True
                                WorkerUpdateCancel = True
                                If WorkerUpdateCancel Then
                                    Exit For
                                End If
                                Exit Sub
                            End If
                        Next

                    End If
                End If

            End If
        Catch ex As Exception
            ValidCloudConnection = False
            BackgroundWorker2.CancelAsync()
            If WorkerUpdateCancel Then
                MsgBox("Cannot fetch data. Please check your internet connection")
            End If
            SendErrorReport(ex.ToString)
            Exit Sub
        End Try
    End Sub
    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            If WorkerUpdateCancel = False Then
                If ValidCloudConnection Then
                    DataGridView2.DataSource = FillDatagridProduct
                    Button3.Enabled = True
                    'UPDATEPRODUCTONLY = False
                    POSISUPDATING = False
                    If DataGridView1.Rows.Count > 0 Or DataGridView2.Rows.Count > 0 Or DataGridView3.Rows.Count > 0 Or DataGridView4.Rows.Count > 0 Or PriceChangeDatatabe.Rows.Count > 0 Or CouponDatatable.Rows.Count > 0 Or CustomProductsApproval.Rows.Count Or DataGridView5.Rows.Count > 0 Or DataGridViewPartners.Rows.Count > 0 Then
                        CheckingForUpdates.CheckingUpdatesUPDATED = True
                        CheckingForUpdates.Close()

                        Dim updatemessage = MessageBox.Show("New Updates are available. Would you like to update now ?", "New Updates", MessageBoxButtons.YesNo, MessageBoxIcon.Information)

                        If updatemessage = DialogResult.Yes Then
                            HASUPDATE = True
                            Enabled = False
                            CheckingForUpdates.Show()
                            CheckingForUpdates.TopMost = True
                            CheckingForUpdates.CheckingUpdatesUPDATED = False
                            If WorkerUpdateCancel = False Then
                                CheckingForUpdates.Instance.Invoke(Sub()
                                                                       CheckingForUpdates.ProgressBar1.Maximum = 0
                                                                       Dim TotalRows = DataGridView1.Rows.Count + DataGridView2.Rows.Count + DataGridView3.Rows.Count + DataGridView4.Rows.Count + PriceChangeDatatabe.Rows.Count + CouponDatatable.Rows.Count + CustomProductsApproval.Rows.Count + DataGridViewPartners.Rows.Count + DataGridView5.Rows.Count
                                                                       CheckingForUpdates.ProgressBar1.Maximum = TotalRows
                                                                   End Sub)
                            End If
                            BackgroundWorkerInstallUpdates.WorkerReportsProgress = True
                            BackgroundWorkerInstallUpdates.WorkerSupportsCancellation = True
                            BackgroundWorkerInstallUpdates.RunWorkerAsync()
                        Else

                            If BegBalanceBool = False Then
                                BegBalance.Show()
                                BegBalance.TopMost = True
                            Else
                                Enabled = True
                            End If
                        End If
                    Else

                        If BegBalanceBool = False Then
                            BegBalance.Show()
                            BegBalance.TopMost = True
                        Else
                            Enabled = True
                        End If
                        'Enabled = True
                        HASUPDATE = False
                        CheckingForUpdates.CheckingUpdatesUPDATED = True
                        CheckingForUpdates.LabelCheckingUpdates.Text = "Complete Checking! No updates found."
                        CheckingForUpdates.Close()
                    End If
                Else
                    Button3.Enabled = True
                End If
            Else
                If BegBalanceBool = False Then
                    BegBalance.Show()
                    BegBalance.TopMost = True
                Else
                    Enabled = True
                End If
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub BackgroundWorkerInstallUpdates_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorkerInstallUpdates.DoWork
        Try
            If ValidLocalConnection Then
                ThreadUpdate = New Thread(AddressOf InstallUpdatesFormula)
                ThreadUpdate.Start()
                THREADLISTUPDATE.Add(ThreadUpdate)
                ThreadUpdate = New Thread(AddressOf InstallUpdatesInventory)
                ThreadUpdate.Start()
                THREADLISTUPDATE.Add(ThreadUpdate)
                ThreadUpdate = New Thread(AddressOf InstallUpdatesCategory)
                ThreadUpdate.Start()
                THREADLISTUPDATE.Add(ThreadUpdate)
                ThreadUpdate = New Thread(AddressOf InstallUpdatesCoupons)
                ThreadUpdate.Start()
                THREADLISTUPDATE.Add(ThreadUpdate)
                ThreadUpdate = New Thread(AddressOf InstallUpdatesProducts)
                ThreadUpdate.Start()
                THREADLISTUPDATE.Add(ThreadUpdate)
                ThreadUpdate = New Thread(AddressOf InstallUpdatesPriceChange)
                ThreadUpdate.Start()
                THREADLISTUPDATE.Add(ThreadUpdate)
                ThreadUpdate = New Thread(AddressOf InstallCoupons)
                ThreadUpdate.Start()
                THREADLISTUPDATE.Add(ThreadUpdate)
                ThreadUpdate = New Thread(AddressOf InstallProducts)
                ThreadUpdate.Start()
                THREADLISTUPDATE.Add(ThreadUpdate)
                ThreadUpdate = New Thread(AddressOf InstallUpdatesPartners)
                ThreadUpdate.Start()
                THREADLISTUPDATE.Add(ThreadUpdate)
                For Each t In THREADLISTUPDATE
                    t.Join()
                    If (BackgroundWorker2.CancellationPending) Then
                        e.Cancel = True
                        WorkerUpdateCancel = True
                        Exit Sub
                    End If
                Next
            End If

        Catch ex As Exception
            ValidCloudConnection = False
            BackgroundWorker2.CancelAsync()
            If WorkerUpdateCancel Then
                MsgBox("Cannot fetch data. Please check your internet connection")
            End If
            SendErrorReport(ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub BackgroundWorkerInstallUpdates_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorkerInstallUpdates.RunWorkerCompleted
        Try
            CheckingForUpdates.CheckingUpdatesUPDATED = True
            CheckingForUpdates.Close()

            If PRICECHANGE = True Then
                MsgBox("Product price changes approved")
                PRICECHANGE = False
            End If
            If CouponApp = True Then
                MsgBox("Coupon Approved")
            End If
            If CustomProdctsAppBool = True Then
                MsgBox("Products Approved")
            End If
            LoadCategory()

            If BegBalanceBool = False Then
                BegBalance.Show()
                BegBalance.TopMost = True
            Else
                Enabled = True
            End If
            If DtMessage.Rows.Count > 0 Then
                Dim ConnectionLocal As MySqlConnection = LocalhostConn()
                For i As Integer = 0 To DtMessage.Rows.Count - 1 Step +1
                    Dim sql = "INSERT INTO loc_message (`server_message_id`,`from`, `subject`, `content`, `guid`, `store_id`, `active`, `created_at`, `origin`, `seen`) VALUES (@1,@2,@3,@4,@5,@6,@7,@8,@9,@10)"
                    Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
                    cmd.Parameters.Add("@1", MySqlDbType.Int64).Value = DtMessage(i)(0).ToString
                    cmd.Parameters.Add("@2", MySqlDbType.Text).Value = DtMessage(i)(1).ToString
                    cmd.Parameters.Add("@3", MySqlDbType.Text).Value = DtMessage(i)(2).ToString
                    cmd.Parameters.Add("@4", MySqlDbType.Text).Value = DtMessage(i)(3).ToString
                    cmd.Parameters.Add("@5", MySqlDbType.Text).Value = DtMessage(i)(4).ToString
                    cmd.Parameters.Add("@6", MySqlDbType.Text).Value = DtMessage(i)(5).ToString
                    cmd.Parameters.Add("@7", MySqlDbType.Int64).Value = DtMessage(i)(6)
                    cmd.Parameters.Add("@8", MySqlDbType.Text).Value = DtMessage(i)(7).ToString
                    cmd.Parameters.Add("@9", MySqlDbType.Text).Value = DtMessage(i)(8).ToString
                    cmd.Parameters.Add("@10", MySqlDbType.Int64).Value = 0
                    cmd.ExecuteNonQuery()
                    cmd.Dispose()
                Next
                Enabled = False
                For i As Integer = 0 To DtMessage.Rows.Count - 1 Step +1
                    If DtMessage(i)(4).ToString = "Server" Then
                        Message.Show()
                    ElseIf DtMessage(i)(4).ToString = ClientGuid Then
                        If DtMessage(i)(5).ToString = ClientStoreID Then
                            Message.Show()
                        End If
                    End If

                Next
            End If
            For Each btn As Button In Panel3.Controls.OfType(Of Button)()
                If btn.Text = "Simply Perfect" Then
                    btn.PerformClick()
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    'Dim UPDATEPRODUCTONLY As Boolean = False
    Dim FillDatagridProduct As DataTable
    Private Sub GetProducts()
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim ConnectionServer As MySqlConnection = ServerCloudCon()

            FillDatagridProduct = New DataTable
            FillDatagridProduct.Columns.Add("product_id")
            FillDatagridProduct.Columns.Add("product_sku")
            FillDatagridProduct.Columns.Add("product_name")
            FillDatagridProduct.Columns.Add("formula_id")
            FillDatagridProduct.Columns.Add("product_barcode")
            FillDatagridProduct.Columns.Add("product_category")
            FillDatagridProduct.Columns.Add("product_price")
            FillDatagridProduct.Columns.Add("product_desc")
            FillDatagridProduct.Columns.Add("product_image")
            FillDatagridProduct.Columns.Add("product_status")
            FillDatagridProduct.Columns.Add("origin")
            FillDatagridProduct.Columns.Add("date_modified")
            FillDatagridProduct.Columns.Add("inventory_id")
            FillDatagridProduct.Columns.Add("addontype")
            FillDatagridProduct.Columns.Add("half_batch")

            Dim Query = "SELECT * FROM loc_admin_products"
            Dim CmdCheck As MySqlCommand = New MySqlCommand(Query, ConnectionLocal)
            Dim DaCheck As MySqlDataAdapter = New MySqlDataAdapter(CmdCheck)
            Dim DtCheck As DataTable = New DataTable
            DaCheck.Fill(DtCheck)

            If DtCheck.Rows.Count < 1 Then
                GetAllProducts()
            Else
                Dim DtCount As DataTable
                Dim DtCountProductIds As DataTable = New DataTable
                Dim SqlCount = "SELECT product_id FROM admin_products_org"
                Dim CmdCount As MySqlCommand = New MySqlCommand(SqlCount, ConnectionServer)
                Dim DataAdapter As MySqlDataAdapter = New MySqlDataAdapter(CmdCount)
                DataAdapter.Fill(DtCountProductIds)
                Dim DaCount As MySqlDataAdapter
                Dim FillDt As DataTable = New DataTable

                For i As Integer = 0 To DtCountProductIds.Rows.Count - 1 Step +1
                    If WorkerUpdateCancel Then
                        Exit For
                    End If
                    Dim Query1 As String = "SELECT date_modified, price_change FROM loc_admin_products WHERE server_product_id = " & DtCountProductIds(i)(0)
                    Dim cmd As MySqlCommand = New MySqlCommand(Query1, ConnectionLocal)
                    DaCount = New MySqlDataAdapter(cmd)
                    FillDt = New DataTable
                    DaCount.Fill(FillDt)
                    Dim Prod As DataRow = FillDatagridProduct.NewRow
                    If FillDt.Rows.Count > 0 Then
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   If DataGridViewUpdate.Rows.Count > 0 Then
                                                                       CheckingForUpdates.ProgressBar1.Value += 1
                                                                       CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                                   End If
                                                               End Sub)
                        End If

                        Dim PriceChange = FillDt(0)(1)
                        'Exist then check for update
                        Query1 = "SELECT * FROM admin_products_org WHERE product_id = " & DtCountProductIds(i)(0)
                        cmd = New MySqlCommand(Query1, ConnectionServer)
                        DaCount = New MySqlDataAdapter(cmd)
                        DtCount = New DataTable
                        DaCount.Fill(DtCount)
                        If FillDt(0)(0).ToString <> DtCount(0)(11) Then
                            Prod("product_id") = DtCount(0)(0)
                            Prod("product_sku") = DtCount(0)(1)
                            Prod("product_name") = DtCount(0)(2)
                            Prod("formula_id") = DtCount(0)(3)
                            Prod("product_barcode") = DtCount(0)(4)
                            Prod("product_category") = DtCount(0)(5)
                            If FillDt(0)(1) = 1 Then
                                Dim sql2 = "SELECT product_price FROM loc_admin_products WHERE server_product_id = " & DtCountProductIds(i)(0)
                                Dim cmd2 As MySqlCommand = New MySqlCommand(sql2, LocalhostConn)
                                Dim da2 As MySqlDataAdapter = New MySqlDataAdapter(cmd2)
                                Dim dt2 As DataTable = New DataTable
                                da2.Fill(dt2)
                                Prod("product_price") = dt2(0)(0)
                            Else
                                Prod("product_price") = DtCount(0)(6)
                            End If
                            Prod("product_desc") = DtCount(0)(7)
                            Prod("product_image") = DtCount(0)(8)
                            Prod("product_status") = DtCount(0)(9)
                            Prod("origin") = DtCount(0)(10)
                            Prod("date_modified") = DtCount(0)(11)
                            Prod("inventory_id") = DtCount(0)(12)
                            Prod("addontype") = DtCount(0)(13)
                            Prod("half_batch") = DtCount(0)(14)
                            FillDatagridProduct.Rows.Add(Prod)
                        End If
                    Else
                        'Insert new product
                        Query1 = "SELECT * FROM admin_products_org WHERE product_id = " & DtCountProductIds(i)(0)

                        cmd = New MySqlCommand(Query1, ConnectionServer)
                        DaCount = New MySqlDataAdapter(cmd)
                        DtCount = New DataTable
                        DaCount.Fill(DtCount)

                        Prod("product_id") = DtCount(0)(0)
                        Prod("product_sku") = DtCount(0)(1)
                        Prod("product_name") = DtCount(0)(2)
                        Prod("formula_id") = DtCount(0)(3)
                        Prod("product_barcode") = DtCount(0)(4)
                        Prod("product_category") = DtCount(0)(5)
                        Prod("product_price") = DtCount(0)(6)
                        Prod("product_desc") = DtCount(0)(7)
                        Prod("product_image") = DtCount(0)(8)
                        Prod("product_status") = DtCount(0)(9)
                        Prod("origin") = DtCount(0)(10)
                        Prod("date_modified") = DtCount(0)(11)
                        Prod("inventory_id") = DtCount(0)(12)
                        Prod("addontype") = DtCount(0)(13)
                        Prod("half_batch") = DtCount(0)(14)
                        FillDatagridProduct.Rows.Add(Prod)

                    End If

                Next


                ConnectionLocal.Close()
                ConnectionServer.Close()
            End If
        Catch ex As Exception
            BackgroundWorker2.CancelAsync()
            SendErrorReport(ex.ToString)
            Exit Sub
        End Try
    End Sub
    Private Sub GetAllProducts()
        Try
            Dim Connection As MySqlConnection = ServerCloudCon()
            Dim SqlCount = "SELECT product_id FROM admin_products_org"
            Dim CmdCount As MySqlCommand = New MySqlCommand(SqlCount, Connection)
            Dim DataAdapter As MySqlDataAdapter = New MySqlDataAdapter(CmdCount)
            Dim DaGet As DataTable = New DataTable
            DataAdapter.Fill(DaGet)

            Dim Cmd As MySqlCommand
            FillDatagridProduct = New DataTable
            FillDatagridProduct.Columns.Add("product_id")
            FillDatagridProduct.Columns.Add("product_sku")
            FillDatagridProduct.Columns.Add("product_name")
            FillDatagridProduct.Columns.Add("formula_id")
            FillDatagridProduct.Columns.Add("product_barcode")
            FillDatagridProduct.Columns.Add("product_category")
            FillDatagridProduct.Columns.Add("product_price")
            FillDatagridProduct.Columns.Add("product_desc")
            FillDatagridProduct.Columns.Add("product_image")
            FillDatagridProduct.Columns.Add("product_status")
            FillDatagridProduct.Columns.Add("origin")
            FillDatagridProduct.Columns.Add("date_modified")
            FillDatagridProduct.Columns.Add("inventory_id")
            FillDatagridProduct.Columns.Add("addontype")
            FillDatagridProduct.Columns.Add("half_batch")

            Dim DaCount As MySqlDataAdapter
            Dim FillDt As DataTable = New DataTable

            If WorkerUpdateCancel = False Then
                CheckingForUpdates.Instance.Invoke(Sub()
                                                       CheckingForUpdates.ProgressBar1.Maximum += DaGet.Rows.Count
                                                       'Console.WriteLine("PRODUCTS : " & CheckingForUpdates.ProgressBar1.Maximum)
                                                   End Sub)
            End If

            For a As Integer = 0 To DaGet.Rows.Count - 1 Step +1
                If WorkerUpdateCancel Then
                    Exit For
                End If
                CheckingForUpdates.Instance.Invoke(Sub()
                                                       If CheckingForUpdates.ProgressBar1.Maximum > 0 Then
                                                           CheckingForUpdates.ProgressBar1.Value += 1
                                                           CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                       End If
                                                   End Sub)
                Dim Query As String = "SELECT * FROM admin_products_org WHERE product_id = " & DaGet(a)(0)
                Cmd = New MySqlCommand(Query, Connection)
                DaCount = New MySqlDataAdapter(Cmd)
                FillDt = New DataTable
                DaCount.Fill(FillDt)
                For i As Integer = 0 To FillDt.Rows.Count - 1 Step +1
                    Dim Prod As DataRow = FillDatagridProduct.NewRow
                    Prod("product_id") = FillDt(i)(0)
                    Prod("product_sku") = FillDt(i)(1)
                    Prod("product_name") = FillDt(i)(2)
                    Prod("formula_id") = FillDt(i)(3)
                    Prod("product_barcode") = FillDt(i)(4)
                    Prod("product_category") = FillDt(i)(5)
                    Prod("product_price") = FillDt(i)(6)
                    Prod("product_desc") = FillDt(i)(7)
                    Prod("product_image") = FillDt(i)(8)
                    Prod("product_status") = FillDt(i)(9)
                    Prod("origin") = FillDt(i)(10)
                    Prod("date_modified") = FillDt(i)(11)
                    Prod("inventory_id") = FillDt(i)(12)
                    Prod("addontype") = FillDt(i)(13)
                    Prod("half_batch") = FillDt(i)(14)
                    FillDatagridProduct.Rows.Add(Prod)
                Next
            Next
        Catch ex As Exception
            BackgroundWorker2.CancelAsync()
            SendErrorReport(ex.ToString)
            Exit Sub
        End Try
    End Sub
#End Region
#Region "Formulas Update"
    Private Function LoadFormulaLocal() As DataTable
        Dim cmdlocal As MySqlCommand
        Dim dalocal As MySqlDataAdapter
        Dim dtlocal As DataTable = New DataTable
        dtlocal.Columns.Add("server_date_modified")
        dtlocal.Columns.Add("server_formula_id")
        Dim dtlocal1 As DataTable = New DataTable
        Try
            Dim sql = "SELECT server_date_modified, server_formula_id FROM loc_product_formula"
            cmdlocal = New MySqlCommand(sql, LocalhostConn)
            dalocal = New MySqlDataAdapter(cmdlocal)
            dalocal.Fill(dtlocal1)
            For i As Integer = 0 To dtlocal1.Rows.Count - 1 Step +1
                Dim Cat As DataRow = dtlocal.NewRow
                Cat("server_date_modified") = dtlocal1(i)(0).ToString
                Cat("server_formula_id") = dtlocal1(i)(1)
                dtlocal.Rows.Add(Cat)
            Next
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return dtlocal
    End Function
    Private Sub Function3()
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim ConnectionServer As MySqlConnection = ServerCloudCon()
            Dim FormulaLocal = LoadFormulaLocal()

            Dim Query = "SELECT * FROM loc_product_formula"
            Dim CmdCheck As MySqlCommand = New MySqlCommand(Query, ConnectionLocal)
            Dim DaCheck As MySqlDataAdapter = New MySqlDataAdapter(CmdCheck)
            Dim DtCheck As DataTable = New DataTable
            DaCheck.Fill(DtCheck)
            Dim cmdserver As MySqlCommand
            Dim daserver As MySqlDataAdapter
            Dim dtserver As DataTable
            If DtCheck.Rows.Count < 1 Then
                Dim sql = "SELECT `server_formula_id`, `product_ingredients`, `primary_unit`, `primary_value`, `secondary_unit`, `secondary_value`, `serving_unit`, `serving_value`, `no_servings`, `status`, `date_modified`, `unit_cost`, `origin` FROM admin_product_formula_org"
                cmdserver = New MySqlCommand(sql, ConnectionServer)
                daserver = New MySqlDataAdapter(cmdserver)
                dtserver = New DataTable
                daserver.Fill(dtserver)
                If WorkerUpdateCancel = False Then
                    CheckingForUpdates.Instance.Invoke(Sub()
                                                           CheckingForUpdates.ProgressBar1.Maximum += dtserver.Rows.Count
                                                           'Console.WriteLine("CATEGORY : " & CheckingForUpdates.ProgressBar1.Maximum)
                                                       End Sub)
                End If

                For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                    DataGridView3.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3), dtserver(i)(4), dtserver(i)(5), dtserver(i)(6), dtserver(i)(7), dtserver(i)(8), dtserver(i)(9), dtserver(i)(10).ToString, dtserver(i)(11), dtserver(i)(12))
                    If WorkerUpdateCancel = False Then
                        CheckingForUpdates.Instance.Invoke(Sub()
                                                               If DataGridViewUpdate.Rows.Count > 0 Then
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End If
                                                           End Sub)
                    End If

                Next
            Else
                Dim Ids As String = ""

                If ValidCloudConnection = True Then
                    For i As Integer = 0 To FormulaLocal.Rows.Count - 1 Step +1
                        If Ids = "" Then
                            Ids = "" & FormulaLocal(i)(1) & ""
                        Else
                            Ids += "," & FormulaLocal(i)(1) & ""
                        End If
                    Next
                    Dim sql = "SELECT `server_formula_id`, `product_ingredients`, `primary_unit`, `primary_value`, `secondary_unit`, `secondary_value`, `serving_unit`, `serving_value`, `no_servings`, `status`, `date_modified`, `unit_cost`, `origin` FROM admin_product_formula_org WHERE server_formula_id  IN (" & Ids & ") "
                    cmdserver = New MySqlCommand(sql, ConnectionServer)
                    daserver = New MySqlDataAdapter(cmdserver)
                    dtserver = New DataTable
                    daserver.Fill(dtserver)
                    For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                        If FormulaLocal(i)(0).ToString <> dtserver(i)(10).ToString Then
                            DataGridView3.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3), dtserver(i)(4), dtserver(i)(5), dtserver(i)(6), dtserver(i)(7), dtserver(i)(8), dtserver(i)(9), dtserver(i)(10).ToString, dtserver(i)(11), dtserver(i)(12))
                        End If
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   If DataGridViewUpdate.Rows.Count > 0 Then
                                                                       CheckingForUpdates.ProgressBar1.Value += 1
                                                                       CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                                   End If
                                                               End Sub)
                        End If

                    Next
                    Dim sql2 = "SELECT `server_formula_id`, `product_ingredients`, `primary_unit`, `primary_value`, `secondary_unit`, `secondary_value`, `serving_unit`, `serving_value`, `no_servings`, `status`, `date_modified`, `unit_cost`, `origin` FROM admin_product_formula_org WHERE server_formula_id NOT IN (" & Ids & ") "
                    cmdserver = New MySqlCommand(sql2, ConnectionServer)
                    daserver = New MySqlDataAdapter(cmdserver)
                    dtserver = New DataTable
                    daserver.Fill(dtserver)
                    For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                        If FormulaLocal(i)(0).ToString <> dtserver(i)(10) Then
                            DataGridView3.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3), dtserver(i)(4), dtserver(i)(5), dtserver(i)(6), dtserver(i)(7), dtserver(i)(8), dtserver(i)(9), dtserver(i)(10).ToString, dtserver(i)(11), dtserver(i)(12))
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            BackgroundWorker2.CancelAsync()
            SendErrorReport(ex.ToString)
            Exit Sub
        End Try
    End Sub
#End Region
#Region "Inventory Update"
    Private Function LoadInventoryLocal() As DataTable
        Dim cmdlocal As MySqlCommand
        Dim dalocal As MySqlDataAdapter
        Dim dtlocal As DataTable = New DataTable
        dtlocal.Columns.Add("server_date_modified")
        dtlocal.Columns.Add("server_inventory_id")
        Dim dtlocal1 As DataTable = New DataTable
        Try
            Dim sql = "SELECT server_date_modified , server_inventory_id FROM loc_pos_inventory"
            cmdlocal = New MySqlCommand(sql, LocalhostConn)
            dalocal = New MySqlDataAdapter(cmdlocal)
            dalocal.Fill(dtlocal)
            For i As Integer = 0 To dtlocal1.Rows.Count - 1 Step +1
                Dim Cat As DataRow = dtlocal.NewRow
                Cat("server_date_modified") = dtlocal1(i)(0).ToString
                Cat("server_inventory_id") = dtlocal1(i)(1)
                dtlocal.Rows.Add(Cat)
            Next
            LocalhostConn.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return dtlocal
    End Function
    Private Sub Function4()
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim ConnectionServer As MySqlConnection = ServerCloudCon()
            Dim InventoryLocal = LoadInventoryLocal()

            Dim Query = "SELECT * FROM loc_pos_inventory"
            Dim CmdCheck As MySqlCommand = New MySqlCommand(Query, ConnectionLocal)
            Dim DaCheck As MySqlDataAdapter = New MySqlDataAdapter(CmdCheck)
            Dim DtCheck As DataTable = New DataTable
            DaCheck.Fill(DtCheck)
            Dim cmdserver As MySqlCommand
            Dim daserver As MySqlDataAdapter
            Dim dtserver As DataTable
            If DtCheck.Rows.Count < 1 Then
                Dim sql = "SELECT `server_inventory_id`, `product_ingredients`, `sku`, `stock_primary`, `stock_secondary`, `stock_no_of_servings`, `stock_status`, `critical_limit`, `date_modified`, `main_inventory_id`, `origin`, `show_stockin` FROM admin_pos_inventory_org"
                cmdserver = New MySqlCommand(sql, ConnectionServer)
                daserver = New MySqlDataAdapter(cmdserver)
                dtserver = New DataTable
                daserver.Fill(dtserver)
                If WorkerUpdateCancel = False Then
                    CheckingForUpdates.Instance.Invoke(Sub()
                                                           CheckingForUpdates.ProgressBar1.Maximum += dtserver.Rows.Count
                                                           'Console.WriteLine("INVENTORY : " & CheckingForUpdates.ProgressBar1.Maximum)
                                                       End Sub)
                End If

                For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                    DataGridView4.Rows.Add(dtserver(i)(0), 0, dtserver(i)(1), dtserver(i)(2), dtserver(i)(3), dtserver(i)(4), dtserver(i)(5), dtserver(i)(6), dtserver(i)(7), dtserver(i)(8).ToString, dtserver(i)(9).ToString, dtserver(i)(10).ToString, dtserver(i)(11))
                    If WorkerUpdateCancel = False Then
                        CheckingForUpdates.Instance.Invoke(Sub()
                                                               If DataGridViewUpdate.Rows.Count > 0 Then
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End If
                                                           End Sub)
                    End If

                Next
            Else
                Dim Ids As String = ""

                If ValidCloudConnection = True Then
                    For i As Integer = 0 To InventoryLocal.Rows.Count - 1 Step +1
                        If Ids = "" Then
                            Ids = "" & InventoryLocal(i)(1) & ""
                        Else
                            Ids += "," & InventoryLocal(i)(1) & ""
                        End If
                    Next
                    Dim sql = "SELECT `server_inventory_id`, `product_ingredients`, `sku`, `stock_primary`, `stock_secondary`, `stock_no_of_servings`, `stock_status`, `critical_limit`, `date_modified`,`main_inventory_id`, `origin`, `show_stockin` FROM admin_pos_inventory_org WHERE server_inventory_id IN (" & Ids & ")"
                    cmdserver = New MySqlCommand(sql, ConnectionServer)
                    daserver = New MySqlDataAdapter(cmdserver)
                    dtserver = New DataTable
                    daserver.Fill(dtserver)
                    For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                        If InventoryLocal(i)(0).ToString <> dtserver(i)(8).ToString Then
                            DataGridView4.Rows.Add(dtserver(i)(0), 0, dtserver(i)(1), dtserver(i)(2), dtserver(i)(3), dtserver(i)(4), dtserver(i)(5), dtserver(i)(6), dtserver(i)(7), dtserver(i)(8).ToString, dtserver(i)(9).ToString, dtserver(i)(10).ToString, dtserver(i)(11))
                        End If
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   If DataGridViewUpdate.Rows.Count > 0 Then
                                                                       CheckingForUpdates.ProgressBar1.Value += 1
                                                                       CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                                   End If
                                                               End Sub)
                        End If

                    Next
                    Dim sql2 = "SELECT `server_inventory_id`, `product_ingredients`, `sku`, `stock_primary`, `stock_secondary`, `stock_no_of_servings`, `stock_status`, `critical_limit`, `date_modified`,`main_inventory_id`, `origin`, `show_stockin` FROM admin_pos_inventory_org WHERE server_inventory_id NOT IN (" & Ids & ")"
                    cmdserver = New MySqlCommand(sql2, ConnectionServer)
                    daserver = New MySqlDataAdapter(cmdserver)
                    dtserver = New DataTable
                    daserver.Fill(dtserver)
                    For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                        If InventoryLocal(i)(0).ToString <> dtserver(i)(8) Then
                            DataGridView4.Rows.Add(dtserver(i)(0), 0, dtserver(i)(1), dtserver(i)(2), dtserver(i)(3), dtserver(i)(4), dtserver(i)(5), dtserver(i)(6), dtserver(i)(7), dtserver(i)(8).ToString, dtserver(i)(9).ToString, dtserver(i)(10).ToString, dtserver(i)(11))
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            BackgroundWorker2.CancelAsync()
            SendErrorReport(ex.ToString)
            Exit Sub
        End Try
    End Sub
    Private Function LoadPartnersCategory() As DataTable
        Dim cmdlocal As MySqlCommand
        Dim dalocal As MySqlDataAdapter
        Dim dtlocal As DataTable = New DataTable
        dtlocal.Columns.Add("date_modified")
        dtlocal.Columns.Add("id")
        Dim dtlocal1 As DataTable = New DataTable
        Try
            Dim sql = "SELECT date_modified, id FROM loc_partners_transaction"
            cmdlocal = New MySqlCommand(sql, LocalhostConn())
            dalocal = New MySqlDataAdapter(cmdlocal)
            dalocal.Fill(dtlocal1)
            For i As Integer = 0 To dtlocal1.Rows.Count - 1 Step +1
                Dim Cat As DataRow = dtlocal.NewRow
                Cat("date_modified") = dtlocal1(i)(0).ToString
                Cat("id") = dtlocal1(i)(1)
                dtlocal.Rows.Add(Cat)
            Next
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return dtlocal
    End Function
    Private Sub Function6()
        Try
            Dim Query = "SELECT * FROM loc_partners_transaction"
            Dim CmdCheck As MySqlCommand = New MySqlCommand(Query, LocalhostConn)
            Dim DaCheck As MySqlDataAdapter = New MySqlDataAdapter(CmdCheck)
            Dim DtCheck As DataTable = New DataTable
            DaCheck.Fill(DtCheck)
            Dim cmdserver As MySqlCommand
            Dim daserver As MySqlDataAdapter
            Dim dtserver As DataTable
            If DtCheck.Rows.Count < 1 Then
                Dim sql = "SELECT `id`, `arrid`, `bankname`, `date_modified`, `active` FROM admin_partners_transaction_org"
                cmdserver = New MySqlCommand(sql, ServerCloudCon())
                daserver = New MySqlDataAdapter(cmdserver)
                dtserver = New DataTable
                daserver.Fill(dtserver)
                If WorkerUpdateCancel = False Then
                    CheckingForUpdates.Instance.Invoke(Sub()
                                                           CheckingForUpdates.ProgressBar1.Maximum += dtserver.Rows.Count
                                                           'Console.WriteLine("PARTNERS : " & CheckingForUpdates.ProgressBar1.Maximum)
                                                       End Sub)
                End If

                For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                    DataGridViewPartners.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3).ToString, dtserver(i)(4))
                    If WorkerUpdateCancel = False Then
                        CheckingForUpdates.Instance.Invoke(Sub()
                                                               If DataGridViewUpdate.Rows.Count > 0 Then
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End If
                                                           End Sub)
                    End If

                Next
            Else
                Dim Ids As String = ""
                If ValidCloudConnection = True Then
                    For i As Integer = 0 To LoadPartnersCategory.Rows.Count - 1 Step +1
                        If Ids = "" Then
                            Ids = "" & LoadPartnersCategory(i)(1) & ""
                        Else
                            Ids += "," & LoadPartnersCategory(i)(1) & ""
                        End If
                    Next
                    Dim sql = "SELECT `id`, `arrid`, `bankname`, `date_modified`, `active` FROM admin_partners_transaction_org WHERE id IN (" & Ids & ")"
                    cmdserver = New MySqlCommand(sql, ServerCloudCon())
                    daserver = New MySqlDataAdapter(cmdserver)
                    dtserver = New DataTable
                    daserver.Fill(dtserver)
                    For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                        If LoadPartnersCategory(i)(0).ToString <> dtserver(i)(3).ToString Then
                            DataGridViewPartners.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3).ToString, dtserver(i)(4))
                        End If
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   If DataGridViewUpdate.Rows.Count > 0 Then
                                                                       CheckingForUpdates.ProgressBar1.Value += 1
                                                                       CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                                   End If
                                                               End Sub)
                        End If

                    Next
                    Dim sql2 = "SELECT `id`, `arrid`, `bankname`, `date_modified`, `active` FROM admin_partners_transaction_org WHERE id NOT IN (" & Ids & ")"
                    cmdserver = New MySqlCommand(sql2, ServerCloudCon())
                    daserver = New MySqlDataAdapter(cmdserver)
                    dtserver = New DataTable
                    daserver.Fill(dtserver)
                    For i As Integer = 0 To dtserver.Rows.Count - 1 Step +1
                        If LoadPartnersCategory(i)(0) <> dtserver(i)(3) Then
                            DataGridViewPartners.Rows.Add(dtserver(i)(0), dtserver(i)(1), dtserver(i)(2), dtserver(i)(3).ToString, dtserver(i)(4))
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            BackgroundWorker2.CancelAsync()
            SendErrorReport(ex.ToString)
            Exit Sub
            'If table doesnt have data
        End Try
    End Sub
    Private Sub GetRowCount()
        Try
            Dim Products = count("product_id", "loc_admin_products")
            Dim Category = count("category_id", "loc_admin_category")
            Dim Inventory = count("inventory_id", "loc_pos_inventory")
            Dim Formula = count("formula_id", "loc_product_formula")
            Dim Coupons = count("ID", "tbcoupon")
            Dim Partners = count("id", "loc_partners_transaction")
            DataGridViewUpdate.Rows.Add(Products)
            DataGridViewUpdate.Rows.Add(Category)
            DataGridViewUpdate.Rows.Add(Inventory)
            DataGridViewUpdate.Rows.Add(Formula)
            DataGridViewUpdate.Rows.Add(Coupons)
            DataGridViewUpdate.Rows.Add(Partners)
            If WorkerUpdateCancel = False Then
                CheckingForUpdates.Instance.Invoke(Sub()
                                                       CheckingForUpdates.ProgressBar1.Maximum = SumOfColumnsToInt(DataGridViewUpdate, 0)
                                                   End Sub)
            End If

        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InstallUpdatesCategory()
        Try
            Dim Connection As MySqlConnection = LocalhostConn()
            Dim cmdlocal As MySqlCommand
            With DataGridView1
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    Dim sql = "SELECT category_id FROM loc_admin_category WHERE category_id = " & .Rows(i).Cells(0).Value
                    cmdlocal = New MySqlCommand(sql, Connection)
                    Dim result As Integer = cmdlocal.ExecuteScalar
                    If result = 0 Then
                        Dim sqlinsert = "INSERT INTO `loc_admin_category`(`category_name`, `brand_name`, `updated_at`, `origin`, `status`) VALUES (@0,@1,@2,@3,@4)"
                        cmdlocal = New MySqlCommand(sqlinsert, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.VarChar).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.VarChar).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.VarChar).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.VarChar).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@4", MySqlDbType.Int64).Value = .Rows(i).Cells(5).Value.ToString()
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If

                    Else
                        Dim sqlupdate = "UPDATE `loc_admin_category` SET `category_name`=@0,`brand_name`=@1,`updated_at`=@2,`origin`=@3,`status`=@4 WHERE category_id = " & .Rows(i).Cells(0).Value
                        cmdlocal = New MySqlCommand(sqlupdate, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.VarChar).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.VarChar).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.VarChar).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.VarChar).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@4", MySqlDbType.Int64).Value = .Rows(i).Cells(5).Value.ToString()
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If

                    End If
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Private Sub InstallUpdatesPartners()
        Try
            Dim Connection As MySqlConnection = LocalhostConn()
            Dim cmdlocal As MySqlCommand
            With DataGridViewPartners
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    Dim sql = "SELECT id FROM loc_partners_transaction WHERE id = " & .Rows(i).Cells(0).Value
                    cmdlocal = New MySqlCommand(sql, Connection)
                    Dim result As Integer = cmdlocal.ExecuteScalar
                    If result = 0 Then
                        Dim sqlinsert = "INSERT INTO `loc_partners_transaction`(`arrid`, `bankname`, `date_modified`, `crew_id`, `store_id`, `guid`, `active`, `synced`) VALUES (@0,@1,@2,@3,@4,@5,@6,@7)"
                        cmdlocal = New MySqlCommand(sqlinsert, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.Int64).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.VarChar).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.Text).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.VarChar).Value = "Admin"
                        cmdlocal.Parameters.Add("@4", MySqlDbType.VarChar).Value = ClientStoreID
                        cmdlocal.Parameters.Add("@5", MySqlDbType.VarChar).Value = ClientGuid
                        cmdlocal.Parameters.Add("@6", MySqlDbType.Int64).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@7", MySqlDbType.VarChar).Value = "Synced"
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If

                    Else
                        Dim sqlupdate = "UPDATE `loc_partners_transaction` SET `arrid`=@0,`bankname`=@1,`date_modified`=@2,`crew_id`=@3,`store_id`=@4,`guid`=@5,`active`=@6,`synced`=@7 WHERE id = " & .Rows(i).Cells(0).Value
                        cmdlocal = New MySqlCommand(sqlupdate, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.Int64).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.VarChar).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.Text).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.VarChar).Value = "Admin"
                        cmdlocal.Parameters.Add("@4", MySqlDbType.VarChar).Value = ClientStoreID
                        cmdlocal.Parameters.Add("@5", MySqlDbType.VarChar).Value = ClientGuid
                        cmdlocal.Parameters.Add("@6", MySqlDbType.Int64).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@7", MySqlDbType.VarChar).Value = "Synced"
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Private Sub InstallUpdatesCoupons()
        Try
            Dim Connection As MySqlConnection = LocalhostConn()
            Dim cmdlocal As MySqlCommand
            With DataGridView5
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    Dim sql = "SELECT ID FROM tbcoupon WHERE ID = " & .Rows(i).Cells(0).Value
                    cmdlocal = New MySqlCommand(sql, Connection)
                    Dim result As Integer = cmdlocal.ExecuteScalar
                    If result = 0 Then
                        Dim sqlinsert = "INSERT INTO `tbcoupon`(`Couponname_`, `Desc_`, `Discountvalue_`, `Referencevalue_`, `Type`, `Bundlebase_`, `BBValue_`, `Bundlepromo_`, `BPValue_`, `Effectivedate`, `Expirydate`, `active`, `store_id`, `crew_id`, `guid`, `origin`, `synced`, `date_created`) VALUES (@0,@1,@2,@3,@4,@5,@6,@7,@8,@9,@10,@11,@12,@13,@14,@15,@16,@17)"
                        cmdlocal = New MySqlCommand(sqlinsert, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.Text).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.Text).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.Text).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.Text).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@4", MySqlDbType.Text).Value = .Rows(i).Cells(5).Value.ToString()
                        cmdlocal.Parameters.Add("@5", MySqlDbType.Text).Value = .Rows(i).Cells(6).Value.ToString()
                        cmdlocal.Parameters.Add("@6", MySqlDbType.Text).Value = .Rows(i).Cells(7).Value.ToString()
                        cmdlocal.Parameters.Add("@7", MySqlDbType.Text).Value = .Rows(i).Cells(8).Value.ToString()
                        cmdlocal.Parameters.Add("@8", MySqlDbType.Text).Value = .Rows(i).Cells(9).Value.ToString()
                        cmdlocal.Parameters.Add("@9", MySqlDbType.Text).Value = .Rows(i).Cells(10).Value.ToString()
                        cmdlocal.Parameters.Add("@10", MySqlDbType.Text).Value = .Rows(i).Cells(11).Value.ToString()
                        cmdlocal.Parameters.Add("@11", MySqlDbType.Text).Value = "1"
                        cmdlocal.Parameters.Add("@12", MySqlDbType.Text).Value = ClientStoreID
                        cmdlocal.Parameters.Add("@13", MySqlDbType.Text).Value = ClientCrewID
                        cmdlocal.Parameters.Add("@14", MySqlDbType.Text).Value = ClientGuid
                        cmdlocal.Parameters.Add("@15", MySqlDbType.Text).Value = "Server"
                        cmdlocal.Parameters.Add("@16", MySqlDbType.Text).Value = "Synced"
                        cmdlocal.Parameters.Add("@17", MySqlDbType.Text).Value = .Rows(i).Cells(12).Value.ToString()
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If

                    Else
                        Dim sqlupdate = "UPDATE `tbcoupon` SET `Couponname_` = @0, `Desc_` = @1, `Discountvalue_` = @2, `Referencevalue_` = @3, `Type` = @4, `Bundlebase_` = @5, `BBValue_` = @6, `Bundlepromo_` = @7, `BPValue_` = @8, `Effectivedate` = @9, `Expirydate` = @10, `active` = @11, `store_id` = @12, `crew_id` = @13, `guid` = @14, `origin` = @15, `synced` = @16, `date_created` = @17 WHERE ID = " & .Rows(i).Cells(0).Value
                        cmdlocal = New MySqlCommand(sqlupdate, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.Text).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.Text).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.Text).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.Text).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@4", MySqlDbType.Text).Value = .Rows(i).Cells(5).Value.ToString()
                        cmdlocal.Parameters.Add("@5", MySqlDbType.Text).Value = .Rows(i).Cells(6).Value.ToString()
                        cmdlocal.Parameters.Add("@6", MySqlDbType.Text).Value = .Rows(i).Cells(7).Value.ToString()
                        cmdlocal.Parameters.Add("@7", MySqlDbType.Text).Value = .Rows(i).Cells(8).Value.ToString()
                        cmdlocal.Parameters.Add("@8", MySqlDbType.Text).Value = .Rows(i).Cells(9).Value.ToString()
                        cmdlocal.Parameters.Add("@9", MySqlDbType.Text).Value = .Rows(i).Cells(10).Value.ToString()
                        cmdlocal.Parameters.Add("@10", MySqlDbType.Text).Value = .Rows(i).Cells(11).Value.ToString()
                        cmdlocal.Parameters.Add("@11", MySqlDbType.Text).Value = "1"
                        cmdlocal.Parameters.Add("@12", MySqlDbType.Text).Value = ClientStoreID
                        cmdlocal.Parameters.Add("@13", MySqlDbType.Text).Value = ClientCrewID
                        cmdlocal.Parameters.Add("@14", MySqlDbType.Text).Value = ClientGuid
                        cmdlocal.Parameters.Add("@15", MySqlDbType.Text).Value = "Server"
                        cmdlocal.Parameters.Add("@16", MySqlDbType.Text).Value = "Synced"
                        cmdlocal.Parameters.Add("@17", MySqlDbType.Text).Value = .Rows(i).Cells(12).Value.ToString()
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If

                    End If
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InstallUpdatesFormula()
        Try
            Dim Connection As MySqlConnection = LocalhostConn()
            Dim cmdlocal As MySqlCommand
            With DataGridView3
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    Dim sql = "SELECT formula_id FROM loc_product_formula WHERE formula_id = " & .Rows(i).Cells(0).Value
                    cmdlocal = New MySqlCommand(sql, Connection)
                    Dim result As Integer = cmdlocal.ExecuteScalar
                    If result = 0 Then
                        Dim sqlinsert = "INSERT INTO loc_product_formula (`server_formula_id`,`product_ingredients`, `primary_unit`, `primary_value`, `secondary_unit`, `secondary_value`, `serving_unit`, `serving_value`, `no_servings`, `status`, `date_modified`, `unit_cost`, `origin`, `store_id`, `guid`, `crew_id`, `server_date_modified`) VALUES
                                        (@0 ,@1, @2, @3, @4, @5, @6, @7, @8, @9, @10, @11 , @12 , @13 , @14, @15, @16)"
                        cmdlocal = New MySqlCommand(sqlinsert, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.Int64).Value = .Rows(i).Cells(0).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.VarChar).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.VarChar).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.VarChar).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@4", MySqlDbType.VarChar).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@5", MySqlDbType.VarChar).Value = .Rows(i).Cells(5).Value.ToString()
                        cmdlocal.Parameters.Add("@6", MySqlDbType.VarChar).Value = .Rows(i).Cells(6).Value.ToString()
                        cmdlocal.Parameters.Add("@7", MySqlDbType.VarChar).Value = .Rows(i).Cells(7).Value.ToString()
                        cmdlocal.Parameters.Add("@8", MySqlDbType.VarChar).Value = .Rows(i).Cells(8).Value.ToString()
                        cmdlocal.Parameters.Add("@9", MySqlDbType.Int64).Value = .Rows(i).Cells(9).Value.ToString()
                        cmdlocal.Parameters.Add("@10", MySqlDbType.VarChar).Value = .Rows(i).Cells(10).Value
                        cmdlocal.Parameters.Add("@11", MySqlDbType.Decimal).Value = .Rows(i).Cells(11).Value.ToString()
                        cmdlocal.Parameters.Add("@12", MySqlDbType.VarChar).Value = .Rows(i).Cells(12).Value.ToString()
                        cmdlocal.Parameters.Add("@13", MySqlDbType.VarChar).Value = ClientStoreID
                        cmdlocal.Parameters.Add("@14", MySqlDbType.VarChar).Value = ClientGuid
                        cmdlocal.Parameters.Add("@15", MySqlDbType.VarChar).Value = "0"
                        cmdlocal.Parameters.Add("@16", MySqlDbType.Text).Value = .Rows(i).Cells(10).Value.ToString()
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If
                    Else
                        Dim sqlupdate = "UPDATE `loc_product_formula` SET `server_formula_id`= @0,`product_ingredients`= @1,`primary_unit`= @2,`primary_value`= @3,`secondary_unit`= @4,`secondary_value`=@5,`serving_unit`=@6,`serving_value`=@7,`no_servings`=@8,`status`=@9,`date_modified`=@10,`unit_cost`=@11,`origin`=@12,`store_id`=@13,`guid`=@14,`crew_id`=@15,`server_date_modified`=@16 WHERE server_formula_id =  " & .Rows(i).Cells(0).Value
                        cmdlocal = New MySqlCommand(sqlupdate, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.Int64).Value = .Rows(i).Cells(0).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.VarChar).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.VarChar).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.VarChar).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@4", MySqlDbType.VarChar).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@5", MySqlDbType.VarChar).Value = .Rows(i).Cells(5).Value.ToString()
                        cmdlocal.Parameters.Add("@6", MySqlDbType.VarChar).Value = .Rows(i).Cells(6).Value.ToString()
                        cmdlocal.Parameters.Add("@7", MySqlDbType.VarChar).Value = .Rows(i).Cells(7).Value.ToString()
                        cmdlocal.Parameters.Add("@8", MySqlDbType.VarChar).Value = .Rows(i).Cells(8).Value.ToString()
                        cmdlocal.Parameters.Add("@9", MySqlDbType.Int64).Value = .Rows(i).Cells(9).Value.ToString()
                        cmdlocal.Parameters.Add("@10", MySqlDbType.VarChar).Value = .Rows(i).Cells(10).Value
                        cmdlocal.Parameters.Add("@11", MySqlDbType.Decimal).Value = .Rows(i).Cells(11).Value.ToString()
                        cmdlocal.Parameters.Add("@12", MySqlDbType.VarChar).Value = .Rows(i).Cells(12).Value.ToString()
                        cmdlocal.Parameters.Add("@13", MySqlDbType.VarChar).Value = ClientStoreID
                        cmdlocal.Parameters.Add("@14", MySqlDbType.VarChar).Value = ClientGuid
                        cmdlocal.Parameters.Add("@15", MySqlDbType.VarChar).Value = "0"
                        cmdlocal.Parameters.Add("@16", MySqlDbType.Text).Value = .Rows(i).Cells(10).Value.ToString()
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InstallUpdatesInventory()
        Try
            Dim Connection As MySqlConnection = LocalhostConn()
            Dim cmdlocal As MySqlCommand
            With DataGridView4
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    Dim sql = "SELECT inventory_id FROM loc_pos_inventory WHERE inventory_id = " & .Rows(i).Cells(0).Value
                    cmdlocal = New MySqlCommand(sql, Connection)
                    Dim result As Integer = cmdlocal.ExecuteScalar
                    If result = 0 Then
                        Dim sqlinsert = "INSERT INTO loc_pos_inventory (`server_inventory_id`,`formula_id`,`product_ingredients`,`sku`,`stock_primary`,`stock_secondary`,`stock_no_of_servings`,`stock_status`,`critical_limit`,`date_modified`,`server_date_modified`,`store_id`,`crew_id`,`guid`,`synced`,`main_inventory_id`,`origin`,`show_stockin`) VALUES
                                        (@0 ,@1, @2, @3, @4, @5, @6, @7, @8, @9, @10, @11, @12, @13, @14, @15, @16, @17)"
                        cmdlocal = New MySqlCommand(sqlinsert, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.Int64).Value = .Rows(i).Cells(0).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.Int64).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.VarChar).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.VarChar).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@4", MySqlDbType.Decimal).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@5", MySqlDbType.Decimal).Value = .Rows(i).Cells(5).Value.ToString()
                        cmdlocal.Parameters.Add("@6", MySqlDbType.Decimal).Value = .Rows(i).Cells(6).Value.ToString()
                        cmdlocal.Parameters.Add("@7", MySqlDbType.Int64).Value = .Rows(i).Cells(7).Value.ToString()
                        cmdlocal.Parameters.Add("@8", MySqlDbType.Int64).Value = .Rows(i).Cells(8).Value.ToString()
                        cmdlocal.Parameters.Add("@9", MySqlDbType.Text).Value = .Rows(i).Cells(9).Value.ToString()
                        cmdlocal.Parameters.Add("@10", MySqlDbType.Text).Value = .Rows(i).Cells(9).Value.ToString()
                        cmdlocal.Parameters.Add("@11", MySqlDbType.VarChar).Value = ClientStoreID
                        cmdlocal.Parameters.Add("@12", MySqlDbType.VarChar).Value = "0"
                        cmdlocal.Parameters.Add("@13", MySqlDbType.VarChar).Value = ClientGuid
                        cmdlocal.Parameters.Add("@14", MySqlDbType.VarChar).Value = "Synced"
                        cmdlocal.Parameters.Add("@15", MySqlDbType.VarChar).Value = .Rows(i).Cells(10).Value.ToString()
                        cmdlocal.Parameters.Add("@16", MySqlDbType.Text).Value = .Rows(i).Cells(11).Value.ToString()
                        cmdlocal.Parameters.Add("@17", MySqlDbType.Int64).Value = .Rows(i).Cells(12).Value
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If
                    Else
                        Dim sqlUpdate = "UPDATE `loc_pos_inventory` SET `server_inventory_id`= @0,`formula_id`=@1,`product_ingredients`=@2,`sku`=@3,`stock_status`=@7,`critical_limit`=@8,`date_modified`=@9,`server_date_modified`=@10,`store_id`=@11,`crew_id`=@12,`guid`=@13,`synced`=@14,`main_inventory_id`=@15,`origin`=@16,`show_stockin`= @17 WHERE `server_inventory_id`= " & .Rows(i).Cells(0).Value
                        cmdlocal = New MySqlCommand(sqlUpdate, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.Int64).Value = .Rows(i).Cells(0).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.Int64).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.VarChar).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.VarChar).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@7", MySqlDbType.Int64).Value = .Rows(i).Cells(7).Value.ToString()
                        cmdlocal.Parameters.Add("@8", MySqlDbType.Int64).Value = .Rows(i).Cells(8).Value.ToString()
                        cmdlocal.Parameters.Add("@9", MySqlDbType.Text).Value = .Rows(i).Cells(9).Value.ToString()
                        cmdlocal.Parameters.Add("@10", MySqlDbType.Text).Value = .Rows(i).Cells(9).Value.ToString()
                        cmdlocal.Parameters.Add("@11", MySqlDbType.VarChar).Value = ClientStoreID
                        cmdlocal.Parameters.Add("@12", MySqlDbType.VarChar).Value = "0"
                        cmdlocal.Parameters.Add("@13", MySqlDbType.VarChar).Value = ClientGuid
                        cmdlocal.Parameters.Add("@14", MySqlDbType.VarChar).Value = "Synced"
                        cmdlocal.Parameters.Add("@15", MySqlDbType.VarChar).Value = .Rows(i).Cells(10).Value.ToString()
                        cmdlocal.Parameters.Add("@16", MySqlDbType.Text).Value = .Rows(i).Cells(11).Value.ToString()
                        cmdlocal.Parameters.Add("@17", MySqlDbType.Int64).Value = .Rows(i).Cells(12).Value
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InstallUpdatesProducts()
        Try
            Dim Connection As MySqlConnection = LocalhostConn()
            Dim cmdlocal As MySqlCommand
            With DataGridView2
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    Dim sql = "SELECT product_id FROM loc_admin_products WHERE server_product_id = " & .Rows(i).Cells(0).Value
                    cmdlocal = New MySqlCommand(sql, Connection)
                    Dim result As Integer = cmdlocal.ExecuteScalar
                    If result = 0 Then
                        Dim sqlinsert = "INSERT INTO loc_admin_products (`server_product_id`, `product_sku`, `product_name`, `formula_id`, `product_barcode`, `product_category`, `product_price`, `product_desc`, `product_image`, `product_status`, `origin`, `date_modified`, `server_inventory_id`, `guid`, `store_id`, `crew_id`, `synced`, `addontype`, `half_batch`) VALUES
                                        (@0 ,@1, @2, @3, @4, @5, @6, @7, @8, @9, @10, @11, @12, @13, @14, @15, @16, @17, @18)"
                        cmdlocal = New MySqlCommand(sqlinsert, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.Int64).Value = .Rows(i).Cells(0).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.VarChar).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.VarChar).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.VarChar).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@4", MySqlDbType.VarChar).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@5", MySqlDbType.VarChar).Value = .Rows(i).Cells(5).Value.ToString()
                        cmdlocal.Parameters.Add("@6", MySqlDbType.Int64).Value = .Rows(i).Cells(6).Value.ToString()
                        cmdlocal.Parameters.Add("@7", MySqlDbType.VarChar).Value = .Rows(i).Cells(7).Value.ToString()
                        cmdlocal.Parameters.Add("@8", MySqlDbType.VarChar).Value = .Rows(i).Cells(8).Value.ToString()
                        cmdlocal.Parameters.Add("@9", MySqlDbType.VarChar).Value = .Rows(i).Cells(9).Value.ToString()
                        cmdlocal.Parameters.Add("@10", MySqlDbType.VarChar).Value = .Rows(i).Cells(10).Value.ToString()
                        cmdlocal.Parameters.Add("@11", MySqlDbType.VarChar).Value = .Rows(i).Cells(11).Value.ToString()
                        cmdlocal.Parameters.Add("@12", MySqlDbType.Text).Value = .Rows(i).Cells(12).Value.ToString()
                        cmdlocal.Parameters.Add("@13", MySqlDbType.VarChar).Value = ClientGuid
                        cmdlocal.Parameters.Add("@14", MySqlDbType.Int64).Value = ClientStoreID
                        cmdlocal.Parameters.Add("@15", MySqlDbType.VarChar).Value = "0"
                        cmdlocal.Parameters.Add("@16", MySqlDbType.VarChar).Value = "Synced"
                        cmdlocal.Parameters.Add("@17", MySqlDbType.Text).Value = .Rows(i).Cells(13).Value.ToString()
                        cmdlocal.Parameters.Add("@18", MySqlDbType.Int64).Value = .Rows(i).Cells(14).Value.ToString()
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If
                    Else
                        Dim sqlupdate = "UPDATE `loc_admin_products` SET `server_product_id`=@0,`product_sku`=@1,`product_name`=@2,`formula_id`=@3,`product_barcode`=@4,`product_category`=@5,`product_price`=@6,`product_desc`=@7,`product_image`=@8,`product_status`=@9,`origin`=@10,`date_modified`=@11,`server_inventory_id`=@12,`guid`=@13,`store_id`=@14,`crew_id`=@15,`synced`=@16,`addontype`=@17,`half_batch`=@18 WHERE server_product_id =  " & .Rows(i).Cells(0).Value
                        cmdlocal = New MySqlCommand(sqlupdate, Connection)
                        cmdlocal.Parameters.Add("@0", MySqlDbType.Int64).Value = .Rows(i).Cells(0).Value.ToString()
                        cmdlocal.Parameters.Add("@1", MySqlDbType.VarChar).Value = .Rows(i).Cells(1).Value.ToString()
                        cmdlocal.Parameters.Add("@2", MySqlDbType.VarChar).Value = .Rows(i).Cells(2).Value.ToString()
                        cmdlocal.Parameters.Add("@3", MySqlDbType.VarChar).Value = .Rows(i).Cells(3).Value.ToString()
                        cmdlocal.Parameters.Add("@4", MySqlDbType.VarChar).Value = .Rows(i).Cells(4).Value.ToString()
                        cmdlocal.Parameters.Add("@5", MySqlDbType.VarChar).Value = .Rows(i).Cells(5).Value.ToString()
                        cmdlocal.Parameters.Add("@6", MySqlDbType.Int64).Value = .Rows(i).Cells(6).Value.ToString()
                        cmdlocal.Parameters.Add("@7", MySqlDbType.VarChar).Value = .Rows(i).Cells(7).Value.ToString()
                        cmdlocal.Parameters.Add("@8", MySqlDbType.VarChar).Value = .Rows(i).Cells(8).Value.ToString()
                        cmdlocal.Parameters.Add("@9", MySqlDbType.VarChar).Value = .Rows(i).Cells(9).Value.ToString()
                        cmdlocal.Parameters.Add("@10", MySqlDbType.VarChar).Value = .Rows(i).Cells(10).Value.ToString()
                        cmdlocal.Parameters.Add("@11", MySqlDbType.VarChar).Value = .Rows(i).Cells(11).Value.ToString()
                        cmdlocal.Parameters.Add("@12", MySqlDbType.Text).Value = .Rows(i).Cells(12).Value.ToString()
                        cmdlocal.Parameters.Add("@13", MySqlDbType.VarChar).Value = ClientGuid
                        cmdlocal.Parameters.Add("@14", MySqlDbType.Int64).Value = ClientStoreID
                        cmdlocal.Parameters.Add("@15", MySqlDbType.VarChar).Value = "0"
                        cmdlocal.Parameters.Add("@16", MySqlDbType.VarChar).Value = "Synced"
                        cmdlocal.Parameters.Add("@17", MySqlDbType.Text).Value = .Rows(i).Cells(13).Value.ToString()
                        cmdlocal.Parameters.Add("@18", MySqlDbType.Int64).Value = .Rows(i).Cells(14).Value.ToString()
                        cmdlocal.ExecuteNonQuery()
                        If WorkerUpdateCancel = False Then
                            CheckingForUpdates.Instance.Invoke(Sub()
                                                                   CheckingForUpdates.ProgressBar1.Value += 1
                                                                   CheckingForUpdates.Label1.Text = CheckingForUpdates.ProgressBar1.Value
                                                               End Sub)
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InstallUpdatesPriceChange()
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim ConnectionServer As MySqlConnection = ServerCloudCon()
            Dim CmdCheck As MySqlCommand
            For i As Integer = 0 To PriceChangeDatatabe.Rows.Count - 1 Step +1
                Dim sql = "UPDATE loc_admin_products SET product_price = " & PriceChangeDatatabe(i)(4) & ", price_change = 1 WHERE server_product_id = " & PriceChangeDatatabe(i)(3) & ""
                CmdCheck = New MySqlCommand(sql, ConnectionLocal)
                CmdCheck.ExecuteNonQuery()
                Dim sql2 = "UPDATE loc_price_request_change SET active = 2 WHERE request_id = " & PriceChangeDatatabe(i)(0) & ""
                CmdCheck = New MySqlCommand(sql2, ConnectionLocal)
                CmdCheck.ExecuteNonQuery()
                Dim sq3 = "UPDATE admin_price_request SET synced = 'Synced' WHERE request_id = " & PriceChangeDatatabe(i)(0) & ""
                CmdCheck = New MySqlCommand(sq3, ConnectionServer)
                CmdCheck.ExecuteNonQuery()
                GLOBAL_SYSTEM_LOGS("UPDATE PRICE CHANGE", "Store ID: " & ClientStoreID & ", Product ID: " & PriceChangeDatatabe(i)(3))
            Next
            ConnectionLocal.Close()
            ConnectionServer.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InstallCoupons()
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim ConnectionServer As MySqlConnection = ServerCloudCon()
            Dim CmdCheck As MySqlCommand
            For i As Integer = 0 To CouponDatatable.Rows.Count - 1 Step +1
                Dim sql = "UPDATE tbcoupon SET active = 1 WHERE ID = " & CouponDatatable(i)(0) & ""
                CmdCheck = New MySqlCommand(sql, ConnectionLocal)
                CmdCheck.ExecuteNonQuery()
                Dim sql2 = "UPDATE admin_custom_coupon SET synced = 'Synced' WHERE ID = " & CouponDatatable(i)(0) & ""
                CmdCheck = New MySqlCommand(sql2, ConnectionServer)
                CmdCheck.ExecuteNonQuery()
                GLOBAL_SYSTEM_LOGS("UPDATE CUSTOM COUPON", "Store ID: " & ClientStoreID & ", Coupon ID: " & CouponDatatable(i)(0))
            Next
            ConnectionLocal.Close()
            ConnectionServer.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub InstallProducts()
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim ConnectionServer As MySqlConnection = ServerCloudCon()
            Dim CmdCheck As MySqlCommand
            For i As Integer = 0 To CustomProductsApproval.Rows.Count - 1 Step +1
                Dim sql = "UPDATE loc_admin_products SET product_status = 1 WHERE product_id = " & CustomProductsApproval(i)(0) & ""
                CmdCheck = New MySqlCommand(sql, ConnectionLocal)
                CmdCheck.ExecuteNonQuery()

                Dim sql1 = "UPDATE loc_product_formula SET status = 1 WHERE formula_id = " & CustomProductsApproval(i)(1) & ""
                CmdCheck = New MySqlCommand(sql1, ConnectionLocal)
                CmdCheck.ExecuteNonQuery()

                Dim sql0 = "UPDATE loc_pos_inventory SET stock_status = 1 WHERE formula_id = " & CustomProductsApproval(i)(1) & ""
                CmdCheck = New MySqlCommand(sql0, ConnectionLocal)
                CmdCheck.ExecuteNonQuery()

                Dim sql2 = "UPDATE loc_product_list SET synced = 'Synced' WHERE loc_product_id = " & CustomProductsApproval(i)(0) & ""
                CmdCheck = New MySqlCommand(sql2, ConnectionServer)
                CmdCheck.ExecuteNonQuery()
                GLOBAL_SYSTEM_LOGS("UPDATE CUSTOM PRODUCTS", "Store ID: " & ClientStoreID & ", Product ID: " & CustomProductsApproval(i)(0))
            Next
            ConnectionLocal.Close()
            ConnectionServer.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
#End Region
#Region "Message"
    Public DtMessage As DataTable


    Private Sub PromptMessage()
        Dim LocMessageDatatable As DataTable

        DtMessage = New DataTable
        DtMessage.Columns.Add("message_id")
        DtMessage.Columns.Add("from")
        DtMessage.Columns.Add("subject")
        DtMessage.Columns.Add("content")
        DtMessage.Columns.Add("guid")
        DtMessage.Columns.Add("store_id")
        DtMessage.Columns.Add("active")
        DtMessage.Columns.Add("created_at")
        DtMessage.Columns.Add("origin")

        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim ConnectionCloud As MySqlConnection = ServerCloudCon()
            Dim Query = "SELECT server_message_id FROM loc_message"
            Dim Command As MySqlCommand = New MySqlCommand(Query, ConnectionLocal)
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(Command)
            LocMessageDatatable = New DataTable
            da.Fill(LocMessageDatatable)
            If LocMessageDatatable.Rows.Count > 0 Then
                Dim MessageIDS = ""
                For i As Integer = 0 To LocMessageDatatable.Rows.Count - 1 Step +1
                    MessageIDS += LocMessageDatatable(i)(0).ToString & ","
                Next
                MessageIDS = MessageIDS.TrimEnd(CChar(","))
                Query = "SELECT * FROM admin_message WHERE message_id NOT IN (" & MessageIDS & ") AND guid = '" & ClientGuid & "' "
                Command = New MySqlCommand(Query, ServerCloudCon)
                da = New MySqlDataAdapter(Command)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                For i As Integer = 0 To dt.Rows.Count - 1 Step +1
                    Dim Mess As DataRow = DtMessage.NewRow
                    Mess("message_id") = dt(i)(0)
                    Mess("from") = dt(i)(1)
                    Mess("subject") = dt(i)(5)
                    Mess("content") = dt(i)(6)
                    Mess("guid") = dt(i)(7)
                    Mess("store_id") = ClientStoreID
                    Mess("active") = 1
                    Mess("created_at") = dt(i)(8)
                    Mess("origin") = "Server"
                    DtMessage.Rows.Add(Mess)
                Next
            Else
                Query = "SELECT * FROM admin_message WHERE guid = '" & ClientGuid & "' "
                Command = New MySqlCommand(Query, ServerCloudCon)
                da = New MySqlDataAdapter(Command)
                Dim dt As DataTable = New DataTable
                da.Fill(dt)
                For i As Integer = 0 To dt.Rows.Count - 1 Step +1
                    Dim Mess As DataRow = DtMessage.NewRow

                    Mess("message_id") = dt(i)(0)
                    Mess("from") = dt(i)(1)
                    Mess("subject") = dt(i)(5)
                    Mess("content") = dt(i)(6)
                    Mess("guid") = dt(i)(7)
                    Mess("store_id") = ClientStoreID
                    Mess("active") = 1
                    Mess("created_at") = dt(i)(8)
                    Mess("origin") = "Server"
                    DtMessage.Rows.Add(Mess)

                Next
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Private Sub TextBoxGRANDTOTAL_TextChanged(sender As Object, e As EventArgs) Handles TextBoxGRANDTOTAL.TextChanged
        Try
            If My.Settings.LedDisplayTrue Then
                LedDisplay(TextBoxGRANDTOTAL.Text, True)
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = "Show" Then
            SynctoCloud.TopMost = True
            SynctoCloud.WindowState = FormWindowState.Normal
            Button2.Text = "Hide"
        Else
            SynctoCloud.TopMost = False
            SynctoCloud.WindowState = FormWindowState.Minimized
            Button2.Text = "Show"
        End If
    End Sub

    Private Sub LabelTransactionType_TextChanged(sender As Object, e As EventArgs) Handles LabelTransactionType.TextChanged
        Try
            LabelTransactionType.Text = LabelTransactionType.Text.ToUpper
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub




#End Region
#End Region
End Class


