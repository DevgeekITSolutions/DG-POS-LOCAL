Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Text
Module RetrieveModule
    Dim user_id As String
    Dim acroo As String
    Dim fullname As String
    Dim Location_control As New Point(10, 10)
    Public Count_control As Integer = 0
    Dim result As Integer
    Dim dr
    Dim municipality
    Dim province
    Dim full_name
    Dim productcode As String
    Dim password As String = Login.txtpassword.Text
    Dim wrapper As New Simple3Des(password)
    Dim returnval
    Dim RowsReturned As Integer
    Dim product_line
    Dim available_stock
    Dim dailysales
    Dim critical_item
    Dim product
    Dim cipherText As String

    Public Sub retrieveLoginDetails()
        Try
            If Login.txtusername.Text = "" Then
                MessageBox.Show("Input username first", "Login Form", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Login.txtusername.Focus()
            ElseIf Login.txtpassword.Text = "" Then
                MessageBox.Show("Input password first", "Login Form", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Login.txtpassword.Focus()
            Else
                Try
                    cipherText = ConvertPassword(SourceString:=Login.txtpassword.Text)
                    sql = "SELECT * FROM loc_users WHERE username = @Username AND password = @Password AND active = 1;"
                    cmd = New MySqlCommand(sql, LocalhostConn())
                    With cmd
                        .Parameters.Clear()
                        .Parameters.AddWithValue("@Username", Login.txtusername.Text)
                        .Parameters.AddWithValue("@UserID", Login.txtusername.Text)
                        .Parameters.AddWithValue("@Password", cipherText)
                        .Parameters.AddWithValue("@StoreID", ClientStoreID)
                        Dim reader As MySqlDataReader
                        reader = .ExecuteReader()
                        While reader.Read()
                            user_id = reader("uniq_id")
                        End While
                        reader.Close()
                    End With
                    da = New MySqlDataAdapter
                    dt = New DataTable
                    da.SelectCommand = cmd
                    da.Fill(dt)
                Catch ex As MySqlException
                    MsgBox(ex.ToString)
                    SendErrorReport(ex.ToString)
                Finally
                    da.Dispose()
                    If dt.Rows.Count > 0 Then
                        Dim crew_id, username, password, fullname, userlevel, active, storeid, franguid, role As String
                        crew_id = dt.Rows(0).Item(0)
                        role = dt.Rows(0).Item(1)
                        username = dt.Rows(0).Item(3)
                        password = dt.Rows(0).Item(4)
                        userlevel = dt.Rows(0).Item(7)
                        fullname = dt.Rows(0).Item(2)
                        active = dt.Rows(0).Item(11)
                        franguid = dt.Rows(0).Item(12)
                        storeid = dt.Rows(0).Item(13)
                        ClientRole = role
                        If Login.txtusername.Text = username And cipherText = password And userlevel = "Crew" And ClientStoreID = storeid And active = 1 And franguid = ClientGuid Then
                            MessageBox.Show("Welcome " + fullname + "!", "Login Successfully", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Login.txtusername.Text = ""
                            Login.txtpassword.Text = ""
                            ClientCrewID = user_id
                            messageboxappearance = True
                            SystemLogType = "LOGIN"
                            SystemLogDesc = "User Login: " & username & " : " & ClientRole
                            GLOBAL_SYSTEM_LOGS(SystemLogType, SystemLogDesc)
                            Shift = ""
                            If S_Layout = "POS" Then
                                Login.Close()
                                POS.Show()
                            ElseIf S_Layout = "GROCERY" Then
                                Login.Close()
                                Grocery.Show()
                            End If

                        ElseIf Login.txtusername.Text = username And cipherText = password And userlevel = "Head Crew" And ClientStoreID = storeid And active = 1 And franguid = ClientGuid Then
                            MessageBox.Show("Welcome " + fullname + "!", "Login Successfully(" & ClientRole & ")", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Login.txtusername.Text = ""
                            Login.txtpassword.Text = ""
                            ClientCrewID = user_id
                            messageboxappearance = True
                            SystemLogType = "LOGIN"
                            SystemLogDesc = "User Login: " & username & " : " & ClientRole
                            GLOBAL_SYSTEM_LOGS(SystemLogType, SystemLogDesc)
                            Shift = ""
                            If S_Layout = "POS" Then
                                Login.Close()
                                POS.Show()
                            ElseIf S_Layout = "GROCERY" Then
                                Login.Close()
                                Grocery.Show()
                            End If

                        ElseIf Login.txtusername.Text = username And cipherText = password And userlevel = "Manager" And ClientStoreID = storeid And active = 1 And franguid = ClientGuid Then
                            MessageBox.Show("Welcome " + fullname + "!", "Login Successfully(" & ClientRole & ")", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Login.txtusername.Text = ""
                            Login.txtpassword.Text = ""
                            ClientCrewID = user_id
                            messageboxappearance = True
                            SystemLogType = "LOGIN"
                            SystemLogDesc = "User Login: " & username & " : " & ClientRole
                            GLOBAL_SYSTEM_LOGS(SystemLogType, SystemLogDesc)
                            Shift = ""
                            If S_Layout = "POS" Then
                                Login.Close()
                                POS.Show()
                            ElseIf S_Layout = "GROCERY" Then
                                Login.Close()
                                Grocery.Show()
                            End If
                        ElseIf Login.txtusername.Text = username And cipherText = password And userlevel = "Admin" Then
                            MessageBox.Show("Welcome " + fullname + "!", "Login Successfully(" & ClientRole & ")", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Login.txtusername.Text = ""
                            Login.txtpassword.Text = ""
                            ClientCrewID = user_id
                            messageboxappearance = True
                            SystemLogType = "LOGIN"
                            SystemLogDesc = "User Login: " & username & " : " & ClientRole
                            GLOBAL_SYSTEM_LOGS(SystemLogType, SystemLogDesc)
                            Shift = ""
                            If S_Layout = "POS" Then
                                Login.Close()
                                POS.Show()
                            ElseIf S_Layout = "GROCERY" Then
                                Login.Close()
                                Grocery.Show()
                            End If
                        Else
                            MessageBox.Show("Incorrect username or password!", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            SystemLogType = "ERROR"
                            SystemLogDesc = "FAILED TO LOGIN: Username: " & Login.txtusername.Text & " Password: " & Login.txtpassword.Text
                            GLOBAL_SYSTEM_LOGS(SystemLogType, SystemLogDesc)
                        End If
                    Else
                        MessageBox.Show("Incorrect username or password!", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Login.txtusername.Focus()
                        messageboxappearance = True
                        SystemLogType = "ERROR"
                        SystemLogDesc = "FAILED TO LOGIN: Username and password input " & Login.txtusername.Text & " " & Login.txtpassword.Text
                        GLOBAL_SYSTEM_LOGS(SystemLogType, SystemLogDesc)
                    End If
                End Try
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Function CheckUserName(Username) As Boolean
        Dim ReturnUsername As Boolean = False
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql = "SELECT username FROM loc_users WHERE username = '" & Username & "'"
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Using reader As MySqlDataReader = cmd.ExecuteReader()
                If reader.HasRows Then
                    ReturnUsername = True
                Else
                    ReturnUsername = False
                End If
            End Using
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return ReturnUsername
    End Function
    Public Function CheckEmail(Email) As Boolean
        Dim ReturnUsername As Boolean = False
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql = "SELECT email FROM loc_users WHERE email = '" & Email & "'"
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Using reader As MySqlDataReader = cmd.ExecuteReader()
                If reader.HasRows Then
                    ReturnUsername = True
                Else
                    ReturnUsername = False
                End If
            End Using
            ConnectionLocal.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return ReturnUsername
    End Function
    Public Function CheckContactNumber(ContactNumber) As Boolean
        Dim ConnectionLocal As MySqlConnection = LocalhostConn()
        Dim ReturnUsername As Boolean = False
        Try
            Dim sql = "SELECT contact_number FROM loc_users WHERE contact_number = '" & ContactNumber & "'"
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Using reader As MySqlDataReader = cmd.ExecuteReader()
                If reader.HasRows Then
                    ReturnUsername = True
                Else
                    ReturnUsername = False
                End If
            End Using
            ConnectionLocal.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return ReturnUsername
    End Function
    Dim r As Random = New Random(Guid.NewGuid().GetHashCode())
    Public Function CheckUserId() As String
        Dim Uniqid = ""
        Try
            Dim ReturnThis As Boolean = True
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Do
                Uniqid = ClientStorename & "-" & r.[Next](1000, 10000)
                Dim sql = "SELECT uniq_id FROM loc_users WHERE uniq_id = '" & Uniqid & "'"
                Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
                Using reader As MySqlDataReader = cmd.ExecuteReader()
                    If reader.HasRows Then
                        ReturnThis = True
                    Else
                        ReturnThis = False
                    End If
                End Using
            Loop Until (ReturnThis = False)
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return Uniqid
    End Function
    'FUNCTION LOADING EXPENCES / POS ==================================================================================== 
    Public Sub listviewproductsshow(ByVal where As String)
        Try
            Dim Query As String = ""
            Dim Connectionlocal As MySqlConnection = LocalhostConn()
            If where = "Others" Then
                Query = "SELECT product_id, product_name, product_image, product_price, formula_id, product_sku FROM loc_admin_products WHERE product_category ='" & where & "' AND product_status = 1 AND store_id = " & ClientStoreID
            Else
                Query = "SELECT product_id, product_name, product_image, product_price, formula_id, product_sku FROM loc_admin_products WHERE product_category ='" & where & "' AND product_status = 1 "
            End If
            With POS
                .PanelProducts.Controls.Clear()
                Dim cmd As MySqlCommand = New MySqlCommand(Query, Connectionlocal)
                Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
                Dim dt As DataTable = New DataTable()
                da.Fill(dt)
                For Each row As DataRow In dt.Rows
                    Count_control += 1
                    Dim new_Button_product As New Button
                    Dim buttonname As String = row("product_sku")
                    Dim newlabel As New Label
                    productprice = row("product_price")
                    productID = row("product_id")
                    With new_Button_product
                        .Name = buttonname
                        .Text = productprice
                        .TextImageRelation = TextImageRelation.ImageBeforeText
                        .TextAlign = ContentAlignment.TopLeft
                        If where = "Premium" Then
                            .ForeColor = Color.White
                        Else
                            .ForeColor = Color.Black
                        End If
                        .Font = New Font("Tahoma", 10)
                        .BackgroundImage = Base64ToImage(row("product_image"))
                        .FlatStyle = FlatStyle.Flat
                        .FlatAppearance.BorderSize = 0
                        .BackgroundImageLayout = ImageLayout.Stretch
                        .Location = New Point(Location_control.X, Location_control.Y)
                        .Width = 148
                        .Height = 120
                        .Cursor = Cursors.Hand
                        With newlabel
                            .Text = buttonname
                            .Font = New Font("Tahoma", 10)
                            .ForeColor = Color.White
                            .Width = 148
                            .Location = New Point(0, 100)
                            .BackColor = Color.SlateGray
                            .Parent = new_Button_product
                            .TextAlign = ContentAlignment.TopCenter
                        End With
                        .Controls.Add(newlabel)
                    End With
                    .PanelProducts.Controls.Add(new_Button_product)
                    AddHandler new_Button_product.Click, AddressOf new_product_button_click
                Next
                Connectionlocal.Close()
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    'MANAGE PRODUCTS PANEL ========================================================
    Public Sub selectmax(ByVal whatform As Integer)
        If whatform = 1 Then
            POS.TextBoxMAXID.Text = Format(Now, "yyddMMHHmmssyy")
        ElseIf whatform = 2 Then
            Addexpense.TextBoxMAXID.Text = Format(Now, "yydd-MMHH-mmssyy")
        ElseIf whatform = 3 Then
            Registration.TextBoxMAXID.Text = Format(Now, "yydd-MMHH-mmssyy")
        End If
    End Sub
    Dim formulaid
    Public Function selectmaxformula(ByVal whatid As String, ByVal fromtable As String, ByVal flds As String)
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql = "Select " & flds & " FROM " & fromtable & " ORDER BY " & whatid & " DESC LIMIT 1"
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            formulaid = cmd.ExecuteScalar()
            cmd.Dispose()
            ConnectionLocal.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return formulaid
    End Function
    Public Function returnfullname(ByVal where As String)
        Dim FullName As String = ""
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim cmd As MySqlCommand = New MySqlCommand("SELECT full_name FROM loc_users WHERE uniq_id = '" + where + "' ", ConnectionLocal)
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            FullName = dt(0)(0).ToString
            LocalhostConn.Close()
            da.Dispose()
            cmd.Dispose()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return FullName
    End Function
    Dim valuetoreturn
    Dim MyLocalConnection As MySqlConnection
    Dim MyCloudConnection As MySqlConnection
    Dim MyCmd As MySqlCommand
    Public Function GLOBAL_RETURN_FUNCTION(tbl As String, flds As String, toreturn As String, thisislocalconn As Boolean)
        Try
            Dim sql = "SELECT " & flds & " FROM " & tbl
            If thisislocalconn = True Then
                MyLocalConnection = New MySqlConnection
                MyLocalConnection.ConnectionString = LocalConnectionString
                MyLocalConnection.Open()
                MyCmd = New MySqlCommand(sql, MyLocalConnection)
            Else
                MyCloudConnection = New MySqlConnection
                MyCloudConnection.ConnectionString = CloudConnectionString
                MyCloudConnection.Open()
                MyCmd = New MySqlCommand(sql, MyCloudConnection)
            End If
            Using reader As MySqlDataReader = MyCmd.ExecuteReader()
                If reader.HasRows Then
                    While reader.Read
                        valuetoreturn = reader(toreturn)
                    End While
                End If
            End Using
            If thisislocalconn = True Then
                MyLocalConnection.Close()
            Else
                MyCloudConnection.Close()
            End If
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return valuetoreturn
    End Function
    Public Function AsDatatable(table, fields, datagridd) As DataTable
        datagridd.rows.clear
        Dim dttable As DataTable = New DataTable
        Dim ConnectionLocal As MySqlConnection = LocalhostConn()
        Try
            Dim sql = "SELECT " & fields & " FROM " & table
            'Console.WriteLine(sql)
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            da.Fill(dttable)
            With datagridd
                .RowHeadersVisible = False
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToOrderColumns = False
                .AllowUserToResizeColumns = False
                .AllowUserToResizeRows = False
                .Font = New Font("tahoma", 10)
                .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
                .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        Finally
            ConnectionLocal.Close()
            cmd.Dispose()
            da.Dispose()
        End Try
        Return dttable
    End Function
    Public Function AsDatatableFontIncrease(table, fields, datagridd) As DataTable
        datagridd.rows.clear
        Dim dttable As DataTable = New DataTable
        Dim ConnectionLocal As MySqlConnection = LocalhostConn()
        Try
            Dim sql = "SELECT " & fields & " FROM " & table
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            da.Fill(dttable)
            With datagridd
                .RowHeadersVisible = False
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToOrderColumns = False
                .AllowUserToResizeColumns = False
                .AllowUserToResizeRows = False
                .Font = New Font("tahoma", 12)
                .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
                .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .rowtemplate.height = 30
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        Finally
            ConnectionLocal.Close()
            cmd.Dispose()
            da.Dispose()
        End Try
        Return dttable
    End Function
    Public Sub GLOBAL_SELECT_ALL_FUNCTION(ByVal table As String, ByVal fields As String, ByRef datagrid As DataGridView)
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql As String = "SELECT " + fields + " FROM " + table
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            With datagrid
                .DataSource = Nothing
                .DataSource = dt
                .RowHeadersVisible = False
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToOrderColumns = False
                .AllowUserToResizeColumns = False
                .AllowUserToResizeRows = False
                .Font = New Font("tahoma", 10)
                .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
                .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            End With
            ConnectionLocal.Close()
            cmd.Dispose()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub GLOBAL_SELECT_ALL_FUNCTION_WHERE(ByVal table As String, ByVal fields As String, ByVal where As String, ByVal successmessage As String, ByVal errormessage As String, ByRef datagrid As DataGridView)
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql As String = "SELECT " + fields + " FROM " + table + " WHERE " + where
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            With datagrid
                .DataSource = Nothing
                .DataSource = dt
                .RowHeadersVisible = False
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToOrderColumns = False
                .AllowUserToResizeColumns = False
                .AllowUserToResizeRows = False
                .Font = New Font("tahoma", 10)
                .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
                .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            End With
            ConnectionLocal.Close()
            cmd.Dispose()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub GLOBAL_SELECT_ALL_FUNCTION_COMBOBOX(table As String, fields As String, combobox As ComboBox, Loccon As Boolean)
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim ConnectionCloud As MySqlConnection = ServerCloudCon()
            Dim sql = "SELECT " + fields + " FROM " + table
            Dim cmd As MySqlCommand
            If Loccon = True Then
                cmd = New MySqlCommand(sql, ConnectionLocal)
            Else
                cmd = New MySqlCommand(sql, ConnectionCloud)
            End If
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            With combobox
                .DataSource = Nothing
                .DataSource = dt
                .ValueMember = fields
                .DisplayMember = fields
            End With
            ConnectionLocal.Close()
            ConnectionCloud.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        Finally
            da.Dispose()
        End Try
    End Sub
    Public Function GLOBAL_SELECT_FUNCTION_RETURN(ByVal table As String, ByVal fields As String, ByVal values As String, ByVal returnvalrow As String)
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql = "SELECT " + fields + " FROM " + table + " WHERE " + values
            Console.WriteLine(sql)
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Using readerObj As MySqlDataReader = cmd.ExecuteReader
                While readerObj.Read
                    returnval = readerObj(returnvalrow).ToString
                End While
            End Using
            cmd.Dispose()
            ConnectionLocal.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return returnval
    End Function
    '=================================================================MDI FORM
    Dim returncount As String
    Public Function count(ByVal tocount As String, ByVal table As String)
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql As String = "SELECT COUNT(" & tocount & ") FROM " & table
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            For Each row As DataRow In dt.Rows
                returncount = row("COUNT(" & tocount & ")")
            Next
            cmd.Dispose()
            ConnectionLocal.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return returncount
    End Function
    Dim returnsum
    Public Function roundsum(tototal As String, table As String, Columncall As String)
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql As String = "SELECT SUM(ROUND(" & tototal & ",0)) AS " & Columncall & " FROM " & table
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            If IsDBNull(dt.Rows(0)(0)) Then
                returnsum = 0
            Else
                For Each row As DataRow In dt.Rows
                    returnsum = row(Columncall)
                Next
            End If
            cmd.Dispose()
            ConnectionLocal.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return returnsum
    End Function
    Public Function sum(ByVal tototal As String, ByVal table As String)
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql As String = "SELECT SUM(" & tototal & ") FROM " & table
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Dim da As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            If IsDBNull(dt.Rows(0)(0)) Then
                returnsum = 0
            Else
                For Each row As DataRow In dt.Rows
                    returnsum = row("SUM(" & tototal & ")")
                Next
            End If
            cmd.Dispose()
            ConnectionLocal.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try

        Return returnsum
    End Function

    Public Function returnselect(toreturn As String, table As String)
        Dim RetunSel As String = ""
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim sql As String = "SELECT `" & toreturn & "` as RETURNSEL FROM " & table
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ConnectionLocal)
            Using reader As MySqlDataReader = cmd.ExecuteReader
                If reader.HasRows Then
                    While reader.Read
                        RetunSel = reader("RETURNSEL")
                    End While
                End If
            End Using
            cmd.Dispose()
            ConnectionLocal.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return RetunSel
    End Function

    Public Function ReturnMunicipalityName(id) As String
        Dim ServerConn As MySqlConnection = ServerCloudCon()
        Dim ReturnMun As String = ""
        Try
            Dim sql = "SELECT mn_name FROM admin_municipality WHERE mn_id = " & id & ""
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ServerConn)
            Using reader As MySqlDataReader = cmd.ExecuteReader()
                If reader.HasRows Then
                    While reader.Read
                        ReturnMun = reader("mn_name")
                    End While
                End If
            End Using
            cmd.Dispose()
            ServerConn.Close()

        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return ReturnMun
    End Function
    Public Function ReturnProvinceName(id) As String
        Dim ServerConn As MySqlConnection = ServerCloudCon()
        Dim ReturnProv As String = ""
        Try
            Dim sql = "SELECT province FROM admin_province WHERE add_id = " & id & ""
            Dim cmd As MySqlCommand = New MySqlCommand(sql, ServerConn)
            Using reader As MySqlDataReader = cmd.ExecuteReader()
                If reader.HasRows Then
                    While reader.Read
                        ReturnProv = reader("province")
                    End While
                End If
            End Using
            ServerConn.Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
        Return ReturnProv
    End Function
    Public Sub GetOldGrandtotal()
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim Query = "SELECT S_Old_Grand_Total FROM loc_settings WHERE settings_id = 1"
            Dim Cmd As MySqlCommand = New MySqlCommand(Query, ConnectionLocal)
            Using reader As MySqlDataReader = Cmd.ExecuteReader
                If reader.HasRows Then
                    While reader.Read
                        'If reader("S_Old_Grand_Total") < 1 Then
                        '    S_OLDGRANDTOTAL = sum("amountdue", "loc_daily_transaction WHERE zreading <> '" & S_Zreading & "'")
                        'Else
                        Dim CurrentZreadTotal As Double = sum("amountdue", "loc_daily_transaction WHERE zreading = '" & S_Zreading & "'")
                            S_OLDGRANDTOTAL = reader("S_Old_Grand_Total") + CurrentZreadTotal
                        'End If
                    End While
                End If
            End Using
            Query = "UPDATE loc_settings SET S_Old_Grand_Total = " & S_OLDGRANDTOTAL & " WHERE settings_id = 1"
            Cmd = New MySqlCommand(Query, ConnectionLocal)
            Cmd.ExecuteReader()
            ConnectionLocal.Close()
            Cmd.Dispose()
            'Console.WriteLine(S_OLDGRANDTOTAL)
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub LoadOldGrandtotal()
        Try
            Dim ConnectionLocal As MySqlConnection = LocalhostConn()
            Dim Query = "SELECT S_Old_Grand_Total FROM loc_settings WHERE settings_id = 1"
            Dim Cmd As MySqlCommand = New MySqlCommand(Query, ConnectionLocal)
            Using reader As MySqlDataReader = Cmd.ExecuteReader
                If reader.HasRows Then
                    While reader.Read
                        S_OLDGRANDTOTAL = reader("S_Old_Grand_Total")
                    End While
                End If
            End Using
            ConnectionLocal.Close()
            Cmd.Dispose()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
End Module