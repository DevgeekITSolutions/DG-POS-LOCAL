Imports MySql.Data.MySqlClient
Imports System.Threading
Public Class CouponCode
    Dim ThreadList As List(Of Thread) = New List(Of Thread)
    Dim Thread As Thread

    Private Sub CouponCode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            CheckForIllegalCrossThreadCalls = False
            BackgroundWorker1.WorkerReportsProgress = True
            BackgroundWorker1.WorkerSupportsCancellation = True
            BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Thread = New Thread(AddressOf LoadCoupons)
            Thread.Start()
            ThreadList.Add(Thread)
            For Each t In ThreadList
                t.Join()
                If (BackgroundWorker1.CancellationPending) Then
                    ' Indicate that the task was canceled.
                    e.Cancel = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub LoadCoupons(sender As Object)
        Try

            Dim LoadCouponTable = AsDatatableFontIncrease("tbcoupon WHERE active = 1 ORDER BY Type ASC", "*", DataGridViewCoupons)
            For Each row As DataRow In LoadCouponTable.Rows
                DataGridViewCoupons.DefaultCellStyle.Font = New Font("Tahoma", 15)
                DataGridViewCoupons.Height = 40
                DataGridViewCoupons.Rows.Add(row("ID"), row("Couponname_"), row("Desc_"), row("Discountvalue_"), row("Referencevalue_"), row("Type"), row("Bundlebase_"), row("BBValue_"), row("Bundlepromo_"), row("BPValue_"), row("Effectivedate"), row("Expirydate"))
            Next
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub CouponCode_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        POS.Enabled = True
    End Sub
    Private Sub CouponDefault()
        Try
            CouponApplied = False
            CouponName = ""
            CouponDesc = ""
            CouponLine = 10
            CouponTotal = 0
            With POS
                For i As Integer = 0 To .DataGridViewOrders.Rows.Count - 1 Step +1
                    If .DataGridViewOrders.Rows(i).Cells(11).Value > 0 Then
                        Dim priceadd = .DataGridViewOrders.Rows(i).Cells(11).Value * S_Upgrade_Price
                        If S_ZeroRated = "0" Then
                            .DataGridViewOrders.Rows(i).Cells(3).Value = .DataGridViewOrders.Rows(i).Cells(1).Value * .DataGridViewOrders.Rows(i).Cells(2).Value + priceadd
                        Else
                            Dim TotalPrice As Double = 0
                            Dim Tax = 1 + Val(S_Tax)
                            Dim Total = Math.Round(.DataGridViewOrders.Rows(i).Cells(1).Value * .DataGridViewOrders.Rows(i).Cells(2).Value + priceadd / Tax, 2, MidpointRounding.AwayFromZero)
                            .DataGridViewOrders.Rows(i).Cells(3).Value = Total
                        End If
                    Else
                        If S_ZeroRated = "0" Then
                            .DataGridViewOrders.Rows(i).Cells(3).Value = .DataGridViewOrders.Rows(i).Cells(1).Value * .DataGridViewOrders.Rows(i).Cells(2).Value
                        Else
                            Dim TotalPrice As Double = 0
                            Dim Tax = 1 + Val(S_Tax)
                            Dim Total = Math.Round(.DataGridViewOrders.Rows(i).Cells(1).Value * .DataGridViewOrders.Rows(i).Cells(2).Value / Tax, 2, MidpointRounding.AwayFromZero)
                            .DataGridViewOrders.Rows(i).Cells(3).Value = Total
                        End If
                    End If
                Next
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub ButtonSubmit_Click(sender As Object, e As EventArgs) Handles ButtonSubmit.Click
        Try


            Dim CountItem As Integer = 0
            With POS
                For i As Integer = 0 To .DataGridViewOrders.Rows.Count - 1 Step +1
                    CountItem += .DataGridViewOrders.Rows(i).Cells(1).Value
                Next
            End With
            If CountItem < 1 Then
                MsgBox("Cannot apply coupon! Minimum product quantity is 1", vbInformation)
                Exit Sub
            ElseIf Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString = "Percentage(w/o vat)" Then
                Enabled = False
                SeniorDetails.Show()
            ElseIf Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString = "Percentage(w/ vat)" Then
                coupondiscountpercentagewvat()
            ElseIf Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString = "Fix-1" Then
                ' MsgBox("Coupon is " & Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value)
                couponfix1()
            ElseIf Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString = "Fix-2" Then
                '  MsgBox("Coupon is " & Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value)
                couponfix2()
            ElseIf Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString = "Bundle-1(Fix)" Then
                ' MsgBox("Coupon is " & Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value)
                couponbundle1()
            ElseIf Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString = "Bundle-2(Fix)" Then
                '   MsgBox("Coupon is " & Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value)
                couponbundle2()
            ElseIf Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString = "Bundle-3(%)" Then
                '  MsgBox("Coupon is " & Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value)
                couponbundle3()
            ElseIf Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString = "Bundle-4(Fix)" Then
                '  MsgBox("Coupon is " & Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value)
                couponbundle4()
            End If
            'Compute()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub

    Public Sub couponpercentage()
        Try
            If SENIORDETAILSBOOL = True Then
                If S_ZeroRated = "0" Then
                    CouponDefault()
                    With POS.DataGridViewOrders

                        Dim AmountDueWaffle As Double = 0
                        Dim AmountDueDrinks As Double = 0
                        Dim SeniorPwdDisk = DataGridViewCoupons.SelectedRows(0).Cells(3).Value / 100
                        Dim Tax = 1 + Val(S_Tax)



                        Dim DISCOUNTAMOUNT As Double = SeniorPWd + SeniorPWdDrinks

                        MsgBox(DISCOUNTAMOUNT)

                        Dim VATEXEMPTSALES As Double = Format(DISCOUNTAMOUNT / Tax, "0.00")
                        'MsgBox("VAT EXEMPT SALES : " & VATEXEMPTSALES)
                        Dim NOTDISCOUNTEDAMOUNT As Double = Double.Parse(POS.TextBoxGRANDTOTAL.Text) - DISCOUNTAMOUNT
                        Dim VATABLESALES As Double = Format(NOTDISCOUNTEDAMOUNT / Tax, "0.00")
                        'MsgBox("VATABLESALES : " & VATABLESALES)
                        Dim ZERORATEDSALES As Double = Format(0, "0.00")
                        'MsgBox("ZERO RATED SALES : " & ZERORATEDSALES)
                        Dim LESSVAT As Double = Format(DISCOUNTAMOUNT - VATEXEMPTSALES, "0.00")
                        'MsgBox("LESS VAT : " & LESSVAT)
                        Dim VAT12PERCENT As Double = Format(NOTDISCOUNTEDAMOUNT - VATABLESALES, "0.00")
                        'MsgBox("VAT 12 : " & VAT12PERCENT)
                        Dim TOTALDISCOUNT As Double = Format(VATEXEMPTSALES * SeniorPwdDisk, "0.00")
                        'MsgBox("TOTAL DISCOUNT : " & TOTALDISCOUNT)
                        Dim TOTALDISCPLUSLESSVAT As Double = TOTALDISCOUNT + LESSVAT
                        Dim TOTALAMOUNTDUE As Double = Double.Parse(POS.TextBoxGRANDTOTAL.Text) - TOTALDISCPLUSLESSVAT
                        'MsgBox("TOTAL AMOUNT DUE : " & TOTALAMOUNTDUE)

                        For i As Integer = 0 To .Rows.Count - 1 Step +1

                            If .Rows(i).Cells(9).Value.ToString = "WAFFLE" Then
                                AmountDueWaffle += .Rows(i).Cells(3).Value
                            Else
                                AmountDueDrinks += .Rows(i).Cells(3).Value
                            End If
                        Next

                        With POS
                            '.GROSSSALE = GROSSSALES
                            .TOTALDISCOUNTEDAMOUNT = DISCOUNTAMOUNT
                            .VATEXEMPTSALES = VATEXEMPTSALES
                            .LESSVAT = LESSVAT
                            .TOTALDISCOUNT = TOTALDISCOUNT
                            .VATABLESALES = VATABLESALES
                            .VAT12PERCENT = VAT12PERCENT
                            .TOTALAMOUNTDUE = TOTALAMOUNTDUE
                            .TextBoxGRANDTOTAL.Text = NUMBERFORMAT(TOTALAMOUNTDUE)
                            .TextBoxDISCOUNT.Text = NUMBERFORMAT(TOTALDISCOUNT)
                        End With
                        SeniorGCDiscount = True
                        CouponTotal = TOTALDISCOUNT

                    End With
                Else
                    'CouponDefault()
                    With POS.DataGridViewOrders
                        Dim SeniorPwdDisk = DataGridViewCoupons.SelectedRows(0).Cells(3).Value / 100
                        Dim Tax = 1 + Val(S_Tax)
                        Dim GROSSSALES As Double = Double.Parse(POS.TextBoxGRANDTOTAL.Text)
                        Dim TotalCountedAmount As Double = SeniorPWd + SeniorPWdDrinks
                        Dim SENIOR20PERCENT As Double = TotalCountedAmount / Tax * SeniorPwdDisk
                        SENIOR20PERCENT = Math.Round(SENIOR20PERCENT, 2, MidpointRounding.AwayFromZero)
                        With POS
                            '.GROSSSALE = GROSSSALES
                            .VATABLESALES = 0
                            .VAT12PERCENT = 0
                            .LESSVAT = 0
                            .TOTALDISCOUNTEDAMOUNT = TotalCountedAmount
                            .VATEXEMPTSALES = 0
                            .TOTALDISCOUNT = SENIOR20PERCENT
                            .ZERORATEDSALES = .Label76.Text
                            .ZERORATEDNETSALES = .Label76.Text - SENIOR20PERCENT
                            .TextBoxGRANDTOTAL.Text = NUMBERFORMAT(.Label76.Text - SENIOR20PERCENT)
                            .TextBoxDISCOUNT.Text = SENIOR20PERCENT
                        End With
                        SeniorGCDiscount = True
                        CouponTotal = SENIOR20PERCENT
                    End With
                End If
                CouponApplied = True
                CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                MsgBox("Applied")
            Else
                SeniorGCDiscount = False
            End If
            With POS
                .DISCOUNTTYPE = Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
            End With
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Public Sub coupondiscountpercentagewvat()
        Try
            If S_ZeroRated = "0" Then
                CouponDefault()

                With POS
                    Dim GROSSSALES As Double = Double.Parse(.Label76.Text)
                    Dim TAX As Double = 1 + Val(S_Tax)
                    Dim TOTALDISCOUNT As Double = DataGridViewCoupons.SelectedRows(0).Cells(3).Value / 100
                    Dim TOTALAMOUNTDUE As Double = 0

                    TOTALDISCOUNT = GROSSSALES * TOTALDISCOUNT
                    TOTALAMOUNTDUE = Format(GROSSSALES - TOTALDISCOUNT, "0.00")

                    Dim VATABLESALES As Double = Format(GROSSSALES / TAX, "0.00")
                    Dim VAT12PERCENT As Double = Format(VATABLESALES * S_Tax, "0.00")

                    '.GROSSSALE = GROSSSALES
                    .TOTALDISCOUNTEDAMOUNT = GROSSSALES
                    .VATEXEMPTSALES = 0
                    .LESSVAT = 0
                    .TOTALDISCOUNT = TOTALDISCOUNT
                    .VATABLESALES = VATABLESALES
                    .VAT12PERCENT = VAT12PERCENT
                    .TOTALAMOUNTDUE = TOTALAMOUNTDUE
                    .TextBoxGRANDTOTAL.Text = NUMBERFORMAT(TOTALAMOUNTDUE)
                    .TextBoxDISCOUNT.Text = NUMBERFORMAT(TOTALDISCOUNT)
                    CouponDesc = ""
                    CouponTotal = TOTALDISCOUNT
                    CouponApplied = True
                    CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                    MsgBox("Applied")

                End With
            Else
                CouponDefault()
                With POS
                    Dim GROSSSALES As Double = Double.Parse(.Label76.Text)
                    Dim TAX As Double = 1 + Val(S_Tax)
                    Dim TOTALDISCOUNT As Double = DataGridViewCoupons.SelectedRows(0).Cells(3).Value / 100
                    Dim TOTALAMOUNTDUE As Double = 0

                    TOTALDISCOUNT = GROSSSALES * TOTALDISCOUNT
                    TOTALAMOUNTDUE = Format(GROSSSALES - TOTALDISCOUNT, "0.00")

                    'Dim VATABLESALES As Double = Format(GROSSSALES / TAX, "0.00")
                    'Dim VAT12PERCENT As Double = Format(VATABLESALES * S_Tax, "0.00")

                    '.GROSSSALE = GROSSSALES
                    .TOTALDISCOUNTEDAMOUNT = GROSSSALES
                    .VATEXEMPTSALES = 0
                    .LESSVAT = 0
                    .TOTALDISCOUNT = TOTALDISCOUNT
                    .VATABLESALES = .Label76.Text
                    .VAT12PERCENT = 0
                    .TOTALAMOUNTDUE = TOTALAMOUNTDUE
                    .TextBoxGRANDTOTAL.Text = NUMBERFORMAT(TOTALAMOUNTDUE)
                    .TextBoxDISCOUNT.Text = NUMBERFORMAT(TOTALDISCOUNT)

                    .ZERORATEDSALES = .Label76.Text
                    .ZERORATEDNETSALES = .Label76.Text - TOTALDISCOUNT
                    .TextBoxGRANDTOTAL.Text = NUMBERFORMAT(.Label76.Text - TOTALDISCOUNT)

                    CouponDesc = ""
                    CouponTotal = TOTALDISCOUNT
                    CouponApplied = True
                    CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                    MsgBox("Applied")
                End With
            End If
            With POS
                .DISCOUNTTYPE = Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
            End With
            Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub couponfix1()
        Try
            Dim GROSSSALES As Double = Double.Parse(POS.Label76.Text)
            Dim TOTALDISCOUNT As Double = DataGridViewCoupons.SelectedRows(0).Cells(3).Value
            Dim DISCOUNTAMOUNT As Double = 0
            Dim LESSVAT As Double = 0
            Dim VATABLE As Double = 0
            Dim TOTALAMOUNTDUE As Double = 0
            Dim Tax = 1 + Val(S_Tax)
            Dim VATABLESALES As Double = 0
            Dim VAT12PERCENT As Double = 0
            'SeniorGCDiscount = True
            With POS
                If Double.Parse(TOTALDISCOUNT) < GROSSSALES Then
                    If SeniorGCDiscount = False Then
                        CouponDefault()
                        If S_ZeroRated = "1" Then
                            TOTALAMOUNTDUE = Val(POS.TextBoxGRANDTOTAL.Text) - TOTALDISCOUNT
                            VATABLESALES = .Label76.Text
                            VAT12PERCENT = 0
                        Else
                            TOTALAMOUNTDUE = GROSSSALES - TOTALDISCOUNT
                            VATABLESALES = Format(GROSSSALES / Tax, "0.00")
                            VAT12PERCENT = Format(VATABLESALES * S_Tax, "0.00")
                        End If
                        '.GROSSSALE = GROSSSALES
                        .TOTALDISCOUNTEDAMOUNT = GROSSSALES
                        .VATEXEMPTSALES = 0
                        .LESSVAT = 0
                        .TOTALDISCOUNT = TOTALDISCOUNT
                        .VATABLESALES = VATABLESALES
                        .VAT12PERCENT = VAT12PERCENT
                        .TOTALAMOUNTDUE = TOTALAMOUNTDUE
                        .TextBoxGRANDTOTAL.Text = NUMBERFORMAT(TOTALAMOUNTDUE)
                        .TextBoxDISCOUNT.Text = NUMBERFORMAT(TOTALDISCOUNT)
                        If S_ZeroRated = "1" Then
                            .ZERORATEDSALES = VATABLESALES
                            .ZERORATEDNETSALES = TOTALAMOUNTDUE
                        End If
                        CouponDesc = ""
                        CouponTotal = TOTALDISCOUNT
                        CouponApplied = True
                        CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                        MsgBox("Applied")
                    Else
                        If S_ZeroRated = "1" Then
                            TOTALAMOUNTDUE = GROSSSALES - TOTALDISCOUNT
                            VATABLESALES = .Label76.Text
                            VAT12PERCENT = 0
                        Else
                            TOTALAMOUNTDUE = GROSSSALES - TOTALDISCOUNT
                            VATABLESALES = Format(GROSSSALES / Tax, "0.00")
                            VAT12PERCENT = Format(VATABLESALES * S_Tax, "0.00")
                        End If
                        .TextBoxDISCOUNT.Text = Val(.TextBoxDISCOUNT.Text) + TOTALDISCOUNT
                        .TOTALDISCOUNT = .TOTALDISCOUNT + TOTALDISCOUNT
                        .TOTALAMOUNTDUE = .TOTALAMOUNTDUE - TOTALDISCOUNT
                        If S_ZeroRated = "1" Then
                            .ZERORATEDSALES = VATABLESALES
                            .ZERORATEDNETSALES = TOTALAMOUNTDUE
                        End If
                        .TextBoxGRANDTOTAL.Text = NUMBERFORMAT(Val(.TextBoxGRANDTOTAL.Text) - TOTALDISCOUNT)
                        CouponTotal = CouponTotal + TOTALDISCOUNT
                        MsgBox("Applied")
                    End If
                Else
                    If SeniorGCDiscount = False Then
                        CouponDefault()
                        If S_ZeroRated = "1" Then
                            TOTALAMOUNTDUE = 0
                            VAT12PERCENT = 0
                            VATABLESALES = Format(GROSSSALES, "0.00")
                        Else
                            TOTALAMOUNTDUE = GROSSSALES - TOTALDISCOUNT
                            VATABLESALES = Format(GROSSSALES / Tax, "0.00")
                            VAT12PERCENT = Format(GROSSSALES - VATABLESALES, "0.00")
                        End If
                        TOTALDISCOUNT = Val(POS.TextBoxGRANDTOTAL.Text)
                        TOTALAMOUNTDUE = 0
                        .GROSSSALE = GROSSSALES
                        .TOTALDISCOUNTEDAMOUNT = GROSSSALES
                        .VATEXEMPTSALES = 0
                        .LESSVAT = 0
                        .TOTALDISCOUNT = TOTALDISCOUNT
                        .VATABLESALES = VATABLESALES
                        .VAT12PERCENT = VAT12PERCENT
                        .TOTALAMOUNTDUE = TOTALAMOUNTDUE
                        .TextBoxGRANDTOTAL.Text = NUMBERFORMAT(0)
                        .TextBoxDISCOUNT.Text = NUMBERFORMAT(TOTALDISCOUNT)
                        If S_ZeroRated = "1" Then
                            .ZERORATEDSALES = VATABLESALES
                            .ZERORATEDNETSALES = TOTALAMOUNTDUE
                        End If
                        CouponDesc = ""
                        CouponTotal = TOTALDISCOUNT
                        CouponApplied = True
                        CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                        MsgBox("Applied")
                    Else
                        Dim TotalDiscounted As Double = Double.Parse(.TextBoxDISCOUNT.Text) + TOTALDISCOUNT
                        .GROSSSALE = 0
                        .TOTALDISCOUNTEDAMOUNT = GROSSSALES
                        .TOTALDISCOUNT = TotalDiscounted
                        .TOTALAMOUNTDUE = 0
                        .TextBoxGRANDTOTAL.Text = 0
                        .TextBoxDISCOUNT.Text = TotalDiscounted
                        CouponDesc = ""
                        CouponTotal = TotalDiscounted
                        CouponApplied = True
                        CouponName = CouponName & " + GC " & TOTALDISCOUNT
                        MsgBox("Applied")
                    End If
                End If
            End With
            If SeniorGCDiscount = False Then
                With POS
                    .DISCOUNTTYPE = Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                End With
            Else
                With POS
                    .DISCOUNTTYPE = .DISCOUNTTYPE & " + GC"
                End With
            End If

            Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub couponfix2()
        Try
            With POS
                If Double.Parse(.Label76.Text) < Double.Parse(Me.DataGridViewCoupons.Item(4, Me.DataGridViewCoupons.CurrentRow.Index).Value) Then
                    MsgBox("Condition not meet")
                    CouponApplied = False
                Else
                    Dim GROSSSALES As Double = POS.Label76.Text
                    Dim TAX As Double = 1 + Val(S_Tax)
                    Dim TOTALDISCOUNT As Double = DataGridViewCoupons.SelectedRows(0).Cells(3).Value
                    Dim TOTALAMOUNTDUE As Double = 0

                    TOTALAMOUNTDUE = Format(GROSSSALES - TOTALDISCOUNT, "0.00")
                    Dim VATABLESALES As Double = 0
                    Dim VAT12PERCENT As Double = 0
                    Dim LESSVAT As Double = 0



                    If S_ZeroRated = "1" Then
                        Dim TotalPrice As Double = 0
                        Dim GrandTotal As Double = 0
                        With POS.DataGridViewOrders
                            For i As Integer = 0 To .Rows.Count - 1 Step +1
                                TotalPrice = .Rows(i).Cells(1).Value * .Rows(i).Cells(2).Value
                                GrandTotal += TotalPrice
                            Next
                        End With

                        Dim SubTotal = Convert.ToDecimal(Double.Parse(POS.TextBoxSUBTOTAL.Text))
                        LESSVAT = Math.Round(GrandTotal - SubTotal, 2, MidpointRounding.AwayFromZero)

                        VATABLESALES = 0
                        VAT12PERCENT = 0
                        .ZERORATEDSALES = .Label76.Text
                        .ZERORATEDNETSALES = TOTALAMOUNTDUE
                    Else
                        VATABLESALES = Format(GROSSSALES / TAX, "0.00")
                        VAT12PERCENT = Format(VATABLESALES * S_Tax, "0.00")
                        .ZERORATEDSALES = 0
                        .ZERORATEDNETSALES = 0
                    End If

                    '.GROSSSALE = GROSSSALES
                    .TOTALDISCOUNTEDAMOUNT = GROSSSALES
                    .VATEXEMPTSALES = 0
                    .LESSVAT = LESSVAT
                    .TOTALDISCOUNT = TOTALDISCOUNT
                    .VATABLESALES = VATABLESALES
                    .VAT12PERCENT = VAT12PERCENT
                    .TOTALAMOUNTDUE = TOTALAMOUNTDUE

                    .TextBoxGRANDTOTAL.Text = TOTALAMOUNTDUE
                    .TextBoxDISCOUNT.Text = TOTALDISCOUNT
                    CouponDesc = ""
                    CouponTotal = TOTALDISCOUNT
                    CouponApplied = True
                    CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                    MsgBox("Applied")
                End If
            End With
            With POS
                .DISCOUNTTYPE = Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
            End With
            Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub couponbundle1()
        Try
            With POS
                Dim ReferenceExist As Boolean = False
                Dim referenceID As String = Me.DataGridViewCoupons.Item(6, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString


                Dim refIds As String() = referenceID.Split(New Char() {","c})
                'Total Discount
                Dim DISCOUNTEDAMOUNT As Double = 0
                Dim GROSSSALES As Double = .Label76.Text
                Dim VATABLESALES As Double = 0
                Dim TAX As Double = 1 + Val(S_Tax)

                Try
                    For Each getRefids In refIds
                        For i As Integer = 0 To .DataGridViewOrders.Rows.Count - 1 Step +1
                            If .DataGridViewOrders.Rows(i).Cells(5).Value.ToString.Contains(getRefids) = True Then
                                If .DataGridViewOrders.Rows(i).Cells(1).Value >= Me.DataGridViewCoupons.Item(7, Me.DataGridViewCoupons.CurrentRow.Index).Value Then

                                    ReferenceExist = True
                                    Exit Try
                                Else
                                    CouponDefault()
                                End If
                            Else
                                ReferenceExist = False
                            End If
                        Next
                    Next

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
                'If ReferenceExist Then
                '    With POS
                '        'If S_ZeroRated = "0" Then
                '        '    DISCOUNTEDAMOUNT = .DataGridViewOrders.Rows(i).Cells(2).Value
                '        '    MsgBox(.DataGridViewOrders.Rows(i).Cells(5).Value)
                '        'Else
                '        '    DISCOUNTEDAMOUNT = .DataGridViewOrders.Rows(i).Cells(2).Value / TAX
                '        'End If
                '    End With
                'End If

                'Gross Sales/ Subtotal
                Dim TOTALAMOUNTDUE As Double = 0

                Dim SUBTOTAL As Double = 0
                If S_ZeroRated = "1" Then
                    SUBTOTAL = Val(POS.TextBoxGRANDTOTAL.Text)
                Else
                    SUBTOTAL = Val(POS.Label76.Text)
                End If

                Dim QtyCondMeet As Boolean = True
                Dim TotalPrice As Double = 0
                Dim BundlepromoID As String = Me.DataGridViewCoupons.Item(8, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                Dim bundIds As String() = BundlepromoID.Split(New Char() {","c})
                Dim BundleIDExist As Boolean = False
                If ReferenceExist = True Then
                    For Each getBundleids In bundIds
                        For i As Integer = 0 To .DataGridViewOrders.Rows.Count - 1 Step +1
                            If POS.DataGridViewOrders.Rows(i).Cells(5).Value.ToString.Contains(getBundleids) = True Then
                                DISCOUNTEDAMOUNT = .DataGridViewOrders.Rows(i).Cells(2).Value
                                TotalPrice += .DataGridViewOrders.Rows(i).Cells(3).Value
                                TOTALAMOUNTDUE = SUBTOTAL - DISCOUNTEDAMOUNT
                                If POS.DataGridViewOrders.Rows(i).Cells(1).Value >= Me.DataGridViewCoupons.SelectedRows(0).Cells(9).Value Then
                                    Dim GetDiscount As Double = 0
                                    Dim OrgPrice = POS.DataGridViewOrders.Rows(i).Cells(2).Value
                                    Dim Qty = Me.DataGridViewCoupons.SelectedRows(0).Cells(9).Value
                                    Dim TotalLess = OrgPrice * Qty
                                    Dim TotalZeroRatedPrice As Double = 0
                                    Dim VAT12PERCENT As Double = 0
                                    If S_ZeroRated = "1" Then
                                        GetDiscount = Math.Round(POS.DataGridViewOrders.Rows(i).Cells(2).Value / TAX, 2, MidpointRounding.AwayFromZero)
                                        TotalZeroRatedPrice = GetDiscount * DataGridViewCoupons.SelectedRows(0).Cells(7).Value
                                        .ZERORATEDSALES = .Label76.Text
                                        .ZERORATEDNETSALES = TOTALAMOUNTDUE
                                        .TextBoxGRANDTOTAL.Text = Format(Val(.TextBoxGRANDTOTAL.Text) - TotalZeroRatedPrice, "0.00")
                                        .TextBoxDISCOUNT.Text = Format(GetDiscount, "0.00")
                                        VAT12PERCENT = 0
                                        VATABLESALES = .Label76.Text
                                    Else
                                        VATABLESALES = GROSSSALES / TAX
                                        VAT12PERCENT = Format(VATABLESALES * S_Tax, "0.00")
                                        .TextBoxGRANDTOTAL.Text = Format(TOTALAMOUNTDUE, "0.00")
                                        .TextBoxDISCOUNT.Text = Format(DISCOUNTEDAMOUNT, "0.00")
                                    End If

                                    '.GROSSSALE = GROSSSALES
                                    .TOTALDISCOUNTEDAMOUNT = DISCOUNTEDAMOUNT
                                    .VATEXEMPTSALES = 0
                                    .LESSVAT = 0
                                    .TOTALDISCOUNT = DISCOUNTEDAMOUNT
                                    .VATABLESALES = VATABLESALES
                                    .VAT12PERCENT = VAT12PERCENT
                                    .TOTALAMOUNTDUE = TOTALAMOUNTDUE

                                    CouponDesc = ""
                                    CouponTotal = DISCOUNTEDAMOUNT
                                    CouponApplied = True
                                    BundleIDExist = True
                                    CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                                Else
                                    CouponDefault()
                                    CouponApplied = False
                                    BundleIDExist = False
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                    If BundleIDExist = False Then
                        MsgBox("Please select bundle promo")
                    End If
                Else
                    CouponDefault()
                    MsgBox("Condition not meet")
                End If
            End With
            With POS
                .DISCOUNTTYPE = Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
            End With
            Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub couponbundle2()
        Try
            CouponDefault()
            Dim ReferenceExist As Boolean = False
            Dim GROSSSALES As Double = 0
            Dim VATABLESALES As Double = 0
            Dim VAT12PERCENT As Double = 0
            Dim TAX As Double = 1 + Val(S_Tax)
            Dim referenceID As String = DataGridViewCoupons.Item(6, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
            Dim TOTALDISCOUNT As Double = DataGridViewCoupons.SelectedRows(0).Cells(3).Value
            Dim refIds As String() = referenceID.Split(New Char() {","c})
            With POS
                Try
                    For Each getRefids In refIds
                        For i As Integer = 0 To .DataGridViewOrders.Rows.Count - 1 Step +1
                            If POS.DataGridViewOrders.Rows(i).Cells(5).Value.ToString.Contains(getRefids) = True Then
                                If POS.DataGridViewOrders.Rows(i).Cells(1).Value >= Me.DataGridViewCoupons.Item(7, Me.DataGridViewCoupons.CurrentRow.Index).Value Then
                                    ReferenceExist = True
                                    Exit Try
                                Else
                                    CouponDefault()
                                End If
                            Else
                                ReferenceExist = False
                            End If
                        Next
                    Next
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
                Dim TOTALAMOUNTDUE As Double = 0
                For i As Integer = 0 To .DataGridViewOrders.Rows.Count - 1 Step +1
                    GROSSSALES += .DataGridViewOrders.Rows(i).Cells(3).Value
                Next
                TOTALAMOUNTDUE = GROSSSALES - TOTALDISCOUNT
                VATABLESALES = GROSSSALES / TAX
                Dim QtyCondMeet As Boolean = True
                Dim TotalPrice As Integer = 0
                Dim BundlepromoID As String = Me.DataGridViewCoupons.Item(8, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                Dim bundIds As String() = BundlepromoID.Split(New Char() {","c})
                Dim BundleIDExist As Boolean = False
                If ReferenceExist = True Then
                    Try
                        For Each getBundleids In bundIds
                            Console.Write(getBundleids)
                            For i As Integer = 0 To .DataGridViewOrders.Rows.Count - 1 Step +1
                                If POS.DataGridViewOrders.Rows(i).Cells(5).Value.ToString.Contains(getBundleids) = True Then
                                    If POS.DataGridViewOrders.Rows(i).Cells(1).Value >= Me.DataGridViewCoupons.SelectedRows(0).Cells(9).Value Then


                                        If S_ZeroRated = "1" Then
                                            VATABLESALES = .Label76.Text
                                            .ZERORATEDSALES = VATABLESALES
                                            .ZERORATEDNETSALES = TOTALAMOUNTDUE

                                            .TextBoxGRANDTOTAL.Text = Format(Val(.TextBoxGRANDTOTAL.Text) - TOTALDISCOUNT, "0.00")
                                            .TextBoxDISCOUNT.Text = Format(TOTALDISCOUNT, "0.00")
                                            .TOTALAMOUNTDUE = .TextBoxGRANDTOTAL.Text
                                        Else
                                            .TextBoxGRANDTOTAL.Text = Format(TOTALAMOUNTDUE, "0.00")
                                            .TextBoxDISCOUNT.Text = Format(TOTALDISCOUNT, "0.00")
                                            .TOTALAMOUNTDUE = TOTALAMOUNTDUE
                                            VATABLESALES = Format(VATABLESALES, "0.00")
                                            VAT12PERCENT = Format(VATABLESALES * S_Tax, "0.00")
                                        End If

                                        .GROSSSALE = GROSSSALES
                                        .TOTALDISCOUNTEDAMOUNT = TOTALAMOUNTDUE
                                        .VATEXEMPTSALES = 0
                                        .LESSVAT = 0
                                        .TOTALDISCOUNT = TOTALDISCOUNT
                                        .VATABLESALES = VATABLESALES
                                        .VAT12PERCENT = VAT12PERCENT


                                        CouponDesc = ""
                                        CouponTotal = TOTALDISCOUNT
                                        CouponApplied = True
                                        CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                                        BundleIDExist = True
                                        Exit Try
                                    Else
                                        CouponApplied = False
                                        BundleIDExist = False
                                    End If
                                Else

                                End If
                            Next
                        Next
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                    If BundleIDExist = False Then
                        MsgBox("Please select bundle promo")
                    End If
                Else
                    MsgBox("Condition not meet")
                End If
            End With
            With POS
                .DISCOUNTTYPE = Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
            End With
            Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub couponbundle3()
        CouponDefault()
        Try
            With POS
                Dim TAX As Double = 1 + Val(S_Tax)
                Dim ReferenceExist As Boolean = False
                Dim referenceID As String = Me.DataGridViewCoupons.Item(6, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                Dim refIds As String() = referenceID.Split(New Char() {","c})
                Dim TotalQtyCount As Integer = 0
                Try
                    For Each getRefids In refIds
                        For i As Integer = 0 To .DataGridViewOrders.Rows.Count - 1 Step +1
                            If POS.DataGridViewOrders.Rows(i).Cells(5).Value = getRefids Then
                                TotalQtyCount += POS.DataGridViewOrders.Rows(i).Cells(1).Value
                            End If
                        Next
                    Next
                Catch ex As Exception
                    SendErrorReport(ex.ToString)
                End Try

                Dim BundlepromoID As String = Me.DataGridViewCoupons.Item(8, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                Dim bundIds As String() = BundlepromoID.Split(New Char() {","c})
                Dim BundpromoID As Boolean = False
                Dim CountQty As Integer = 0
                If TotalQtyCount >= Me.DataGridViewCoupons.Item(7, Me.DataGridViewCoupons.CurrentRow.Index).Value Then
                    Try
                        For Each getBundleids In bundIds
                            For i As Integer = 0 To .DataGridViewOrders.Rows.Count - 1 Step +1
                                If POS.DataGridViewOrders.Rows(i).Cells(5).Value.ToString.Contains(getBundleids) = True Then
                                    If POS.DataGridViewOrders.Rows(i).Cells(1).Value >= Me.DataGridViewCoupons.SelectedRows(0).Cells(9).Value Then
                                        BundpromoID = True
                                        Exit Try
                                    Else
                                        BundpromoID = False
                                    End If
                                End If
                            Next
                        Next
                    Catch ex As Exception
                        SendErrorReport(ex.ToString)
                    End Try
                    Dim GROSSSALES As Double = 0
                    Dim DISCOUNTAMOUNT As Double = 0
                    Dim Percentage As Double = DataGridViewCoupons.SelectedRows(0).Cells(3).Value / 100
                    Dim TOTALDISCOUNT As Double = 0
                    Dim VATABLESALES As Double = 0
                    Dim TOTALAMOUNTDUE As Double = 0
                    Dim VAT12PERCENT As Double = 0
                    If BundpromoID = True Then
                        For i As Integer = 0 To .DataGridViewOrders.Rows.Count - 1 Step +1
                            GROSSSALES += .DataGridViewOrders.Rows(i).Cells(2).Value
                            VATABLESALES = GROSSSALES / TAX
                            If .DataGridViewOrders.Rows(i).Cells(9).Value.ToString = "WAFFLE" Then
                                CountQty += .DataGridViewOrders.Rows(i).Cells(1).Value
                            End If

                            If CountQty = 3 Then
                                'MsgBox(.DataGridViewOrders.Rows(i).Cells(2).Value)
                                'MsgBox(Percentage)
                                If S_ZeroRated = "0" Then
                                    VAT12PERCENT = Format(VATABLESALES * S_Tax, "0.00")
                                    DISCOUNTAMOUNT = .DataGridViewOrders.Rows(i).Cells(2).Value
                                    VATABLESALES = Format(TOTALAMOUNTDUE / TAX, "0.00")
                                    TOTALDISCOUNT = Percentage * DISCOUNTAMOUNT
                                    .TextBoxDISCOUNT.Text = Format(TOTALDISCOUNT, "0.00")
                                    TOTALAMOUNTDUE = Double.Parse(POS.TextBoxGRANDTOTAL.Text) - Double.Parse(POS.TextBoxDISCOUNT.Text)
                                Else
                                    DISCOUNTAMOUNT = .DataGridViewOrders.Rows(i).Cells(2).Value / TAX
                                    VATABLESALES = .Label76.Text
                                    TOTALDISCOUNT = Percentage * DISCOUNTAMOUNT
                                    .TextBoxDISCOUNT.Text = Format(TOTALDISCOUNT, "0.00")
                                    TOTALAMOUNTDUE = Double.Parse(POS.TextBoxGRANDTOTAL.Text) - Double.Parse(POS.TextBoxDISCOUNT.Text)
                                    .ZERORATEDSALES = VATABLESALES
                                    .ZERORATEDNETSALES = TOTALAMOUNTDUE
                                    VAT12PERCENT = 0
                                End If
                                .GROSSSALE = GROSSSALES
                                .TOTALDISCOUNTEDAMOUNT = DISCOUNTAMOUNT
                                .VATEXEMPTSALES = 0
                                .LESSVAT = 0
                                .TOTALDISCOUNT = TOTALDISCOUNT
                                .VATABLESALES = VATABLESALES
                                .VAT12PERCENT = VAT12PERCENT
                                .TOTALAMOUNTDUE = TOTALAMOUNTDUE
                                .TextBoxGRANDTOTAL.Text = Format(TOTALAMOUNTDUE, "0.00")

                                CouponDesc = ""
                                CouponTotal = TOTALDISCOUNT
                                CouponApplied = True
                                CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                            End If

                        Next
                        If CountQty < 3 Then
                            MsgBox("Buy 3 waffles to apply the coupon")
                        End If
                    End If
                Else
                    MsgBox("Condition not meet")
                End If
            End With
            With POS
                .DISCOUNTTYPE = Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
            End With
            Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    Private Sub couponbundle4()
        Try
            If S_ZeroRated = "0" Then
                CouponDefault()
                With POS
                    Dim GROSSSALES As Double = 0
                    Dim TAX As Double = 1 + Val(S_Tax)
                    Dim TOTALDISCOUNT As Double = DataGridViewCoupons.SelectedRows(0).Cells(3).Value / 100
                    Dim TOTALAMOUNTDUE As Double = 0
                    With .DataGridViewOrders
                        For i As Integer = 0 To .Rows.Count - 1 Step +1
                            GROSSSALES += .Rows(i).Cells(3).Value
                        Next
                    End With
                    TOTALDISCOUNT = GROSSSALES * TOTALDISCOUNT
                    TOTALAMOUNTDUE = Format(GROSSSALES - TOTALDISCOUNT, "0.00")

                    Dim VATABLESALES As Double = Format(GROSSSALES / TAX, "0.00")
                    Dim VAT12PERCENT As Double = Format(VATABLESALES * S_Tax, "0.00")

                    .GROSSSALE = GROSSSALES
                    .TOTALDISCOUNTEDAMOUNT = GROSSSALES
                    .VATEXEMPTSALES = 0
                    .LESSVAT = 0
                    .TOTALDISCOUNT = TOTALDISCOUNT
                    .VATABLESALES = VATABLESALES
                    .VAT12PERCENT = VAT12PERCENT
                    .TOTALAMOUNTDUE = TOTALAMOUNTDUE
                    .TextBoxGRANDTOTAL.Text = TOTALAMOUNTDUE
                    .TextBoxDISCOUNT.Text = TOTALDISCOUNT
                    CouponDesc = ""
                    CouponTotal = TOTALDISCOUNT
                    CouponApplied = True
                    CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                    MsgBox("Applied")

                End With
            Else
                CouponDefault()
                With POS
                    Dim GROSSSALES As Double = Double.Parse(.TextBoxGRANDTOTAL.Text)
                    Dim TAX As Double = 1 + Val(S_Tax)
                    Dim TOTALDISCOUNT As Double = DataGridViewCoupons.SelectedRows(0).Cells(3).Value / 100
                    Dim TOTALAMOUNTDUE As Double = 0

                    TOTALDISCOUNT = GROSSSALES * TOTALDISCOUNT
                    TOTALAMOUNTDUE = Format(GROSSSALES - TOTALDISCOUNT, "0.00")

                    'Dim VATABLESALES As Double = Format(GROSSSALES / TAX, "0.00")
                    'Dim VAT12PERCENT As Double = Format(VATABLESALES * S_Tax, "0.00")

                    .GROSSSALE = GROSSSALES
                    .TOTALDISCOUNTEDAMOUNT = GROSSSALES
                    .VATEXEMPTSALES = 0
                    .LESSVAT = 0
                    .TOTALDISCOUNT = TOTALDISCOUNT
                    .VATABLESALES = .Label76.Text
                    .VAT12PERCENT = 0
                    .TOTALAMOUNTDUE = TOTALAMOUNTDUE
                    .TextBoxGRANDTOTAL.Text = TOTALAMOUNTDUE
                    .TextBoxDISCOUNT.Text = NUMBERFORMAT(TOTALDISCOUNT)

                    .ZERORATEDSALES = .Label76.Text
                    .ZERORATEDNETSALES = .Label76.Text - TOTALDISCOUNT
                    .TextBoxGRANDTOTAL.Text = NUMBERFORMAT(.Label76.Text - TOTALDISCOUNT)

                    CouponDesc = ""
                    CouponTotal = TOTALDISCOUNT
                    CouponApplied = True
                    CouponName = Me.DataGridViewCoupons.Item(1, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
                    MsgBox("Applied")
                End With
            End If
            With POS
                .DISCOUNTTYPE = Me.DataGridViewCoupons.Item(5, Me.DataGridViewCoupons.CurrentRow.Index).Value.ToString
            End With
            Close()
        Catch ex As Exception
            SendErrorReport(ex.ToString)
        End Try
    End Sub
    'Private Sub DataGridViewCoupons_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridViewCoupons.CellClick
    '    Try
    '        LabelDesc.Text = DataGridViewCoupons.SelectedRows(0).Cells(2).Value.ToString
    '    Catch ex As Exception
    '        SendErrorReport(ex.ToString)
    '    End Try
    'End Sub
End Class