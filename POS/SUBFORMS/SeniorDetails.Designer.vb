<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SeniorDetails
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SeniorDetails))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxSENIORID = New System.Windows.Forms.TextBox()
        Me.TextBoxSENIORNAME = New System.Windows.Forms.TextBox()
        Me.ButtonCANCEL = New System.Windows.Forms.Button()
        Me.ButtonSubmit = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(130, 16)
        Me.Label1.TabIndex = 104
        Me.Label1.Text = "Identification Number"
        '
        'TextBoxSENIORID
        '
        Me.TextBoxSENIORID.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBoxSENIORID.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxSENIORID.Location = New System.Drawing.Point(16, 31)
        Me.TextBoxSENIORID.Name = "TextBoxSENIORID"
        Me.TextBoxSENIORID.Size = New System.Drawing.Size(360, 16)
        Me.TextBoxSENIORID.TabIndex = 0
        '
        'TextBoxSENIORNAME
        '
        Me.TextBoxSENIORNAME.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBoxSENIORNAME.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxSENIORNAME.Location = New System.Drawing.Point(16, 83)
        Me.TextBoxSENIORNAME.Name = "TextBoxSENIORNAME"
        Me.TextBoxSENIORNAME.Size = New System.Drawing.Size(360, 16)
        Me.TextBoxSENIORNAME.TabIndex = 1
        '
        'ButtonCANCEL
        '
        Me.ButtonCANCEL.BackColor = System.Drawing.Color.Red
        Me.ButtonCANCEL.FlatAppearance.BorderSize = 0
        Me.ButtonCANCEL.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonCANCEL.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonCANCEL.ForeColor = System.Drawing.Color.White
        Me.ButtonCANCEL.Location = New System.Drawing.Point(262, 112)
        Me.ButtonCANCEL.Name = "ButtonCANCEL"
        Me.ButtonCANCEL.Size = New System.Drawing.Size(114, 31)
        Me.ButtonCANCEL.TabIndex = 106
        Me.ButtonCANCEL.Text = "Cancel"
        Me.ButtonCANCEL.UseVisualStyleBackColor = False
        '
        'ButtonSubmit
        '
        Me.ButtonSubmit.BackColor = System.Drawing.Color.FromArgb(CType(CType(248, Byte), Integer), CType(CType(160, Byte), Integer), CType(CType(34, Byte), Integer))
        Me.ButtonSubmit.FlatAppearance.BorderSize = 0
        Me.ButtonSubmit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonSubmit.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSubmit.ForeColor = System.Drawing.Color.White
        Me.ButtonSubmit.Location = New System.Drawing.Point(15, 112)
        Me.ButtonSubmit.Name = "ButtonSubmit"
        Me.ButtonSubmit.Size = New System.Drawing.Size(242, 31)
        Me.ButtonSubmit.TabIndex = 103
        Me.ButtonSubmit.Text = "Submit"
        Me.ButtonSubmit.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 16)
        Me.Label2.TabIndex = 105
        Me.Label2.Text = "Full Name"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(13, 33)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(365, 16)
        Me.Label6.TabIndex = 252
        Me.Label6.Text = "___________________________________________________"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(13, 85)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(365, 16)
        Me.Label3.TabIndex = 253
        Me.Label3.Text = "___________________________________________________"
        '
        'SeniorDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(388, 157)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBoxSENIORID)
        Me.Controls.Add(Me.TextBoxSENIORNAME)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ButtonCANCEL)
        Me.Controls.Add(Me.ButtonSubmit)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label6)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "SeniorDetails"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "POS | COUPON DETAILS"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBoxSENIORID As TextBox
    Friend WithEvents TextBoxSENIORNAME As TextBox
    Friend WithEvents ButtonCANCEL As Button
    Friend WithEvents ButtonSubmit As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label3 As Label
End Class
