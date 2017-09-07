<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DirectBillImportForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DirectBillImportForm))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbCarrier = New System.Windows.Forms.ComboBox()
        Me.btnMakeWhole = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.cbFileSelect = New System.Windows.Forms.ComboBox()
        Me.btnMove = New System.Windows.Forms.Button()
        Me.tbAComm = New System.Windows.Forms.TextBox()
        Me.tbGComm = New System.Windows.Forms.TextBox()
        Me.tbPremium = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tbAPAmt = New System.Windows.Forms.TextBox()
        Me.tbRevAmt = New System.Windows.Forms.TextBox()
        Me.tbTPay = New System.Windows.Forms.TextBox()
        Me.tbAmount = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.tbPolicyCount = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.tbBatId = New System.Windows.Forms.TextBox()
        Me.btnGArchive = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Carrier"
        '
        'cbCarrier
        '
        Me.cbCarrier.FormattingEnabled = True
        Me.cbCarrier.Location = New System.Drawing.Point(59, 22)
        Me.cbCarrier.Name = "cbCarrier"
        Me.cbCarrier.Size = New System.Drawing.Size(156, 21)
        Me.cbCarrier.TabIndex = 1
        '
        'btnMakeWhole
        '
        Me.btnMakeWhole.Enabled = False
        Me.btnMakeWhole.Location = New System.Drawing.Point(59, 111)
        Me.btnMakeWhole.Name = "btnMakeWhole"
        Me.btnMakeWhole.Size = New System.Drawing.Size(156, 25)
        Me.btnMakeWhole.TabIndex = 2
        Me.btnMakeWhole.Text = "Make Whole"
        Me.btnMakeWhole.UseVisualStyleBackColor = True
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(59, 142)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(156, 25)
        Me.btnRun.TabIndex = 3
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'cbFileSelect
        '
        Me.cbFileSelect.Enabled = False
        Me.cbFileSelect.FormattingEnabled = True
        Me.cbFileSelect.Location = New System.Drawing.Point(59, 84)
        Me.cbFileSelect.Name = "cbFileSelect"
        Me.cbFileSelect.Size = New System.Drawing.Size(156, 21)
        Me.cbFileSelect.TabIndex = 4
        '
        'btnMove
        '
        Me.btnMove.Enabled = False
        Me.btnMove.Location = New System.Drawing.Point(59, 53)
        Me.btnMove.Name = "btnMove"
        Me.btnMove.Size = New System.Drawing.Size(156, 25)
        Me.btnMove.TabIndex = 5
        Me.btnMove.Text = "Move Files"
        Me.btnMove.UseVisualStyleBackColor = True
        '
        'tbAComm
        '
        Me.tbAComm.Enabled = False
        Me.tbAComm.Location = New System.Drawing.Point(382, 55)
        Me.tbAComm.Name = "tbAComm"
        Me.tbAComm.Size = New System.Drawing.Size(100, 20)
        Me.tbAComm.TabIndex = 12
        '
        'tbGComm
        '
        Me.tbGComm.Enabled = False
        Me.tbGComm.Location = New System.Drawing.Point(382, 81)
        Me.tbGComm.Name = "tbGComm"
        Me.tbGComm.Size = New System.Drawing.Size(100, 20)
        Me.tbGComm.TabIndex = 13
        '
        'tbPremium
        '
        Me.tbPremium.Enabled = False
        Me.tbPremium.Location = New System.Drawing.Point(382, 27)
        Me.tbPremium.Name = "tbPremium"
        Me.tbPremium.Size = New System.Drawing.Size(100, 20)
        Me.tbPremium.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(329, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(47, 13)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Premium"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(257, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(119, 13)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Total Gross Commission"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(256, 62)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(120, 13)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Total Agent Commission"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(305, 114)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 13)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Total Payable"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(286, 168)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(90, 13)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Revenue Amount"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(316, 194)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(60, 13)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "AP Amount"
        '
        'tbAPAmt
        '
        Me.tbAPAmt.Enabled = False
        Me.tbAPAmt.Location = New System.Drawing.Point(382, 187)
        Me.tbAPAmt.Name = "tbAPAmt"
        Me.tbAPAmt.Size = New System.Drawing.Size(100, 20)
        Me.tbAPAmt.TabIndex = 29
        '
        'tbRevAmt
        '
        Me.tbRevAmt.Enabled = False
        Me.tbRevAmt.Location = New System.Drawing.Point(382, 161)
        Me.tbRevAmt.Name = "tbRevAmt"
        Me.tbRevAmt.Size = New System.Drawing.Size(100, 20)
        Me.tbRevAmt.TabIndex = 30
        '
        'tbTPay
        '
        Me.tbTPay.Enabled = False
        Me.tbTPay.Location = New System.Drawing.Point(382, 107)
        Me.tbTPay.Name = "tbTPay"
        Me.tbTPay.Size = New System.Drawing.Size(100, 20)
        Me.tbTPay.TabIndex = 31
        '
        'tbAmount
        '
        Me.tbAmount.Enabled = False
        Me.tbAmount.Location = New System.Drawing.Point(382, 135)
        Me.tbAmount.Name = "tbAmount"
        Me.tbAmount.Size = New System.Drawing.Size(100, 20)
        Me.tbAmount.TabIndex = 33
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(334, 138)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(43, 13)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "Amount"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(42, 221)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(67, 13)
        Me.Label9.TabIndex = 35
        Me.Label9.Text = "Import Count"
        '
        'tbPolicyCount
        '
        Me.tbPolicyCount.Enabled = False
        Me.tbPolicyCount.Location = New System.Drawing.Point(115, 218)
        Me.tbPolicyCount.Name = "tbPolicyCount"
        Me.tbPolicyCount.Size = New System.Drawing.Size(100, 20)
        Me.tbPolicyCount.TabIndex = 34
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(34, 195)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(75, 13)
        Me.Label10.TabIndex = 37
        Me.Label10.Text = "Batch Number"
        '
        'tbBatId
        '
        Me.tbBatId.Enabled = False
        Me.tbBatId.Location = New System.Drawing.Point(115, 188)
        Me.tbBatId.Name = "tbBatId"
        Me.tbBatId.Size = New System.Drawing.Size(100, 20)
        Me.tbBatId.TabIndex = 36
        '
        'btnGArchive
        '
        Me.btnGArchive.Enabled = False
        Me.btnGArchive.Location = New System.Drawing.Point(372, 215)
        Me.btnGArchive.Name = "btnGArchive"
        Me.btnGArchive.Size = New System.Drawing.Size(110, 23)
        Me.btnGArchive.TabIndex = 38
        Me.btnGArchive.Tag = ""
        Me.btnGArchive.Text = "Geovera Archive"
        Me.btnGArchive.UseVisualStyleBackColor = True
        '
        'DirectBillImportForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(505, 269)
        Me.Controls.Add(Me.btnGArchive)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.tbBatId)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.tbPolicyCount)
        Me.Controls.Add(Me.tbAmount)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbTPay)
        Me.Controls.Add(Me.tbRevAmt)
        Me.Controls.Add(Me.tbAPAmt)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbPremium)
        Me.Controls.Add(Me.tbGComm)
        Me.Controls.Add(Me.tbAComm)
        Me.Controls.Add(Me.btnMove)
        Me.Controls.Add(Me.cbFileSelect)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.btnMakeWhole)
        Me.Controls.Add(Me.cbCarrier)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "DirectBillImportForm"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Direct Bill Imports"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbCarrier As System.Windows.Forms.ComboBox
    Friend WithEvents btnMakeWhole As System.Windows.Forms.Button
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents cbFileSelect As System.Windows.Forms.ComboBox
    Friend WithEvents btnMove As System.Windows.Forms.Button
    Friend WithEvents tbAComm As System.Windows.Forms.TextBox
    Friend WithEvents tbGComm As System.Windows.Forms.TextBox
    Friend WithEvents tbPremium As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents tbAPAmt As System.Windows.Forms.TextBox
    Friend WithEvents tbRevAmt As System.Windows.Forms.TextBox
    Friend WithEvents tbTPay As System.Windows.Forms.TextBox
    Friend WithEvents tbAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents tbPolicyCount As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents tbBatId As System.Windows.Forms.TextBox
    Friend WithEvents btnGArchive As System.Windows.Forms.Button

End Class
