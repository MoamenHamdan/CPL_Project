<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.txtMoney = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.dtglist = New System.Windows.Forms.DataGridView()
        Me.txtRemarks = New System.Windows.Forms.RichTextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpTransDate = New System.Windows.Forms.DateTimePicker()
        Me.cboType = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dtpfrom = New System.Windows.Forms.DateTimePicker()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.btnclear = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtRecieved = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtWidthraw = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtBalance = New System.Windows.Forms.TextBox()
        Me.dtpto = New System.Windows.Forms.DateTimePicker()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        CType(Me.dtglist, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtMoney
        '
        Me.txtMoney.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.txtMoney.Location = New System.Drawing.Point(111, 9)
        Me.txtMoney.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtMoney.Name = "txtMoney"
        Me.txtMoney.Size = New System.Drawing.Size(180, 26)
        Me.txtMoney.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 9)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Money :"
        '
        'btnSave
        '
        Me.btnSave.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSave.Location = New System.Drawing.Point(12, 232)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(90, 48)
        Me.btnSave.TabIndex = 2
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dtglist
        '
        Me.dtglist.AllowUserToAddRows = False
        Me.dtglist.AllowUserToDeleteRows = False
        Me.dtglist.AllowUserToResizeColumns = False
        Me.dtglist.AllowUserToResizeRows = False
        Me.dtglist.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtglist.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dtglist.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dtglist.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dtglist.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dtglist.Location = New System.Drawing.Point(3, 295)
        Me.dtglist.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtglist.Name = "dtglist"
        Me.dtglist.RowHeadersWidth = 62
        Me.dtglist.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dtglist.Size = New System.Drawing.Size(1043, 462)
        Me.dtglist.TabIndex = 3
        '
        'txtRemarks
        '
        Me.txtRemarks.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.txtRemarks.Location = New System.Drawing.Point(111, 45)
        Me.txtRemarks.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.Size = New System.Drawing.Size(454, 62)
        Me.txtRemarks.TabIndex = 4
        Me.txtRemarks.Text = ""
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 65)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 20)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Remarks :"
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(22, 178)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(135, 20)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Transaction Date:"
        '
        'dtpTransDate
        '
        Me.dtpTransDate.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.dtpTransDate.CustomFormat = "yyyy-MM-dd hh:mm:ss"
        Me.dtpTransDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTransDate.Location = New System.Drawing.Point(177, 178)
        Me.dtpTransDate.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtpTransDate.Name = "dtpTransDate"
        Me.dtpTransDate.Size = New System.Drawing.Size(198, 26)
        Me.dtpTransDate.TabIndex = 7
        '
        'cboType
        '
        Me.cboType.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.cboType.FormattingEnabled = True
        Me.cboType.Items.AddRange(New Object() {"Received", "Withdraw"})
        Me.cboType.Location = New System.Drawing.Point(177, 122)
        Me.cboType.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cboType.Name = "cboType"
        Me.cboType.Size = New System.Drawing.Size(194, 28)
        Me.cboType.TabIndex = 8
        Me.cboType.Text = "Select"
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 125)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(156, 20)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Type of Transaction :"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(551, 260)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 20)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Search  :"
        '
        'dtpfrom
        '
        Me.dtpfrom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpfrom.CustomFormat = "yyyy-MM-dd"
        Me.dtpfrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpfrom.Location = New System.Drawing.Point(631, 255)
        Me.dtpfrom.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtpfrom.Name = "dtpfrom"
        Me.dtpfrom.Size = New System.Drawing.Size(142, 26)
        Me.dtpfrom.TabIndex = 12
        '
        'btnSearch
        '
        Me.btnSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSearch.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSearch.Location = New System.Drawing.Point(937, 251)
        Me.btnSearch.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(92, 35)
        Me.btnSearch.TabIndex = 13
        Me.btnSearch.Text = "Find"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'btnclear
        '
        Me.btnclear.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnclear.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnclear.Location = New System.Drawing.Point(111, 232)
        Me.btnclear.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnclear.Name = "btnclear"
        Me.btnclear.Size = New System.Drawing.Size(90, 48)
        Me.btnclear.TabIndex = 14
        Me.btnclear.Text = "Clear"
        Me.btnclear.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDelete.Location = New System.Drawing.Point(210, 232)
        Me.btnDelete.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(90, 48)
        Me.btnDelete.TabIndex = 15
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(446, 125)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(192, 20)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Over All Recieved Money :"
        '
        'txtRecieved
        '
        Me.txtRecieved.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtRecieved.Location = New System.Drawing.Point(669, 119)
        Me.txtRecieved.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtRecieved.Name = "txtRecieved"
        Me.txtRecieved.ReadOnly = True
        Me.txtRecieved.Size = New System.Drawing.Size(284, 26)
        Me.txtRecieved.TabIndex = 16
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(446, 165)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(192, 20)
        Me.Label7.TabIndex = 19
        Me.Label7.Text = "Over All Withdraw Money :"
        '
        'txtWidthraw
        '
        Me.txtWidthraw.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtWidthraw.Location = New System.Drawing.Point(669, 162)
        Me.txtWidthraw.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtWidthraw.Name = "txtWidthraw"
        Me.txtWidthraw.ReadOnly = True
        Me.txtWidthraw.Size = New System.Drawing.Size(284, 26)
        Me.txtWidthraw.TabIndex = 18
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(446, 212)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(133, 20)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Over All Balance :"
        '
        'txtBalance
        '
        Me.txtBalance.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtBalance.Location = New System.Drawing.Point(669, 198)
        Me.txtBalance.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtBalance.Name = "txtBalance"
        Me.txtBalance.ReadOnly = True
        Me.txtBalance.Size = New System.Drawing.Size(284, 26)
        Me.txtBalance.TabIndex = 20
        '
        'dtpto
        '
        Me.dtpto.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dtpto.CustomFormat = "yyyy-MM-dd"
        Me.dtpto.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpto.Location = New System.Drawing.Point(784, 255)
        Me.dtpto.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtpto.Name = "dtpto"
        Me.dtpto.Size = New System.Drawing.Size(142, 26)
        Me.dtpto.TabIndex = 22
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(627, 232)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(54, 20)
        Me.Label9.TabIndex = 23
        Me.Label9.Text = "From :"
        '
        'Label10
        '
        Me.Label10.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(779, 231)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(35, 20)
        Me.Label10.TabIndex = 24
        Me.Label10.Text = "To :"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1047, 885)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.dtpto)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtBalance)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtWidthraw)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtRecieved)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnclear)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.dtpfrom)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cboType)
        Me.Controls.Add(Me.dtpTransDate)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtRemarks)
        Me.Controls.Add(Me.dtglist)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtMoney)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Budget Monitoring System"
        CType(Me.dtglist, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtMoney As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents btnSave As Button
    Friend WithEvents dtglist As DataGridView
    Friend WithEvents txtRemarks As RichTextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents dtpTransDate As DateTimePicker
    Friend WithEvents cboType As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents dtpfrom As DateTimePicker
    Friend WithEvents btnSearch As Button
    Friend WithEvents btnclear As Button
    Friend WithEvents btnDelete As Button
    Friend WithEvents Label6 As Label
    Friend WithEvents txtRecieved As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents txtWidthraw As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents txtBalance As TextBox
    Friend WithEvents dtpto As DateTimePicker
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
End Class
