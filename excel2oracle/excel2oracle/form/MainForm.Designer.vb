<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
  Inherits System.Windows.Forms.Form

  'Form 覆寫 Dispose 以清除元件清單。
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

  '為 Windows Form 設計工具的必要項
  Private components As System.ComponentModel.IContainer

  '注意: 以下為 Windows Form 設計工具所需的程序
  '可以使用 Windows Form 設計工具進行修改。
  '請不要使用程式碼編輯器進行修改。
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.DataGridView1 = New System.Windows.Forms.DataGridView()
    Me.btnOpenFile = New System.Windows.Forms.Button()
    Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
    Me.ListBox1 = New System.Windows.Forms.ListBox()
    Me.btnEnd = New System.Windows.Forms.Button()
    Me.txtPath = New System.Windows.Forms.TextBox()
    CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'DataGridView1
    '
    Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.DataGridView1.Location = New System.Drawing.Point(114, 79)
    Me.DataGridView1.Name = "DataGridView1"
    Me.DataGridView1.RowTemplate.Height = 24
    Me.DataGridView1.Size = New System.Drawing.Size(605, 340)
    Me.DataGridView1.TabIndex = 0
    '
    'btnOpenFile
    '
    Me.btnOpenFile.Location = New System.Drawing.Point(12, 22)
    Me.btnOpenFile.Name = "btnOpenFile"
    Me.btnOpenFile.Size = New System.Drawing.Size(75, 23)
    Me.btnOpenFile.TabIndex = 1
    Me.btnOpenFile.Text = "開啟檔案"
    Me.btnOpenFile.UseVisualStyleBackColor = True
    '
    'OpenFileDialog1
    '
    Me.OpenFileDialog1.FileName = "OpenFileDialog1"
    '
    'ListBox1
    '
    Me.ListBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.ListBox1.FormattingEnabled = True
    Me.ListBox1.ItemHeight = 12
    Me.ListBox1.Location = New System.Drawing.Point(12, 79)
    Me.ListBox1.Name = "ListBox1"
    Me.ListBox1.Size = New System.Drawing.Size(96, 340)
    Me.ListBox1.TabIndex = 3
    '
    'btnEnd
    '
    Me.btnEnd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.btnEnd.Location = New System.Drawing.Point(644, 22)
    Me.btnEnd.Name = "btnEnd"
    Me.btnEnd.Size = New System.Drawing.Size(75, 23)
    Me.btnEnd.TabIndex = 4
    Me.btnEnd.Text = "結束程式"
    Me.btnEnd.UseVisualStyleBackColor = True
    '
    'txtPath
    '
    Me.txtPath.Location = New System.Drawing.Point(114, 22)
    Me.txtPath.Name = "txtPath"
    Me.txtPath.Size = New System.Drawing.Size(487, 22)
    Me.txtPath.TabIndex = 5
    '
    'MainForm
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(731, 431)
    Me.Controls.Add(Me.txtPath)
    Me.Controls.Add(Me.btnEnd)
    Me.Controls.Add(Me.ListBox1)
    Me.Controls.Add(Me.btnOpenFile)
    Me.Controls.Add(Me.DataGridView1)
    Me.Name = "MainForm"
    Me.Text = "NPOI 存取範例"
    CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
  Friend WithEvents btnOpenFile As System.Windows.Forms.Button
  Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
  Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
  Friend WithEvents btnEnd As System.Windows.Forms.Button
  Friend WithEvents txtPath As System.Windows.Forms.TextBox

End Class
