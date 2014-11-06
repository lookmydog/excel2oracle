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
    Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
    Me.ListBox1 = New System.Windows.Forms.ListBox()
    Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
    Me.檔案ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.menuItemOpenFile = New System.Windows.Forms.ToolStripMenuItem()
    Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
    Me.menuItemEndProgram = New System.Windows.Forms.ToolStripMenuItem()
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.statusLblFilePath = New System.Windows.Forms.ToolStripStatusLabel()
    Me.menuItemShowFilePath = New System.Windows.Forms.ToolStripMenuItem()
    CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.MenuStrip1.SuspendLayout()
    Me.StatusStrip1.SuspendLayout()
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
    Me.DataGridView1.Size = New System.Drawing.Size(605, 328)
    Me.DataGridView1.TabIndex = 0
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
    Me.ListBox1.Size = New System.Drawing.Size(96, 328)
    Me.ListBox1.TabIndex = 3
    '
    'MenuStrip1
    '
    Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.檔案ToolStripMenuItem})
    Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
    Me.MenuStrip1.Name = "MenuStrip1"
    Me.MenuStrip1.Size = New System.Drawing.Size(731, 24)
    Me.MenuStrip1.TabIndex = 6
    Me.MenuStrip1.Text = "MenuStrip1"
    '
    '檔案ToolStripMenuItem
    '
    Me.檔案ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuItemOpenFile, Me.menuItemShowFilePath, Me.ToolStripSeparator1, Me.menuItemEndProgram})
    Me.檔案ToolStripMenuItem.Name = "檔案ToolStripMenuItem"
    Me.檔案ToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
    Me.檔案ToolStripMenuItem.Text = "檔案"
    '
    'menuItemOpenFile
    '
    Me.menuItemOpenFile.Name = "menuItemOpenFile"
    Me.menuItemOpenFile.Size = New System.Drawing.Size(152, 22)
    Me.menuItemOpenFile.Text = "開啟檔案"
    '
    'ToolStripSeparator1
    '
    Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
    Me.ToolStripSeparator1.Size = New System.Drawing.Size(149, 6)
    '
    'menuItemEndProgram
    '
    Me.menuItemEndProgram.Name = "menuItemEndProgram"
    Me.menuItemEndProgram.Size = New System.Drawing.Size(152, 22)
    Me.menuItemEndProgram.Text = "結束程式"
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statusLblFilePath})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 409)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(731, 22)
    Me.StatusStrip1.TabIndex = 7
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'statusLblFilePath
    '
    Me.statusLblFilePath.Name = "statusLblFilePath"
    Me.statusLblFilePath.Size = New System.Drawing.Size(0, 17)
    '
    'menuItemShowFilePath
    '
    Me.menuItemShowFilePath.Name = "menuItemShowFilePath"
    Me.menuItemShowFilePath.Size = New System.Drawing.Size(152, 22)
    Me.menuItemShowFilePath.Text = "顯示檔案路徑"
    '
    'MainForm
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(731, 431)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.ListBox1)
    Me.Controls.Add(Me.DataGridView1)
    Me.Controls.Add(Me.MenuStrip1)
    Me.MainMenuStrip = Me.MenuStrip1
    Me.Name = "MainForm"
    Me.Text = "NPOI 存取範例"
    CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.MenuStrip1.ResumeLayout(False)
    Me.MenuStrip1.PerformLayout()
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
  Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
  Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
  Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
  Friend WithEvents 檔案ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents menuItemOpenFile As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents menuItemEndProgram As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents statusLblFilePath As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents menuItemShowFilePath As System.Windows.Forms.ToolStripMenuItem

End Class
