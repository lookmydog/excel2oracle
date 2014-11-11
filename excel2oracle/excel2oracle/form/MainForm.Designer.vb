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
    Me.components = New System.ComponentModel.Container()
    Me.DataGridView1 = New System.Windows.Forms.DataGridView()
    Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
    Me.ListBox1 = New System.Windows.Forms.ListBox()
    Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
    Me.檔案ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.menuItemOpenFile = New System.Windows.Forms.ToolStripMenuItem()
    Me.menuItemShowFilePath = New System.Windows.Forms.ToolStripMenuItem()
    Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
    Me.menuItemEndProgram = New System.Windows.Forms.ToolStripMenuItem()
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.statusLblFilePath = New System.Windows.Forms.ToolStripStatusLabel()
    Me.rightClickMenuDV_All = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
    Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
    Me.ctxMenu_All_Delete = New System.Windows.Forms.ToolStripMenuItem()
    Me.ctxMenu_All_Clear = New System.Windows.Forms.ToolStripMenuItem()
    Me.rigthClickMenuDV_Row = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.TestRowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.rightClickMenuDV_Col = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.TestColToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.rightClickMenuDV_Cell = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.TestCellToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
    Me.rightClickMenuListBox = New System.Windows.Forms.ContextMenuStrip(Me.components)
    Me.ctxListBox_delete = New System.Windows.Forms.ToolStripMenuItem()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.lblCellPos = New System.Windows.Forms.Label()
    Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
    Me.ctxMenu_Row_DeleteChoose = New System.Windows.Forms.ToolStripMenuItem()
    CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.MenuStrip1.SuspendLayout()
    Me.StatusStrip1.SuspendLayout()
    Me.rightClickMenuDV_All.SuspendLayout()
    Me.rigthClickMenuDV_Row.SuspendLayout()
    Me.rightClickMenuDV_Col.SuspendLayout()
    Me.rightClickMenuDV_Cell.SuspendLayout()
    Me.rightClickMenuListBox.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
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
    Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
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
    'menuItemShowFilePath
    '
    Me.menuItemShowFilePath.Name = "menuItemShowFilePath"
    Me.menuItemShowFilePath.Size = New System.Drawing.Size(152, 22)
    Me.menuItemShowFilePath.Text = "顯示檔案路徑"
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
    'rightClickMenuDV_All
    '
    Me.rightClickMenuDV_All.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.ToolStripSeparator2, Me.ctxMenu_All_Delete, Me.ctxMenu_All_Clear})
    Me.rightClickMenuDV_All.Name = "rightClickMenuDV_All"
    Me.rightClickMenuDV_All.Size = New System.Drawing.Size(125, 76)
    '
    'ToolStripMenuItem1
    '
    Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
    Me.ToolStripMenuItem1.Size = New System.Drawing.Size(124, 22)
    Me.ToolStripMenuItem1.Text = "Test_All"
    '
    'ToolStripSeparator2
    '
    Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
    Me.ToolStripSeparator2.Size = New System.Drawing.Size(121, 6)
    '
    'ctxMenu_All_Delete
    '
    Me.ctxMenu_All_Delete.Name = "ctxMenu_All_Delete"
    Me.ctxMenu_All_Delete.Size = New System.Drawing.Size(124, 22)
    Me.ctxMenu_All_Delete.Text = "全部刪除"
    '
    'ctxMenu_All_Clear
    '
    Me.ctxMenu_All_Clear.Name = "ctxMenu_All_Clear"
    Me.ctxMenu_All_Clear.Size = New System.Drawing.Size(124, 22)
    Me.ctxMenu_All_Clear.Text = "全部清空"
    '
    'rigthClickMenuDV_Row
    '
    Me.rigthClickMenuDV_Row.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TestRowToolStripMenuItem, Me.ToolStripSeparator3, Me.ctxMenu_Row_DeleteChoose})
    Me.rigthClickMenuDV_Row.Name = "rigthClickMenuDV_Row"
    Me.rigthClickMenuDV_Row.Size = New System.Drawing.Size(153, 76)
    '
    'TestRowToolStripMenuItem
    '
    Me.TestRowToolStripMenuItem.Name = "TestRowToolStripMenuItem"
    Me.TestRowToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
    Me.TestRowToolStripMenuItem.Text = "Test_Row"
    '
    'rightClickMenuDV_Col
    '
    Me.rightClickMenuDV_Col.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TestColToolStripMenuItem})
    Me.rightClickMenuDV_Col.Name = "rightClickMenuDV_Col"
    Me.rightClickMenuDV_Col.Size = New System.Drawing.Size(124, 26)
    '
    'TestColToolStripMenuItem
    '
    Me.TestColToolStripMenuItem.Name = "TestColToolStripMenuItem"
    Me.TestColToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
    Me.TestColToolStripMenuItem.Text = "Test_Col"
    '
    'rightClickMenuDV_Cell
    '
    Me.rightClickMenuDV_Cell.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TestCellToolStripMenuItem})
    Me.rightClickMenuDV_Cell.Name = "rightClickMenuDV_Cell"
    Me.rightClickMenuDV_Cell.Size = New System.Drawing.Size(126, 26)
    '
    'TestCellToolStripMenuItem
    '
    Me.TestCellToolStripMenuItem.Name = "TestCellToolStripMenuItem"
    Me.TestCellToolStripMenuItem.Size = New System.Drawing.Size(125, 22)
    Me.TestCellToolStripMenuItem.Text = "Test_Cell"
    '
    'rightClickMenuListBox
    '
    Me.rightClickMenuListBox.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ctxListBox_delete})
    Me.rightClickMenuListBox.Name = "rightClickMenuListBox"
    Me.rightClickMenuListBox.Size = New System.Drawing.Size(101, 26)
    '
    'ctxListBox_delete
    '
    Me.ctxListBox_delete.Name = "ctxListBox_delete"
    Me.ctxListBox_delete.Size = New System.Drawing.Size(100, 22)
    Me.ctxListBox_delete.Text = "刪除"
    '
    'GroupBox1
    '
    Me.GroupBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox1.Controls.Add(Me.lblCellPos)
    Me.GroupBox1.Location = New System.Drawing.Point(590, 28)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(129, 45)
    Me.GroupBox1.TabIndex = 8
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Cell Position"
    '
    'lblCellPos
    '
    Me.lblCellPos.AutoSize = True
    Me.lblCellPos.Location = New System.Drawing.Point(7, 22)
    Me.lblCellPos.Name = "lblCellPos"
    Me.lblCellPos.Size = New System.Drawing.Size(0, 12)
    Me.lblCellPos.TabIndex = 0
    '
    'ToolStripSeparator3
    '
    Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
    Me.ToolStripSeparator3.Size = New System.Drawing.Size(149, 6)
    '
    'ctxMenu_Row_DeleteChoose
    '
    Me.ctxMenu_Row_DeleteChoose.Name = "ctxMenu_Row_DeleteChoose"
    Me.ctxMenu_Row_DeleteChoose.Size = New System.Drawing.Size(152, 22)
    Me.ctxMenu_Row_DeleteChoose.Text = "刪除所選列"
    '
    'MainForm
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(731, 431)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.ListBox1)
    Me.Controls.Add(Me.DataGridView1)
    Me.Controls.Add(Me.MenuStrip1)
    Me.MainMenuStrip = Me.MenuStrip1
    Me.Name = "MainForm"
    Me.Text = "Excel2Oracle"
    CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.MenuStrip1.ResumeLayout(False)
    Me.MenuStrip1.PerformLayout()
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.rightClickMenuDV_All.ResumeLayout(False)
    Me.rigthClickMenuDV_Row.ResumeLayout(False)
    Me.rightClickMenuDV_Col.ResumeLayout(False)
    Me.rightClickMenuDV_Cell.ResumeLayout(False)
    Me.rightClickMenuListBox.ResumeLayout(False)
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
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
  Friend WithEvents rightClickMenuDV_All As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents rigthClickMenuDV_Row As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents TestRowToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents rightClickMenuDV_Col As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents rightClickMenuDV_Cell As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents TestColToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents TestCellToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ctxMenu_All_Delete As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents ctxMenu_All_Clear As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents rightClickMenuListBox As System.Windows.Forms.ContextMenuStrip
  Friend WithEvents ctxListBox_delete As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents lblCellPos As System.Windows.Forms.Label
  Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents ctxMenu_Row_DeleteChoose As System.Windows.Forms.ToolStripMenuItem

End Class
