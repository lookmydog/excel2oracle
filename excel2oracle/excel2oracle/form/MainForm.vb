'------------------imports Item--------------
Imports System.IO
Imports System.Data
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
'---------------------------------------------

Public Class MainForm

  '----------------class variable ------------
  Private strExcelName As String = ""
  Private dsExcel As DataSet
  '------------end class variable -------------

  Private Sub MainForm_Load(sender As Object, e As System.EventArgs) Handles Me.Load

    '過濾附檔名,only xls , xlsx
    OpenFileDialog1.Filter = "Excel|*.xls|Excel 2007|*.xlsx"
    '預設檔名名稱
    OpenFileDialog1.FileName = "test.xls"

    '把DataGridView 的新增row 關閉
    DataGridView1.AllowUserToAddRows = False

    '指定ContextMenuStrip 給 listbox1
    ListBox1.ContextMenuStrip = rightClickMenuListBox

  End Sub

  Private Sub ListBox1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles ListBox1.KeyDown
    Select Case e.KeyCode
      Case Keys.Delete
        DeleteListItems(Me.ListBox1, dsExcel)
    End Select
  End Sub

  Private Sub ListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
    If ListBox1.SelectedIndex < 0 Then Return
    DataGridView1.DataSource = dsExcel.Tables(ListBox1.SelectedIndex)

    '當選擇sheet 的時候去 auto size 此sheet
    DataGridView1.AutoResizeColumns()
    ' DataGridView1.AutoResizeRows()

  End Sub

  '結束程式
  Private Sub menuItemEndProgram_Click(sender As System.Object, e As System.EventArgs) Handles menuItemEndProgram.Click
    End
  End Sub

  Private Sub menuItemOpenFile_Click(sender As System.Object, e As System.EventArgs) Handles menuItemOpenFile.Click
    '選取檔案
    If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
      strExcelName = OpenFileDialog1.FileName.Trim
      '將檔名列在 status 上
      statusLblFilePath.Text = OpenFileDialog1.InitialDirectory & strExcelName
      '建立新的 dataset
      dsExcel = New DataSet
      '清空listbox 上的 item
      ListBox1.Items.Clear()
      '讀取 excel
      LoadExcelData(strExcelName, OpenFileDialog1.FilterIndex, Me.ListBox1, dsExcel)
      DataGridView1.DataSource = Nothing
    ElseIf OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.Cancel Then
    Else
      MsgBox(MessageFormat("", "menuItemOpenFile_Click"))
      Return
    End If
  End Sub

  '顯示EXCEL 檔案路徑
  Private Sub menuItemShowFilePath_Click(sender As System.Object, e As System.EventArgs) Handles menuItemShowFilePath.Click
    Dim frmShowPath As New frmShowFilePath(OpenFileDialog1.InitialDirectory & strExcelName)
    frmShowPath.Show()
  End Sub

  'release source
  Protected Overrides Sub Finalize()
    dsExcel = Nothing
    MyBase.Finalize()
  End Sub

  '選擇哪種contextMenuStrip (按右鍵出現的選單)
  Private Sub DataGridView1_CellContextMenuStripNeeded(sender As Object, e As System.Windows.Forms.DataGridViewCellContextMenuStripNeededEventArgs) Handles DataGridView1.CellContextMenuStripNeeded
    Try
      Dim dgv As DataGridView = DirectCast(sender, DataGridView)

      If e.ColumnIndex < 0 And e.RowIndex < 0 Then
        e.ContextMenuStrip = Me.rightClickMenuDV_All
      ElseIf e.RowIndex < 0 Then
        e.ContextMenuStrip = Me.rightClickMenuDV_Col
      ElseIf e.ColumnIndex < 0 Then
        e.ContextMenuStrip = Me.rigthClickMenuDV_Row
      ElseIf dgv(e.ColumnIndex, e.RowIndex).Value <> Nothing Then
        e.ContextMenuStrip = Me.rightClickMenuDV_Cell
      End If
    Catch ex As Exception
    End Try
  End Sub

  '將DataGridView table 整個刪掉(初始化)
  '2014/11/10 感覺沒什麼意義,未來考慮刪除功能
  Private Sub ctxMenu_All_Delete_Click(sender As System.Object, e As System.EventArgs) Handles ctxMenu_All_Delete.Click
    Try
      DataGridView1.DataSource = Nothing
      dsExcel.Tables(ListBox1.SelectedIndex).BeginInit()

      '需求
      '欄位的位數還是固定，不是很符合要求，想要改變成能修改（新增或是刪除）
    Catch ex As Exception
      MsgBox(MessageFormat(ex.ToString, "ctxMenu_All_Delete_Click"))
    End Try
  End Sub

  '清空 DataGridView 內容
  Private Sub ctxMenu_All_Clear_Click(sender As System.Object, e As System.EventArgs) Handles ctxMenu_All_Clear.Click
    DataGridViewAllClear(Me.DataGridView1)
  End Sub

  '刪除listbox items
  Private Sub ctxListBox_delete_Click(sender As System.Object, e As System.EventArgs) Handles ctxListBox_delete.Click
    DeleteListItems(Me.ListBox1, dsExcel)
  End Sub

  Private Sub ctxMenu_Row_DeleteChoose_Click(sender As System.Object, e As System.EventArgs) Handles ctxMenu_Row_DeleteChoose.Click
    For Each row As DataGridViewRow In DataGridView1.SelectedRows
      If Not row.IsNewRow Then
        DataGridView1.Rows.Remove(row)
      End If
    Next
  End Sub

  'show position 
  Private Sub DataGridView1_CellEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
    Me.lblCellPos.Text = String.Format("({0} , {1})", e.RowIndex + 1, e.ColumnIndex + 1)
  End Sub
End Class
