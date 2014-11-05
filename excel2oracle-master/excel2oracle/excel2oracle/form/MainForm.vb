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

  Private Sub btnOpenFile_Click(sender As System.Object, e As System.EventArgs) Handles btnOpenFile.Click
    '選取檔案
    If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
      strExcelName = OpenFileDialog1.FileName.Trim
      '將檔名列在 textbox 上
      txtPath.Text = OpenFileDialog1.InitialDirectory & strExcelName
      '建立新的 dataset
      dsExcel = New DataSet
      '清空listbox 上的 item
      ListBox1.Items.Clear()
      '讀取 excel
      LoadExcelData(strExcelName, OpenFileDialog1.FilterIndex)
      DataGridView1.DataSource = Nothing
    Else
      MsgBox("Open FileDialog Error.", MsgBoxStyle.Critical)
      Return
    End If
  End Sub

  Private Sub LoadExcelData(ByVal FileName As String, ByVal FileType As Integer)
    Try
      Dim fsFile As FileStream = New FileStream(FileName, FileMode.Open, FileAccess.Read)
      Dim wbXLS As IWorkbook = Nothing
      '1 is xls , 2 is xlsx , 選擇不同的 workbook
      If FileType = 1 Then
        wbXLS = New HSSFWorkbook(fsFile)
      ElseIf FileType = 2 Then
        wbXLS = New XSSFWorkbook(fsFile)
      End If

      '讀取有多少sheet，建立 listbox items and dsExcel tables
      For i As Integer = 0 To wbXLS.NumberOfSheets - 1
        ListBox1.Items.Add(wbXLS.GetSheetName(i))
        dsExcel.Tables.Add(LoadSheetData(wbXLS, wbXLS.GetSheetName(i)))
      Next

      '關閉檔案
      fsFile.Close()
      fsFile.Dispose()
    Catch ex As Exception
      MsgBox(ex.ToString)
    End Try

  End Sub

  '讀取 sheet data 
  Private Function LoadSheetData(ByRef wbXLS As IWorkbook, ByVal SheetName As String) As DataTable
    Dim dtSheetData As New DataTable(SheetName)
    Dim wbSheet As ISheet = wbXLS.GetSheet(SheetName)

    '沒有資料直接回傳 空的 dataTable
    If wbSheet.LastRowNum = 0 Then
      Return dtSheetData
    End If

    'get the max cell number
    Dim maxCellCount As Integer = 0
    For i As Integer = 0 To wbSheet.LastRowNum - 1
      If (Not wbSheet.GetRow(i) Is Nothing) AndAlso (wbSheet.GetRow(i).LastCellNum > maxCellCount) Then
        maxCellCount = wbSheet.GetRow(i).LastCellNum
      End If
    Next i

    'build datatable column , A.B.C.D...etc
    For i As Integer = 0 To maxCellCount - 1
      dtSheetData.Columns.Add(Convert10to26(i + 1), GetType(Object))
    Next

    Dim tmpAddRow As DataRow
    Try
      For i As Integer = 0 To wbSheet.LastRowNum
        tmpAddRow = dtSheetData.NewRow
        If wbSheet.GetRow(i) Is Nothing Then
          dtSheetData.Rows.Add(tmpAddRow)
          Continue For
        End If

        Dim tRow As IRow = wbSheet.GetRow(i)

        For Each tCell As ICell In tRow.Cells
          If tCell.CellType = CellType.Formula Then

            Dim iFormula As IFormulaEvaluator
            Dim obj As Object

            Try
              iFormula = WorkbookFactory.CreateFormulaEvaluator(wbXLS)
              obj = iFormula.Evaluate(tCell).CellType

              If obj = CellType.Numeric Then
                tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = tCell.NumericCellValue
              Else
                tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = tCell.ToString
              End If
            Catch
              tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = tCell.StringCellValue
            End Try

          Else

            tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = tCell.ToString
          End If

        Next

        dtSheetData.Rows.Add(tmpAddRow)
      Next
    Catch ex As Exception
      MsgBox(ex.ToString, MsgBoxStyle.Critical)
    End Try
    Return dtSheetData
  End Function

  Private Function Convert10to26(ByVal originalNum As Integer) As String
    If originalNum <= 0 Then Return ""
    Dim count As Integer = 0
    Dim strReturn As String = ""

    While originalNum > 0
      count = originalNum Mod 26
      If count = 0 Then
        count = 26
        originalNum = originalNum - 1
      End If
      strReturn = Convert.ToChar(64 + count) & strReturn
      originalNum = originalNum / 26

    End While
    Return strReturn
  End Function

  Private Sub MainForm_Load(sender As Object, e As System.EventArgs) Handles Me.Load
    '過濾附檔名,only xls , xlsx
    OpenFileDialog1.Filter = "Excel|*.xls|Excel2007|*.xlsx"
    '預設檔名名稱
    OpenFileDialog1.FileName = "test.xls"
  End Sub

  Private Sub ListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
    If ListBox1.SelectedIndex < 0 Then Return
    DataGridView1.DataSource = dsExcel.Tables(ListBox1.SelectedIndex)
    '當選擇sheet 的時候去 auto size 此sheet
    DataGridView1.AutoResizeColumns()
    DataGridView1.AutoResizeRows()
  End Sub

  '結束程式
  Private Sub btnEnd_Click(sender As System.Object, e As System.EventArgs) Handles btnEnd.Click
    dsExcel = Nothing
    End
  End Sub

End Class
