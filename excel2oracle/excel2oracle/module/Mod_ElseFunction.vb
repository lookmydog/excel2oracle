'------------------imports Item--------------
Imports System.IO
Imports System.Data
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
'---------------------------------------------

Module Mod_ElseFunction

  Public Function Convert10to26(ByVal originalNum As Integer) As String
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

  Public Sub UpdateDataSet(ByVal index As Integer, ByRef dsExcel As DataSet)

    Try
      dsExcel.Tables.RemoveAt(index)
    Catch ex As Exception
      MsgBox(MessageFormat(ex.ToString, "UpdateDataSet"))
    End Try
  End Sub

  Public Sub DeleteListItems(ByRef ListBox1 As ListBox, ByRef dsExcel As DataSet)
    If ListBox1.SelectedItems.Count < 1 Then Return
    Try
      For i As Integer = ListBox1.SelectedIndices.Count - 1 To 0 Step -1
        Dim Itemp As Integer = ListBox1.SelectedIndices(i)
        ListBox1.Items.RemoveAt(Itemp)
        UpdateDataSet(Itemp, dsExcel)
      Next
    Catch ex As Exception
      MsgBox(MessageFormat(ex.ToString, "DeleteListItems"))
    End Try
  End Sub

  Public Sub DataGridViewAllClear(ByRef DataGridView1 As DataGridView)
    If DataGridView1 Is Nothing Then Return
    '很單純全部清空
    For Each dt As DataGridViewRow In DataGridView1.Rows
      For Each cell As DataGridViewCell In dt.Cells
        cell.Value = ""
      Next
    Next
  End Sub

  Public Sub LoadExcelData(ByVal FileName As String, ByVal FileType As Integer, ByRef ListBox1 As ListBox, ByRef dsExcel As DataSet)
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
      MsgBox(MessageFormat(ex.ToString, "LoadExcelData"))
    End Try

  End Sub

  '讀取 sheet data 
  Public Function LoadSheetData(ByRef wbXLS As IWorkbook, ByVal SheetName As String) As DataTable
    Dim dtSheetData As New DataTable(SheetName)
    Dim wbSheet As ISheet = wbXLS.GetSheet(SheetName)

    '沒有資料直接回傳 空的 dataTable
    If wbSheet.LastRowNum = 0 Then
      Return dtSheetData
    End If

    'get the max cell number
    Dim maxCellCount As Integer = 0
    For i As Integer = 0 To wbSheet.LastRowNum
      If (Not wbSheet.GetRow(i) Is Nothing) AndAlso (wbSheet.GetRow(i).LastCellNum > maxCellCount) Then
        'Dim nullValue As Boolean = False
        'For Each cells As ICell In wbSheet.GetRow(i).Cells
        '  If Not cells Is Nothing Then
        '    cells.SetCellType(CellType.String)
        '    If cells.ToString <> "" Then
        '      nullValue = True
        '      'Exit For
        '    End If
        '  Else
        '    Continue For
        '  End If
        'Next

        'If nullValue Then maxCellCount = wbSheet.GetRow(i).LastCellNum
        maxCellCount = wbSheet.GetRow(i).LastCellNum
      End If
    Next

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



        '取sheet 的 row
        Dim tRow As IRow = wbSheet.GetRow(i)

        '判別cell 是什麼格式
        For Each tCell As ICell In tRow.Cells

          Select Case tCell.CellType
            Case CellType.Formula

              Dim iFormula As IFormulaEvaluator
              Dim obj As Object

              Try
                'interface evaluate formula in cell
                iFormula = WorkbookFactory.CreateFormulaEvaluator(wbXLS)
                obj = iFormula.Evaluate(tCell).CellType

                If obj = CellType.Numeric Then
                  tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = tCell.NumericCellValue
                Else
                  tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = tCell.ToString
                End If
              Catch
                'the special formula : "W" & formula => string value
                tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = tCell.StringCellValue
              End Try
            Case CellType.Numeric
              tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = tCell.NumericCellValue
            Case CellType.String
              '不是公式直接輸出字串
              tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = tCell.ToString
            Case CellType.Error
              tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = "[Error]"
            Case CellType.Blank
              tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = "[Blank]"
            Case CellType.Unknown
              tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = "[Unknown]"
            Case Else
              tmpAddRow(Convert10to26(tCell.ColumnIndex + 1)) = "[Else]"
          End Select

        Next
        'datatable 新增 row
        dtSheetData.Rows.Add(tmpAddRow)
      Next
    Catch ex As Exception
      MsgBox(MessageFormat(ex.ToString, "LoadSheetData"))
    End Try

    Return dtSheetData
  End Function

  Public Function MessageFormat(ByVal ex As String, ByVal funName As String) As String
    Dim strResult As String = ""
    strResult = String.Format("<{0}> Error." & vbNewLine & vbNewLine & "{1}", funName, ex)
    Return strResult
  End Function

End Module
