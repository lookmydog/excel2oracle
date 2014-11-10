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
      MsgBox("UpdateDataSet Error.")
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
      MsgBox("ctxListBox_delete_Click Error.")
    End Try
  End Sub

  Public Sub DataGridViewAllClear(ByRef DataGridView1 As DataGridView)

    '很單純全部清空
    For Each dt As DataGridViewRow In DataGridView1.Rows
      For Each cell As DataGridViewCell In dt.Cells
        cell.Value = ""
      Next
    Next
  End Sub


End Module
