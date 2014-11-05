Imports System.Data.OleDb
'Imports Microsoft.Office.Interop.Excel

Public Class Form1

  '定義OLE====================================================== 
  '1.檔案位置 
  Private Const FileName As String = "D:\test.xls"
  '2.提供者名稱 
  Private Const ProviderName As String = "Microsoft.Jet.OLEDB.4.0;"
  '3.Excel版本，Excel 8.0 針對Excel2000及以上版本，Excel5.0 針對Excel97。 
  Private Const ExtendedString As String = "'Excel 8.0;"
  '4.第一行是否為標題 
  Private Const Hdr As String = "Yes;"
  '5.IMEX=1 通知驅動程序始終將「互混」數據列作為文本讀取 
  Private Const IMEX As String = "0';"
  '============================================================= 

  '連線字串 
  Private cs As String = (((("Data Source=" & FileName & ";" & "Provider=") + ProviderName & "Extended Properties=") + ExtendedString & "HDR=") + Hdr & "IMEX=") + IMEX


  Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
    Dim sheetName As String = "Cover"

    Dim myConn As OleDb.OleDbConnection = New OleDb.OleDbConnection(cs)
    Dim strCom As String = "SELECT * FROM [" & sheetName & "$]"
    myConn.Open()
    Dim myCommand As OleDbDataAdapter = New OleDbDataAdapter(strCom, myConn)
    Dim myDataSet As New DataSet
    myCommand.Fill(myDataSet, "[" & sheetName & "$]")
    myConn.Close()

    DataGridView1.DataMember = "[" & sheetName & "$]"
    DataGridView1.DataSource = myDataSet
  End Sub

End Class
