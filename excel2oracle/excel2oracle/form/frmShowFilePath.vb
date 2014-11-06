Public Class frmShowFilePath
  Public Sub New()
    InitializeComponent()
  End Sub
  Public Sub New(ByVal FilePath As String)
    InitializeComponent()
    txtFilePath.Text = FilePath
  End Sub
End Class