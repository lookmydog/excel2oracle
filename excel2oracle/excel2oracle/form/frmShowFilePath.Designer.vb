<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmShowFilePath
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
    Me.txtFilePath = New System.Windows.Forms.TextBox()
    Me.SuspendLayout()
    '
    'txtFilePath
    '
    Me.txtFilePath.Location = New System.Drawing.Point(12, 12)
    Me.txtFilePath.Multiline = True
    Me.txtFilePath.Name = "txtFilePath"
    Me.txtFilePath.Size = New System.Drawing.Size(260, 238)
    Me.txtFilePath.TabIndex = 0
    '
    'frmShowFilePath
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(284, 262)
    Me.Controls.Add(Me.txtFilePath)
    Me.MaximizeBox = False
    Me.Name = "frmShowFilePath"
    Me.Text = "檔案路徑"
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
End Class
