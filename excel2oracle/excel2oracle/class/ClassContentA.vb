Public Class ClassContentA

  Private str As String
  'Private comment As String

  Public Sub New()
    str = ""
    'comment = ""
  End Sub

  Public Property content() As String
    Get
      Return Me.str
    End Get
    Set(value As String)
      Me.str = value
    End Set
  End Property

End Class
