Public Class KeyValuePair
    Public m_key As Object
    Public m_value As String

    Public Sub New(ByVal newKey As Object, ByVal newValue As String)
        m_key = newKey
        m_value = newValue
    End Sub

    Public Overrides Function ToString() As String
        Return m_value
    End Function
End Class
