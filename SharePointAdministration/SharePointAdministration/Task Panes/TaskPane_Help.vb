Public Class TaskPane_Help

    Private Sub TaskPane_Help_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' This should come from config file / globals
        Me.RichTextBox1.LoadFile("C:\myrichtext.rtf")
    End Sub

    Private Sub btnGetContextAwareHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetContextAwareHelp.Click
        Dim row As Integer = Globals.ThisAddIn.Application.ActiveCell.Row
        Dim column As Integer = Globals.ThisAddIn.Application.ActiveCell.Column
        Dim ws As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        Me.TextBox1.Clear()
        Me.TextBox1.Text = "Karnac says you are on Row: " & row & _
            " Column: " & column & _
            " of WorkSheet: " & ws.Name & _
            " but says he doesn't know much else"
    End Sub

End Class
