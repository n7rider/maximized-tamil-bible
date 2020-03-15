Public Class FormDisplay



    Private Sub FormDisplay_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        RichTextBox1.Width = Me.Width - 25
        RichTextBox1.Height = Me.Height - 50
    End Sub
End Class