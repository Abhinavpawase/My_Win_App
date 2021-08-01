

Public Class Form1

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Len(Me.TextBox1.Text) <> 0 Or Len(Me.TextBox2.Text) <> 0 Then
            Module1.username = Me.TextBox1.Text
            Dim frm As New Form9
            frm.Show()
            Me.Hide()
        Else
            MessageBox.Show("Enter Username")
        End If
    End Sub

End Class