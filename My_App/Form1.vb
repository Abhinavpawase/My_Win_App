

Imports System.Windows.Forms


Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim cls As New string_comp
        'Dim bool1 As Boolean
        'bool1 = cls.string_capital2("heloo")
        'MsgBox(bool1)
    End Sub
    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Me.Dispose()
        Application.Exit()
    End Sub

End Class