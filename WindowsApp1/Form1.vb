

Imports System.Data.SqlClient

Public Class Form1

    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Len(Me.TextBox1.Text) <> 0 Or Len(Me.TextBox2.Text) <> 0 Then
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "select password FROM user_data_table where username = '" + Me.TextBox1.Text + "';"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Dim password As String = ""
            Do While myReader.Read()
                password = myReader.GetString(0).Trim
            Loop
            myReader.Close()
            myConn.Close()
            If password <> "" Then
                If password = Me.TextBox2.Text Then
                    Module1.username = Me.TextBox1.Text
                    Dim frm As New Form9
                    frm.Show()
                    Me.Hide()
                Else
                    MessageBox.Show("Password is Wrong")
                End If
            Else
                MessageBox.Show("Username is Wrong")
            End If
        Else
            MessageBox.Show("Enter Username")
        End If
    End Sub

End Class