
Imports System.Data.SqlClient

Public Class Form8
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TextBox4.Text = ""
        Me.ComboBox1.SelectedIndex = 0
    End Sub
    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim category1 As String = Me.ComboBox1.SelectedItem.ToString
        Dim code1 As String = ""
        myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
        myConn.Open()
        myCmd = myConn.CreateCommand
        Dim str As String
        str = "SELECT code FROM Category_Table where Category = '" + category1 + "' ;"
        myCmd.CommandText = str
        myReader = myCmd.ExecuteReader()
        Do While myReader.Read()
            code1 = myReader.GetString(0).Trim
        Loop
        myReader.Close()
        str = "select part from Inventory_Table_1 where part like '" + code1 + "%';;"
        myCmd.CommandText = str
        myReader = myCmd.ExecuteReader()
        Dim Next_part As Integer = 0
        Do While myReader.Read()
            Dim inst1 As Integer = Mid(myReader.GetString(0), 3, 2)
            If inst1 > Next_part Then
                Next_part = inst1
            End If
        Loop
        myReader.Close()
        myConn.Close()
        Next_part = Next_part + 1
        If Next_part < 10 Then
            Me.TextBox4.Text = "" & code1 & "0" & Next_part
        Else
            Me.TextBox4.Text = "" & code1 & "" & Next_part
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Dispose()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not (String.IsNullOrEmpty(Me.TextBox1.Text)) And Not (String.IsNullOrEmpty(Me.TextBox2.Text)) And Not (String.IsNullOrEmpty(Me.TextBox4.Text)) Then
            If IsNumeric(Me.TextBox2.Text) Then
                myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
                myConn.Open()
                myCmd = myConn.CreateCommand
                Dim str As String
                str = "insert into Inventory_Table_1 (Part,Revision,Quantity,description) values ('" + Me.TextBox4.Text + "','A','" + Me.TextBox2.Text + "','" + Me.TextBox1.Text + "') ;"
                myCmd.CommandText = str
                myReader = myCmd.ExecuteReader()
                myReader.Close()
                myConn.Close()
                Me.TextBox1.Text = ""
                Me.TextBox2.Text = ""
                Me.TextBox4.Text = ""
                MessageBox.Show("Part Number Registered")
            Else
                MessageBox.Show("Enter Number in Quantity")
            End If
        Else
            MessageBox.Show("Enter All Details")
        End If
    End Sub

End Class