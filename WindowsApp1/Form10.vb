
Imports System.Data.SqlClient

Public Class Form10
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader

    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Size = New Size(439, 376)
        Me.Panel2.Location = New Point(12, 12)
        Me.Panel2.BringToFront()
        Call load_users()
    End Sub
    Private Sub load_users()
        myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
        myConn.Open()
        myCmd = myConn.CreateCommand
        Dim str As String
        str = "SELECT username FROM User_Data_Table ;"
        myCmd.CommandText = str
        myReader = myCmd.ExecuteReader()
        Do While myReader.Read()
            Me.ComboBox2.Items.Add(myReader.GetString(0).Trim)
        Loop
        myReader.Close()
        myConn.Close()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.TextBox8.Text = "Abhinav" And Me.TextBox7.Text = "Pawase@123" Then
            Me.Panel2.Location = New Point(438, 12)
            Me.Panel1.BringToFront()
        End If
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call reg_user()
    End Sub
    Private Sub reg_user()
        If Not (String.IsNullOrEmpty(Me.TextBox1.Text)) _
          And Not (String.IsNullOrEmpty(Me.TextBox2.Text)) _
            And Not (String.IsNullOrEmpty(Me.TextBox3.Text)) _
              And Not (String.IsNullOrEmpty(Me.TextBox4.Text)) _
                And Not Convert.ToInt32(ComboBox2.SelectedIndex) <> -1 Then
            If Me.TextBox2.Text = Me.TextBox3.Text Then
                myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
                myConn.Open()
                myCmd = myConn.CreateCommand
                Dim str As String
                str = "Insert into User_Data_Table (username,password,Email,department) values ('" + Me.TextBox1.Text + "','" + Me.TextBox2.Text + "','" + Me.TextBox4.Text + "','" + Me.ComboBox1.SelectedItem.ToString + "') ;"
                myCmd.CommandText = str
                myReader = myCmd.ExecuteReader()
                myReader.Close()
                myConn.Close()
                MessageBox.Show("User Registered")
                Me.ComboBox2.Items.Add(Me.TextBox1.Text)
                Me.TextBox1.Text = ""
                Me.TextBox2.Text = ""
                Me.TextBox3.Text = ""
                Me.TextBox4.Text = ""
            Else
                MessageBox.Show("Password Mismatched")
            End If
        Else
            MessageBox.Show("Enter All Fields")
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Convert.ToInt32(ComboBox2.SelectedIndex) <> -1 Then
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "delete from User_Data_Table where username = '" + Me.ComboBox2.SelectedItem.ToString + "' ;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            myReader.Close()
            myConn.Close()
            Me.ComboBox2.Items.Remove(Me.ComboBox2.SelectedItem.ToString)
            MessageBox.Show("User Removed")
        Else
            MessageBox.Show("Select User")
        End If
    End Sub


End Class