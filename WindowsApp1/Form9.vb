

Public Class Form9

    Dim frm2 As Form2 = Nothing
    Dim frm3 As Form3 = Nothing
    Dim frm5 As Form5 = Nothing
    Dim frm6 As Form6 = Nothing
    Dim frm7 As Form7 = Nothing

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load




    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If frm2 Is Nothing Then
            frm2 = New Form2
            frm2.frm9 = Me
            frm2.Show()
            AddHandler frm2.FormClosed, AddressOf frm2_close
        Else
            frm2.TopMost = True
            frm2.TopMost = False
        End If
    End Sub
    Private Sub frm2_close()
        frm2 = Nothing
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If frm3 Is Nothing Then
            frm3 = New Form3
            frm3.frm9 = Me
            frm3.Show()

            AddHandler frm3.FormClosed, AddressOf frm3_close
        Else
            frm3.TopMost = True
            frm3.TopMost = False
        End If
    End Sub
    Private Sub frm3_close()
        frm3 = Nothing
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If frm5 Is Nothing Then
            frm5 = New Form5
            frm5.frm9 = Me
            frm5.Show()

            AddHandler frm5.FormClosed, AddressOf frm5_close
        Else
            frm5.TopMost = True
            frm5.TopMost = False
        End If
    End Sub
    Private Sub frm5_close()
        frm5 = Nothing
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If frm6 Is Nothing Then
            frm6 = New Form6
            frm6.frm9 = Me
            frm6.Show()

            AddHandler frm6.FormClosed, AddressOf frm6_close
        Else
            frm6.TopMost = True
            frm6.TopMost = False
        End If
    End Sub
    Private Sub frm6_close()
        frm6 = Nothing
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If frm7 Is Nothing Then
            frm7 = New Form7
            frm7.frm9 = Me
            frm7.Show()

            AddHandler frm7.FormClosed, AddressOf frm7_close
        Else
            frm7.TopMost = True
            frm7.TopMost = False
        End If
    End Sub
    Private Sub frm7_close()
        frm7 = Nothing
    End Sub
    Dim create_no_frm As Form8
    Private Sub CreatePartNumber_Click(sender As Object, e As EventArgs) Handles CreatePartNumber.Click
        If create_no_frm Is Nothing Then
            create_no_frm = New Form8
            AddHandler create_no_frm.FormClosed, AddressOf create_no_frm_close
            create_no_frm.Show()
        End If
    End Sub
    Private Sub create_no_frm_close()
        create_no_frm = Nothing
    End Sub
    Private Sub Form9_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Call close_app()
    End Sub
    Private Sub close_app()
        Me.Hide()
        If Not frm2 Is Nothing Then
            frm2.Hide()
        End If
        If Not frm3 Is Nothing Then
            frm3.Hide()
        End If
        If Not frm5 Is Nothing Then
            frm5.Hide()
        End If
        If Not frm6 Is Nothing Then
            frm6.Hide()
        End If
        If Not frm7 Is Nothing Then
            frm7.Hide()
        End If
        Close()
    End Sub
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        close_app()
    End Sub
End Class