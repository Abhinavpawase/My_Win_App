
Public Class Form4
    Public Property partno As String
    Dim part_name As String
    Dim revision As String
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        part_name = Mid(partno, 1, Len(partno) - 2)
        revision = Mid(partno, Len(partno), 1)
        Me.Label5.Text = part_name
        Me.Label6.Text = revision
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.Label1.Text = "Request Date" Then
            Dim cls As New request_part
            Dim req_part_data1 As New req_part_data
            req_part_data1.Part = part_name
            req_part_data1.Revision = revision
            req_part_data1.req_date = Me.TextBox1.Text
            Dim dty As Integer = Me.TextBox2.Text
            req_part_data1.req_qty = dty.ToString
            req_part_data1.req_before = Me.TextBox3.Text
            cls.req_part_data1 = req_part_data1
            Call cls.insert_data()
            Me.Dispose()
            MsgBox("Request Placed")
        End If
        If Me.Label1.Text = "Ordered Date" Then
            Dim cls As New ordered_part
            Dim req_part_data1 As New ordered_part_data
            req_part_data1.Part = part_name
            req_part_data1.Revision = revision
            req_part_data1.ordered_date = Me.TextBox1.Text
            Dim dty As Integer = Me.TextBox2.Text
            req_part_data1.ordered_qty = dty.ToString
            req_part_data1.Threshold_date = Me.TextBox3.Text
            cls.ordered_part_data1 = req_part_data1
            Call cls.insert_data()
            Me.Dispose()
            MsgBox("Ordered Placed")
        End If
        If Me.Label1.Text = "In Date" Then
            Dim cls As New In_part
            Dim req_part_data1 As New In_part_data
            req_part_data1.Part = part_name
            req_part_data1.Revision = revision
            req_part_data1.in_date = Me.TextBox1.Text
            Dim dty As Integer = Me.TextBox2.Text
            req_part_data1.in_qty = dty.ToString
            cls.In_part_data1 = req_part_data1
            Call cls.insert_data()
            Me.Dispose()
            MsgBox("Part In")
        End If
        If Me.Label1.Text = "Out Date" Then
            Dim cls As New Out_part
            Dim req_part_data1 As New Out_part_data
            req_part_data1.Part = part_name
            req_part_data1.Revision = revision
            req_part_data1.out_date = Me.TextBox1.Text
            Dim dty As Integer = Me.TextBox2.Text
            req_part_data1.out_qty = dty.ToString
            cls.Out_part_data1 = req_part_data1
            Call cls.insert_data()
            MsgBox("Part Out")
        End If

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Dispose()
    End Sub
    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        Call get_date()
    End Sub
    Private Sub get_date()
        Dim str As String = Me.TextBox1.Text
        If Len(str) > 0 Then
            Dim date1 As String = ""
            Call get_date12(str, date1)
            If date1 <> "" Then
                Me.TextBox1.Text = date1
            End If
        End If
        str = Me.TextBox3.Text
        If Len(str) > 0 Then
            Dim date1 As String = ""
            Call get_date12(str, date1)
            If date1 <> "" Then
                Me.TextBox3.Text = date1
            End If
        End If
    End Sub
    Private Sub get_date12(str As String, ByRef str2 As String)
        Dim y As Integer
        Dim bool1 As Boolean = False
        For y = 0 To Len(str) - 1
            If str(y) = "-" Then
                bool1 = True
            End If
        Next
        If bool1 = False Then
            Dim date1 As String = ""
            Dim month1 As String = ""
            Dim year1 As String = ""
            If IsNumeric(Mid(str, Len(str), 1)) Then
                Dim i As Integer
                Dim j As Integer = 1
                For i = 1 To Len(str)
                    If Not IsNumeric(Mid(str, i, 1)) Then
                        j = i
                        Exit For
                    End If
                Next
                date1 = Mid(str, 1, j - 1)
                month1 = Mid(str, j, 3)
                year1 = Mid(str, j + 3, 2)
            End If
            If Not IsNumeric(Mid(str, Len(str), 1)) Then
                Dim i As Integer
                Dim j As Integer = 1
                For i = 1 To Len(str)
                    If Not IsNumeric(Mid(str, i, 1)) Then
                        j = i
                        Exit For
                    End If
                Next
                date1 = Mid(str, 1, j - 1)
                month1 = Mid(str, j, 3)
                year1 = "21"
            End If
            str2 = date1 + "-" + month1 + "-" + year1
        End If
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Call get_date()
        End If
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Call get_date()
        End If
    End Sub
    Private Sub DateTimePicker1_TextChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.TextChanged
        Me.TextBox1.Text = DateTimePicker1.Value.ToString("dd - MMM - yyyy")
    End Sub
    Private Sub DateTimePicker2_TextChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.TextChanged
        Me.TextBox3.Text = DateTimePicker2.Value.ToString("dd - MMM - yyyy")
    End Sub

End Class