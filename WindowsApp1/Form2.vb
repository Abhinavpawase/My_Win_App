

Public Class Form2

    Dim panel2_y As Double = 0
    Dim uctl As UserControl1 = Nothing
    Dim ctl8 As Control
    Dim ctl12 As Control

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call load_form()
    End Sub
    Private Sub load_form()
        panel2_y = Me.Panel2.Location.Y
        Dim _data1 As New ArrayList
        Dim cnt As Integer
        Dim cls As New get_data
        cls.get_data()
        _data1 = cls._data1
        cnt = cls.cnt
        Dim i As Integer
        For i = 0 To cnt - 1
            Dim k As Integer
            For k = 0 To 5
                Dim data2 As New data1
                data2 = _data1.Item(i)
                Dim btn As New Button
                btn.Location = New Point(4, 3 + 28 * i)
                btn.Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top
                btn.AutoSize = False
                btn.AutoSizeMode = AutoSizeMode.GrowAndShrink
                btn.Font = New Font("Microsoft Sans Serif", 11, FontStyle.Regular)
                btn.FlatAppearance.BorderSize = 0
                btn.FlatStyle = FlatStyle.Flat
                btn.TextAlign = ContentAlignment.MiddleLeft
                Select Case k
                    Case 0
                        btn.Size = New Size(Me.Panel11.Width - 8, 25)
                        btn.Text = data2.Part
                        Me.Panel11.Controls.Add(btn)
                        btn.Name = data2.Part + "!" + data2.Revision
                        AddHandler btn.MouseDown, AddressOf form2_Click
                    Case 1
                        btn.Size = New Size(Me.Panel12.Width - 8, 25)
                        btn.Text = data2.Revision
                        Me.Panel12.Controls.Add(btn)
                    Case 2
                        btn.Size = New Size(Me.Panel13.Width - 8, 25)
                        btn.Text = data2.description
                        Me.Panel13.Controls.Add(btn)
                    Case 3
                        btn.Size = New Size(Me.Panel14.Width - 8, 25)
                        btn.Text = data2.Quantity
                        Me.Panel14.Controls.Add(btn)
                End Select
            Next
        Next
        Me.Panel11.Size = New Size(Me.Panel11.Width, cnt * 31)
        Me.Panel12.Size = New Size(Me.Panel12.Width, cnt * 31)
        Me.Panel13.Size = New Size(Me.Panel13.Width, cnt * 31)
        Me.Panel14.Size = New Size(Me.Panel14.Width, cnt * 31)
        Me.Panel2.Size = New Size(Me.Panel2.Width, cnt * 31)
        Call mouse_hover()
        Call add_handlers()
    End Sub
    Private Sub mouse_hover()
        AddHandler Me.MouseHover, AddressOf Form2_MouseHover
        Dim ctl As Control
        For Each ctl In Me.Controls
            AddHandler ctl.MouseHover, AddressOf Form2_MouseHover
            Dim ctl2 As Control
            For Each ctl2 In ctl.Controls
                AddHandler ctl2.MouseHover, AddressOf Form2_MouseHover
                Dim ctl3 As Control
                For Each ctl3 In ctl2.Controls
                    AddHandler ctl3.MouseHover, AddressOf Form2_MouseHover
                    Dim ctl4 As Control
                    For Each ctl4 In ctl3.Controls
                        AddHandler ctl4.MouseHover, AddressOf Form2_MouseHover
                    Next
                Next
            Next
        Next
    End Sub
    Private Sub add_handlers()
        AddHandler Panel11.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel12.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel13.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel14.SizeChanged, AddressOf panel_size_changed

        AddHandler Panel11.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel12.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel13.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel14.LocationChanged, AddressOf panel_Location_changed

        Call panel_size_changed()
        Call panel_Location_changed()
    End Sub
    Private Sub panel_size_changed()
        Me.Button1.Size = New Size(Me.Panel11.Width, Me.Button1.Height)
        Me.Button2.Size = New Size(Me.Panel12.Width, Me.Button2.Height)
        Me.Button3.Size = New Size(Me.Panel13.Width, Me.Button3.Height)
        Me.Button4.Size = New Size(Me.Panel14.Width, Me.Button4.Height)

        Me.Button1.Location = New Point(Me.Panel11.Location.X, Me.Button1.Location.Y)
        Me.Button2.Location = New Point(Me.Panel12.Location.X, Me.Button2.Location.Y)
        Me.Button3.Location = New Point(Me.Panel13.Location.X, Me.Button3.Location.Y)
        Me.Button4.Location = New Point(Me.Panel14.Location.X, Me.Button4.Location.Y)
    End Sub
    Private Sub panel_Location_changed()
        Me.Button1.Size = New Size(Me.Panel11.Width, Me.Button1.Height)
        Me.Button2.Size = New Size(Me.Panel12.Width, Me.Button2.Height)
        Me.Button3.Size = New Size(Me.Panel13.Width, Me.Button3.Height)
        Me.Button4.Size = New Size(Me.Panel14.Width, Me.Button4.Height)

        Me.Button1.Location = New Point(Me.Panel11.Location.X, Me.Button1.Location.Y)
        Me.Button2.Location = New Point(Me.Panel12.Location.X, Me.Button2.Location.Y)
        Me.Button3.Location = New Point(Me.Panel13.Location.X, Me.Button3.Location.Y)
        Me.Button4.Location = New Point(Me.Panel14.Location.X, Me.Button4.Location.Y)
    End Sub
    Private Sub VScrollBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles VScrollBar1.Scroll
        Me.Panel2.Location = New Point(Me.Panel2.Location.X, panel2_y - VScrollBar1.Value * 28)
    End Sub
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Call Close()
    End Sub
    Private Sub form2_Click(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Me.Panel1.Controls.Remove(uctl)
            ctl8 = DirectCast(sender, Button)
            ctl12 = DirectCast(sender, Button).Parent
            uctl = New UserControl1
            uctl.Location = New Point(ctl8.Location.X + 70, ctl8.Location.Y + 10)
            Me.Panel1.Controls.Add(uctl)
            uctl.BringToFront()
            AddHandler uctl.Button1.Click, AddressOf Place_order
            AddHandler uctl.Button2.Click, AddressOf request_part
            AddHandler uctl.Button3.Click, AddressOf in_part
            AddHandler uctl.Button4.Click, AddressOf out_part
        End If

    End Sub
    Private Sub request_part()
        Call remove_uctl()
        Dim str As String = ctl8.Name
        Dim frm As New Form4
        frm.Label1.Text = "Request Date"
        frm.Label3.Text = "Required Before"
        frm.partno = str
        frm.Show()
    End Sub
    Private Sub Place_order()
        Call remove_uctl()
        Dim str As String = ctl8.Name
        Dim frm As New Form4
        frm.Label1.Text = "Ordered Date"
        frm.Label3.Text = "Threshold Date"
        frm.partno = str
        frm.Show()
    End Sub
    Private Sub in_part()
        Call remove_uctl()
        Dim str As String = ctl8.Name
        Dim frm As New Form4
        frm.Label1.Text = "In Date"
        frm.Label3.Visible = False
        frm.TextBox3.Visible = False
        frm.DateTimePicker2.Visible = False
        frm.Button2.Location = New Point(57, 255 - 50)
        frm.Button1.Location = New Point(188, 255 - 50)
        frm.Size = New Size(frm.Width, frm.Height - 50)
        frm.partno = str
        frm.Show()
    End Sub
    Private Sub out_part()
        Call remove_uctl()
        Dim str As String = ctl8.Name
        Dim frm As New Form4
        frm.Label1.Text = "Out Date"
        frm.Label3.Visible = False
        frm.TextBox3.Visible = False
        frm.DateTimePicker2.Visible = False
        frm.Button2.Location = New Point(57, 255 - 50)
        frm.Button1.Location = New Point(188, 255 - 50)
        frm.Size = New Size(frm.Width, frm.Height - 50)
        frm.partno = str
        frm.Show()
    End Sub
    Private Sub remove_uctl()
        uctl.BackColor = Color.Transparent
        uctl.Controls.Remove(uctl.Button1)
        uctl.Controls.Remove(uctl.Button2)
        uctl.Controls.Remove(uctl.Button3)
        uctl.Controls.Remove(uctl.Button4)
        uctl.SendToBack()
        uctl.Hide()
        Me.Panel1.Controls.Remove(uctl)
        uctl = Nothing
    End Sub
    Private Sub Form2_MouseHover(sender As Object, e As EventArgs)
        If Not uctl Is Nothing Then
            Call remove_uctl()
        End If
    End Sub

End Class