

Public Class Form5
    Dim panel2_y As Double = 0
    Dim uctl As UserControl1
    Dim ctl8 As Control
    Dim ctl12 As Control

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call load_form()
    End Sub
    Private Sub load_form()
        panel2_y = Me.Panel2.Location.Y
        Dim data_arr As New ArrayList
        Dim cnt As Integer
        Dim cls As New get_data3
        cls.get_data()
        data_arr = cls._data3
        cnt = cls.cnt
        Dim i As Integer
        For i = 0 To cnt - 1
            Dim k As Integer
            For k = 0 To 7
                Dim data3 As New data3
                data3 = data_arr.Item(i)
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
                        btn.Text = data3.Part
                        Me.Panel11.Controls.Add(btn)
                        btn.Name = data3.Part + "!" + data3.Revision
                        AddHandler btn.MouseDown, AddressOf form2_Click
                    Case 1
                        btn.Size = New Size(Me.Panel12.Width - 8, 25)
                        btn.Text = data3.Revision
                        Me.Panel12.Controls.Add(btn)
                    Case 2
                        btn.Size = New Size(Me.Panel13.Width - 8, 25)
                        btn.Text = data3.description
                        Me.Panel13.Controls.Add(btn)
                    Case 3
                        btn.Size = New Size(Me.Panel14.Width - 8, 25)
                        btn.Text = data3.ordered_qty
                        Me.Panel14.Controls.Add(btn)
                    Case 4
                        btn.Size = New Size(Me.Panel15.Width - 8, 25)
                        btn.Text = data3.in_qty
                        Me.Panel15.Controls.Add(btn)
                    Case 5
                        btn.Size = New Size(Me.Panel16.Width - 8, 25)
                        btn.Text = data3.ordered_date
                        Me.Panel16.Controls.Add(btn)
                    Case 6
                        btn.Size = New Size(Me.Panel17.Width - 8, 25)
                        btn.Text = data3.threshold_date
                        Me.Panel17.Controls.Add(btn)
                    Case 7
                        btn.Size = New Size(Me.Panel18.Width - 8, 25)
                        btn.Text = data3.ordered_by
                        Me.Panel18.Controls.Add(btn)
                End Select
            Next
        Next
        Me.Panel11.Size = New Size(Me.Panel11.Width, cnt * 31)
        Me.Panel12.Size = New Size(Me.Panel12.Width, cnt * 31)
        Me.Panel13.Size = New Size(Me.Panel13.Width, cnt * 31)
        Me.Panel14.Size = New Size(Me.Panel14.Width, cnt * 31)
        Me.Panel15.Size = New Size(Me.Panel15.Width, cnt * 31)
        Me.Panel16.Size = New Size(Me.Panel16.Width, cnt * 31)
        Me.Panel17.Size = New Size(Me.Panel17.Width, cnt * 31)
        Me.Panel18.Size = New Size(Me.Panel18.Width, cnt * 31)
        Me.Panel2.Size = New Size(Me.Panel2.Width, cnt * 31)
        Call add_handlers()
    End Sub
    Private Sub form2_Click(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Me.Panel1.Controls.Remove(uctl)
            ctl8 = DirectCast(sender, Button)
            ctl12 = DirectCast(sender, Button).Parent
            uctl = New UserControl1
            uctl.Button1.Text = "In Part"
            uctl.Size = New Size(uctl.Width, 30)
            uctl.Location = New Point(ctl8.Location.X + 70, ctl8.Location.Y + 10)
            Me.Panel1.Controls.Add(uctl)
            uctl.BringToFront()
            AddHandler uctl.Button1.Click, AddressOf in_part
        End If
    End Sub
    Private Sub add_handlers()
        AddHandler Panel11.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel12.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel13.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel14.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel15.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel16.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel17.SizeChanged, AddressOf panel_size_changed

        AddHandler Panel11.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel12.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel13.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel14.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel15.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel16.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel17.LocationChanged, AddressOf panel_Location_changed

        Call panel_size_changed()
        Call panel_Location_changed()
    End Sub
    Private Sub panel_size_changed()
        Me.Button1.Size = New Size(Me.Panel11.Width, Me.Button1.Height)
        Me.Button2.Size = New Size(Me.Panel12.Width, Me.Button2.Height)
        Me.Button3.Size = New Size(Me.Panel13.Width, Me.Button3.Height)
        Me.Button4.Size = New Size(Me.Panel14.Width, Me.Button4.Height)
        Me.Button5.Size = New Size(Me.Panel15.Width, Me.Button5.Height)
        Me.Button6.Size = New Size(Me.Panel16.Width, Me.Button6.Height)
        Me.Button7.Size = New Size(Me.Panel17.Width, Me.Button7.Height)
        Me.Button8.Size = New Size(Me.Panel18.Width, Me.Button8.Height)

        Me.Button1.Location = New Point(Me.Panel11.Location.X, Me.Button1.Location.Y)
        Me.Button2.Location = New Point(Me.Panel12.Location.X, Me.Button2.Location.Y)
        Me.Button3.Location = New Point(Me.Panel13.Location.X, Me.Button3.Location.Y)
        Me.Button4.Location = New Point(Me.Panel14.Location.X, Me.Button4.Location.Y)
        Me.Button5.Location = New Point(Me.Panel15.Location.X, Me.Button5.Location.Y)
        Me.Button6.Location = New Point(Me.Panel16.Location.X, Me.Button6.Location.Y)
        Me.Button7.Location = New Point(Me.Panel17.Location.X, Me.Button7.Location.Y)
        Me.Button8.Location = New Point(Me.Panel18.Location.X, Me.Button8.Location.Y)
    End Sub
    Private Sub panel_Location_changed()
        Me.Button1.Size = New Size(Me.Panel11.Width, Me.Button1.Height)
        Me.Button2.Size = New Size(Me.Panel12.Width, Me.Button2.Height)
        Me.Button3.Size = New Size(Me.Panel13.Width, Me.Button3.Height)
        Me.Button4.Size = New Size(Me.Panel14.Width, Me.Button4.Height)
        Me.Button5.Size = New Size(Me.Panel15.Width, Me.Button5.Height)
        Me.Button6.Size = New Size(Me.Panel16.Width, Me.Button6.Height)
        Me.Button7.Size = New Size(Me.Panel17.Width, Me.Button7.Height)
        Me.Button8.Size = New Size(Me.Panel18.Width, Me.Button8.Height)

        Me.Button1.Location = New Point(Me.Panel11.Location.X, Me.Button1.Location.Y)
        Me.Button2.Location = New Point(Me.Panel12.Location.X, Me.Button2.Location.Y)
        Me.Button3.Location = New Point(Me.Panel13.Location.X, Me.Button3.Location.Y)
        Me.Button4.Location = New Point(Me.Panel14.Location.X, Me.Button4.Location.Y)
        Me.Button5.Location = New Point(Me.Panel15.Location.X, Me.Button5.Location.Y)
        Me.Button6.Location = New Point(Me.Panel16.Location.X, Me.Button6.Location.Y)
        Me.Button7.Location = New Point(Me.Panel17.Location.X, Me.Button7.Location.Y)
        Me.Button8.Location = New Point(Me.Panel18.Location.X, Me.Button8.Location.Y)
    End Sub
    Private Sub VScrollBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles VScrollBar1.Scroll
        Me.Panel2.Location = New Point(Me.Panel2.Location.X, panel2_y - VScrollBar1.Value * 28)
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
    Private Sub remove_uctl()
        uctl.BackColor = Color.Transparent
        uctl.Controls.Remove(uctl.Button1)
        uctl.Controls.Remove(uctl.Button2)
        uctl.Controls.Remove(uctl.Button3)
        uctl.Controls.Remove(uctl.Button4)
        uctl.SendToBack()
        uctl.Hide()
        Me.Panel1.Controls.Remove(uctl)
    End Sub


End Class