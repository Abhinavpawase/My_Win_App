

Public Class Form7
    Dim panel2_y As Double = 0

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call load_form()
    End Sub
    Private Sub load_form()
        panel2_y = Me.Panel2.Location.Y
        Dim data_arr As New ArrayList
        Dim cnt As Integer
        Dim cls As New get_data5
        cls.get_data()
        data_arr = cls._data5
        cnt = cls.cnt
        Dim i As Integer
        For i = 0 To cnt - 1
            Dim k As Integer
            For k = 0 To 4
                Dim data5 As New data5
                data5 = data_arr.Item(i)
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
                        btn.Text = data5.Part
                        Me.Panel11.Controls.Add(btn)
                        btn.Name = data5.Part + "!" + data5.Revision
                        'AddHandler btn.MouseDown, AddressOf form2_Click
                    Case 1
                        btn.Size = New Size(Me.Panel12.Width - 8, 25)
                        btn.Text = data5.Revision
                        Me.Panel12.Controls.Add(btn)
                    Case 2
                        btn.Size = New Size(Me.Panel13.Width - 8, 25)
                        btn.Text = data5.description
                        Me.Panel13.Controls.Add(btn)
                    Case 3
                        btn.Size = New Size(Me.Panel14.Width - 8, 25)
                        btn.Text = data5.Out_qty
                        Me.Panel14.Controls.Add(btn)
                    Case 4
                        btn.Size = New Size(Me.Panel15.Width - 8, 25)
                        btn.Text = data5.Out_On
                        Me.Panel15.Controls.Add(btn)
                End Select
            Next
        Next
        Me.Panel11.Size = New Size(Me.Panel11.Width, cnt * 31)
        Me.Panel12.Size = New Size(Me.Panel12.Width, cnt * 31)
        Me.Panel13.Size = New Size(Me.Panel13.Width, cnt * 31)
        Me.Panel14.Size = New Size(Me.Panel14.Width, cnt * 31)
        Me.Panel15.Size = New Size(Me.Panel15.Width, cnt * 31)
        Me.Panel2.Size = New Size(Me.Panel2.Width, cnt * 31)
        Call add_handlers()
    End Sub
    Private Sub add_handlers()
        AddHandler Panel11.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel12.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel13.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel14.SizeChanged, AddressOf panel_size_changed
        AddHandler Panel15.SizeChanged, AddressOf panel_size_changed

        AddHandler Panel11.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel12.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel13.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel14.LocationChanged, AddressOf panel_Location_changed
        AddHandler Panel15.LocationChanged, AddressOf panel_Location_changed

        Call panel_size_changed()
        Call panel_Location_changed()
    End Sub
    Private Sub panel_size_changed()
        Me.Button1.Size = New Size(Me.Panel11.Width, Me.Button1.Height)
        Me.Button2.Size = New Size(Me.Panel12.Width, Me.Button2.Height)
        Me.Button3.Size = New Size(Me.Panel13.Width, Me.Button3.Height)
        Me.Button4.Size = New Size(Me.Panel14.Width, Me.Button4.Height)
        Me.Button5.Size = New Size(Me.Panel15.Width, Me.Button5.Height)

        Me.Button1.Location = New Point(Me.Panel11.Location.X, Me.Button1.Location.Y)
        Me.Button2.Location = New Point(Me.Panel12.Location.X, Me.Button2.Location.Y)
        Me.Button3.Location = New Point(Me.Panel13.Location.X, Me.Button3.Location.Y)
        Me.Button4.Location = New Point(Me.Panel14.Location.X, Me.Button4.Location.Y)
        Me.Button5.Location = New Point(Me.Panel15.Location.X, Me.Button5.Location.Y)
    End Sub
    Private Sub panel_Location_changed()
        Me.Button1.Size = New Size(Me.Panel11.Width, Me.Button1.Height)
        Me.Button2.Size = New Size(Me.Panel12.Width, Me.Button2.Height)
        Me.Button3.Size = New Size(Me.Panel13.Width, Me.Button3.Height)
        Me.Button4.Size = New Size(Me.Panel14.Width, Me.Button4.Height)
        Me.Button5.Size = New Size(Me.Panel15.Width, Me.Button5.Height)

        Me.Button1.Location = New Point(Me.Panel11.Location.X, Me.Button1.Location.Y)
        Me.Button2.Location = New Point(Me.Panel12.Location.X, Me.Button2.Location.Y)
        Me.Button3.Location = New Point(Me.Panel13.Location.X, Me.Button3.Location.Y)
        Me.Button4.Location = New Point(Me.Panel14.Location.X, Me.Button4.Location.Y)
        Me.Button5.Location = New Point(Me.Panel15.Location.X, Me.Button5.Location.Y)
    End Sub
    Private Sub VScrollBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles VScrollBar1.Scroll
        Me.Panel2.Location = New Point(Me.Panel2.Location.X, panel2_y - VScrollBar1.Value * 28)
    End Sub

End Class