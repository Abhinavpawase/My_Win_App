
Imports System.Data.SqlClient

Public Class get_data

    Public data_1(1) As data1
    Public cnt As Integer
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Public Sub get_data()
        Try
            cnt = 0
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "SELECT COUNT(part) FROM inventory_table_1"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Dim count As Integer = 0
            Do While myReader.Read()
                count = myReader.GetInt32(0)
            Loop
            myReader.Close()

            Array.Resize(data_1, count)

            str = "select Part,Revision,description,Quantity from Inventory_Table_1 "
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Do While myReader.Read()
                Dim data2 As New data1
                data2.Part = myReader.GetString(0).Trim()
                data2.Revision = myReader.GetString(1).Trim()
                data2.description = myReader.GetString(2).Trim()
                data2.Quantity = myReader.GetString(3).Trim()
                data_1(cnt) = data2
                cnt += 1
            Loop
            myReader.Close()
            myConn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

End Class
Public Structure data1
    Public Part As String
    Public Revision As String
    Public Quantity As String
    Public description As String
End Structure

Public Class get_data2
    Public data_1(1) As data2
    Public cnt As Integer
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Public Sub get_data()
        Try
            cnt = 0
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "SELECT COUNT(part) FROM Required_Parts_Table"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Dim count As Integer = 0
            Do While myReader.Read()
                count = myReader.GetInt32(0)
            Loop
            myReader.Close()

            Array.Resize(data_1, count)

            str = "select Part,Revision,description,Req_Qty,Out_Qty,Req_date,Req_before,requested_by from Required_Parts_Table;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Do While myReader.Read()
                Dim data2 As New data2
                data2.Part = myReader.GetString(0).Trim()
                data2.Revision = myReader.GetString(1).Trim()
                data2.description = myReader.GetString(2).Trim()
                data2.req_qty = myReader.GetString(3).Trim()
                data2.out_qty = myReader.GetString(4).Trim()
                data2.req_date = myReader.GetDateTime(5).ToShortDateString.Trim
                data2.req_before = myReader.GetDateTime(6).ToShortDateString.Trim
                data2.req_by = myReader.GetString(7).Trim()
                data_1(cnt) = data2
                cnt += 1
            Loop
            myReader.Close()
            myConn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

End Class
Public Structure data2
    Public Part As String
    Public Revision As String
    Public req_qty As String
    Public out_qty As String
    Public req_date As String
    Public req_before As String
    Public description As String
    Public req_by As String
End Structure

Public Class get_data3
    Public data_1(1) As data3
    Public cnt As Integer
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Public Sub get_data()
        Try
            cnt = 0
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "SELECT COUNT(part) FROM Ordered_Parts_Table"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Dim count As Integer = 0
            Do While myReader.Read()
                count = myReader.GetInt32(0)
            Loop
            myReader.Close()

            Array.Resize(data_1, count)

            str = "select Part,Revision,description,Ordered_qty,In_qty,Ordered_Date,Threshold_Date,Ordered_by from Ordered_Parts_Table;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Do While myReader.Read()
                Dim data2 As New data3
                data2.Part = myReader.GetString(0).Trim()
                data2.Revision = myReader.GetString(1).Trim()
                data2.description = myReader.GetString(2).Trim()
                data2.ordered_qty = myReader.GetString(3).Trim()
                data2.in_qty = myReader.GetString(4).Trim()
                data2.ordered_date = myReader.GetDateTime(5).ToShortDateString.Trim
                data2.threshold_date = myReader.GetDateTime(6).ToShortDateString.Trim
                data2.ordered_by = myReader.GetString(7).Trim()
                data_1(cnt) = data2
                cnt += 1
            Loop
            myReader.Close()
            myConn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

End Class
Public Structure data3
    Public Part As String
    Public Revision As String
    Public ordered_qty As String
    Public in_qty As String
    Public ordered_date As String
    Public threshold_date As String
    Public description As String
    Public ordered_by As String
End Structure

Public Class get_data4
    Public data_1(1) As data4
    Public cnt As Integer
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Public Sub get_data()
        Try
            cnt = 0
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "SELECT COUNT(part) FROM In_Parts_Table"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Dim count As Integer = 0
            Do While myReader.Read()
                count = myReader.GetInt32(0)
            Loop
            myReader.Close()

            Array.Resize(data_1, count)

            str = "select Part,Revision,description,In_qty,In_On from In_Parts_Table;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Do While myReader.Read()
                Dim data2 As New data4
                data2.Part = myReader.GetString(0).Trim()
                data2.Revision = myReader.GetString(1).Trim()
                data2.description = myReader.GetString(2).Trim()
                data2.In_qty = myReader.GetString(3).Trim()
                data2.in_On = myReader.GetDateTime(4).ToShortDateString.Trim
                data_1(cnt) = data2
                cnt += 1
            Loop
            myReader.Close()
            myConn.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

End Class
Public Structure data4
    Public Part As String
    Public Revision As String
    Public In_qty As String
    Public in_On As String
    Public description As String
End Structure

Public Class get_data5
    Public data_1(1) As data5
    Public cnt As Integer
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Public Sub get_data()
        Try
            cnt = 0
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "SELECT COUNT(part) FROM Out_Parts_Table"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Dim count As Integer = 0
            Do While myReader.Read()
                count = myReader.GetInt32(0)
            Loop
            myReader.Close()

            Array.Resize(data_1, count)

            str = "select Part,Revision,description, Out_Qty	, Out_On from Out_Parts_Table;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Do While myReader.Read()
                Dim data2 As New data5
                data2.Part = myReader.GetString(0).Trim()
                data2.Revision = myReader.GetString(1).Trim()
                data2.description = myReader.GetString(2).Trim()
                data2.Out_qty = myReader.GetString(3).Trim()
                data2.Out_On = myReader.GetDateTime(4).ToShortDateString.Trim
                data_1(cnt) = data2
                cnt += 1
            Loop
            myReader.Close()
            myConn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

End Class
Public Structure data5
    Public Part As String
    Public Revision As String
    Public Out_qty As String
    Public Out_On As String
    Public description As String
End Structure

Public Class request_part
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Public req_part_data1 As req_part_data
    Public Sub insert_data()
        Try
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "Select description from inventory_table_1 where part = '" + req_part_data1.Part + "' and revision = '" + req_part_data1.Revision + "' ;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Dim desc As String = ""
            Do While myReader.Read
                desc = myReader.GetString(0).Trim
            Loop
            myReader.Close()

            str = "insert into Required_Parts_Table (Part,Revision,Req_Qty,Out_Qty,Req_date,Req_before,description, requested_by) values ('" + req_part_data1.Part + "','" + req_part_data1.Revision + "','" + req_part_data1.req_qty + "','0','" + req_part_data1.req_date + "','" + req_part_data1.req_before + "','" + desc + "','" + Module1.username + "') ;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            myReader.Close()
            myConn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
Public Structure req_part_data
    Public Part As String
    Public Revision As String
    Public req_qty As String
    Public req_date As String
    Public req_before As String
End Structure

Public Class ordered_part
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Public ordered_part_data1 As ordered_part_data
    Public Sub insert_data()
        Try
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "Select description from inventory_table_1 where part = '" + ordered_part_data1.Part + "' and revision = '" + ordered_part_data1.Revision + "' ;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Dim desc As String = ""
            Do While myReader.Read
                desc = myReader.GetString(0).Trim
            Loop
            myReader.Close()

            str = "insert into Ordered_Parts_Table (Part,Revision,Ordered_qty,In_qty,Ordered_Date,Threshold_Date,description,Ordered_by) values ('" + ordered_part_data1.Part + "','" + ordered_part_data1.Revision + "','" + ordered_part_data1.ordered_qty + "','0','" + ordered_part_data1.ordered_date + "','" + ordered_part_data1.Threshold_date + "','" + desc + "','" + Module1.username + "') ;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            myReader.Close()
            myConn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
Public Structure ordered_part_data
    Public Part As String
    Public Revision As String
    Public ordered_qty As String
    Public ordered_date As String
    Public Threshold_date As String
End Structure

Public Class In_part
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Public In_part_data1 As In_part_data
    Public Sub insert_data()
        Try
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "Select description from inventory_table_1 where part = '" + In_part_data1.Part + "' and revision = '" + In_part_data1.Revision + "' ;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Dim desc As String = ""
            Do While myReader.Read
                desc = myReader.GetString(0).Trim
            Loop
            myReader.Close()

            str = "insert into In_Parts_Table (Part,Revision,In_qty,In_On,description) values ('" + In_part_data1.Part + "','" + In_part_data1.Revision + "','" + In_part_data1.in_qty + "','" + In_part_data1.in_date + "','" + desc + "') ;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            myReader.Close()
            myConn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
Public Structure In_part_data
    Public Part As String
    Public Revision As String
    Public in_qty As String
    Public in_date As String
End Structure

Public Class Out_part
    Public myConn As SqlConnection
    Public myCmd As SqlCommand
    Public myReader As SqlDataReader
    Public Out_part_data1 As Out_part_data
    Public Sub insert_data()
        Try
            myConn = New SqlConnection("Data Source = LAPTOP-U89PQ616\SQLEXPRESS ; Initial Catalog = Inventory ; Integrated Security=SSPI ; ")
            myConn.Open()
            myCmd = myConn.CreateCommand
            Dim str As String
            str = "Select description from inventory_table_1 where part = '" + Out_part_data1.Part + "' and revision = '" + Out_part_data1.Revision + "' ;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            Dim desc As String = ""
            Do While myReader.Read
                desc = myReader.GetString(0).Trim
            Loop
            myReader.Close()

            str = "insert into Out_Parts_Table (Part,Revision,Out_Qty,Out_On,description) values ('" + Out_part_data1.Part + "','" + Out_part_data1.Revision + "','" + Out_part_data1.out_qty + "','" + Out_part_data1.out_date + "','" + desc + "') ;"
            myCmd.CommandText = str
            myReader = myCmd.ExecuteReader()
            myReader.Close()
            myConn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
Public Structure Out_part_data
    Public Part As String
    Public Revision As String
    Public out_qty As String
    Public out_date As String
End Structure