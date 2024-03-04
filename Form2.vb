Imports System.Data.SqlClient

Public Class Form2

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.AcceptButton = Button2

        ComboBox2.SelectedIndex = 0
        ComboBox1.SelectedIndex = 0

        Dim strConnection As String = "server=" & Trim(Form1.TextBox4.Text) & "; " &
             "user id=" & Trim(Form1.TextBox2.Text) & "; " &
             "password=" & Trim(Form1.TextBox3.Text) & ";" &
             "database=" & Trim(Form1.TextBox1.Text) & ""

        Dim conn As New SqlConnection(strConnection)

        Try
            conn.Open()

            Dim cmd As New SqlCommand
            Dim data As New DataTable
            cmd.Connection = conn
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "Select * From ITEM_UPGRADE Order By nIndex Desc"
            cmd.ExecuteNonQuery()
            Dim adp As New SqlDataAdapter(cmd)
            adp.Fill(data)
            TextBox1.Text = data.Rows(0).Item(0) + 100
            conn.Close()

            'MsgBox("Successfully connected to the database.")

        Catch ex As Exception

            'MsgBox("Unable to connect to database, please re-check the connection information.")

        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            ' Create method of writing item values into text.
            Dim objWriter As New System.IO.StreamWriter(Application.StartupPath & "/ITEM_UPGRADE.sql", True)
            ' Item_Upgrade query to +1
            objWriter.WriteLine("Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 1 & ",5001,'" & TextBox2.Text & "','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox4.Text & "," & TextBox20.Text & "," & TextBox19.Text & "," & TextBox18.Text & "," & TextBox17.Text & "," & TextBox16.Text & "," & TextBox15.Text & "," & TextBox14.Text & "," & TextBox21.Text & "," & TextBox22.Text & ",0," & TextBox115.Text & "," & TextBox32.Text & ")")
            ' Item_Upgrade query to +2
            objWriter.WriteLine("Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 2 & ",5001,'" & TextBox2.Text & "','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox5.Text & "," & TextBox50.Text & "," & TextBox49.Text & "," & TextBox48.Text & "," & TextBox47.Text & "," & TextBox46.Text & "," & TextBox45.Text & "," & TextBox44.Text & "," & TextBox43.Text & "," & TextBox23.Text & ",0," & TextBox116.Text & "," & TextBox33.Text & ")")
            ' Item_Upgrade query to +3
            objWriter.WriteLine("Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 3 & ",5001,'" & TextBox2.Text & "','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox6.Text & "," & TextBox58.Text & "," & TextBox57.Text & "," & TextBox56.Text & "," & TextBox55.Text & "," & TextBox54.Text & "," & TextBox53.Text & "," & TextBox52.Text & "," & TextBox51.Text & "," & TextBox24.Text & ",0," & TextBox117.Text & "," & TextBox34.Text & ")")
            ' Item_Upgrade query to +4
            objWriter.WriteLine("Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 4 & ",5001,'" & TextBox2.Text & "','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox7.Text & "," & TextBox66.Text & "," & TextBox65.Text & "," & TextBox64.Text & "," & TextBox63.Text & "," & TextBox62.Text & "," & TextBox61.Text & "," & TextBox60.Text & "," & TextBox59.Text & "," & TextBox25.Text & ",0," & TextBox118.Text & "," & TextBox35.Text & ")")
            ' Item_Upgrade query to +5
            objWriter.WriteLine("Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 5 & ",5001,'" & TextBox2.Text & "','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox8.Text & "," & TextBox74.Text & "," & TextBox73.Text & "," & TextBox72.Text & "," & TextBox71.Text & "," & TextBox70.Text & "," & TextBox69.Text & "," & TextBox68.Text & "," & TextBox67.Text & "," & TextBox26.Text & ",0," & TextBox119.Text & "," & TextBox36.Text & ")")
            ' Item_Upgrade query to +6
            objWriter.WriteLine("Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 6 & ",5001,'" & TextBox2.Text & "','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox9.Text & "," & TextBox82.Text & "," & TextBox81.Text & "," & TextBox80.Text & "," & TextBox79.Text & "," & TextBox78.Text & "," & TextBox77.Text & "," & TextBox76.Text & "," & TextBox75.Text & "," & TextBox27.Text & ",0," & TextBox120.Text & "," & TextBox37.Text & ")")
            ' Item_Upgrade query to +7
            objWriter.WriteLine("Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 7 & ",5001,'" & TextBox2.Text & "','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox10.Text & "," & TextBox90.Text & "," & TextBox89.Text & "," & TextBox88.Text & "," & TextBox87.Text & "," & TextBox86.Text & "," & TextBox85.Text & "," & TextBox84.Text & "," & TextBox83.Text & "," & TextBox28.Text & ",0," & TextBox121.Text & "," & TextBox38.Text & ")")
            ' Item_Upgrade query to +8
            objWriter.WriteLine("Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 8 & ",5001,'" & TextBox2.Text & "','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox11.Text & "," & TextBox98.Text & "," & TextBox97.Text & "," & TextBox96.Text & "," & TextBox95.Text & "," & TextBox94.Text & "," & TextBox93.Text & "," & TextBox92.Text & "," & TextBox91.Text & "," & TextBox29.Text & ",0," & TextBox122.Text & "," & TextBox39.Text & ")")
            ' Item_Upgrade query to +9
            objWriter.WriteLine("Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 9 & ",5001,'" & TextBox2.Text & "','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox12.Text & "," & TextBox106.Text & "," & TextBox105.Text & "," & TextBox104.Text & "," & TextBox103.Text & "," & TextBox102.Text & "," & TextBox101.Text & "," & TextBox100.Text & "," & TextBox99.Text & "," & TextBox30.Text & ",0," & TextBox123.Text & "," & TextBox40.Text & ")")
            ' Item_Upgrade query to +10
            objWriter.WriteLine("Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 10 & ",5001,'" & TextBox2.Text & "','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox13.Text & "," & TextBox114.Text & "," & TextBox113.Text & "," & TextBox112.Text & "," & TextBox111.Text & "," & TextBox110.Text & "," & TextBox109.Text & "," & TextBox108.Text & "," & TextBox107.Text & "," & TextBox31.Text & ",0," & TextBox124.Text & "," & TextBox41.Text & ")")
            objWriter.Close()
        Catch ex As Exception
            MsgBox("File has been updated!")
        Finally
            MsgBox("ITEM sql script has been created!")
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim strConnection As String = "server=" & Trim(Form1.TextBox4.Text) & "; " &
             "user id=" & Trim(Form1.TextBox2.Text) & "; " &
             "password=" & Trim(Form1.TextBox3.Text) & ";" &
             "database=" & Trim(Form1.TextBox1.Text) & ""

        Dim conn1 As New SqlConnection(strConnection)

        Try
            conn1.Open()
            Dim cmd1 As New SqlCommand
            Dim data1 As New DataTable
            cmd1.Connection = conn1
            cmd1.CommandType = CommandType.Text
            cmd1.CommandText = "Select strName, Num From ITEM Where StrName Like '%" & TextBox2.Text & "%'"
            cmd1.ExecuteNonQuery()
            Dim adp As New SqlDataAdapter(cmd1)
            adp.Fill(data1)
            Dim ds As New DataSet
            adp.Fill(ds, "StrName")
            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox("Item not found!")
            End If
            TextBox2.DataBindings.Add("Text", ds.Tables("strName"), "Subject")
            conn1.Close()

        Catch ex As Exception

            MsgBox("Unable to connect to database, please re-check the connection information.")

        End Try

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        TextBox32.Text = Val(TextBox5.Text) - Val(TextBox4.Text)
    End Sub
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        TextBox33.Text = Val(TextBox6.Text) - Val(TextBox5.Text)
    End Sub
    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        TextBox34.Text = Val(TextBox7.Text) - Val(TextBox6.Text)
    End Sub
    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        TextBox35.Text = Val(TextBox8.Text) - Val(TextBox7.Text)
    End Sub
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        TextBox36.Text = Val(TextBox9.Text) - Val(TextBox8.Text)
    End Sub
    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        TextBox37.Text = Val(TextBox10.Text) - Val(TextBox9.Text)
    End Sub
    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        TextBox38.Text = Val(TextBox11.Text) - Val(TextBox10.Text)
    End Sub
    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        TextBox39.Text = Val(TextBox12.Text) - Val(TextBox11.Text)
    End Sub
    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        TextBox40.Text = Val(TextBox13.Text) - Val(TextBox12.Text)
    End Sub
    Private Sub TextBox42_TextChanged(sender As Object, e As EventArgs) Handles TextBox42.TextChanged
        TextBox41.Text = Val(TextBox42.Text) - Val(TextBox13.Text)
    End Sub

    'Item Group Start
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        'Custom preset.
        If ComboBox2.SelectedIndex = 0 Then
            'Required Money
            TextBox22.Text = 0
            TextBox23.Text = 0
            TextBox24.Text = 0
            TextBox25.Text = 0
            TextBox26.Text = 0
            TextBox27.Text = 0
            TextBox28.Text = 0
            TextBox29.Text = 0
            TextBox30.Text = 0
            TextBox31.Text = 0
            'Required Items
            TextBox20.Text = 0
            TextBox50.Text = 0
            TextBox58.Text = 0
            TextBox66.Text = 0
            TextBox74.Text = 0
            TextBox82.Text = 0
            TextBox90.Text = 0
            TextBox98.Text = 0
            TextBox106.Text = 0
            TextBox114.Text = 0
            'Upgrade Possibility
            TextBox115.Text = 10000
            TextBox116.Text = 10000
            TextBox117.Text = 10000
            TextBox118.Text = 10000
            TextBox119.Text = 10000
            TextBox120.Text = 10000
            TextBox121.Text = 10000
            TextBox122.Text = 10000
            TextBox123.Text = 10000
            TextBox124.Text = 10000
        End If
        'Mid Upgrade preset.
        If ComboBox2.SelectedIndex = 1 Then
            'Required Money
            TextBox22.Text = 240000
            TextBox23.Text = 240000
            TextBox24.Text = 240000
            TextBox25.Text = 240000
            TextBox26.Text = 240000
            TextBox27.Text = 240000
            TextBox28.Text = 240000
            TextBox29.Text = 240000
            TextBox30.Text = 240000
            TextBox31.Text = 240000
            'Required Items
            TextBox20.Text = 379016000
            TextBox50.Text = 379016000
            TextBox58.Text = 379016000
            TextBox66.Text = 379016000
            TextBox74.Text = 379016000
            TextBox82.Text = 379016000
            TextBox90.Text = 379016000
            TextBox98.Text = 379016000
            TextBox106.Text = 379016000
            TextBox114.Text = 379016000
            'Upgrade Possibility
            TextBox115.Text = 10000
            TextBox116.Text = 10000
            TextBox117.Text = 10000
            TextBox118.Text = 10000
            TextBox119.Text = 10000
            TextBox120.Text = 10000
            TextBox121.Text = 10000
            TextBox122.Text = 10000
            TextBox123.Text = 10000
            TextBox124.Text = 10000
        End If
        'HIGH Upgrade preset.
        If ComboBox2.SelectedIndex = 2 Then
            'Required Money
            TextBox22.Text = 500000
            TextBox23.Text = 500000
            TextBox24.Text = 500000
            TextBox25.Text = 500000
            TextBox26.Text = 500000
            TextBox27.Text = 500000
            TextBox28.Text = 500000
            TextBox29.Text = 500000
            TextBox30.Text = 500000
            TextBox31.Text = 500000
            'Required Items
            TextBox20.Text = 379021000
            TextBox50.Text = 379021000
            TextBox58.Text = 379021000
            TextBox66.Text = 379021000
            TextBox74.Text = 379021000
            TextBox82.Text = 379021000
            TextBox90.Text = 379021000
            TextBox98.Text = 379021000
            TextBox106.Text = 379021000
            TextBox114.Text = 379021000
            'Upgrade Possibility
            TextBox115.Text = 10000
            TextBox116.Text = 10000
            TextBox117.Text = 10000
            TextBox118.Text = 10000
            TextBox119.Text = 10000
            TextBox120.Text = 10000
            TextBox121.Text = 10000
            TextBox122.Text = 10000
            TextBox123.Text = 10000
            TextBox124.Text = 10000
        End If
        'UNIQUE Upgrade preset.
        If ComboBox2.SelectedIndex = 3 Then
            'Required Money
            TextBox22.Text = 3200000
            TextBox23.Text = 2000000
            TextBox24.Text = 2000000
            TextBox25.Text = 2000000
            TextBox26.Text = 2000000
            TextBox27.Text = 2000000
            TextBox28.Text = 2000000
            TextBox29.Text = 2000000
            TextBox30.Text = 2000000
            TextBox31.Text = 2000000
            'Required Items
            TextBox20.Text = 379025000
            TextBox50.Text = 379021000
            TextBox58.Text = 379021000
            TextBox66.Text = 379021000
            TextBox74.Text = 379021000
            TextBox82.Text = 379021000
            TextBox90.Text = 379021000
            TextBox98.Text = 379021000
            TextBox106.Text = 379021000
            TextBox114.Text = 379021000
            'Upgrade Possibility
            TextBox115.Text = 10000
            TextBox116.Text = 10000
            TextBox117.Text = 10000
            TextBox118.Text = 10000
            TextBox119.Text = 10000
            TextBox120.Text = 10000
            TextBox121.Text = 10000
            TextBox122.Text = 10000
            TextBox123.Text = 10000
            TextBox124.Text = 10000
        End If
        'JEWELERY Upgrade preset.
        If ComboBox2.SelectedIndex = 4 Then
            'Required Money
            TextBox22.Text = 2000000
            TextBox23.Text = 2000000
            TextBox24.Text = 2000000
            TextBox25.Text = 2000000
            TextBox26.Text = 2000000
            TextBox27.Text = 2000000
            TextBox28.Text = 2000000
            TextBox29.Text = 2000000
            TextBox30.Text = 2000000
            TextBox31.Text = 2000000
            'Required Items
            TextBox20.Text = 379159000
            TextBox50.Text = 379159000
            TextBox58.Text = 379159000
            TextBox66.Text = 379159000
            TextBox74.Text = 379159000
            TextBox82.Text = 379159000
            TextBox90.Text = 379159000
            TextBox98.Text = 379159000
            TextBox106.Text = 379159000
            TextBox114.Text = 379159000
            'Upgrade Possibility
            TextBox115.Text = 10000
            TextBox116.Text = 10000
            TextBox117.Text = 10000
            TextBox118.Text = 10000
            TextBox119.Text = 10000
            TextBox120.Text = 10000
            TextBox121.Text = 10000
            TextBox122.Text = 10000
            TextBox123.Text = 10000
            TextBox124.Text = 10000
        End If
        'DMG Change preset.
        If ComboBox2.SelectedIndex = 5 Then
            'Required Money
            TextBox22.Text = 700000
            TextBox23.Text = 700000
            TextBox24.Text = 700000
            TextBox25.Text = 700000
            TextBox26.Text = 700000
            TextBox27.Text = 700000
            TextBox28.Text = 700000
            TextBox29.Text = 700000
            TextBox30.Text = 700000
            TextBox31.Text = 700000
            'Upgrade Possibility
            TextBox115.Text = 10000
            TextBox116.Text = 10000
            TextBox117.Text = 10000
            TextBox118.Text = 10000
            TextBox119.Text = 10000
            TextBox120.Text = 10000
            TextBox121.Text = 10000
            TextBox122.Text = 10000
            TextBox123.Text = 10000
            TextBox124.Text = 10000
        End If
        'STAT Change preset.
        If ComboBox2.SelectedIndex = 6 Then
            'Required Money
            TextBox22.Text = 700000
            TextBox23.Text = 700000
            TextBox24.Text = 700000
            TextBox25.Text = 700000
            TextBox26.Text = 700000
            TextBox27.Text = 700000
            TextBox28.Text = 700000
            TextBox29.Text = 700000
            TextBox30.Text = 700000
            TextBox31.Text = 700000
            'Upgrade Possibility
            TextBox115.Text = 10000
            TextBox116.Text = 10000
            TextBox117.Text = 10000
            TextBox118.Text = 10000
            TextBox119.Text = 10000
            TextBox120.Text = 10000
            TextBox121.Text = 10000
            TextBox122.Text = 10000
            TextBox123.Text = 10000
            TextBox124.Text = 10000
        End If

    End Sub

    'Item Group End

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Using conn As New SqlConnection("server=" & Trim(Form1.TextBox4.Text) & "; " &
            "user id=" & Trim(Form1.TextBox2.Text) & "; " &
            "password=" & Trim(Form1.TextBox3.Text) & ";" &
            "database=" & Trim(Form1.TextBox1.Text) & "")

            conn.Open()

            Dim comm As New SqlCommand
            Dim comm1 As New SqlCommand
            Dim comm2 As New SqlCommand
            Dim comm3 As New SqlCommand
            Dim comm4 As New SqlCommand
            Dim comm5 As New SqlCommand
            Dim comm6 As New SqlCommand
            Dim comm7 As New SqlCommand
            Dim comm8 As New SqlCommand
            Dim comm9 As New SqlCommand

            If conn.State = 1 Then
                With comm
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "Insert Into ITEM_UPGRADE Values (" & TextBox1.Text & ",5001,'" & TextBox2.Text & "+0 -> +1','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox4.Text & "," & TextBox20.Text & "," & TextBox19.Text & "," & TextBox18.Text & "," & TextBox17.Text & "," & TextBox16.Text & "," & TextBox15.Text & "," & TextBox14.Text & "," & TextBox21.Text & "," & TextBox22.Text & ",0," & TextBox115.Text & "," & TextBox32.Text & ")"
                End With
                With comm1
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 1 & ",5001,'" & TextBox2.Text & " +1 -> +2','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox5.Text & "," & TextBox50.Text & "," & TextBox49.Text & "," & TextBox48.Text & "," & TextBox47.Text & "," & TextBox46.Text & "," & TextBox45.Text & "," & TextBox44.Text & "," & TextBox43.Text & "," & TextBox23.Text & ",0," & TextBox116.Text & "," & TextBox33.Text & ")"
                End With
                With comm2
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 2 & ",5001,'" & TextBox2.Text & " +2 -> +3','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox6.Text & "," & TextBox58.Text & "," & TextBox57.Text & "," & TextBox56.Text & "," & TextBox55.Text & "," & TextBox54.Text & "," & TextBox53.Text & "," & TextBox52.Text & "," & TextBox51.Text & "," & TextBox24.Text & ",0," & TextBox117.Text & "," & TextBox34.Text & ")"
                End With
                With comm3
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 3 & ",5001,'" & TextBox2.Text & " +3 -> +4','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox7.Text & "," & TextBox66.Text & "," & TextBox65.Text & "," & TextBox64.Text & "," & TextBox63.Text & "," & TextBox62.Text & "," & TextBox61.Text & "," & TextBox60.Text & "," & TextBox59.Text & "," & TextBox25.Text & ",0," & TextBox118.Text & "," & TextBox35.Text & ")"
                End With
                With comm4
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 4 & ",5001,'" & TextBox2.Text & " +4 -> +5','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox8.Text & "," & TextBox74.Text & "," & TextBox73.Text & "," & TextBox72.Text & "," & TextBox71.Text & "," & TextBox70.Text & "," & TextBox69.Text & "," & TextBox68.Text & "," & TextBox67.Text & "," & TextBox26.Text & ",0," & TextBox119.Text & "," & TextBox36.Text & ")"
                End With
                With comm5
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 5 & ",5001,'" & TextBox2.Text & " +5 -> +6','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox9.Text & "," & TextBox82.Text & "," & TextBox81.Text & "," & TextBox80.Text & "," & TextBox79.Text & "," & TextBox78.Text & "," & TextBox77.Text & "," & TextBox76.Text & "," & TextBox75.Text & "," & TextBox27.Text & ",0," & TextBox120.Text & "," & TextBox37.Text & ")"
                End With
                With comm6
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 6 & ",5001,'" & TextBox2.Text & " +6 -> +7','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox10.Text & "," & TextBox90.Text & "," & TextBox89.Text & "," & TextBox88.Text & "," & TextBox87.Text & "," & TextBox86.Text & "," & TextBox85.Text & "," & TextBox84.Text & "," & TextBox83.Text & "," & TextBox28.Text & ",0," & TextBox121.Text & "," & TextBox38.Text & ")"
                End With
                With comm7
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 7 & ",5001,'" & TextBox2.Text & " +7 -> +8','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox11.Text & "," & TextBox98.Text & "," & TextBox97.Text & "," & TextBox96.Text & "," & TextBox95.Text & "," & TextBox94.Text & "," & TextBox93.Text & "," & TextBox92.Text & "," & TextBox91.Text & "," & TextBox29.Text & ",0," & TextBox122.Text & "," & TextBox39.Text & ")"
                End With
                With comm8
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 8 & ",5001,'" & TextBox2.Text & " +8 -> +9','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox12.Text & "," & TextBox106.Text & "," & TextBox105.Text & "," & TextBox104.Text & "," & TextBox103.Text & "," & TextBox102.Text & "," & TextBox101.Text & "," & TextBox100.Text & "," & TextBox99.Text & "," & TextBox30.Text & ",0," & TextBox123.Text & "," & TextBox40.Text & ")"
                End With
                With comm9
                    .Connection = conn
                    .CommandType = CommandType.Text
                    .CommandText = "Insert Into ITEM_UPGRADE Values (" & TextBox1.Text + 9 & ",5001,'" & TextBox2.Text & " +9 -> +10','" & TextBox3.Text & "'," & ComboBox1.SelectedIndex & "," & TextBox13.Text & "," & TextBox114.Text & "," & TextBox113.Text & "," & TextBox112.Text & "," & TextBox111.Text & "," & TextBox110.Text & "," & TextBox109.Text & "," & TextBox108.Text & "," & TextBox107.Text & "," & TextBox31.Text & ",0," & TextBox124.Text & "," & TextBox41.Text & ")"
                End With
                comm.ExecuteNonQuery()
                comm1.ExecuteNonQuery()
                comm2.ExecuteNonQuery()
                comm3.ExecuteNonQuery()
                comm4.ExecuteNonQuery()
                comm5.ExecuteNonQuery()
                comm6.ExecuteNonQuery()
                comm7.ExecuteNonQuery()
                comm8.ExecuteNonQuery()
                comm9.ExecuteNonQuery()
                MsgBox("ITEM_UPGRADE Updated!")
                conn.Close()

                Me.Close()

            Else
                MsgBox("Unable to connect to database, please re-check the connection information.")
            End If

        End Using

    End Sub

End Class