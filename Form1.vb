Imports System.Data.SqlClient

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If TextBox1.TextLength = 0 Or TextBox2.TextLength = 0 Or TextBox3.TextLength = 0 Or TextBox4.TextLength = 0 Then

            MsgBox("Please fill the connection information!")

        Else

            Dim strConnection As String = "server=" & Trim(TextBox4.Text) & "; " & _
                         "user id=" & Trim(TextBox2.Text) & "; " & _
                         "password=" & Trim(TextBox3.Text) & ";" & _
                         "database=" & Trim(TextBox1.Text) & ""
            Dim conn As New SqlConnection(strConnection)

            Try
                conn.Open()

                MsgBox("Successfully connected to the database.")

                conn.Close()

                Form2.Show()

                Me.Hide()

            Catch ex As Exception

                MsgBox("Unable to connect to database, please re-check the connection information.")

            End Try

        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AcceptButton = Button1
    End Sub
End Class
