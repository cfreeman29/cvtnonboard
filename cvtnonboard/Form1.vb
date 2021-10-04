Imports System.IO
Imports System.Text

Public Class Form1
    Dim dt As DataTable = New DataTable
    Dim fName As String = ""
    Dim usernum As Integer
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub
    Private Sub Checkedipi(edipival As String)
        If Not edipival.Length = 16 Then
            edipi.Clear()
            MessageBox.Show("EDIPI must be 16 characters.")
        Else
        End If
    End Sub
    Private Sub Checkstatusedipi(edipistatus As String)
        If Not edipistatus.Length = 16 Then
            statuscheck.Clear()
            MessageBox.Show("EDIPI must be 16 characters.")
        Else
        End If
    End Sub
    Private Sub Edipi_Leave(sender As Object, e As EventArgs) Handles edipi.Leave
        Checkedipi(edipi.Text)
    End Sub
    Private Sub Statuscheck_Leave(sender As Object, e As EventArgs) Handles statuscheck.Leave
        Checkstatusedipi(statuscheck.Text)
    End Sub
    Private Sub Statuscheck_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles statuscheck.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub Edipi_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles edipi.KeyPress
        If Not IsNumeric(e.KeyChar) And Not e.KeyChar = ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub Firstname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles firstname.KeyPress
        If Not (Asc(e.KeyChar) = 8) Then
            Dim allowedChars As String = "abcdefghijklmnopqrstuvwxyz"
            If Not allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub Lastname_KeyPress(sender As Object, e As KeyPressEventArgs) Handles lastname.KeyPress
        If Not (Asc(e.KeyChar) = 8) Then
            Dim allowedChars As String = "abcdefghijklmnopqrstuvwxyz"
            If Not allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
                e.KeyChar = ChrW(0)
                e.Handled = True
            End If
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox4.Text = OpenFileDialog1.FileName
            fName = OpenFileDialog1.FileName
            Dim TextLine As String = ""
            Dim SplitLine() As String

            If System.IO.File.Exists(fName) = True Then
                dt.Clear()
                Dim objReader As New System.IO.StreamReader(fName, Encoding.ASCII)
                Dim index As Integer = 0
                Do While objReader.Peek() <> -1
                    If index > 0 Then
                        TextLine = objReader.ReadLine()
                        SplitLine = Split(TextLine, ",")
                        If dt.Columns.Contains("firstname") Then
                        Else
                            dt.Columns.Add("firstname", GetType(String))
                        End If
                        If dt.Columns.Contains("lastname") Then
                        Else
                            dt.Columns.Add("lastname", GetType(String))
                        End If
                        If dt.Columns.Contains("rank") Then
                        Else
                            dt.Columns.Add("rank", GetType(String))
                        End If
                        If dt.Columns.Contains("edipi") Then
                        Else
                            dt.Columns.Add("edipi", GetType(String))
                        End If
                        If dt.Columns.Contains("mos") Then
                        Else
                            dt.Columns.Add("mos", GetType(String))
                        End If
                        dt.Rows.Add(SplitLine)
                        Else
                            TextLine = objReader.ReadLine()
                    End If
                    index = index + 1
                Loop
                DataGridView1.DataSource = dt
            Else
                MsgBox("File Does Not Exist")
            End If

        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Application.Restart()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim url As String = "https://dod.teams.microsoft.us/"

        Process.Start(url)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Process.Start("mailto: christopher.m.freeman37.ctr@army.mil?subject=Hello&body=")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim validate As Integer
        If String.IsNullOrEmpty(firstname.Text) Then
            MsgBox("First Name is Empty")
            lastname.Focus()
        Else
            validate += 1
        End If
        If String.IsNullOrEmpty(lastname.Text) Then
            MsgBox("Last Name is Empty")
            lastname.Focus()
        Else
            validate += 1
        End If
        If String.IsNullOrEmpty(edipi.Text) Then
            MsgBox("EDIPI is Empty")
            lastname.Focus()
        Else
            validate += 1
        End If
        If String.IsNullOrEmpty(rank.Text) Then
            MsgBox("Rank is Empty")
            lastname.Focus()
        Else
            validate += 1
        End If
        If String.IsNullOrEmpty(mos.Text) Then
            MsgBox("MOS is Empty")
            lastname.Focus()
        Else
            validate += 1
        End If
        If validate = 5 Then
            Dim first = firstname.Text
            Dim last = lastname.Text
            Dim rankbox = rank.Text
            Dim edipibox = edipi.Text
            Dim mosbox = mos.Text
            File.AppendAllText("C:\Users\cfree\Documents\master.csv", Environment.NewLine + first + "," + last + "," + rankbox + "," + edipibox + "," + mosbox)
            MsgBox("Student Added to Queue")
            edipi.Clear()
            firstname.Clear()
            lastname.Clear()
        Else
            MsgBox("Please fix values and try again.")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    Private Sub statuscheck_TextChanged(sender As Object, e As EventArgs) Handles statuscheck.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If String.IsNullOrEmpty(fName) Then
            MsgBox("Please choose a file and try again.")
        Else
            usernum = 0
            For Each dr As DataRow In dt.Rows
                usernum += 1
                Dim str1 As String = dr(0).ToString()
                Dim str2 As String = dr(1).ToString()
                Dim str3 As String = dr(2).ToString()
                Dim str4 As String = dr(3).ToString()
                Dim str5 As String = dr(4).ToString()
                File.AppendAllText("C:\Users\cfree\Documents\master.csv", Environment.NewLine + str1 + "," + str2 + "," + str3 + "," + str4 + "," + str5)
            Next
            Dim total As String
            total = usernum.ToString()
            MsgBox(total + " users have been added successfully.")
            dt.Clear()
        End If
    End Sub
End Class
