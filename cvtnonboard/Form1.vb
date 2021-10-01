Imports System.Text

Public Class Form1
    Dim dt As DataTable = New DataTable
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
            Dim fName As String = ""
            fName = OpenFileDialog1.FileName
            Dim TextLine As String = ""
            Dim SplitLine() As String

            If System.IO.File.Exists(fName) = True Then
                Dim objReader As New System.IO.StreamReader(fName, Encoding.ASCII)
                Dim index As Integer = 0
                Do While objReader.Peek() <> -1
                    If index > 0 Then
                        TextLine = objReader.ReadLine()
                        SplitLine = Split(TextLine, ",")
                        dt.Columns.Add("firstname", GetType(String))
                        dt.Columns.Add("lastname", GetType(String))
                        dt.Columns.Add("rank", GetType(String))
                        dt.Columns.Add("edipi", GetType(String))
                        dt.Columns.Add("mos", GetType(String))
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
End Class
