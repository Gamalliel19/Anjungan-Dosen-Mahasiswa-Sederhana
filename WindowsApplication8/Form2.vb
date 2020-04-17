Imports System.Data.OleDb
Public Class Form2

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim login As Integer
        Dim User As String
        Dim Password As String
        User = TextBox1.Text
        Password = TextBox2.Text
        If User = "admin" And Password = "admin" Then
            MsgBox("Selamat Datang Dosen!")
            Datajenis.Show()
            Me.Hide()
        Else If User = "siswa" And Password = "siswa" Then
            MsgBox("Selamat Datang Siswa!")
            Form3.Show()
            Me.Hide()
        Else
            MsgBox("Anda Salah Memasukkan Username & Password!")
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim pesan As String
        pesan = MsgBox("Anda Yakin Mau Keluar ??", vbQuestion + vbYesNo, "Question")
        If pesan = vbYes Then


        End If
    End Sub
End Class