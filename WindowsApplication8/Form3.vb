Imports System.Data.OleDb
Public Class Form3
    Sub kosong()
        TextBox1.Text = ""
    End Sub
    Sub TampilJenis()
        da = New OleDbDataAdapter("select * from Jadwal", conn)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "Jadwal")
        DataGridView1.DataSource = ds.Tables("Jadwal")
        DataGridView1.Refresh()
    End Sub
    Sub AturGrid()
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 200
        DataGridView1.Columns(3).Width = 200
        DataGridView1.Columns(4).Width = 100
        DataGridView1.Columns(5).Width = 100
        DataGridView1.Columns(6).Width = 100
        DataGridView1.Columns(7).Width = 100
        DataGridView1.Columns(0).HeaderText = "Kode_Dosen"
        DataGridView1.Columns(1).HeaderText = "Kode_Matkul"
        DataGridView1.Columns(2).HeaderText = "Nama_Dosen"
        DataGridView1.Columns(3).HeaderText = "Mata_Kuliah"
        DataGridView1.Columns(4).HeaderText = "Hari"
        DataGridView1.Columns(5).HeaderText = "Jam"
        DataGridView1.Columns(6).HeaderText = "Kelas"
        DataGridView1.Columns(7).HeaderText = "Ruangan"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Kode Dosen belum diisi")
            TextBox1.Focus()
            Exit Sub
        Else
            Dim cari As String = "Select * from Jadwal where " & _
            "nama_dosen='" & TextBox1.Text & "'"
            cmd = New OleDbCommand(cari, conn)
            cmd.ExecuteNonQuery()
            MsgBox("Cari data sukses", MsgBoxStyle.Information, "Perhatian")
            Call TampilJenis()
            Call kosong()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call koneksi()
        Call TampilJenis()
        Call AturGrid()
        Call kosong()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Call koneksi()
        da = New OleDbDataAdapter("SELECT * from Jadwal where nama_dosen LIKE '%" & TextBox1.Text & "%'", conn)
        ds = New DataSet
        da.Fill(ds, "nama_dosen")
        DataGridView1.DataSource = ds.Tables("nama_dosen")
        DataGridView1.DataSource = ds.Tables(0)
    End Sub
End Class