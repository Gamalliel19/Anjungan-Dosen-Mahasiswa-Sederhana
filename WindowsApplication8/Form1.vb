Imports System.Data.OleDb
Public Class Datajenis
    Sub kosong()
        Textbox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox1.Focus()
    End Sub
    Sub isi()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox2.Focus()

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
    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Call koneksi()
        Call TampilJenis()
        Call AturGrid()
        Call kosong()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data belum lengkap")
            TextBox1.Focus()
            Exit Sub
        Else
            cmd = New OleDbCommand("select * from Jadwal where kode_dosen='" & TextBox1.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where kode_matkul='" & TextBox2.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where nama_dosen='" & TextBox3.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where mata_kuliah='" & TextBox4.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where kode_dosen='" & TextBox6.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where kode_matkul='" & TextBox7.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where nama_dosen='" & TextBox8.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where mata_kuliah='" & TextBox9.Text & "'", conn)


            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                Dim simpan As String = "insert into Jadwal values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "' )"
                cmd = New OleDbCommand(simpan, conn)
                cmd.ExecuteNonQuery()
                MsgBox("Simpan Data Sukses", MsgBoxStyle.Information, "Perhatian")

            End If
            Call TampilJenis()
            Call kosong()
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("Kode Dosen & Kode Matakuliah belum diisi")
            TextBox1.Focus()
            Exit Sub
        Else
            Dim ubah As String = "Update Jadwal set " & _
            "nama_dosen='" & TextBox3.Text & "' " & _
            "where kode_dosen='" & TextBox1.Text & "'"
            cmd = New OleDbCommand(ubah, conn)
            cmd.ExecuteNonQuery()
            MsgBox("Ubah data sukses", MsgBoxStyle.Information, "Perhatian")
            Call TampilJenis()
            Call kosong()
            TextBox1.Focus()


        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("kode Dosen belum diisi")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("Yakin akan menghapus data jenis " & TextBox1.Text &
                               "?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                cmd = New OleDbCommand("delete * from Jadwal where kode_dosen='" & TextBox1.Text & "'", conn)
                cmd.ExecuteNonQuery()
                Call kosong()
                Call TampilJenis()
            Else
                Call kosong()
            End If
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call kosong()
    End Sub

    Private Sub Keluar_Click(sender As Object, e As EventArgs) Handles Keluar.Click
        End
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
