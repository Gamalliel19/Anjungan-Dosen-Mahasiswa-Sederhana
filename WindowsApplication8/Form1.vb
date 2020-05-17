Imports System.Data.OleDb
Public Class Datajenis
    Sub kosong()
        Textbox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.Text = ""
        TextBox4.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox1.Focus()
    End Sub
    Sub isi()
        TextBox2.Clear()
        TextBox4.Clear()
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
        DataGridView1.Columns(0).HeaderText = "Nim"
        DataGridView1.Columns(1).HeaderText = "Nama"
        DataGridView1.Columns(2).HeaderText = "Sex"
        DataGridView1.Columns(3).HeaderText = "Alamat"
        DataGridView1.Columns(4).HeaderText = "Agama"
        DataGridView1.Columns(5).HeaderText = "Jurusan"
        DataGridView1.Columns(6).HeaderText = "Hobby"
        DataGridView1.Columns(7).HeaderText = "Kode Kelas"
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
            cmd = New OleDbCommand("select * from Jadwal where Nim='" & TextBox1.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where Nama='" & TextBox2.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where Sex='" & ComboBox1.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where Alamat='" & TextBox4.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where Agama='" & ComboBox2.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where Jurusan='" & ComboBox3.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where Hobby='" & TextBox8.Text & "'", conn)
            cmd = New OleDbCommand("select * from Jadwal where Kd_kelas='" & TextBox9.Text & "'", conn)


            rd = cmd.ExecuteReader
            rd.Read()
            If Not rd.HasRows Then
                Dim simpan As String = "insert into Jadwal values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox4.Text & "','" & ComboBox2.Text & "','" & ComboBox3.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "' )"
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
            Dim ubah As String = "Update Jadwal set " &
            "Nama='" & TextBox2.Text & "' " &
            ",Sex='" & ComboBox1.Text & "' " &
            ",Alamat='" & TextBox4.Text & "' " &
            ",Agama='" & ComboBox2.Text & "' " &
            ",Jurusan='" & ComboBox3.Text & "' " &
            ",Hobby='" & TextBox8.Text & "' " &
            ",Kd_kelas='" & TextBox9.Text & "' " &
            "where Nim='" & TextBox1.Text & "'"
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
            MsgBox("Nim belum diisi")
            TextBox1.Focus()
            Exit Sub
        Else
            If MessageBox.Show("Yakin akan menghapus data jenis " & TextBox1.Text &
                               "?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                cmd = New OleDbCommand("delete * from Jadwal where Nim='" & TextBox1.Text & "'", conn)
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

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Hari_Click(sender As Object, e As EventArgs) Handles Hari.Click

    End Sub
End Class
