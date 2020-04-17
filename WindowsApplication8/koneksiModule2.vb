Imports System.Data.OleDb
Module koneksiModule2
    Public strg As String
    Public connn As OleDbConnection
    Public cmdd As OleDbCommand
    Public dr As OleDbDataReader

    Sub koneksifitas()
        strg = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath &
              "\Login.accdb"
        connn = New OleDbConnection(strg)
        If connn.State = ConnectionState.Closed Then
            connn.Open()
        Else
            MsgBox("Koneksi Gagal!")
        End If
    End Sub

End Module
